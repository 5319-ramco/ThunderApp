import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import fs from 'fs';
import path from 'path';
import App from './App';
import { createWorkbook, parseWorkbook } from './workbook';
import { buildStandings, getMatchWinner } from './tournament';

beforeEach(() => {
  window.localStorage.clear();
});

test('signs in and shows the team admin dashboard', () => {
  const { container } = render(<App />);

  expect(container.querySelector('.brand-image')).toHaveAttribute('src', '/logo192.png');
  expect(container.querySelector('.promo-ball')).toHaveAttribute('src', '/logo192.png');
  fireEvent.change(screen.getByLabelText(/email address/i), {
    target: { value: 'coach@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/password/i), {
    target: { value: 'demo-password' },
  });
  fireEvent.click(screen.getByRole('button', { name: /sign in/i }));

  expect(screen.getByRole('heading', { name: /good to see you, coach/i })).toBeInTheDocument();
  expect(window.localStorage.getItem('thunderbat-user')).toContain('coach@example.com');
});

test('adds a player with roster information', () => {
  render(<App />);

  fireEvent.change(screen.getByLabelText(/email address/i), {
    target: { value: 'coach@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/password/i), {
    target: { value: 'demo-password' },
  });
  fireEvent.click(screen.getByRole('button', { name: /sign in/i }));
  fireEvent.click(within(screen.getByRole('navigation')).getByRole('button', { name: /team/i }));
  fireEvent.click(screen.getByRole('button', { name: /add a player/i }));

  fireEvent.change(screen.getByLabelText(/player name/i), {
    target: { value: 'Taylor Smith' },
  });
  fireEvent.change(screen.getByLabelText(/^email address$/i), {
    target: { value: 'taylor@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/position or role/i), {
    target: { value: 'Goalkeeper' },
  });
  fireEvent.click(screen.getByRole('button', { name: /^add player$/i }));
  fireEvent.click(within(screen.getByRole('navigation')).getByRole('button', { name: /team/i }));

  expect(screen.getByText('Taylor Smith')).toBeInTheDocument();
  expect(screen.getByText('taylor@example.com')).toBeInTheDocument();
  expect(screen.getByText('Goalkeeper')).toBeInTheDocument();
});

test('edits a team member name', () => {
  render(<App />);

  fireEvent.change(screen.getByLabelText(/email address/i), {
    target: { value: 'coach@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/password/i), {
    target: { value: 'demo-password' },
  });
  fireEvent.click(screen.getByRole('button', { name: /sign in/i }));
  fireEvent.click(within(screen.getByRole('navigation')).getByRole('button', { name: /team/i }));
  fireEvent.click(screen.getAllByRole('button', { name: /edit/i })[0]);
  fireEvent.change(screen.getByLabelText(/player name/i), {
    target: { value: 'Alex Thunder' },
  });
  fireEvent.click(screen.getByRole('button', { name: /save changes/i }));

  expect(screen.getByText('Alex Thunder')).toBeInTheDocument();
  expect(window.localStorage.getItem('thunderbat-members')).toContain('Alex Thunder');
});

test('records an expense and shows its equal split', () => {
  render(<App />);

  fireEvent.change(screen.getByLabelText(/email address/i), {
    target: { value: 'coach@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/password/i), {
    target: { value: 'demo-password' },
  });
  fireEvent.click(screen.getByRole('button', { name: /sign in/i }));
  fireEvent.click(screen.getAllByRole('button', { name: /add an expense/i })[0]);

  fireEvent.change(screen.getByLabelText(/what was it for/i), {
    target: { value: 'Match snacks' },
  });
  fireEvent.change(screen.getByLabelText(/total cost/i), {
    target: { value: '900' },
  });
  fireEvent.click(screen.getByRole('button', { name: /save expense/i }));

  expect(screen.getByText('Match snacks')).toBeInTheDocument();
  expect(screen.getByRole('status')).toHaveTextContent(/split equally across 3 players/i);

  fireEvent.click(screen.getByRole('button', { name: /all expenses/i }));
  const expensePanel = screen.getByText(/all expenses/i).closest('section');
  expect(within(expensePanel).getByText(/split across 3 players · ₹300\.00 per share/i)).toBeInTheDocument();
});

test('switches to dark mode', () => {
  render(<App />);

  fireEvent.click(screen.getByRole('button', { name: /switch to dark mode/i }));

  expect(document.documentElement).toHaveAttribute('data-theme', 'dark');
  expect(window.localStorage.getItem('thunderbat-theme')).toBe('"dark"');
});

test('exports and reads members, expenses, and the receipts worksheet', async () => {
  const data = {
    members: [{ id: 'p1', name: 'Alex', email: 'alex@example.com', position: 'Captain' }],
    expenses: [{ id: 'e1', date: '2026-10-02', title: 'Team lunch', amount: 1250, paidBy: 'p1', participantIds: ['p1'] }],
    receipts: [{ id: 'r1', expenseId: 'e1', date: '2026-10-02', amount: 1250, personId: 'p1', description: 'Team lunch', fileName: 'lunch.jpg', ocrText: 'Total 1250' }],
    documents: [{ id: 'd1', memberId: 'p1', fileName: 'profile.pdf', contentType: 'application/pdf', kind: 'member' }],
  };

  const buffer = await createWorkbook(data);
  const parsed = await parseWorkbook(buffer);

  expect(parsed).toMatchObject({
    ...data,
    matches: [],
    hasMembersWorksheet: true,
    hasTournamentWorksheet: false,
  });
});

test('reads the starter workbook from public', async () => {
  const workbookPath = path.join(process.cwd(), 'public', 'thunder.xlsx');
  const parsed = await parseWorkbook(fs.readFileSync(workbookPath));

  expect(parsed.members).toHaveLength(3);
  expect(parsed.members[0]).toMatchObject({ id: 'player-1', name: 'Alex Morgan' });
  expect(parsed.expenses).toEqual([]);
  expect(parsed.receipts).toEqual([]);
});

test('imports doubles fixtures, game scores, winners, and league standings from the tournament workbook', async () => {
  const workbookPath = path.join(process.cwd(), 'public', 'Badminton_Doubles_League_Tournament.xlsx');
  const parsed = await parseWorkbook(fs.readFileSync(workbookPath));

  expect(parsed.hasMembersWorksheet).toBe(false);
  expect(parsed.hasTournamentWorksheet).toBe(true);
  expect(parsed.matches).toHaveLength(6);
  expect(parsed.matches[0]).toMatchObject({
    name: 'Match 1',
    team1: 'Team A',
    players1: 'Pradab & Pandi',
    team2: 'Team B',
    players2: 'Gowtham & Deepak',
    winner: 'Team A',
  });
  expect(parsed.matches[1].games).toEqual([
    { team1Score: 8, team2Score: 15 },
    { team1Score: 15, team2Score: 13 },
    { team1Score: 15, team2Score: 14 },
  ]);
  expect(getMatchWinner(parsed.matches[1])).toBe('Team C');
  expect(getMatchWinner(parsed.matches[2])).toBe('');

  const standings = buildStandings(parsed.matches);
  expect(standings).toEqual([
    { rank: 1, team: 'Team C', players: 'Lokesh & Karuna', played: 2, won: 2, lost: 0, points: 4 },
    { rank: 2, team: 'Team D', players: 'Selva & Kumar', played: 3, won: 2, lost: 1, points: 4 },
    { rank: 3, team: 'Team A', players: 'Pradab & Pandi', played: 2, won: 1, lost: 1, points: 2 },
    { rank: 4, team: 'Team B', players: 'Gowtham & Deepak', played: 3, won: 0, lost: 3, points: 0 },
  ]);
});

test('imports a selected Excel file back into the app', async () => {
  const workbook = await createWorkbook({
    members: [{ id: 'imported-1', name: 'Morgan Import', email: 'morgan@example.com', position: 'Captain' }],
    expenses: [{ id: 'imported-expense', date: '2026-10-02', title: 'Imported practice fee', amount: 600, paidBy: 'imported-1', participantIds: ['imported-1'] }],
    receipts: [],
    documents: [],
  });
  const file = new File([workbook], 'team-update.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  Object.defineProperty(file, 'arrayBuffer', {
    value: async () => workbook.buffer.slice(workbook.byteOffset, workbook.byteOffset + workbook.byteLength),
  });

  render(<App />);
  fireEvent.change(screen.getByLabelText(/email address/i), {
    target: { value: 'coach@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/password/i), {
    target: { value: 'demo-password' },
  });
  fireEvent.click(screen.getByRole('button', { name: /sign in/i }));
  fireEvent.change(screen.getByLabelText(/import excel/i), {
    target: { files: [file] },
  });

  expect(await screen.findByRole('status')).toHaveTextContent(/loaded team-update\.xlsx: 1 players, 1 expenses/i);
  expect(screen.getByRole('heading', { name: /good to see you, coach/i })).toBeInTheDocument();
  expect(screen.getByText('Imported practice fee')).toBeInTheDocument();
  await waitFor(() => {
    expect(window.localStorage.getItem('thunderbat-members')).toContain('Morgan Import');
    expect(window.localStorage.getItem('thunderbat-expenses')).toContain('Imported practice fee');
  });
});

test('loads the tournament workbook and records a match result', async () => {
  const buffer = fs.readFileSync(path.join(process.cwd(), 'public', 'Badminton_Doubles_League_Tournament.xlsx'));
  const file = new File([buffer], 'Badminton_Doubles_League_Tournament.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  Object.defineProperty(file, 'arrayBuffer', {
    value: async () => buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
  });

  render(<App />);
  fireEvent.change(screen.getByLabelText(/email address/i), {
    target: { value: 'coach@example.com' },
  });
  fireEvent.change(screen.getByLabelText(/password/i), {
    target: { value: 'demo-password' },
  });
  fireEvent.click(screen.getByRole('button', { name: /sign in/i }));
  fireEvent.change(screen.getByLabelText(/import excel/i), {
    target: { files: [file] },
  });

  expect(await screen.findByRole('heading', { name: 'Tournament' })).toBeInTheDocument();
  expect(screen.getByRole('heading', { name: 'League matches' })).toBeInTheDocument();
  expect(screen.getAllByText('Team C').length).toBeGreaterThan(0);
  expect(screen.getByText(/loaded badminton_doubles_league_tournament\.xlsx: 6 doubles matches/i)).toBeInTheDocument();

  const matchThree = screen.getByText('Match 3').closest('article');
  fireEvent.click(within(matchThree).getByRole('button', { name: /record score/i }));
  fireEvent.change(screen.getByLabelText(/game 1 team a points/i), { target: { value: '21' } });
  fireEvent.change(screen.getByLabelText(/game 1 team c points/i), { target: { value: '12' } });
  fireEvent.change(screen.getByLabelText(/game 2 team a points/i), { target: { value: '21' } });
  fireEvent.change(screen.getByLabelText(/game 2 team c points/i), { target: { value: '18' } });
  fireEvent.click(screen.getByRole('button', { name: /save match result/i }));

  expect(within(screen.getByText('Match 3').closest('article')).getByText('Team A won')).toBeInTheDocument();
  expect(screen.getByRole('status')).toHaveTextContent(/league standings have been updated/i);
});
