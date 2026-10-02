import { useEffect, useMemo, useState } from 'react';
import { recognize } from 'tesseract.js';
import { getAttachment, listAttachments, saveAttachment } from './attachmentStore';
import { createWorkbook, parseWorkbook, PUBLIC_WORKBOOK_PATH } from './workbook';
import './App.css';

const STORAGE_KEYS = {
  user: 'thunderbat-user',
  members: 'thunderbat-members',
  expenses: 'thunderbat-expenses',
  theme: 'thunderbat-theme',
};
const TEAM_NAME = 'Thunder';

const starterMembers = [
  { id: 'player-1', name: 'Alex Morgan', email: 'alex@example.com', position: 'Captain' },
  { id: 'player-2', name: 'Jordan Lee', email: 'jordan@example.com', position: 'Forward' },
  { id: 'player-3', name: 'Sam Rivera', email: 'sam@example.com', position: 'Midfielder' },
];

function readStored(key, fallback) {
  try {
    const storedValue = window.localStorage.getItem(key);
    return storedValue ? JSON.parse(storedValue) : fallback;
  } catch {
    return fallback;
  }
}

function writeStored(key, value) {
  window.localStorage.setItem(key, JSON.stringify(value));
}

function createId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function today() {
  const date = new Date();
  const offset = date.getTimezoneOffset();
  return new Date(date.getTime() - offset * 60 * 1000).toISOString().slice(0, 10);
}

function money(amount) {
  return new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 2,
  }).format(Number(amount) || 0);
}

function formatDate(value) {
  if (!value) return 'No date';
  return new Intl.DateTimeFormat('en-IN', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
  }).format(new Date(`${value}T00:00:00`));
}

function initials(name) {
  return name
    .trim()
    .split(/\s+/)
    .slice(0, 2)
    .map((part) => part[0])
    .join('')
    .toUpperCase();
}

function splitExpense(expense) {
  const participantIds = expense.participantIds || [];
  if (!participantIds.length) return [];

  const totalPaise = Math.round(Number(expense.amount) * 100);
  const baseShare = Math.floor(totalPaise / participantIds.length);
  const remainder = totalPaise % participantIds.length;

  return participantIds.map((memberId, index) => ({
    memberId,
    amount: (baseShare + (index < remainder ? 1 : 0)) / 100,
  }));
}

function extractReceiptDetails(text) {
  const dateMatch = text.match(/\b(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\b|\b(\d{1,2})[-/.](\d{1,2})[-/.](20\d{2})\b/);
  let date = '';
  if (dateMatch) {
    const year = dateMatch[1] || dateMatch[6];
    const month = (dateMatch[2] || dateMatch[5]).padStart(2, '0');
    const day = (dateMatch[3] || dateMatch[4]).padStart(2, '0');
    date = `${year}-${month}-${day}`;
  }

  const amounts = [...text.matchAll(/(?:₹|INR|Rs\.?)?\s*(\d{1,7}(?:,\d{3})*(?:\.\d{1,2})?)/gi)]
    .map((match) => Number(match[1].replace(/,/g, '')))
    .filter((amount) => Number.isFinite(amount) && amount > 0);
  const totalLine = text.split(/\r?\n/).find((line) => /\b(total|amount due|grand total|net amount)\b/i.test(line));
  const totalMatch = totalLine && totalLine.match(/(\d[\d,]*(?:\.\d{1,2})?)/);

  return {
    date: date || today(),
    amount: totalMatch ? totalMatch[1].replace(/,/g, '') : amounts.length ? String(Math.max(...amounts)) : '',
  };
}

function downloadWorkbook(buffer, fileName = 'thunder.xlsx') {
  const workbookBlob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(workbookBlob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function Icon({ name, size = 18 }) {
  const paths = {
    grid: <><rect x="3" y="3" width="7" height="7" rx="1.5" /><rect x="14" y="3" width="7" height="7" rx="1.5" /><rect x="3" y="14" width="7" height="7" rx="1.5" /><rect x="14" y="14" width="7" height="7" rx="1.5" /></>,
    users: <><path d="M16 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" /><circle cx="10" cy="7" r="4" /><path d="M20 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" /></>,
    receipt: <><path d="M4 3v18l3-2 3 2 3-2 3 2 4-2V3l-4 2-3-2-3 2-3-2-3 2Z" /><path d="M8 9h8M8 13h6" /></>,
    moon: <path d="M20.9 13A9 9 0 0 1 11 3.1 9 9 0 1 0 20.9 13Z" />,
    sun: <><circle cx="12" cy="12" r="4" /><path d="M12 2v2m0 16v2M4.93 4.93l1.41 1.41m11.32 11.32 1.41 1.41M2 12h2m16 0h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41" /></>,
    arrow: <><path d="M5 12h14" /><path d="m12 5 7 7-7 7" /></>,
    plus: <><path d="M12 5v14M5 12h14" /></>,
    logout: <><path d="M10 17l5-5-5-5" /><path d="M15 12H3" /><path d="M12 3h6a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-6" /></>,
    trend: <><path d="m3 17 6-6 4 4 8-8" /><path d="M15 7h6v6" /></>,
    shield: <><path d="M12 22s8-4 8-11V5l-8-3-8 3v6c0 7 8 11 8 11Z" /><path d="m9 12 2 2 4-4" /></>,
    camera: <><path d="M14 5h-4l-2 2H5a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2-2Z" /><circle cx="12" cy="13" r="3" /></>,
  };

  return (
    <svg
      aria-hidden="true"
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {paths[name]}
    </svg>
  );
}

function ThemeButton({ theme, onClick }) {
  return (
    <button className="theme-button" type="button" onClick={onClick} aria-label={`Switch to ${theme === 'light' ? 'dark' : 'light'} mode`}>
      <Icon name={theme === 'light' ? 'moon' : 'sun'} />
      <span>{theme === 'light' ? 'Dark mode' : 'Light mode'}</span>
    </button>
  );
}

function Login({ onLogin, theme, onToggleTheme }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');

  function handleSubmit(event) {
    event.preventDefault();
    onLogin(email.trim());
  }

  return (
    <main className="login-page">
      <section className="login-aside">
        <div className="brand brand-light">
          <img className="brand-image" src={`${process.env.PUBLIC_URL}/logo192.png`} alt="" />
          <span>thunderbat</span>
        </div>
        <div className="login-promo">
          <span className="eyebrow light-eyebrow">YOUR TEAM, IN SYNC</span>
          <h1>Spend less time splitting.<br /><span>Play more together.</span></h1>
          <p>One simple home for your team roster, shared costs, and every little detail in between.</p>
          <div className="promo-art" aria-hidden="true">
            <img className="promo-ball" src={`${process.env.PUBLIC_URL}/logo192.png`} alt="" />
            <div className="promo-orbit orbit-one" />
            <div className="promo-orbit orbit-two" />
            <span className="promo-tag tag-one">TEAM, SORTED</span>
            <span className="promo-tag tag-two">FAIR SHARES ✦</span>
          </div>
        </div>
        <div className="login-aside-foot"><span className="status-dot" /> The better way to manage your team</div>
      </section>

      <section className="login-main">
        <div className="login-top"><span>Team admin</span><ThemeButton theme={theme} onClick={onToggleTheme} /></div>
        <div className="login-card">
          <span className="eyebrow">WELCOME BACK</span>
          <h2>Sign in to your<br />team space.</h2>
          <p className="login-subtitle">Your squad's next great season starts here.</p>
          <form onSubmit={handleSubmit}>
            <label htmlFor="login-email">Email address</label>
            <input id="login-email" type="email" placeholder="you@yourteam.com" value={email} onChange={(event) => setEmail(event.target.value)} required autoComplete="email" />
            <div className="password-label"><label htmlFor="login-password">Password</label><span>Demo sign-in</span></div>
            <input id="login-password" type="password" placeholder="Enter any password" value={password} onChange={(event) => setPassword(event.target.value)} required autoComplete="current-password" />
            <button className="primary-button login-submit" type="submit">Sign in <Icon name="arrow" /></button>
          </form>
          <p className="demo-note"><Icon name="shield" size={15} /> Demo mode · any email and password will work</p>
        </div>
        <p className="login-legal">Made for teams that do life together.</p>
      </section>
    </main>
  );
}

function App() {
  const [user, setUser] = useState(() => readStored(STORAGE_KEYS.user, null));
  const [members, setMembers] = useState(() => readStored(STORAGE_KEYS.members, starterMembers));
  const [expenses, setExpenses] = useState(() => readStored(STORAGE_KEYS.expenses, []));
  const [receipts, setReceipts] = useState(() => readStored('thunderbat-receipts', []));
  const [documents, setDocuments] = useState([]);
  const [theme, setTheme] = useState(() => readStored(STORAGE_KEYS.theme, 'light'));
  const [activePage, setActivePage] = useState('overview');
  const [showMemberForm, setShowMemberForm] = useState(false);
  const [showExpenseForm, setShowExpenseForm] = useState(false);
  const [showReceiptForm, setShowReceiptForm] = useState(false);
  const [editingMember, setEditingMember] = useState(null);
  const [workbookBusy, setWorkbookBusy] = useState(false);
  const [notice, setNotice] = useState('');

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    writeStored(STORAGE_KEYS.theme, theme);
  }, [theme]);

  useEffect(() => writeStored(STORAGE_KEYS.members, members), [members]);
  useEffect(() => writeStored(STORAGE_KEYS.expenses, expenses), [expenses]);
  useEffect(() => writeStored('thunderbat-receipts', receipts), [receipts]);

  useEffect(() => {
    if (!user || !window.indexedDB) return;
    listAttachments()
      .then(setDocuments)
      .catch((error) => setNotice(`Could not load saved documents: ${error.message}`));
  }, [user]);

  const totals = useMemo(() => {
    const paidByMember = {};
    const owedByMember = {};
    let totalSpend = 0;
    expenses.forEach((expense) => {
      totalSpend += Number(expense.amount) || 0;
      paidByMember[expense.paidBy] = (paidByMember[expense.paidBy] || 0) + Number(expense.amount);
      splitExpense(expense).forEach(({ memberId, amount }) => {
        owedByMember[memberId] = (owedByMember[memberId] || 0) + amount;
      });
    });
    return { totalSpend, paidByMember, owedByMember };
  }, [expenses]);

  const monthExpenses = useMemo(() => {
    const now = new Date();
    return expenses.filter((expense) => {
      const date = new Date(`${expense.date}T00:00:00`);
      return date.getMonth() === now.getMonth() && date.getFullYear() === now.getFullYear();
    }).reduce((sum, expense) => sum + Number(expense.amount), 0);
  }, [expenses]);

  function login(email) {
    const nextUser = { email, name: email.split('@')[0].replace(/[._-]+/g, ' ').replace(/\b\w/g, (letter) => letter.toUpperCase()) };
    setUser(nextUser);
    writeStored(STORAGE_KEYS.user, nextUser);
  }

  function logout() {
    setUser(null);
    window.localStorage.removeItem(STORAGE_KEYS.user);
    setActivePage('overview');
  }

  function toggleTheme() {
    setTheme((currentTheme) => currentTheme === 'light' ? 'dark' : 'light');
  }

  function updateMember(event) {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const updated = {
      ...editingMember,
      name: String(formData.get('name')).trim(),
      email: String(formData.get('email')).trim(),
      position: String(formData.get('position')).trim() || 'Player',
    };
    if (!updated.name || !updated.email) return;
    setMembers((current) => current.map((member) => member.id === updated.id ? updated : member));
    setEditingMember(null);
    setNotice(`${updated.name}'s details were updated.`);
  }

  function addMember(event) {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const member = {
      id: createId(),
      name: String(formData.get('name')).trim(),
      email: String(formData.get('email')).trim(),
      position: String(formData.get('position')).trim() || 'Player',
    };
    if (!member.name || !member.email) return;
    setMembers((current) => [...current, member]);
    setShowMemberForm(false);
    setNotice(`${member.name} added to your team.`);
    event.currentTarget.reset();
  }

  function addExpense(event) {
    event.preventDefault();
    const formData = new FormData(event.currentTarget);
    const amount = Number(formData.get('amount'));
    const participantIds = members.map((member) => member.id);
    if (!participantIds.length || !Number.isFinite(amount) || amount <= 0) {
      setNotice('Add a team member and a valid amount before recording an expense.');
      return;
    }
    const expense = {
      id: createId(),
      title: String(formData.get('title')).trim(),
      amount: Math.round(amount * 100) / 100,
      paidBy: String(formData.get('paidBy')),
      date: String(formData.get('date')),
      participantIds,
    };
    if (!expense.title || !expense.date) return;
    setExpenses((current) => [expense, ...current]);
    setShowExpenseForm(false);
    setNotice(`${expense.title} added and split equally across ${participantIds.length} ${participantIds.length === 1 ? 'player' : 'players'}.`);
    event.currentTarget.reset();
  }

  async function loadWorkbook(buffer, sourceName) {
    const data = await parseWorkbook(buffer);
    if (data.members.length === 0 && data.expenses.length === 0 && data.receipts.length === 0) {
      throw new Error('The workbook has no rows in Members, Expenses, or receipts.');
    }
    setMembers(data.members);
    setExpenses(data.expenses);
    setReceipts(data.receipts);
    setDocuments(data.documents);
    writeStored(STORAGE_KEYS.members, data.members);
    writeStored(STORAGE_KEYS.expenses, data.expenses);
    writeStored('thunderbat-receipts', data.receipts);
    setNotice(`Loaded ${sourceName}: ${data.members.length} players, ${data.expenses.length} expenses, and ${data.receipts.length} receipts.`);
  }

  async function loadPublicWorkbook() {
    setWorkbookBusy(true);
    try {
      const response = await fetch(PUBLIC_WORKBOOK_PATH, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(response.status === 404
          ? 'public/thunder.xlsx was not found. Import a workbook or download one and place it in public as thunder.xlsx.'
          : `The workbook request failed (${response.status}).`);
      }
      await loadWorkbook(await response.arrayBuffer(), 'public/thunder.xlsx');
    } catch (error) {
      setNotice(`Excel load failed: ${error.message}`);
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function importWorkbook(event) {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = '';
    if (!file) return;
    setWorkbookBusy(true);
    try {
      await loadWorkbook(await file.arrayBuffer(), file.name);
    } catch (error) {
      setNotice(`Could not import ${file.name}: ${error.message}`);
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function exportWorkbook() {
    setWorkbookBusy(true);
    try {
      const data = await createWorkbook({ members, expenses, receipts, documents });
      downloadWorkbook(data);
      setNotice('Updated thunder.xlsx downloaded. To use it from public, replace public/thunder.xlsx with this file.');
    } catch (error) {
      setNotice(`Could not create the Excel workbook: ${error.message}`);
    } finally {
      setWorkbookBusy(false);
    }
  }

  async function addMemberDocument(member, files) {
    if (!window.indexedDB) {
      setNotice('This browser does not support local document storage.');
      return;
    }
    try {
      const added = [];
      for (const file of files) {
        const record = {
          id: createId(),
          kind: 'member',
          memberId: member.id,
          fileName: file.name,
          contentType: file.type || 'application/octet-stream',
          createdAt: new Date().toISOString(),
          file,
        };
        added.push(await saveAttachment(record));
      }
      setDocuments((current) => [...current, ...added]);
      setNotice(`${added.length} document${added.length === 1 ? '' : 's'} saved for ${member.name}.`);
    } catch (error) {
      setNotice(`Could not save the document: ${error.message}`);
    }
  }

  async function openDocument(documentRecord) {
    try {
      const attachment = await getAttachment(documentRecord.id);
      if (!attachment?.file) {
        setNotice(`${documentRecord.fileName} is listed in the workbook, but its file is not stored on this device.`);
        return;
      }
      const url = URL.createObjectURL(attachment.file);
      const link = document.createElement('a');
      link.href = url;
      link.download = documentRecord.fileName;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      setNotice(`Could not open ${documentRecord.fileName}: ${error.message}`);
    }
  }

  async function addScannedReceipt(receiptInput, file) {
    const receiptId = createId();
    const expenseId = createId();
    try {
      const attachment = {
        id: receiptId,
        kind: 'receipt',
        memberId: receiptInput.personId,
        fileName: file.name,
        contentType: file.type || 'image/jpeg',
        createdAt: new Date().toISOString(),
        file,
      };
      const metadata = await saveAttachment(attachment);
      const receipt = {
        id: receiptId,
        expenseId,
        date: receiptInput.date,
        amount: Math.round(Number(receiptInput.amount) * 100) / 100,
        personId: receiptInput.personId,
        description: receiptInput.description.trim(),
        fileName: file.name,
        ocrText: receiptInput.ocrText.slice(0, 30000),
      };
      const expense = {
        id: expenseId,
        title: receipt.description,
        amount: receipt.amount,
        paidBy: receipt.personId,
        date: receipt.date,
        participantIds: members.map((member) => member.id),
      };
      setReceipts((current) => [receipt, ...current]);
      setExpenses((current) => [expense, ...current]);
      setDocuments((current) => [...current, metadata]);
      setShowReceiptForm(false);
      setNotice(`${receipt.description} saved as a receipt and split across ${members.length} team ${members.length === 1 ? 'member' : 'members'}.`);
    } catch (error) {
      setNotice(`Could not save the scanned receipt: ${error.message}`);
      throw error;
    }
  }

  const pageTitle = activePage === 'team' ? 'Your team' : activePage === 'expenses' ? 'Expenses' : 'Admin dashboard';
  const recentExpenses = [...expenses].sort((first, second) => second.date.localeCompare(first.date)).slice(0, 5);
  const currentMonth = new Intl.DateTimeFormat('en-US', { month: 'long' }).format(new Date());

  if (!user) return <Login onLogin={login} theme={theme} onToggleTheme={toggleTheme} />;

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <a className="brand sidebar-brand" href="#overview" onClick={() => setActivePage('overview')}>
          <img className="brand-image" src={`${process.env.PUBLIC_URL}/logo192.png`} alt="" /><span>thunderbat</span>
        </a>
        <div className="workspace-switcher"><span className="workspace-avatar">T</span><span className="workspace-copy"><strong>{TEAM_NAME}</strong><small>Team workspace</small></span><span className="switcher-dots">•••</span></div>
        <span className="nav-label">WORKSPACE</span>
        <nav className="side-nav" aria-label="Workspace">
          <button className={activePage === 'overview' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('overview')}><Icon name="grid" />Overview</button>
          <button className={activePage === 'team' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('team')}><Icon name="users" />Team <span className="nav-count">{members.length}</span></button>
          <button className={activePage === 'expenses' ? 'nav-item active' : 'nav-item'} onClick={() => setActivePage('expenses')}><Icon name="receipt" />Expenses</button>
        </nav>
        <div className="sidebar-spacer" />
        <div className="sidebar-tip"><span className="tip-icon">✦</span><strong>Better together.</strong><p>Keep the team on the same page, on and off the field.</p></div>
        <div className="sidebar-user"><span className="avatar avatar-orange">{initials(user.name)}</span><span className="sidebar-user-copy"><strong>{user.name}</strong><small>Administrator</small></span><button className="logout-button" onClick={logout} aria-label="Sign out"><Icon name="logout" /></button></div>
      </aside>

      <main className="main-area">
        <header className="topbar">
          <div className="breadcrumbs"><span>{TEAM_NAME}</span><span className="crumb-slash">/</span><strong>{pageTitle}</strong></div>
          <div className="topbar-actions">
            <span className="topbar-date">{new Intl.DateTimeFormat('en-IN', { weekday: 'short', day: 'numeric', month: 'short' }).format(new Date())}</span>
            <div className="workbook-actions">
              <button className="workbook-button" type="button" onClick={loadPublicWorkbook} disabled={workbookBusy} title="Load public/thunder.xlsx">Load public</button>
              <label className="workbook-button import-workbook" htmlFor="workbook-file" title="Import an Excel workbook into the app">Import Excel</label>
              <input id="workbook-file" className="visually-hidden" type="file" accept=".xlsx" onChange={importWorkbook} />
              <button className="workbook-button workbook-save" type="button" onClick={exportWorkbook} disabled={workbookBusy}>{workbookBusy ? 'Exporting…' : 'Export Excel'}</button>
            </div>
            <ThemeButton theme={theme} onClick={toggleTheme} /><span className="avatar avatar-small">{initials(user.name)}</span>
          </div>
        </header>

        <div className="page-content">
          {notice && <div className="notice" role="status">{notice}<button onClick={() => setNotice('')} aria-label="Dismiss notification">×</button></div>}

          {activePage === 'overview' && (
            <>
              <div className="page-heading-row">
                <div><span className="eyebrow">OVERVIEW</span><h1>Good to see you, {user.name.split(' ')[0]} <span className="wave">✳</span></h1><p>Here's what's happening with your team.</p></div>
                  <div className="heading-actions"><button className="secondary-button" onClick={() => setShowReceiptForm(true)}><Icon name="camera" /> Scan receipt</button><button className="primary-button" onClick={() => setShowExpenseForm(true)}><Icon name="plus" /> Add an expense</button></div>
              </div>
              <section className="stats-grid" aria-label="Team statistics">
                <article className="stat-card"><div className="stat-top"><span className="stat-label">TOTAL SPEND</span><span className="stat-icon"><Icon name="receipt" /></span></div><strong className="stat-value">{money(totals.totalSpend)}</strong><span className="stat-foot"><span className="stat-muted">Across all recorded expenses</span></span></article>
                <article className="stat-card"><div className="stat-top"><span className="stat-label">THIS MONTH</span><span className="stat-icon orange-icon"><Icon name="trend" /></span></div><strong className="stat-value">{money(monthExpenses)}</strong><span className="stat-foot"><span className="stat-muted">{currentMonth} team spend</span></span></article>
                <article className="stat-card"><div className="stat-top"><span className="stat-label">TEAM PLAYERS</span><span className="stat-icon"><Icon name="users" /></span></div><strong className="stat-value">{members.length}</strong><span className="stat-foot"><span className="stat-muted">On your roster</span></span></article>
                <article className="stat-card"><div className="stat-top"><span className="stat-label">EXPENSES LOGGED</span><span className="stat-icon"><Icon name="grid" /></span></div><strong className="stat-value">{expenses.length}</strong><span className="stat-foot"><span className="stat-muted">Every share, accounted for</span></span></article>
              </section>
              <section className="welcome-banner">
                <div className="banner-copy"><span className="eyebrow light-eyebrow">{TEAM_NAME.toUpperCase()} · TEAM ADMIN</span><h2>A good team looks out<br />for each other.</h2><p>Keep your roster organized and shared costs fair. You've got this.</p><button className="banner-link" onClick={() => setActivePage('team')}>Meet your team <Icon name="arrow" size={16} /></button></div>
                <div className="banner-art" aria-hidden="true"><span className="banner-sun" /><span className="banner-big-circle" /><span className="banner-small-circle"><img src={`${process.env.PUBLIC_URL}/logo192.png`} alt="" /></span><span className="banner-stripe stripe-one" /><span className="banner-stripe stripe-two" /><span className="banner-stripe stripe-three" /></div>
              </section>
              <div className="content-grid">
                <section className="panel expense-panel">
                  <div className="panel-heading"><div><span className="eyebrow">THE LATEST</span><h2>Recent expenses</h2></div><button className="text-button" onClick={() => setActivePage('expenses')}>All expenses <Icon name="arrow" size={15} /></button></div>
                  {recentExpenses.length ? <ExpenseList expenses={recentExpenses} members={members} /> : <EmptyState title="No expenses just yet" text="Add your first team expense and we'll split it fairly." action="Add an expense" onAction={() => setShowExpenseForm(true)} />}
                </section>
                <section className="panel balance-panel">
                  <div className="panel-heading"><div><span className="eyebrow">FAIR SHARES</span><h2>Team balance</h2></div><span className="balance-mark">↗</span></div>
                  {members.length ? <div className="balance-list">{members.slice(0, 5).map((member) => {
                    const balance = (totals.paidByMember[member.id] || 0) - (totals.owedByMember[member.id] || 0);
                    return <div className="balance-row" key={member.id}><span className="avatar">{initials(member.name)}</span><span className="balance-name"><strong>{member.name}</strong><small>{member.position}</small></span><span className={balance > 0.005 ? 'balance-amount positive' : balance < -0.005 ? 'balance-amount negative' : 'balance-amount'}>{balance > 0.005 ? '+' : ''}{money(balance)}</span></div>;
                  })}</div> : <EmptyState title="Your roster is empty" text="Add your players to start tracking fair shares." action="Add a player" onAction={() => setShowMemberForm(true)} />}
                  <button className="panel-footer-link" onClick={() => setActivePage('team')}>View full team balance <Icon name="arrow" size={15} /></button>
                </section>
              </div>
            </>
          )}

          {activePage === 'team' && (
            <>
              <div className="page-heading-row"><div><span className="eyebrow">THE SQUAD</span><h1>Your team</h1><p>Player details and each person's share of team expenses.</p></div><button className="primary-button" onClick={() => setShowMemberForm(true)}><Icon name="plus" /> Add a player</button></div>
              <section className="panel team-panel">
                <div className="panel-heading"><div><h2>Team roster <span className="heading-count">{members.length}</span></h2><p className="panel-subtitle">Everyone on the roster is included in new expense splits.</p></div></div>
                {members.length ? <div className="table-wrap"><table><thead><tr><th>PLAYER</th><th>ROLE</th><th>PAID</th><th>SHARE</th><th>NET BALANCE</th><th>DOCUMENTS</th><th>ACTIONS</th></tr></thead><tbody>{members.map((member) => {
                  const balance = (totals.paidByMember[member.id] || 0) - (totals.owedByMember[member.id] || 0);
                  return <tr key={member.id}>
                    <td><div className="player-cell"><span className="avatar">{initials(member.name)}</span><span><strong>{member.name}</strong><small>{member.email}</small></span></div></td>
                    <td><span className="role-pill">{member.position}</span></td>
                    <td>{money(totals.paidByMember[member.id] || 0)}</td>
                    <td>{money(totals.owedByMember[member.id] || 0)}</td>
                    <td><span className={balance > 0.005 ? 'balance-amount positive' : balance < -0.005 ? 'balance-amount negative' : 'balance-amount'}>{balance > 0.005 ? '+' : ''}{money(balance)}</span></td>
                    <td><MemberDocuments member={member} documents={documents.filter((item) => item.memberId === member.id && item.kind !== 'receipt')} onUpload={addMemberDocument} onOpen={openDocument} /></td>
                    <td><button className="secondary-button edit-member-button" type="button" onClick={() => setEditingMember(member)}>Edit</button></td>
                  </tr>;
                })}</tbody></table></div> : <EmptyState title="Start with your first player" text="Add your teammates so you can keep the roster and expense splits in one place." action="Add a player" onAction={() => setShowMemberForm(true)} />}
              </section>
              <section className="team-note"><span className="team-note-icon">✦</span><span><strong>How the split works</strong><small>Every new expense is divided equally among everyone on the roster at the time it's added. Past splits stay unchanged.</small></span></section>
            </>
          )}

          {activePage === 'expenses' && (
            <>
              <div className="page-heading-row"><div><span className="eyebrow">TEAM FINANCES</span><h1>Expenses</h1><p>Every team cost, split equally and kept in one place.</p></div><div className="heading-actions"><button className="secondary-button" onClick={() => setShowReceiptForm(true)}><Icon name="camera" /> Scan receipt</button><button className="primary-button" onClick={() => setShowExpenseForm(true)}><Icon name="plus" /> Add an expense</button></div></div>
              <section className="panel expense-page-panel"><div className="panel-heading"><div><h2>All expenses <span className="heading-count">{expenses.length}</span></h2><p className="panel-subtitle">Each share is calculated from the roster when the expense was recorded.</p></div><span className="expense-total">{money(totals.totalSpend)} total</span></div>{expenses.length ? <ExpenseList expenses={[...expenses].sort((first, second) => second.date.localeCompare(first.date))} members={members} expanded /> : <EmptyState title="No expenses just yet" text="Keep track of team costs and we'll work out everyone's equal share." action="Add an expense" onAction={() => setShowExpenseForm(true)} />}</section>
              <section className="panel receipt-page-panel">
                <div className="panel-heading"><div><h2>Scanned receipts <span className="heading-count">{receipts.length}</span></h2><p className="panel-subtitle">Receipt details are also written to the Excel receipts sheet.</p></div><button className="secondary-button" onClick={() => setShowReceiptForm(true)}>Scan receipt</button></div>
                {receipts.length ? <ReceiptList receipts={receipts} members={members} documents={documents} onOpen={openDocument} /> : <EmptyState title="No scanned receipts" text="Scan a receipt image to read and review its details before adding it." action="Scan a receipt" onAction={() => setShowReceiptForm(true)} />}
              </section>
            </>
          )}
        </div>
        <footer className="main-footer"><span>{TEAM_NAME.toUpperCase()}</span><span>Teamwork makes the dream work <span className="footer-star">✳</span></span></footer>
      </main>

      {showMemberForm && <Modal title="Add a player" onClose={() => setShowMemberForm(false)}><p className="modal-intro">Add someone to the {TEAM_NAME} roster.</p><form className="modal-form" onSubmit={addMember}><label htmlFor="player-name">Player name</label><input id="player-name" name="name" placeholder="e.g. Taylor Smith" required autoFocus /><label htmlFor="player-email">Email address</label><input id="player-email" name="email" type="email" placeholder="taylor@example.com" required /><label htmlFor="player-position">Position or role</label><input id="player-position" name="position" placeholder="e.g. Goalkeeper" /><div className="modal-actions"><button className="secondary-button" type="button" onClick={() => setShowMemberForm(false)}>Cancel</button><button className="primary-button" type="submit"><Icon name="plus" /> Add player</button></div></form></Modal>}

      {editingMember && <Modal title="Edit player" onClose={() => setEditingMember(null)}><p className="modal-intro">Update this player's team details.</p><form className="modal-form" onSubmit={updateMember}><label htmlFor="edit-player-name">Player name</label><input id="edit-player-name" name="name" defaultValue={editingMember.name} required autoFocus /><label htmlFor="edit-player-email">Email address</label><input id="edit-player-email" name="email" type="email" defaultValue={editingMember.email} required /><label htmlFor="edit-player-position">Position or role</label><input id="edit-player-position" name="position" defaultValue={editingMember.position} /><div className="modal-actions"><button className="secondary-button" type="button" onClick={() => setEditingMember(null)}>Cancel</button><button className="primary-button" type="submit">Save changes</button></div></form></Modal>}

      {showExpenseForm && <Modal title="Add an expense" onClose={() => setShowExpenseForm(false)}><p className="modal-intro">Log a team cost. The total is split equally across everyone on the roster.</p>{members.length ? <ExpenseForm members={members} onSubmit={addExpense} onCancel={() => setShowExpenseForm(false)} /> : <EmptyState title="Add your players first" text="You need at least one player on the roster before an expense can be split." action="Add a player" onAction={() => { setShowExpenseForm(false); setShowMemberForm(true); }} />}</Modal>}
      {showReceiptForm && <Modal title="Scan a receipt" onClose={() => setShowReceiptForm(false)}><p className="modal-intro">Capture a receipt image. Text recognition runs in your browser; review the detected details before saving.</p>{members.length ? <ReceiptScanner members={members} onSave={addScannedReceipt} /> : <EmptyState title="Add your players first" text="A receipt expense needs a person to record who paid." action="Add a player" onAction={() => { setShowReceiptForm(false); setShowMemberForm(true); }} />}</Modal>}
    </div>
  );
}

function MemberDocuments({ member, documents, onUpload, onOpen }) {
  const inputId = `member-documents-${member.id}`;
  return <div className="member-documents">
    <label className="document-upload" htmlFor={inputId}>+ Add document</label>
    <input id={inputId} className="visually-hidden" type="file" multiple onChange={(event) => {
      const files = Array.from(event.currentTarget.files || []);
      event.currentTarget.value = '';
      if (files.length) onUpload(member, files);
    }} />
    {documents.map((item) => <button className="document-link" type="button" key={item.id} title={item.fileName} onClick={() => onOpen(item)}>{item.fileName}</button>)}
  </div>;
}

function ExpenseList({ expenses, members, expanded = false }) {
  const memberById = Object.fromEntries(members.map((member) => [member.id, member]));
  return <div className="expense-list">{expenses.map((expense) => {
    const payer = memberById[expense.paidBy];
    const split = splitExpense(expense);
    const shareAmount = split.reduce((sum, part) => sum + part.amount, 0) / (split.length || 1);
    return <article className="expense-row" key={expense.id}><span className="expense-icon"><Icon name="receipt" /></span><div className="expense-info"><strong>{expense.title}</strong><span>{formatDate(expense.date)} <i>·</i> Paid by {payer ? payer.name : 'former player'}</span>{expanded && <small className="expense-share-note">Split across {split.length} {split.length === 1 ? 'player' : 'players'} · {money(shareAmount)} per share</small>}</div><strong className="expense-amount">{money(expense.amount)}</strong></article>;
  })}</div>;
}

function ReceiptList({ receipts, members, documents, onOpen }) {
  const memberById = Object.fromEntries(members.map((member) => [member.id, member]));
  const documentById = Object.fromEntries(documents.map((item) => [item.id, item]));
  return <div className="receipt-list">{[...receipts].sort((first, second) => second.date.localeCompare(first.date)).map((receipt) => (
    <article className="receipt-row" key={receipt.id}>
      <span className="expense-icon"><Icon name="receipt" /></span>
      <span className="receipt-description"><strong>{receipt.description || 'Receipt'}</strong><small>{formatDate(receipt.date)} · {memberById[receipt.personId]?.name || 'Unknown person'}</small></span>
      <strong className="expense-amount">{money(receipt.amount)}</strong>
      {documentById[receipt.id] && <button className="document-link receipt-open" type="button" onClick={() => onOpen(documentById[receipt.id])}>Open image</button>}
    </article>
  ))}</div>;
}

function ReceiptScanner({ members, onSave }) {
  const [file, setFile] = useState(null);
  const [ocrText, setOcrText] = useState('');
  const [date, setDate] = useState(today());
  const [amount, setAmount] = useState('');
  const [description, setDescription] = useState('');
  const [personId, setPersonId] = useState(members[0].id);
  const [scanning, setScanning] = useState(false);
  const [error, setError] = useState('');
  const [saving, setSaving] = useState(false);

  async function scanFile(event) {
    const nextFile = event.currentTarget.files?.[0];
    event.currentTarget.value = '';
    if (!nextFile) return;
    setFile(null);
    if (!nextFile.type.startsWith('image/')) {
      setError('Choose a receipt photo or image. PDF receipt scanning is not supported.');
      return;
    }
    setFile(nextFile);
    setScanning(true);
    setError('');
    try {
      const result = await recognize(nextFile, 'eng');
      const text = result.data.text || '';
      const detected = extractReceiptDetails(text);
      setOcrText(text);
      setDate(detected.date);
      setAmount(detected.amount);
      const firstTextLine = text.split(/\r?\n/).map((line) => line.trim()).find((line) => line.length > 2);
      if (firstTextLine) setDescription(firstTextLine.slice(0, 100));
    } catch (scanError) {
      setError(`Could not read the receipt image: ${scanError.message}`);
    } finally {
      setScanning(false);
    }
  }

  async function submit(event) {
    event.preventDefault();
    if (!file) {
      setError('Choose or capture a receipt image before saving.');
      return;
    }
    if (!Number.isFinite(Number(amount)) || Number(amount) <= 0 || !description.trim() || !date || !personId) {
      setError('Check the date, amount, description, and person before saving.');
      return;
    }
    setSaving(true);
    setError('');
    try {
      await onSave({ date, amount, description, personId, ocrText }, file);
    } catch (saveError) {
      setError(saveError.message);
    } finally {
      setSaving(false);
    }
  }

  return <form className="modal-form" onSubmit={submit}>
    <label htmlFor="receipt-image">Receipt image</label>
    <input id="receipt-image" type="file" accept="image/*" capture="environment" onChange={scanFile} />
    {file && <p className="selected-file">{file.name}</p>}
    {scanning && <p className="scan-status" role="status">Reading receipt in this browser… The first scan may take a little longer.</p>}
    <label htmlFor="receipt-description">Description / merchant</label>
    <input id="receipt-description" value={description} onChange={(event) => setDescription(event.target.value)} required />
    <div className="form-row">
      <div><label htmlFor="receipt-date">Receipt date</label><input id="receipt-date" type="date" value={date} onChange={(event) => setDate(event.target.value)} required /></div>
      <div><label htmlFor="receipt-amount">Amount (₹)</label><input id="receipt-amount" type="number" min="0.01" step="0.01" value={amount} onChange={(event) => setAmount(event.target.value)} required /></div>
    </div>
    <label htmlFor="receipt-person">Paid by</label>
    <select id="receipt-person" value={personId} onChange={(event) => setPersonId(event.target.value)}>{members.map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}</select>
    {ocrText && <details className="ocr-details"><summary>Review text read from receipt</summary><pre>{ocrText}</pre></details>}
    {error && <p className="form-error" role="alert">{error}</p>}
    <div className="modal-actions"><button className="secondary-button" type="button" onClick={() => { setFile(null); setOcrText(''); setDescription(''); setAmount(''); setDate(today()); setError(''); }}>Clear</button><button className="primary-button" type="submit" disabled={scanning || saving}>{saving ? 'Saving…' : 'Add receipt & expense'}</button></div>
  </form>;
}

function ExpenseForm({ members, onSubmit, onCancel }) {
  const [amount, setAmount] = useState('');
  return <form className="modal-form" onSubmit={onSubmit}>
    <label htmlFor="expense-title">What was it for?</label>
    <input id="expense-title" name="title" placeholder="e.g. Match day snacks" required autoFocus />
    <div className="form-row">
      <div><label htmlFor="expense-amount">Total cost (₹)</label><input id="expense-amount" name="amount" type="number" min="0.01" step="0.01" placeholder="0.00" value={amount} onChange={(event) => setAmount(event.target.value)} required /></div>
      <div><label htmlFor="expense-date">Date</label><input id="expense-date" name="date" type="date" defaultValue={today()} required /></div>
    </div>
    <label htmlFor="expense-payer">Paid by</label>
    <select id="expense-payer" name="paidBy" defaultValue={members[0].id}>{members.map((member) => <option key={member.id} value={member.id}>{member.name}</option>)}</select>
    <div className="split-preview"><span className="split-preview-icon">÷</span><span><strong>Split equally between {members.length} {members.length === 1 ? 'player' : 'players'}</strong><small>About {money((Number(amount) || 0) / members.length)} per person</small></span></div>
    <div className="modal-actions"><button className="secondary-button" type="button" onClick={onCancel}>Cancel</button><button className="primary-button" type="submit">Save expense <Icon name="arrow" /></button></div>
  </form>;
}

function EmptyState({ title, text, action, onAction }) {
  return <div className="empty-state"><span className="empty-ornament">✳</span><h3>{title}</h3><p>{text}</p><button className="secondary-button" onClick={onAction}>{action} <Icon name="arrow" size={15} /></button></div>;
}

function Modal({ title, onClose, children }) {
  useEffect(() => {
    function handleKeyDown(event) {
      if (event.key === 'Escape') onClose();
    }
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);

  return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><section className="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title"><div className="modal-heading"><div><span className="eyebrow">THUNDERBAT FC</span><h2 id="modal-title">{title}</h2></div><button className="modal-close" onClick={onClose} aria-label="Close dialog">×</button></div>{children}</section></div>;
}

export default App;
