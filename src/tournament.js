export function getMatchWinner(match) {
  const completedGames = match.games.filter((game) => Number.isInteger(game.team1Score)
    && Number.isInteger(game.team2Score)
    && game.team1Score !== game.team2Score);
  const team1Wins = completedGames.filter((game) => game.team1Score > game.team2Score).length;
  const team2Wins = completedGames.filter((game) => game.team2Score > game.team1Score).length;

  if (team1Wins > team2Wins) return match.team1;
  if (team2Wins > team1Wins) return match.team2;
  return match.winner || '';
}

export function buildStandings(matches) {
  const standings = new Map();

  matches.forEach((match) => {
    [ [match.team1, match.players1], [match.team2, match.players2] ].forEach(([name, players]) => {
      if (!standings.has(name)) {
        standings.set(name, { team: name, players, played: 0, won: 0, lost: 0, points: 0 });
      }
    });

    const winner = getMatchWinner(match);
    if (!winner) return;

    const team1 = standings.get(match.team1);
    const team2 = standings.get(match.team2);
    team1.played += 1;
    team2.played += 1;
    if (winner === match.team1) {
      team1.won += 1;
      team1.points += 2;
      team2.lost += 1;
    } else if (winner === match.team2) {
      team2.won += 1;
      team2.points += 2;
      team1.lost += 1;
    }
  });

  return [...standings.values()]
    .sort((first, second) => second.points - first.points
      || second.won - first.won
      || first.team.localeCompare(second.team))
    .map((team, index) => ({ rank: index + 1, ...team }));
}
