export const TEAM_NAMES = ['Mobile', 'BE', 'QA', 'Product', 'Product Marketing', 'Design', 'DS'];
export const STATUSES = {
  planned: 'Planned',
  todo: 'To do',
  progress: 'In progress',
  done: 'Done',
  dropped: 'Dropped',
};

export const id = () => crypto.randomUUID();

export function belongsToTeam(item, teamId) {
  return item.teamId === teamId || (item.dependencies || []).includes(teamId);
}

export function parseDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error('Enter a valid date.');
  const date = new Date(`${value}T12:00:00Z`);
  if (!Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== value) {
    throw new Error('Enter a valid date.');
  }
  return date;
}

export function addDays(value, days) {
  const date = parseDate(value);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

export function sprintDates(start) {
  const weekday = parseDate(start).getUTCDay();
  const daysSinceMonday = (weekday + 6) % 7;
  const cutoff = addDays(start, -(daysSinceMonday || 7));
  return { start, end: addDays(start, 13), cutoff, release: addDays(cutoff, 7) };
}

export function nextVersion(version, offset) {
  const match = /^(\d+)\.(\d+)(?:\.(\d+))?$/.exec(version.trim());
  if (!match) throw new Error('Use a mobile version such as 4.20 or 4.20.0.');
  const [, major, minor, patch] = match;
  return `${major}.${Number(minor) + offset}${patch === undefined ? '' : `.${patch}`}`;
}

export function createSheet({ name, start, count = 6, firstSprint = 1, version = '4.20', teams = TEAM_NAMES }) {
  parseDate(start);
  if (!name.trim()) throw new Error('Give your sheet a name.');
  if (!Number.isInteger(count) || count < 1 || count > 14) throw new Error('Choose between 1 and 14 sprints.');
  if (!Number.isInteger(firstSprint) || firstSprint < 1 || firstSprint > 9999) throw new Error('Enter a sprint number between 1 and 9999.');
  return {
    id: id(), name: name.trim(),
    teams: teams.map(name => ({ id: id(), name })),
    sprints: Array.from({ length: count }, (_, index) => {
      const sprintStart = addDays(start, index * 14);
      return { id: id(), number: firstSprint + index, version: nextVersion(version, index), ...sprintDates(sprintStart) };
    }),
    initiatives: [],
  };
}

export function moveInitiative(sheet, initiativeId, sprintId, teamId, keepSpillover = false) {
  const item = sheet.initiatives.find(item => item.id === initiativeId);
  if (!item || item.spilloverOf) throw new Error('This initiative cannot be moved.');
  const sourceIndex = sheet.sprints.findIndex(sprint => sprint.id === item.sprintId);
  const targetIndex = sheet.sprints.findIndex(sprint => sprint.id === sprintId);
  if (targetIndex < 0 || !sheet.teams.some(team => team.id === teamId)) throw new Error('Choose a valid sprint and team.');
  if (keepSpillover && targetIndex > sourceIndex) {
    sheet.initiatives.push({
      ...item, id: id(), stakeholders: [...item.stakeholders],
      ...(item.dependencies ? { dependencies: [...item.dependencies] } : {}),
      spilloverOf: item.id, movedTo: sprintId,
    });
  }
  item.sprintId = sprintId;
  item.teamId = teamId;
  // Every historical marker points to the initiative's current destination.
  for (const marker of sheet.initiatives.filter(entry => entry.spilloverOf === item.id)) {
    marker.movedTo = sprintId;
  }
  // A move back to an earlier sprint replaces its old marker.
  sheet.initiatives = sheet.initiatives.filter(entry => entry.spilloverOf !== item.id || entry.sprintId !== sprintId);
  return item;
}

export function validateState(state) {
  if (!state || state.schema !== 1 || !Array.isArray(state.sheets) || state.sheets.length > 100) {
    throw new Error('This is not a supported Common Ground backup.');
  }
  const ids = new Set();
  const checkId = value => {
    if (typeof value !== 'string' || !/^[a-zA-Z0-9-]{1,80}$/.test(value) || ids.has(value)) throw new Error('Invalid or duplicate ID in backup.');
    ids.add(value);
  };
  const text = (value, max = 500) => typeof value === 'string' && value.length <= max;
  for (const sheet of state.sheets) {
    checkId(sheet.id);
    if (!text(sheet.name, 100) || !sheet.name.trim() || !Array.isArray(sheet.teams) || !Array.isArray(sheet.sprints) || !Array.isArray(sheet.initiatives) || !sheet.teams.length || !sheet.sprints.length || sheet.teams.length > 100 || sheet.sprints.length > 100 || sheet.initiatives.length > 10000) throw new Error('Invalid sheet in backup.');
    for (const team of sheet.teams) {
      checkId(team.id);
      if (!text(team.name, 100) || !team.name.trim()) throw new Error('Invalid team in backup.');
    }
    for (const sprint of sheet.sprints) {
      checkId(sprint.id);
      if (!Number.isInteger(sprint.number) || sprint.number < 1 || !text(sprint.version, 30)) throw new Error('Invalid sprint in backup.');
      for (const field of ['start', 'end', 'cutoff', 'release']) parseDate(sprint[field]);
      if (sprint.end < sprint.start || sprint.release < sprint.cutoff) throw new Error('Invalid sprint dates in backup.');
    }
    for (const item of sheet.initiatives) {
      checkId(item.id);
      if (!text(item.title, 200) || !item.title.trim() || !text(item.description, 5000) || !text(item.owner, 100) || !Object.hasOwn(STATUSES, item.status) || !Array.isArray(item.stakeholders) || !item.stakeholders.every(teamId => sheet.teams.some(team => team.id === teamId)) || !sheet.teams.some(team => team.id === item.teamId) || !sheet.sprints.some(sprint => sprint.id === item.sprintId)) throw new Error('Invalid initiative in backup.');
      if (item.spilloverOf && (!sheet.initiatives.some(entry => entry.id === item.spilloverOf && !entry.spilloverOf) || !sheet.sprints.some(sprint => sprint.id === item.movedTo))) throw new Error('Invalid spillover in backup.');
      if (item.dependencies !== undefined && (!Array.isArray(item.dependencies) || !item.dependencies.every(teamId => sheet.teams.some(team => team.id === teamId)))) throw new Error('Invalid team dependencies in backup.');
    }
  }
  if (state.sheets.length ? !state.sheets.some(sheet => sheet.id === state.activeSheetId) : state.activeSheetId !== null) throw new Error('Invalid active sheet in backup.');
  return state;
}

export function initialState(today = new Date()) {
  const quarter = Math.floor(today.getMonth() / 3);
  let start = `${today.getFullYear()}-${String(quarter * 3 + 1).padStart(2, '0')}-01`;
  start = addDays(start, (3 - parseDate(start).getUTCDay() + 7) % 7);
  const sheet = createSheet({ name: `Q${quarter + 1} ${today.getFullYear()}`, start });
  return { schema: 1, activeSheetId: sheet.id, sheets: [sheet] };
}

export function repairSprintNumbering(state) {
  for (const sheet of state.sheets) {
    if (new Set(sheet.sprints.map(sprint => sprint.number)).size === sheet.sprints.length) continue;
    let number = sheet.sprints[0].number;
    let previousQuarter;
    const legacyPattern = sheet.sprints.every(sprint => {
      const quarter = `${sprint.start.slice(0, 4)}-${Math.floor((Number(sprint.start.slice(5, 7)) - 1) / 3)}`;
      if (previousQuarter && quarter !== previousQuarter) number = 1;
      previousQuarter = quarter;
      return sprint.number === number++;
    });
    if (legacyPattern) {
      const first = sheet.sprints[0].number;
      sheet.sprints.forEach((sprint, index) => { sprint.number = first + index; });
    }
  }
  return state;
}

// Older installations saved unmarked sample records. Remove only unchanged samples.
export function removeLegacySamples(state) {
  const samples = [
    [0, 0, 'A smoother first five minutes', 'progress', 'Alex Chen', 'Simplify onboarding and reduce the steps to first value.', [3, 5]],
    [0, 1, 'Offline mode foundations', 'planned', 'Alex Chen', 'Local storage architecture and sync queue for offline-first experiences.', [1]],
    [0, 2, 'Ship the new home experience', 'planned', 'Alex Chen', 'Roll out the refreshed mobile home screen.', [5, 2]],
    [1, 0, 'Permissions API v2', 'done', 'Sam Rivera', 'Scoped roles and team-level access controls.', [0, 2]],
    [1, 1, 'Make search feel instant', 'progress', 'Sam Rivera', 'Index optimization and query performance improvements.', [6]],
    [1, 2, 'Event pipeline migration', 'todo', 'Sam Rivera', 'Move product events to the new ingestion pipeline.', [6]],
    [2, 0, 'Onboarding regression suite', 'progress', 'Taylor Kim', 'Cover the new sign-up and activation paths.', [0]],
    [2, 1, 'Release readiness checks', 'todo', 'Taylor Kim', 'Validate critical flows before the mobile release.', [0, 1]],
    [3, 0, 'Define our activation moment', 'done', 'Jordan Lee', 'Align the team on what meaningful activation looks like.', [6]],
    [3, 1, 'Discovery: team workspaces', 'progress', 'Jordan Lee', 'Customer interviews and opportunity mapping.', [5]],
    [3, 3, 'QBR & next-quarter priorities', 'planned', 'Jordan Lee', 'Review outcomes and align on next quarter.', []],
    [4, 1, 'Launch story & messaging', 'todo', 'Morgan Bell', 'Positioning, launch narrative, and key messages.', [3]],
    [4, 2, 'Customer launch kit', 'planned', 'Morgan Bell', 'Email, help center, and release notes.', [5]],
    [5, 0, 'Onboarding, reimagined', 'done', 'Casey Park', 'High-fidelity designs and handoff for mobile onboarding.', [0, 3]],
    [5, 1, 'Workspace interaction study', 'progress', 'Casey Park', 'Explore how teams organize and share their work.', [3]],
    [5, 2, 'Design system: empty states', 'planned', 'Casey Park', 'Reusable patterns for helpful empty states.', [0]],
    [6, 0, 'Activation baseline', 'done', 'Jamie Patel', 'Establish the baseline and agree on success metrics.', [3]],
    [6, 1, 'Experiment measurement plan', 'todo', 'Jamie Patel', 'Define tracking, guardrails, and experiment readout.', [3, 1]],
    [6, 3, 'Retention cohort dashboard', 'planned', 'Jamie Patel', 'A shared view of retention by activation cohort.', [3]],
  ];
  for (const sheet of state.sheets) {
    if (!/^Q[1-4] \d{4}$/.test(sheet.name)) continue;
    const removed = new Set();
    sheet.initiatives = sheet.initiatives.filter(item => {
      const unchanged = !item.spilloverOf && !sheet.initiatives.some(marker => marker.spilloverOf === item.id) && samples.some(([team, sprint, title, status, owner, description, stakeholders]) =>
        item.title === title && item.status === status && item.owner === owner && item.description === description &&
        sheet.teams.find(entry => entry.id === item.teamId)?.name === TEAM_NAMES[team] &&
        item.sprintId === sheet.sprints[sprint]?.id &&
        JSON.stringify(item.stakeholders.map(teamId => sheet.teams.find(entry => entry.id === teamId)?.name)) === JSON.stringify(stakeholders.map(index => TEAM_NAMES[index])));
      if (unchanged) removed.add(item.id);
      return !unchanged;
    });
    if (removed.size) {
      const oldFirst = (Number(sheet.name[1]) - 1) * 6 + 1;
      if (sheet.sprints.every((sprint, index) => sprint.number === oldFirst + index)) {
        sheet.sprints.forEach((sprint, index) => { sprint.number = index + 1; });
      }
    }
  }
  return state;
}
