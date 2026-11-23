import { STATUSES, TEAM_NAMES, id, belongsToTeam, createSheet, initialState, removeLegacySamples, repairSprintNumbering, moveInitiative, validateState, parseDate } from './model.js';
import { buildWorkbook } from './excel.js';

const STORAGE = 'common-ground-v1';
const icons = {
  plus: '<path d="M12 5v14M5 12h14"/>',
  grid: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
  calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 11h18"/>',
  arrow: '<path d="M5 12h14m-5-5 5 5-5 5"/>',
  search: '<circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 4 4"/>',
  settings: '<path d="M4 7h16M4 17h16"/><circle cx="9" cy="7" r="2"/><circle cx="15" cy="17" r="2"/>',
  down: '<path d="m7 10 5 5 5-5"/>',
  download: '<path d="M12 3v12m-5-5 5 5 5-5M5 16v5h14v-5"/>',
  expand: '<path d="M8 3H3v5m13-5h5v5M3 16v5h5m8 0h5v-5"/>',
  collapse: '<path d="M3 8h5V3m8 0v5h5M8 21v-5H3m18 0h-5v5"/>',
  check: '<path d="m5 12 4 4L19 6"/>',
  close: '<path d="m6 6 12 12M6 18 18 6"/>',
  link: '<path d="m10 13 4-4m-5 7-2 2a4 4 0 0 1-6-6l4-4a4 4 0 0 1 6 0m2 0 2-2a4 4 0 0 1 6 6l-4 4a4 4 0 0 1-6 0"/>',
};
const icon = name => `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${icons[name] || icons.grid}</svg>`;
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
const formatDate = value => new Intl.DateTimeFormat('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' }).format(parseDate(value));
const today = () => new Date().toLocaleDateString('en-CA');
const app = document.querySelector('#app');
const modal = document.querySelector('#modal');
let state;
let storageError = '';
try {
  const stored = localStorage.getItem(STORAGE);
  state = stored ? repairSprintNumbering(removeLegacySamples(validateState(JSON.parse(stored)))) : initialState();
} catch {
  state = initialState();
  storageError = 'Saved data could not be loaded. Export any work before closing this tab.';
}
let search = '';
let statusFilter = 'all';
let teamFilter = 'all';
let dragId = null;
let toastTimer;
let boardExpanded = false;
let excelLoading;
let sidebarCollapsed = false;
try { sidebarCollapsed = localStorage.getItem('common-ground-sidebar-collapsed') === 'true'; } catch {}
document.body.classList.toggle('sidebar-collapsed', sidebarCollapsed);
const sheet = () => state.sheets.find(entry => entry.id === state.activeSheetId);
const realItems = () => sheet()?.initiatives.filter(item => !item.spilloverOf) || [];

function toast(message) {
  const target = document.querySelector('#toast');
  target.textContent = message;
  target.classList.add('visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => target.classList.remove('visible'), 4000);
}

function save(message) {
  try {
    localStorage.setItem(STORAGE, JSON.stringify(state));
    storageError = '';
  } catch {
    storageError = 'Browser storage is unavailable or full. Export a backup to keep your changes.';
  }
  render();
  if (message) toast(storageError || message);
}

function filtered(item) {
  const s = sheet();
  const teamName = s.teams.find(team => team.id === item.teamId)?.name || '';
  const dependencies = (item.dependencies || []).map(id => s.teams.find(team => team.id === id)?.name || '').join(' ');
  return (statusFilter === 'all' || (statusFilter === 'spillover' ? !!item.spilloverOf : !item.spilloverOf && item.status === statusFilter)) &&
    (!search || `${item.title} ${item.description} ${teamName} ${dependencies}`.toLowerCase().includes(search.toLowerCase()));
}

function card(item, rowTeamId) {
  const s = sheet();
  const dependencyCopy = rowTeamId !== item.teamId;
  const responsibleTeam = s.teams.find(team => team.id === item.teamId)?.name || '';
  const destination = s.sprints.find(sprint => sprint.id === item.movedTo);
  return `<button class="initiative ${item.spilloverOf ? 'spillover-card' : ''} ${item.status === 'dropped' ? 'dropped-card' : ''}" data-action="edit-item" data-id="${item.id}" draggable="${!item.spilloverOf}" aria-label="${esc(item.title)}, ${item.spilloverOf ? 'Spillover' : STATUSES[item.status]}. Open details.">
    <span class="card-top"><span class="status ${item.spilloverOf ? 'spillover' : item.status}"><i></i>${item.spilloverOf ? 'Spillover' : STATUSES[item.status]}</span><span class="card-menu" aria-hidden="true">&middot;&middot;&middot;</span></span>
    <span class="card-title">${esc(item.title)}</span>
    ${dependencyCopy ? `<span class="dependency-source">${icon('link')} Linked from ${esc(responsibleTeam)}</span>` : ''}
    ${item.spilloverOf ? `<span class="spillover-target">${icon('arrow')} Moved to Sprint ${destination?.number ?? '?'}</span>` : ''}
    ${item.dependencies?.length ? `<span class="dependency-tags"><span class="dependency-label">Depends on</span>${item.dependencies.map(id => s.teams.find(team => team.id === id)).filter(Boolean).map(team => `<span class="team-tag">${esc(team.name)}</span>`).join('')}</span>` : ''}
  </button>`;
}

function renderBoard() {
  const s = sheet();
  const teams = s.teams.filter(team => teamFilter === 'all' || team.id === teamFilter);
  const items = realItems();
  const current = s.sprints.find(sprint => sprint.start <= today() && sprint.end >= today());
  const container = document.querySelector('#board');
  const scroll = container?.scrollLeft || 0;
  container.innerHTML = `<table class="planning-table"><caption class="sr-only">${esc(s.name)} sprint initiatives organized by team. Dates can be edited from each sprint header.</caption><thead><tr>
    <th class="team-column" scope="col"><span class="eyebrow">WORKSTREAMS</span><span class="team-header">Teams <span class="count">${s.teams.length}</span></span><button class="text-button" data-action="manage-teams">${icon('settings')} Manage teams</button></th>
    ${s.sprints.map((sprint, index) => {
      const sprintItems = items.filter(item => item.sprintId === sprint.id);
      const done = sprintItems.filter(item => item.status === 'done').length;
      return `<th scope="col" class="sprint-column ${current?.id === sprint.id ? 'current-sprint' : ''}"><button class="sprint-edit" data-action="edit-sprint" data-id="${sprint.id}" title="Edit sprint dates and version">
        <span class="sprint-label">SPRINT ${String(sprint.number).padStart(2, '0')} ${current?.id === sprint.id ? '<span class="current-label">CURRENT</span>' : `<span class="sprint-index">${String(index + 1).padStart(2, '0')}</span>`}</span>
        <span class="sprint-range">${formatDate(sprint.start)} <span>&ndash;</span> ${formatDate(sprint.end)}</span>
        <span class="version-label">${icon('link')} Mobile ${esc(sprint.version)} <span class="edit-hint">Edit</span></span>
        <span class="date-line"><span>Mobile cutoff</span><strong>${formatDate(sprint.cutoff)}</strong></span>
        <span class="date-line"><span>Mobile release</span><strong>${formatDate(sprint.release)}</strong></span>
        <span class="sprint-progress"><span><span style="width:${sprintItems.length ? done / sprintItems.length * 100 : 0}%"></span></span><small>${done}/${sprintItems.length} done</small></span>
      </button></th>`;
    }).join('')}</tr></thead><tbody>
    ${teams.map((team, index) => {
      const teamItems = items.filter(item => belongsToTeam(item, team.id));
      return `<tr><th scope="row" class="team-column team-row"><span class="team-symbol team-color-${s.teams.indexOf(team) % 7}">${esc(team.name === 'Product Marketing' ? 'PMM' : team.name === 'Product' ? 'PM' : team.name === 'Design' ? 'De' : team.name === 'Mobile' ? 'Mo' : team.name.slice(0, 2))}</span><span class="team-name">${esc(team.name)}</span><span class="team-item-count">${teamItems.length} initiative${teamItems.length === 1 ? '' : 's'}</span><button class="row-add" data-action="add-item" data-team="${team.id}" data-sprint="${s.sprints[0].id}" aria-label="Add initiative for ${esc(team.name)}">${icon('plus')}</button></th>
        ${s.sprints.map(sprint => {
          const cellItems = s.initiatives.filter(item => belongsToTeam(item, team.id) && item.sprintId === sprint.id && filtered(item));
          return `<td class="board-cell ${current?.id === sprint.id ? 'current-cell' : ''}" data-team="${team.id}" data-sprint="${sprint.id}"><div class="cell-content">${cellItems.map(item => card(item, team.id)).join('')}<button class="cell-add ${cellItems.length ? '' : 'empty-add'}" data-action="add-item" data-team="${team.id}" data-sprint="${sprint.id}">${icon('plus')} Add initiative</button></div></td>`;
        }).join('')}</tr>`;
    }).join('')}
    </tbody></table>`;
  container.scrollLeft = scroll;
  const visible = s.initiatives.filter(item => (teamFilter === 'all' || belongsToTeam(item, teamFilter)) && filtered(item)).length;
  document.querySelector('#board-caption').textContent = `Showing ${visible} initiative${visible === 1 ? '' : 's'} across ${teams.length} team${teams.length === 1 ? '' : 's'} and ${s.sprints.length} sprints`;
}

function render() {
  const s = sheet();
  if (!s) {
    boardExpanded = false;
    document.body.classList.remove('board-expanded');
  }
  const items = realItems();
  const done = items.filter(item => item.status === 'done').length;
  const progress = items.filter(item => item.status === 'progress').length;
  const planned = items.filter(item => ['planned', 'todo'].includes(item.status)).length;
  const pct = items.length ? Math.round(done / items.length * 100) : 0;
  app.innerHTML = `<aside class="sidebar" id="sidebar">
    <a class="brand" href="./" aria-label="Common Ground home"><span class="brand-mark"><i></i><i></i><i></i><i></i></span><span>common<br>ground<span class="brand-dot">.</span></span></a>
    <div class="personal-space">Your personal planning space</div>
    <span class="nav-label">PLANNING</span>
    <button class="nav-item active" data-action="board">${icon('grid')} Team alignment <span>${state.sheets.length}</span></button>
    <button class="nav-item" data-action="schedule" ${s ? '' : 'disabled'}>${icon('calendar')} Release calendar</button>
    <div class="sidebar-divider"></div><div class="nav-label sheets-label">YOUR SHEETS <button data-action="new-sheet" aria-label="Create sheet">${icon('plus')}</button></div>
    <div class="sheet-links">${state.sheets.map(entry => `<button class="sheet-link ${entry.id === s.id ? 'selected' : ''}" data-action="switch-sheet" data-id="${entry.id}"><span class="sheet-dot"></span>${esc(entry.name)}${entry.id === s.id ? icon('check') : ''}</button>`).join('')}</div>
    <button class="new-sheet-link" data-action="new-sheet">${icon('plus')} New planning sheet</button>
    <div class="sidebar-bottom"><div class="local-note"><span class="local-status"></span> Saved on this browser<small>Export a backup to keep your plan safe.</small></div><button class="backup-link" data-action="backup">${icon('download')} Backups & restore</button></div>
  </aside>
   <main><header class="topbar"><div class="topbar-left"><button class="icon-button" data-action="toggle-sidebar" aria-controls="sidebar" aria-expanded="${!sidebarCollapsed}" aria-label="${sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}" title="${sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}">${icon('grid')}</button><div class="breadcrumb">My planning <span>/</span> <strong>Team alignment</strong></div></div><div class="topbar-right"><span class="autosave">${icon('check')} Autosaved locally</span><button class="icon-button" data-action="backup" title="Export or import a backup">${icon('download')}</button></div></header>
    <div class="main-content"><section class="page-heading"><div><div class="eyebrow heading-eyebrow"><span></span> YOUR PLAN. AT A GLANCE.</div><h1>Team alignment</h1><p>Your personal view of initiatives, sprints, and releases.</p></div><button class="button primary" data-action="new-sheet">${icon('plus')} New sheet</button></section>
    ${storageError ? `<div class="warning">${esc(storageError)} <button data-action="backup">Export backup</button></div>` : ''}
    ${s ? `
    <section class="overview" aria-label="Sheet overview"><div class="overview-intro"><span class="overview-label">THE BIG PICTURE</span><strong>${esc(s.name)} <span>at a glance</span></strong><span>${s.sprints.length} sprints <i>&middot;</i> ${s.teams.length} teams</span></div><div class="metric"><span class="metric-icon neutral">${icon('grid')}</span><span><strong>${items.length}</strong><small>Total initiatives</small></span></div><div class="metric"><span class="metric-icon amber"><span class="progress-ring"></span></span><span><strong>${progress}</strong><small>In progress</small></span></div><div class="metric"><span class="metric-icon green">${icon('check')}</span><span><strong>${done}</strong><small>Completed</small></span></div><div class="completion"><span><strong>${pct}%</strong> complete</span><div class="completion-track"><span style="width:${pct}%"></span></div><small>${planned} planned or ready to start</small></div></section>
    <section class="board-section" aria-label="Sprint board"><div class="board-toolbar"><div class="board-title">${icon('grid')} <h2>Sprint board</h2><span class="quarter-pill">${esc(s.name)}</span></div><div class="board-actions"><button class="button compact" data-action="expand-board" aria-expanded="${boardExpanded}" aria-controls="board" title="${boardExpanded ? 'Exit expanded view (Esc)' : 'Expand board to fill the screen'}">${icon(boardExpanded ? 'collapse' : 'expand')} ${boardExpanded ? 'Exit expanded view' : 'Expand board'}</button><button class="button compact" data-action="export-excel" title="Export the full sheet, including hidden or filtered initiatives">${icon('download')} Export Excel</button><button class="text-button" data-action="manage-teams">${icon('settings')} Teams</button><button class="button compact" data-action="add-item">${icon('plus')} Add initiative</button></div></div>
      <div class="filterbar"><div class="filter-group"><label class="search-field">${icon('search')}<input id="search" type="search" placeholder="Find an initiative..." value="${esc(search)}" aria-label="Search initiatives"></label><select id="status-filter" aria-label="Filter by status"><option value="all">All statuses</option>${Object.entries(STATUSES).map(([value, label]) => `<option value="${value}" ${statusFilter === value ? 'selected' : ''}>${label}</option>`).join('')}<option value="spillover" ${statusFilter === 'spillover' ? 'selected' : ''}>Spillover</option></select><select id="team-filter" aria-label="Filter by team"><option value="all">All teams</option>${s.teams.map(team => `<option value="${team.id}" ${teamFilter === team.id ? 'selected' : ''}>${esc(team.name)}</option>`).join('')}</select></div><span class="drag-tip">${icon('arrow')} Drag initiatives to rearrange</span></div>
      <div id="board" class="board-scroll" tabindex="0" aria-label="Sprint planning board, scroll horizontally to see more sprints"></div>
      <div class="board-footer"><span id="board-caption"></span><button class="text-button danger-text" data-action="delete-sheet">Delete sheet</button></div>
    </section>` : `<section class="empty-sheets"><span class="eyebrow">A CLEAN SLATE</span><h2>No planning sheets yet</h2><p>Create a sheet to plan your next quarter, or restore a saved backup.</p><div><button class="button primary" data-action="new-sheet">${icon('plus')} Create your first sheet</button><button class="button" data-action="backup">Import backup</button></div></section>`}<footer class="page-footer"><span>Good alignment makes great work possible.</span><button data-action="backup">${icon('download')} Export & import</button></footer></div>
  </main>`;
  if (s) renderBoard();
}

function openDialog(title, subtitle, content, onSubmit, { wide = false, submit = 'Save changes', extra = '' } = {}) {
  modal.innerHTML = `<form id="dialog-form"><header class="dialog-header"><div><span class="eyebrow">COMMON GROUND</span><h2 id="dialog-title">${esc(title)}</h2>${subtitle ? `<p>${esc(subtitle)}</p>` : ''}</div><button type="button" class="icon-button" data-close aria-label="Close dialog">${icon('close')}</button></header><div class="dialog-body">${content}<p class="form-error" role="alert" id="form-error"></p></div><footer class="dialog-footer">${extra}<div class="dialog-buttons"><button type="button" class="button" data-close>Cancel</button>${onSubmit ? `<button type="submit" class="button primary">${esc(submit)}</button>` : ''}</div></footer></form>`;
  modal.classList.toggle('wide', wide);
  modal.querySelectorAll('[data-close]').forEach(button => button.onclick = () => modal.close());
  modal.querySelector('form').onsubmit = event => {
    event.preventDefault();
    if (!onSubmit) return;
    try {
      const result = onSubmit(new FormData(event.target));
      if (result !== false) modal.close();
    } catch (error) {
      modal.querySelector('#form-error').textContent = error.message;
    }
  };
  if (!modal.open) modal.showModal();
}

const field = (label, name, value, type = 'text', attrs = '') => `<label class="field"><span>${label}</span><input name="${name}" type="${type}" value="${esc(value)}" ${attrs}></label>`;
const teamOptions = selected => sheet().teams.map(team => `<option value="${team.id}" ${team.id === selected ? 'selected' : ''}>${esc(team.name)}</option>`).join('');
const sprintOptions = selected => sheet().sprints.map(sprint => `<option value="${sprint.id}" ${sprint.id === selected ? 'selected' : ''}>Sprint ${sprint.number} &middot; ${formatDate(sprint.start)}</option>`).join('');

function newSheet() {
  const date = new Date();
  const quarter = Math.floor(date.getMonth() / 3) + 1;
  const year = date.getFullYear();
  let start = `${year}-${String((quarter - 1) * 3 + 1).padStart(2, '0')}-01`;
  const offset = (3 - parseDate(start).getUTCDay() + 7) % 7;
  const next = parseDate(start); next.setUTCDate(next.getUTCDate() + offset); start = next.toISOString().slice(0, 10);
  openDialog('Plan a fresh quarter.', 'Set the first sprint to generate your schedule.', `
    ${field('Sheet name', 'name', `Q${quarter} ${year}`, 'text', 'required maxlength="100"')}
    <div class="field-grid">${field('First sprint starts', 'start', start, 'date', 'required')}${field('Number of sprints', 'count', 6, 'number', 'required min="1" max="14"')}</div>
    ${field('First mobile version', 'version', '4.20', 'text', 'required maxlength="30" placeholder="4.20 or 4.20.0"')}
    <p class="field-help">Each quarterly sheet starts at Sprint 1 and counts forward, even if its dates cross a calendar-quarter boundary.</p>
    <div class="schedule-note">${icon('calendar')}<div><strong>Your schedule, automatically.</strong><p>Sprints run for 14 days. Mobile cutoff is the Monday before the start; release is the following Monday. Versions increase by one minor version per sprint. All dates and versions can be edited.</p></div></div>
    ${sheet() ? '<label class="checkbox-line"><input type="checkbox" name="currentTeams" checked> Use teams from the current sheet</label>' : ''}
    <p class="field-help">Wednesday starts are recommended, but any start date is supported. Six sprints cover 12 weeks; use seven to cover a full 13-week quarter.</p>`, data => {
    const s = createSheet({ name: data.get('name'), start: data.get('start'), count: Number(data.get('count')), version: data.get('version'), teams: data.has('currentTeams') ? sheet().teams.map(team => team.name) : TEAM_NAMES });
    state.sheets.push(s); state.activeSheetId = s.id;
    search = ''; statusFilter = 'all'; teamFilter = 'all'; save('Your new planning sheet is ready.');
  }, { submit: 'Create sheet' });
}

function editSprint(sprintId) {
  const sprint = sheet().sprints.find(entry => entry.id === sprintId);
  openDialog(`Sprint ${sprint.number}`, 'Fine-tune this sprint without changing the rest of your schedule.', `
    <div class="field-grid">${field('Sprint number', 'number', sprint.number, 'number', 'required min="1" max="9999"')}${field('Mobile version', 'version', sprint.version, 'text', 'required maxlength="30"')}</div>
    <div class="field-grid">${field('Sprint start', 'start', sprint.start, 'date', 'required')}${field('Sprint end', 'end', sprint.end, 'date', 'required')}</div>
    <div class="field-grid">${field('Mobile cutoff', 'cutoff', sprint.cutoff, 'date', 'required')}${field('Mobile release', 'release', sprint.release, 'date', 'required')}</div>
    <p class="field-help">Date overrides only affect this sprint. Custom cutoff and release days are welcome.</p>`, data => {
    const values = Object.fromEntries(data);
    for (const key of ['start', 'end', 'cutoff', 'release']) parseDate(values[key]);
    if (values.end < values.start) throw new Error('Sprint end must be on or after its start.');
    if (values.release < values.cutoff) throw new Error('Mobile release must be on or after cutoff.');
    if (sheet().sprints.some(entry => entry.id !== sprintId && entry.number === Number(values.number))) throw new Error('That sprint number is already used in this sheet.');
    Object.assign(sprint, values, { number: Number(values.number), version: values.version.trim() });
    save('Sprint schedule updated.');
  });
}

function editItem(itemId, teamId, sprintId) {
  const s = sheet();
  const original = s.initiatives.find(entry => entry.id === itemId);
  if (original?.spilloverOf) {
    const target = s.initiatives.find(entry => entry.id === original.spilloverOf);
    openDialog('A trace of the original plan', 'This initiative moved to another sprint. Its original place is preserved here.', `<div class="spillover-detail"><span class="status spillover"><i></i>Spillover</span><h3>${esc(original.title)}</h3><p>Now in Sprint ${s.sprints.find(sprint => sprint.id === target?.sprintId)?.number ?? '?'} with ${esc(s.teams.find(team => team.id === target?.teamId)?.name || 'another team')}.</p><button type="button" class="button primary" id="open-current">Open current initiative ${icon('arrow')}</button><button type="button" class="text-button danger-text" id="remove-marker">Remove this marker</button></div>`, null);
    modal.querySelector('#open-current').onclick = () => editItem(target.id);
    modal.querySelector('#remove-marker').onclick = () => {
      confirmAction('Remove spillover marker?', 'The current initiative will not be changed.', () => { s.initiatives = s.initiatives.filter(item => item.id !== original.id); save('Spillover marker removed.'); }, 'Remove marker');
    };
    return;
  }
  const item = original || { title: '', description: '', owner: '', status: 'planned', teamId: teamId || (teamFilter !== 'all' ? teamFilter : s.teams[0].id), sprintId: sprintId || s.sprints[0].id, stakeholders: [] };
  openDialog(original ? 'Initiative details' : 'Make space for an initiative.', 'Track the work, its responsible team, and its next step.', `
    ${field('Initiative title', 'title', item.title, 'text', 'required maxlength="200" placeholder="What are we working toward?"')}
    <label class="field"><span>Description</span><textarea name="description" rows="3" maxlength="5000" placeholder="Context, expected outcomes, or useful links...">${esc(item.description)}</textarea></label>
    <div class="field-grid"><label class="field"><span>Responsible team</span><select name="teamId">${teamOptions(item.teamId)}</select></label><label class="field"><span>Sprint</span><select name="sprintId">${sprintOptions(item.sprintId)}</select></label></div>
    <label class="field"><span>Status</span><select name="status">${Object.entries(STATUSES).map(([value, label]) => `<option value="${value}" ${item.status === value ? 'selected' : ''}>${label}</option>`).join('')}</select></label>
    <fieldset class="stakeholder-field"><legend>Team dependencies (optional)</legend><div class="checkbox-pills">${s.teams.map(team => `<label><input type="checkbox" name="dependencies" value="${team.id}" ${(item.dependencies || []).includes(team.id) ? 'checked' : ''}><span>${esc(team.name)}</span></label>`).join('')}</div><p class="field-help">Selected teams get a linked copy in their row. All copies share the same status, details, and sprint.</p></fieldset>`, data => {
    const values = { title: data.get('title').trim(), description: data.get('description').trim(), status: data.get('status'), teamId: data.get('teamId'), sprintId: data.get('sprintId'), dependencies: data.getAll('dependencies') };
    if (!values.title) throw new Error('Give your initiative a title.');
    values.dependencies = values.dependencies.filter(teamId => teamId !== values.teamId);
    if (original) {
      const later = s.sprints.findIndex(entry => entry.id === values.sprintId) > s.sprints.findIndex(entry => entry.id === original.sprintId);
      if (later) {
        promptMove(original.id, values.sprintId, values.teamId, values);
        return false;
      }
      if (original.sprintId !== values.sprintId || original.teamId !== values.teamId) moveInitiative(s, original.id, values.sprintId, values.teamId);
      Object.assign(original, values);
    } else s.initiatives.push({ id: id(), owner: '', stakeholders: [], ...values });
    save(original ? 'Initiative updated.' : 'Initiative added.');
  }, { submit: original ? 'Save changes' : 'Add initiative', extra: original ? '<button type="button" class="text-button danger-text" id="delete-item">Delete initiative</button>' : '' });
  if (original) modal.querySelector('#delete-item').onclick = () => confirmAction('Delete this initiative?', 'This will also remove its spillover markers. This action cannot be undone.', () => { s.initiatives = s.initiatives.filter(entry => entry.id !== original.id && entry.spilloverOf !== original.id); save('Initiative deleted.'); }, 'Delete initiative');
}

function promptMove(itemId, sprintId, teamId, values) {
  const s = sheet();
  const item = s.initiatives.find(entry => entry.id === itemId);
  const from = s.sprints.find(entry => entry.id === item.sprintId);
  const to = s.sprints.find(entry => entry.id === sprintId);
  openDialog('Keep a trace of the move?', `Moving "${item.title}" from Sprint ${from.number} to Sprint ${to.number}.`, `<div class="move-visual"><span>Sprint ${from.number}</span>${icon('arrow')}<span>Sprint ${to.number}</span></div><label class="spillover-option"><input type="checkbox" name="spillover" checked><span><strong>Mark as spillover in Sprint ${from.number}</strong><small>Leave a linked marker in the previous column so the original commitment stays visible.</small></span></label><p class="field-help">Uncheck to move it without leaving a marker. Its status will stay the same.</p>`, data => {
    moveInitiative(s, itemId, sprintId, teamId, data.has('spillover'));
    if (values) Object.assign(item, values);
    save(data.has('spillover') ? 'Initiative moved. Spillover marked in the previous sprint.' : 'Initiative moved.');
  }, { submit: 'Move initiative' });
}

function confirmAction(title, message, callback, submit = 'Confirm') {
  openDialog(title, message, '<p class="field-help">Please confirm before continuing.</p>', () => callback(), { submit });
}

function deleteSheet() {
  const current = sheet();
  confirmAction(`Delete "${current.name}"?`, 'This permanently deletes the sheet, its initiatives, and spillover markers. Export a backup first if you need to keep it. This cannot be undone.', () => {
    const index = state.sheets.findIndex(entry => entry.id === current.id);
    state.sheets.splice(index, 1);
    state.activeSheetId = state.sheets[Math.min(index, state.sheets.length - 1)]?.id || null;
    search = ''; statusFilter = 'all'; teamFilter = 'all';
    save('Sheet deleted.');
  }, 'Delete sheet');
}

function manageTeams() {
  const s = sheet();
  openDialog('Organize your workstreams.', 'Add, rename, or remove teams from this sheet.', `<div class="team-manager">${s.teams.map(team => `<div class="manage-team-row"><input aria-label="Team name" maxlength="100" value="${esc(team.name)}" data-team-name="${team.id}"><button class="icon-button" type="button" data-remove-team="${team.id}" aria-label="Remove ${esc(team.name)}">${icon('close')}</button></div>`).join('')}</div><div class="add-team-field"><input id="new-team-name" placeholder="New team name" maxlength="100" aria-label="New team name"><button type="button" class="button" id="add-team">${icon('plus')} Add team</button></div><p class="field-help">Removing a team with initiatives lets you reassign the work first.</p>`, () => {
    const inputs = [...modal.querySelectorAll('[data-team-name]')];
    const names = inputs.map(input => input.value.trim());
    if (names.some(name => !name)) throw new Error('Every team needs a name.');
    if (new Set(names.map(name => name.toLowerCase())).size !== names.length) throw new Error('Team names must be unique.');
    inputs.forEach((input, index) => { s.teams.find(team => team.id === input.dataset.teamName).name = names[index]; });
    save('Teams updated.');
  });
  modal.querySelector('#add-team').onclick = () => {
    const name = modal.querySelector('#new-team-name').value.trim();
    if (!name || s.teams.some(team => team.name.toLowerCase() === name.toLowerCase())) { modal.querySelector('#form-error').textContent = 'Enter a unique team name.'; return; }
    // Keep any names edited in this dialog when adding a new team.
    const inputs = [...modal.querySelectorAll('[data-team-name]')];
    const names = inputs.map(input => input.value.trim());
    if (names.some(name => !name) || new Set([...names, name].map(name => name.toLowerCase())).size !== names.length + 1) { modal.querySelector('#form-error').textContent = 'Team names must be nonempty and unique.'; return; }
    inputs.forEach((input, index) => { s.teams.find(team => team.id === input.dataset.teamName).name = names[index]; });
    s.teams.push({ id: id(), name }); save('Team added.'); manageTeams();
  };
  modal.querySelectorAll('[data-remove-team]').forEach(button => button.onclick = () => {
    if (s.teams.length === 1) { modal.querySelector('#form-error').textContent = 'Keep at least one team in your sheet.'; return; }
    const team = s.teams.find(entry => entry.id === button.dataset.removeTeam);
    const affected = s.initiatives.filter(item => item.teamId === team.id);
    openDialog(`Remove ${team.name}?`, `${affected.length} initiative or spillover marker${affected.length === 1 ? '' : 's'} belong to this team.`, affected.length ? `<label class="field"><span>Reassign work to</span><select name="replacement">${s.teams.filter(entry => entry.id !== team.id).map(entry => `<option value="${entry.id}">${esc(entry.name)}</option>`).join('')}</select></label>` : '<p class="field-help">This removes the team row from your plan.</p>', data => {
      for (const item of s.initiatives) {
        if (item.teamId === team.id) item.teamId = data.get('replacement');
        item.stakeholders = item.stakeholders.filter(id => id !== team.id);
        if (item.dependencies) item.dependencies = item.dependencies.filter(id => id !== team.id);
      }
      s.teams = s.teams.filter(entry => entry.id !== team.id); if (teamFilter === team.id) teamFilter = 'all'; save('Team removed.');
    }, { submit: 'Remove team' });
  });
}

function releaseCalendar() {
  openDialog('Every release, in one place.', 'Edit any sprint to adjust its schedule.', `<div class="calendar-list">${sheet().sprints.map(sprint => `<button type="button" class="calendar-entry" data-calendar-sprint="${sprint.id}"><span class="calendar-number">${String(sprint.number).padStart(2, '0')}</span><span><strong>Sprint ${sprint.number}</strong><small>${formatDate(sprint.start)} &ndash; ${formatDate(sprint.end)}</small></span><span><strong>Mobile ${esc(sprint.version)}</strong><small>Cutoff ${formatDate(sprint.cutoff)} &middot; Release ${formatDate(sprint.release)}</small></span>${icon('arrow')}</button>`).join('')}</div>`, null, { wide: true });
  modal.querySelectorAll('[data-calendar-sprint]').forEach(button => button.onclick = () => editSprint(button.dataset.calendarSprint));
}

function backup() {
  openDialog('Keep your plan safe.', 'Your sheets are saved locally in this browser.', `<div class="backup-options"><button type="button" class="backup-option" id="export-backup">${icon('download')}<span><strong>Export all sheets</strong><small>Download a JSON backup for safekeeping.</small></span>${icon('arrow')}</button><button type="button" class="backup-option" id="import-backup">${icon('plus')}<span><strong>Import a backup</strong><small>Replaces your current sheets after confirmation.</small></span>${icon('arrow')}</button></div><p class="field-help">Keep a backup before clearing browser data or moving to another device.</p>`, null);
  modal.querySelector('#export-backup').onclick = () => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a'); link.href = url; link.download = `common-ground-${today()}.json`; link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000); toast('Backup exported.');
  };
  modal.querySelector('#import-backup').onclick = () => document.querySelector('#import-file').click();
}

async function exportExcel() {
  const currentSheet = structuredClone(sheet());
  const button = document.querySelector('[data-action="export-excel"]');
  button.disabled = true;
  button.textContent = 'Exporting...';
  try {
    if (!window.ExcelJS) {
      // Load the locally served browser bundle only when an export is requested.
      excelLoading ||= new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = new URL('./vendor/exceljs.js', import.meta.url).href;
        script.onload = resolve;
        script.onerror = () => { script.remove(); reject(new Error('Excel export could not load. Run npm install and try again.')); };
        document.head.append(script);
      });
      await excelLoading;
    }
    const workbook = await buildWorkbook(currentSheet, window.ExcelJS.Workbook);
    const buffer = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `${currentSheet.name.replace(/[^a-zA-Z0-9_-]+/g, '-').slice(0, 80) || 'sprint-board'}.xlsx`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast('Excel exported: full board, initiative details, and sprint schedule.');
  } catch (error) {
    excelLoading = undefined;
    toast(error.message);
  } finally {
    button.disabled = false;
    button.innerHTML = `${icon('download')} Export Excel`;
  }
}

function toggleExpandedBoard() {
  boardExpanded = !boardExpanded;
  document.body.classList.toggle('board-expanded', boardExpanded);
  const button = document.querySelector('[data-action="expand-board"]');
  button.setAttribute('aria-expanded', String(boardExpanded));
  button.title = boardExpanded ? 'Exit expanded view (Esc)' : 'Expand board to fill the screen';
  button.innerHTML = `${icon(boardExpanded ? 'collapse' : 'expand')} ${boardExpanded ? 'Exit expanded view' : 'Expand board'}`;
  button.focus({ preventScroll: true });
}

function toggleSidebar() {
  sidebarCollapsed = !sidebarCollapsed;
  document.body.classList.toggle('sidebar-collapsed', sidebarCollapsed);
  try { localStorage.setItem('common-ground-sidebar-collapsed', String(sidebarCollapsed)); } catch {}
  const button = document.querySelector('[data-action="toggle-sidebar"]');
  button.setAttribute('aria-expanded', String(!sidebarCollapsed));
  button.setAttribute('aria-label', sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar');
  button.title = sidebarCollapsed ? 'Expand sidebar' : 'Collapse sidebar';
}

document.querySelector('#import-file').onchange = async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    if (file.size > 10 * 1024 * 1024) throw new Error('Please choose a backup smaller than 10 MB.');
    const incoming = repairSprintNumbering(validateState(JSON.parse(await file.text())));
    confirmAction('Replace your current sheets?', `This backup contains ${incoming.sheets.length} sheet${incoming.sheets.length === 1 ? '' : 's'}. Export your current work first if you need to keep it.`, () => { state = incoming; search = ''; statusFilter = 'all'; teamFilter = 'all'; save('Backup imported.'); }, 'Replace sheets');
  } catch (error) { toast(`Import failed: ${error.message}`); }
  event.target.value = '';
};

app.addEventListener('click', event => {
  const button = event.target.closest('[data-action]');
  if (!button) return;
  const { action, id: itemId, team, sprint } = button.dataset;
  switch (action) {
    case 'new-sheet': newSheet(); break;
    case 'toggle-sidebar': toggleSidebar(); break;
    case 'delete-sheet': deleteSheet(); break;
    case 'switch-sheet': state.activeSheetId = itemId; search = ''; teamFilter = 'all'; statusFilter = 'all'; save(); break;
    case 'add-item': editItem(null, team, sprint); break;
    case 'edit-item': editItem(itemId); break;
    case 'edit-sprint': editSprint(itemId); break;
    case 'manage-teams': manageTeams(); break;
    case 'schedule': releaseCalendar(); break;
    case 'backup': backup(); break;
    case 'expand-board': toggleExpandedBoard(); break;
    case 'export-excel': exportExcel(); break;
    case 'board': document.querySelector('#board')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); break;
  }
});

app.addEventListener('input', event => {
  if (event.target.id === 'search') { search = event.target.value; renderBoard(); }
});
app.addEventListener('change', event => {
  if (event.target.id === 'status-filter') { statusFilter = event.target.value; renderBoard(); }
  if (event.target.id === 'team-filter') { teamFilter = event.target.value; renderBoard(); }
});
app.addEventListener('dragstart', event => {
  const target = event.target.closest('.initiative[draggable="true"]');
  if (!target) return;
  dragId = target.dataset.id;
  target.dataset.sourceTeam = target.closest('.board-cell').dataset.team;
  event.dataTransfer.setData('text/plain', dragId);
  event.dataTransfer.effectAllowed = 'move';
  target.classList.add('dragging');
});
app.addEventListener('dragover', event => {
  const cell = event.target.closest('.board-cell');
  if (!cell || !dragId) return;
  event.preventDefault(); event.dataTransfer.dropEffect = 'move';
  document.querySelectorAll('.drop-target').forEach(entry => entry.classList.remove('drop-target'));
  cell.classList.add('drop-target');
});
app.addEventListener('drop', event => {
  const cell = event.target.closest('.board-cell');
  if (!cell || !dragId) return;
  event.preventDefault();
  const s = sheet();
  const item = s.initiatives.find(entry => entry.id === dragId);
  const { sprint, team: targetTeam } = cell.dataset;
  const sourceTeam = document.querySelector('.dragging')?.dataset.sourceTeam;
  // Dragging a linked copy changes the shared sprint, not its responsible team.
  const team = sourceTeam && sourceTeam !== item?.teamId ? item.teamId : targetTeam;
  document.querySelectorAll('.drop-target').forEach(entry => entry.classList.remove('drop-target'));
  if (!item || (item.sprintId === sprint && item.teamId === team)) return;
  const later = s.sprints.findIndex(entry => entry.id === sprint) > s.sprints.findIndex(entry => entry.id === item.sprintId);
  if (later) promptMove(item.id, sprint, team);
  else { moveInitiative(s, item.id, sprint, team); save('Initiative moved.'); }
  dragId = null;
});
app.addEventListener('dragend', () => {
  dragId = null;
  document.querySelectorAll('.dragging, .drop-target').forEach(entry => entry.classList.remove('dragging', 'drop-target'));
});
modal.addEventListener('click', event => { if (event.target === modal) { const rect = modal.getBoundingClientRect(); if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) modal.close(); } });
document.addEventListener('keydown', event => {
  if (event.key === 'Escape' && boardExpanded && !modal.open) toggleExpandedBoard();
});
render();
try {
  if (!storageError) localStorage.setItem(STORAGE, JSON.stringify(state));
} catch {
  storageError = 'Browser storage is unavailable. Export your work to keep it.';
  render();
}
