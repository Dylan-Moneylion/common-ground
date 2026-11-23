import { STATUSES, belongsToTeam } from './model.js';

// Inject the constructor so this module needs no Node-only or browser-only imports.
export async function buildWorkbook(sheet, Workbook) {
  const workbook = new Workbook();
  const teams = new Map(sheet.teams.map(team => [team.id, team.name]));
  const sprints = new Map(sheet.sprints.map(sprint => [sprint.id, sprint]));
  const items = new Map(sheet.initiatives.map(item => [item.id, item]));
  const date = value => new Date(`${value}T00:00:00Z`);
  const version = value => `v${String(value).trim().replace(/^v/i, '')}`;
  const sprintColors = sheet.sprints.map((_, index) => {
    // Golden-angle spacing keeps neighboring columns visually distinct.
    const channels = [0, 2, 4].map(phase => Math.round(224 + 23 * Math.cos(index * 2.399963 + phase * Math.PI / 3)));
    const argb = scale => 'FF' + channels.map(channel => Math.round(channel * scale).toString(16).padStart(2, '0')).join('').toUpperCase();
    return { body: argb(1), header: argb(.32) };
  });
  const destinationTeam = item => teams.get(items.get(item.spilloverOf)?.teamId) ?? '';
  const dependencies = item => (item.dependencies || []).map(teamId => teams.get(teamId)).join(', ');

  const board = workbook.addWorksheet('Board');
  const period = sprint => {
    const start = date(sprint.start);
    return `Q${Math.floor(start.getUTCMonth() / 3) + 1} ${start.getUTCFullYear()}`;
  };
  const crossesQuarter = new Set(sheet.sprints.map(period)).size > 1;
  board.addRow(['Product', ...sheet.sprints.map(sprint =>
    `${crossesQuarter ? period(sprint) : sheet.name} Sprint ${sprint.number}`)]);
  board.addRow(['Mobile Version', ...sheet.sprints.map(sprint => version(sprint.version))]);
  for (const [label, field] of [
    ['Sprint Start', 'start'], ['Sprint End', 'end'],
    ['Cutoff Dates (ends)', 'cutoff'], ['Release date', 'release'],
  ]) {
    const row = board.addRow([label, ...sheet.sprints.map(sprint => date(sprint[field]))]);
    for (let column = 2; column <= sheet.sprints.length + 1; column++) {
      row.getCell(column).numFmt = 'm/d';
    }
  }
  const separators = new Set();
  for (const team of sheet.teams) {
    const separator = board.addRow(Array(sheet.sprints.length + 1).fill(''));
    separators.add(separator.number);
    const columns = sheet.sprints.map(sprint =>
      sheet.initiatives.filter(item => belongsToTeam(item, team.id) && item.sprintId === sprint.id));
    const height = Math.max(3, ...columns.map(items => items.length));
    const firstRow = board.rowCount + 1;
    for (let index = 0; index < height; index++) {
      const row = board.addRow([index === 0 ? String(team.name) : '', ...columns.map(items => {
        const item = items[index];
        return item ? String(item.title) + (item.spilloverOf
          ? `\nSpillover -> Sprint ${sprints.get(item.movedTo).number}` : '') : '';
      })]);
      columns.forEach((items, column) => {
        const item = items[index];
        if (!item) return;
        row.getCell(column + 2).note = [
          `Status: ${STATUSES[item.status]}`,
          ...(item.teamId !== team.id ? [`Linked dependency from: ${teams.get(item.teamId)}`] : []),
          ...(item.dependencies?.length ? [`Team dependencies: ${dependencies(item)}`] : []),
          `Description: ${item.description}`,
          `Spillover destination: ${item.spilloverOf
            ? `Sprint ${sprints.get(item.movedTo).number} (${destinationTeam(item)})` : 'None'}`,
        ].join('\n');
      });
    }
    board.mergeCells(firstRow, 1, firstRow + height - 1, 1);
  }
  board.getColumn(1).width = 28;
  for (let column = 2; column <= Math.max(2, sheet.sprints.length + 1); column++) {
    board.getColumn(column).width = 40;
  }
  board.views = [{ state: 'frozen', xSplit: 1, ySplit: 6, showGridLines: false }];
  board.pageSetup = {
    orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    printTitlesRow: '1:6',
  };

  const initiatives = workbook.addWorksheet('Initiatives');
  initiatives.addRow([
    'ID', 'Title', 'Description', 'Status', 'Status code', 'Team', 'Sprint',
    'Record type', 'Spillover of', 'Moved to sprint', 'Destination team',
    'Team ID', 'Sprint ID', 'Moved to sprint ID', 'Team dependencies',
  ]);
  for (const item of sheet.initiatives) {
    initiatives.addRow([
      String(item.id), String(item.title), String(item.description), STATUSES[item.status],
      String(item.status), String(teams.get(item.teamId)),
      sprints.get(item.sprintId).number,
      item.spilloverOf ? 'Spillover' : 'Initiative', String(item.spilloverOf ?? ''),
      item.spilloverOf ? sprints.get(item.movedTo).number : null,
      item.spilloverOf ? destinationTeam(item) : '',
      String(item.teamId), String(item.sprintId), String(item.movedTo ?? ''), dependencies(item),
    ]);
  }
  const widths = [38, 44, 60, 18, 16, 24, 12, 16, 38, 18, 24, 38, 38, 38, 30];
  widths.forEach((width, index) => { initiatives.getColumn(index + 1).width = width; });
  initiatives.views = [{ state: 'frozen', ySplit: 1 }];
  initiatives.autoFilter = { from: 'A1', to: `O${initiatives.rowCount}` };

  const schedule = workbook.addWorksheet('Sprint schedule');
  schedule.addRow(['Sprint ID', 'Sprint number', 'Mobile version', 'Sprint start', 'Sprint end', 'Mobile cutoff', 'Mobile release']);
  for (const sprint of sheet.sprints) {
    const row = schedule.addRow([
      String(sprint.id), sprint.number, version(sprint.version),
      ...['start', 'end', 'cutoff', 'release'].map(field => date(sprint[field])),
    ]);
    for (let column = 4; column <= 7; column++) row.getCell(column).numFmt = 'yyyy-mm-dd';
  }
  [38, 18, 20, 18, 18, 18, 18].forEach((width, index) => { schedule.getColumn(index + 1).width = width; });
  schedule.views = [{ state: 'frozen', ySplit: 1 }];
  schedule.autoFilter = { from: 'A1', to: `G${schedule.rowCount}` };

  for (const tab of workbook.worksheets) {
    tab.eachRow((row, rowNumber) => {
      const isBoard = tab === board;
      const header = rowNumber === 1;
      const metadata = isBoard && rowNumber > 1 && rowNumber <= 6;
      const separator = isBoard && separators.has(rowNumber);
      let lines = 1;
      row.eachCell({ includeEmpty: true }, (cell, column) => {
        const teamLabel = isBoard && rowNumber > 6 && !separator && column === 1;
        cell.font = { name: 'Calibri', size: 11, bold: header || (isBoard && column === 1), color: { argb: header || separator ? 'FFFFFFFF' : 'FF243746' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: header || separator ? 'FF243746'
          : metadata ? 'FFF0F4F7' : teamLabel ? 'FFDDEBF7' : isBoard ? 'FFFFFFFF' : rowNumber % 2 ? 'FFF0F4F7' : 'FFFFFFFF' } };
        if (isBoard && column > 1 && rowNumber <= 6) {
          cell.fill.fgColor.argb = sprintColors[column - 2][header ? 'header' : 'body'];
        }
        cell.alignment = teamLabel
          ? { vertical: 'middle', horizontal: 'center', wrapText: true }
          : { vertical: 'top', wrapText: true };
        if (isBoard && !separator) {
          const edge = { style: 'thin', color: { argb: 'FFD3DCE4' } };
          cell.border = { top: edge, bottom: edge, left: edge, right: edge };
        }
        if (typeof cell.value === 'string') {
          const width = tab.getColumn(column).width - 2;
          lines = Math.max(lines, cell.value.split('\n').reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / width)), 0));
        }
      });
      row.height = separator ? 8 : Math.min(409, Math.max(isBoard ? 32 : header ? 30 : 36, lines * 16 + 8));
    });
  }
  return workbook;
}
