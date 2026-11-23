import test from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { buildWorkbook } from './excel.js';
import { moveInitiative, STATUSES } from './model.js';

function fixture() {
  const sheet = {
    id: 'sheet', name: '=Q2 / [planning]:*?',
    teams: [{ id: 'mobile', name: '=Mobile' }, { id: 'be', name: '+BE' }, { id: 'qa', name: '@QA' }],
    sprints: [
      { id: 's7', number: 7, version: '4.20.0', start: '2026-04-01', end: '2026-04-14', cutoff: '2026-03-30', release: '2026-04-06' },
      { id: 's8', number: 8, version: '+4.21.0', start: '2026-04-15', end: '2026-04-28', cutoff: '2026-04-13', release: '2026-04-20' },
      { id: 's9', number: 9, version: '4.22.0', start: '2026-04-29', end: '2026-05-12', cutoff: '2026-04-27', release: '2026-05-04' },
    ],
    initiatives: [],
  };
  sheet.initiatives = Object.keys(STATUSES).map((status, index) => ({
    id: `item-${index}`, title: ['=1+1', '+SUM(A1:A2)', '-1+2', '@SUM(A1)', 'Dropped work'][index],
    description: '=HYPERLINK("https://example.com", "text")\nDetails',
    owner: '=Owner', status, teamId: 'mobile', sprintId: 's7',
    stakeholders: ['be', 'qa'],
  }));
  sheet.initiatives.push({
    ...sheet.initiatives[0], id: 'later', title: 'A long initiative title '.repeat(8).trim(),
    sprintId: 's8', owner: '', stakeholders: [],
  });
  return sheet;
}

async function roundTrip(sheet) {
  const before = structuredClone(sheet);
  const workbook = await buildWorkbook(sheet, ExcelJS.Workbook);
  assert.ok(workbook instanceof ExcelJS.Workbook);
  const buffer = await workbook.xlsx.writeBuffer();
  assert.equal(buffer[0], 0x50);
  assert.equal(buffer[1], 0x4b);
  const loaded = new ExcelJS.Workbook();
  await loaded.xlsx.load(buffer);
  assert.deepEqual(sheet, before);
  loaded.worksheets.forEach(tab => tab.eachRow(row => row.eachCell(cell => {
    assert.notEqual(cell.type, ExcelJS.ValueType.Formula);
  })));
  return loaded;
}

function noteText(cell) {
  return typeof cell.note === 'string' ? cell.note : cell.note.texts.map(part => part.text).join('');
}

function checkBlocks(board, sheet) {
  let separatorRow = 7;
  const merges = [];
  for (const team of sheet.teams) {
    const columns = sheet.sprints.map(sprint =>
      sheet.initiatives.filter(item => item.teamId === team.id && item.sprintId === sprint.id));
    const height = Math.max(3, ...columns.map(items => items.length));
    const firstRow = separatorRow + 1;
    const lastRow = firstRow + height - 1;
    merges.push(`A${firstRow}:A${lastRow}`);
    assert.equal(board.getRow(separatorRow).height, 8);
    for (let column = 1; column <= sheet.sprints.length + 1; column++) {
      const cell = board.getCell(separatorRow, column);
      assert.equal(cell.value, '');
      assert.equal(cell.fill.pattern, 'solid');
      assert.equal(cell.fill.fgColor.argb, 'FF243746');
    }
    for (let index = 0; index < height; index++) {
      const row = board.getRow(firstRow + index);
      const label = row.getCell(1);
      assert.equal(label.value, team.name);
      assert.equal(label.master.address, `A${firstRow}`);
      assert.equal(label.isMerged, true);
      assert.equal(label.alignment.horizontal, 'center');
      assert.equal(label.alignment.vertical, 'middle');
      assert.equal(label.fill.fgColor.argb, 'FFDDEBF7');
      assert.ok(row.height >= 32);
      columns.forEach((items, column) => {
        const cell = row.getCell(column + 2);
        const item = items[index];
        const destination = item?.spilloverOf
          ? sheet.sprints.find(sprint => sprint.id === item.movedTo).number : null;
        assert.equal(cell.value, item ? item.title + (destination === null ? '' : `\nSpillover -> Sprint ${destination}`) : '');
        assert.equal(cell.alignment.wrapText, true);
        assert.equal(cell.fill.fgColor.argb, 'FFFFFFFF');
        for (const edge of ['top', 'bottom', 'left', 'right']) assert.equal(cell.border[edge].style, 'thin');
        if (!item) {
          assert.equal(cell.note, undefined);
          return;
        }
        const destinationTeam = destination === null ? ''
          : sheet.teams.find(team => team.id === sheet.initiatives.find(entry => entry.id === item.spilloverOf).teamId).name;
        assert.equal(noteText(cell), [
          `Status: ${STATUSES[item.status]}`,
          `Description: ${item.description}`,
          `Spillover destination: ${destination === null ? 'None' : `Sprint ${destination} (${destinationTeam})`}`,
        ].join('\n'));
      });
    }
    separatorRow = lastRow + 1;
  }
  assert.deepEqual(board.model.merges, merges);
  assert.equal(board.rowCount, separatorRow - 1);
}

function checkDetails(workbook, sheet) {
  const details = workbook.getWorksheet('Initiatives');
  assert.equal(details.rowCount, sheet.initiatives.length + 1);
    assert.equal(details.autoFilter, `A1:O${details.rowCount}`);
    assert.equal(details.columnCount, 15);
  assert.deepEqual(details.getRow(1).values.slice(1), [
    'ID', 'Title', 'Description', 'Status', 'Status code', 'Team', 'Sprint',
    'Record type', 'Spillover of', 'Moved to sprint', 'Destination team',
    'Team ID', 'Sprint ID', 'Moved to sprint ID', 'Team dependencies',
  ]);
  assert.deepEqual(details.columns.map(column => column.width),
    [38, 44, 60, 18, 16, 24, 12, 16, 38, 18, 24, 38, 38, 38, 30]);
  for (const [index, item] of sheet.initiatives.entries()) {
    const destination = item.spilloverOf ? sheet.sprints.find(sprint => sprint.id === item.movedTo) : null;
    const destinationTeam = destination
      ? sheet.teams.find(team => team.id === sheet.initiatives.find(entry => entry.id === item.spilloverOf).teamId).name : '';
    assert.deepEqual(Array.from({ length: 15 }, (_, column) => details.getCell(index + 2, column + 1).value), [
      item.id, item.title, item.description, STATUSES[item.status], item.status,
      sheet.teams.find(team => team.id === item.teamId).name,
      sheet.sprints.find(sprint => sprint.id === item.sprintId).number,
      item.spilloverOf ? 'Spillover' : 'Initiative', item.spilloverOf ?? '',
      destination?.number ?? null, destinationTeam,
      item.teamId, item.sprintId, item.movedTo ?? '', (item.dependencies || []).map(id => sheet.teams.find(team => team.id === id).name).join(', '),
    ]);
  }
  const schedule = workbook.getWorksheet('Sprint schedule');
  assert.equal(schedule.rowCount, sheet.sprints.length + 1);
  for (const [index, sprint] of sheet.sprints.entries()) {
    assert.equal(schedule.getCell(index + 2, 1).value, sprint.id);
    assert.equal(schedule.getCell(index + 2, 2).value, sprint.number);
    assert.equal(schedule.getCell(index + 2, 3).value, `v${sprint.version.replace(/^v/i, '')}`);
    for (const [offset, field] of ['start', 'end', 'cutoff', 'release'].entries()) {
      const cell = schedule.getCell(index + 2, offset + 4);
      assert.ok(cell.value instanceof Date);
      assert.equal(cell.value.toISOString(), `${sprint[field]}T00:00:00.000Z`);
      assert.equal(cell.numFmt, 'yyyy-mm-dd');
    }
  }
}

test('serialized XLSX preserves individual initiatives, merged padded team blocks, notes and literal text', async () => {
  const sheet = fixture();
  const workbook = await roundTrip(sheet);
  assert.deepEqual(workbook.worksheets.map(tab => tab.name), ['Board', 'Initiatives', 'Sprint schedule']);
  const board = workbook.getWorksheet('Board');
  assert.equal(board.getCell('A1').value, 'Product');
  assert.deepEqual([2, 3, 4, 5, 6].map(row => board.getCell(row, 1).value),
    ['Mobile Version', 'Sprint Start', 'Sprint End', 'Cutoff Dates (ends)', 'Release date']);
  assert.equal(board.getColumn(1).width, 28);
  for (const [index, sprint] of sheet.sprints.entries()) {
    assert.equal(board.getCell(1, index + 2).value, `${sheet.name} Sprint ${sprint.number}`);
    assert.equal(board.getCell(1, index + 2).isMerged, false);
    assert.equal(board.getCell(2, index + 2).value, `v${sprint.version.replace(/^v/i, '')}`);
    assert.equal(board.getColumn(index + 2).width, 40);
    for (const [offset, field] of ['start', 'end', 'cutoff', 'release'].entries()) {
      const cell = board.getCell(offset + 3, index + 2);
      assert.ok(cell.value instanceof Date);
      assert.equal(cell.value.toISOString(), `${sprint[field]}T00:00:00.000Z`);
      assert.equal(cell.numFmt, 'm/d');
    }
  }
  for (let row = 1; row <= 6; row++) {
    for (let column = 1; column <= 4; column++) {
      const cell = board.getCell(row, column);
      if (column === 1) assert.equal(cell.fill.fgColor.argb, row === 1 ? 'FF243746' : 'FFF0F4F7');
      else assert.equal(cell.fill.fgColor.argb, board.getCell(row === 1 ? 1 : 2, column).fill.fgColor.argb);
      assert.equal(cell.font.color.argb, row === 1 ? 'FFFFFFFF' : 'FF243746');
      assert.equal(cell.alignment.wrapText, true);
    }
  }
  assert.equal(board.getRow(9).height, 32);
  assert.ok(board.getRow(8).height >= Math.ceil(sheet.initiatives.at(-1).title.length / 38) * 16 + 8);
  assert.equal(board.views[0].state, 'frozen');
  assert.equal(board.views[0].xSplit, 1);
  assert.equal(board.views[0].ySplit, 6);
  assert.equal(board.views[0].showGridLines, false);
  assert.equal(board.pageSetup.orientation, 'landscape');
  assert.equal(board.pageSetup.fitToPage, true);
  assert.equal(board.pageSetup.fitToWidth, 1);
  assert.equal(board.pageSetup.fitToHeight, 0);
  assert.equal(board.pageSetup.printTitlesRow, '1:6');
  checkBlocks(board, sheet);
  checkDetails(workbook, sheet);
});

test('optional dependencies export in notes and detail rows without restoring ownership fields', async () => {
  const sheet = fixture();
  sheet.initiatives[0].dependencies = ['be', 'qa'];
  const workbook = await roundTrip(sheet);
  assert.match(noteText(workbook.getWorksheet('Board').getCell('B8')), /Team dependencies: \+BE, @QA/);
  assert.equal(workbook.getWorksheet('Initiatives').getCell('O2').value, '+BE, @QA');
  assert.equal(workbook.getWorksheet('Initiatives').getCell('O3').value, '');
  const board = workbook.getWorksheet('Board');
  assert.equal(board.getCell('B14').value, sheet.initiatives[0].title);
  assert.match(noteText(board.getCell('B14')), /Linked dependency from: =Mobile/);
  assert.equal(board.getCell('B18').value, sheet.initiatives[0].title);
  assert.equal(workbook.getWorksheet('Initiatives').rowCount, sheet.initiatives.length + 1);
});

test('version prefixes avoid numeric text and sprint colors are limited to header rows', async () => {
  const sheet = fixture();
  sheet.sprints[0].version = '7.162';
  sheet.sprints[1].version = 'v7.163';
  sheet.sprints[2].version = 'V7.164';
  const workbook = await roundTrip(sheet);
  const board = workbook.getWorksheet('Board');
  assert.deepEqual(['B2', 'C2', 'D2'].map(address => board.getCell(address).value), ['v7.162', 'v7.163', 'v7.164']);
  const fills = [2, 3, 4].map(column => board.getCell(2, column).fill.fgColor.argb);
  const headers = [2, 3, 4].map(column => board.getCell(1, column).fill.fgColor.argb);
  assert.equal(new Set(fills).size, 3);
  assert.equal(new Set(headers).size, 3);
  for (const column of [2, 3, 4]) {
    for (const row of [3, 4, 5, 6]) assert.equal(board.getCell(row, column).fill.fgColor.argb, fills[column - 2]);
    for (const row of [8, 9, 10]) assert.equal(board.getCell(row, column).fill.fgColor.argb, 'FFFFFFFF');
    assert.equal(workbook.getWorksheet('Sprint schedule').getCell(column, 3).value, board.getCell(2, column).value);
  }
});

test('cross-quarter headers use sprint start calendar quarters, including year boundaries', async () => {
  const sheet = fixture();
  sheet.name = 'Q4 2026';
  sheet.sprints[0].start = '2026-12-23';
  sheet.sprints[0].end = '2027-01-05';
  sheet.sprints[1].start = '2027-01-06';
  sheet.sprints[2].start = '2027-04-01';
  const board = (await roundTrip(sheet)).getWorksheet('Board');
  assert.deepEqual(['B1', 'C1', 'D1'].map(address => board.getCell(address).value),
    ['Q4 2026 Sprint 7', 'Q1 2027 Sprint 8', 'Q2 2027 Sprint 9']);
});

test('exports omit persisted ownership, collaboration and author metadata without changing the input', async () => {
  const sheet = fixture();
  sheet.author = 'Private sheet author';
  sheet.creator = 'Private sheet creator';
  for (const item of sheet.initiatives) {
    item.author = 'Private initiative author';
    item.creator = 'Private initiative creator';
  }
  const workbook = await roundTrip(sheet);
  const exported = JSON.stringify(workbook.model);
  for (const value of ['=Owner', sheet.author, sheet.creator,
    sheet.initiatives[0].author, sheet.initiatives[0].creator]) {
    assert.ok(!exported.includes(value), `Export must not contain ${value}`);
  }
  checkBlocks(workbook.getWorksheet('Board'), sheet);
  checkDetails(workbook, sheet);
});

test('team block height follows the busiest sprint, not just the first column', async () => {
  const sheet = fixture();
  sheet.initiatives = sheet.initiatives.slice(0, 4).map(item => ({ ...item, sprintId: 's9' }));
  const workbook = await roundTrip(sheet);
  const board = workbook.getWorksheet('Board');
  assert.equal(board.getCell('A11').master.address, 'A8');
  assert.equal(board.getRow(12).height, 8);
  assert.equal(board.getCell('D11').value, '@SUM(A1)');
  checkBlocks(board, sheet);
  checkDetails(workbook, sheet);
});

test('spillover snapshots retain individual source cells and notes pointing to the latest destination', async () => {
  const sheet = fixture();
  const item = sheet.initiatives[0];
  moveInitiative(sheet, item.id, 's8', 'mobile', true);
  moveInitiative(sheet, item.id, 's9', 'be', true);
  item.title = 'Current title';
  item.status = 'done';
  const workbook = await roundTrip(sheet);
  const board = workbook.getWorksheet('Board');
  assert.equal(board.getCell('B12').value, '=1+1\nSpillover -> Sprint 9');
  assert.equal(board.getCell('C9').value, '=1+1\nSpillover -> Sprint 9');
  assert.equal(board.getCell('D14').value, 'Current title');
  for (const address of ['B12', 'C9']) {
    assert.ok(noteText(board.getCell(address)).includes('Status: Planned'));
    assert.ok(noteText(board.getCell(address)).includes('Spillover destination: Sprint 9 (+BE)'));
  }
  assert.ok(noteText(board.getCell('D14')).includes('Status: Done'));
  checkBlocks(board, sheet);
  checkDetails(workbook, sheet);
});

test('an empty planning sheet still exports three-row team blocks, separators and its full schedule', async () => {
  const sheet = fixture();
  sheet.initiatives = [];
  const workbook = await roundTrip(sheet);
  const board = workbook.getWorksheet('Board');
  assert.equal(board.rowCount, 6 + sheet.teams.length * 4);
  checkBlocks(board, sheet);
  checkDetails(workbook, sheet);
  assert.equal(workbook.getWorksheet('Initiatives').getCell('B1').value, 'Title');
});
