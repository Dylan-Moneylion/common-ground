import test from 'node:test';
import assert from 'node:assert/strict';
import { addDays, sprintDates, createSheet, nextVersion, moveInitiative, validateState, initialState, removeLegacySamples, repairSprintNumbering } from './model.js';

test('Wednesday sprints have a preceding Monday cutoff and following Monday release', () => {
  assert.deepEqual(sprintDates('2026-04-01'), {
    start: '2026-04-01', end: '2026-04-14', cutoff: '2026-03-30', release: '2026-04-06',
  });
});

test('Monday starts use a strictly preceding Monday cutoff', () => {
  assert.deepEqual(sprintDates('2026-04-06'), {
    start: '2026-04-06', end: '2026-04-19', cutoff: '2026-03-30', release: '2026-04-06',
  });
});

test('date arithmetic handles leap years and year boundaries', () => {
  assert.equal(addDays('2028-02-28', 2), '2028-03-01');
  assert.equal(addDays('2026-12-30', 14), '2027-01-13');
  assert.throws(() => addDays('2026-02-30', 1));
});

test('quarter generation maintains fortnightly starts and increases mobile minor versions', () => {
  const s = createSheet({ name: 'Q2', start: '2026-04-01', firstSprint: 7, count: 7, version: '4.20.0' });
  assert.equal(s.sprints.length, 7);
  assert.equal(s.sprints[1].start, '2026-04-15');
  assert.equal(s.sprints[6].start, '2026-06-24');
  assert.equal(s.sprints[6].number, 13);
  assert.equal(s.sprints[6].version, '4.26.0');
  assert.equal(s.teams.length, 7);
  assert.equal(s.initiatives.length, 0);
  assert.equal(nextVersion('1.99', 1), '1.100');
  assert.throws(() => nextVersion('invalid', 1));
  assert.throws(() => createSheet({ name: '', start: '2026-04-01' }));
  assert.throws(() => createSheet({ name: 'Q2', start: '2026-04-01', count: 0 }));
});

function fixture() {
  const s = createSheet({ name: 'Q2', start: '2026-04-01' });
  s.initiatives.push({ id: crypto.randomUUID(), title: 'Build it', description: '', owner: 'Alex', status: 'progress', teamId: s.teams[0].id, sprintId: s.sprints[0].id, stakeholders: [s.teams[1].id] });
  return s;
}

test('later moves preserve a spillover snapshot and update linked destinations', () => {
  const s = fixture();
  const item = s.initiatives[0];
  moveInitiative(s, item.id, s.sprints[1].id, s.teams[0].id, true);
  assert.equal(s.initiatives.length, 2);
  const marker = s.initiatives[1];
  assert.equal(marker.spilloverOf, item.id);
  assert.equal(marker.sprintId, s.sprints[0].id);
  assert.equal(marker.movedTo, s.sprints[1].id);
  assert.equal(item.status, 'progress');
  item.stakeholders.push(s.teams[2].id);
  assert.equal(marker.stakeholders.length, 1);
  moveInitiative(s, item.id, s.sprints[2].id, s.teams[1].id, true);
  assert.equal(s.initiatives.length, 3);
  assert.equal(marker.movedTo, s.sprints[2].id);
  moveInitiative(s, item.id, s.sprints[0].id, s.teams[0].id);
  assert.equal(s.initiatives.length, 2);
  assert.ok(!s.initiatives.some(entry => entry.spilloverOf && entry.sprintId === item.sprintId));
});

test('moving without spillover or within a sprint does not create markers', () => {
  const s = fixture();
  const item = s.initiatives[0];
  moveInitiative(s, item.id, s.sprints[1].id, s.teams[0].id);
  moveInitiative(s, item.id, s.sprints[1].id, s.teams[1].id, true);
  assert.equal(s.initiatives.length, 1);
  assert.equal(item.teamId, s.teams[1].id);
  assert.throws(() => moveInitiative(s, item.id, 'missing', s.teams[0].id));
});

test('backup validation accepts round trips and rejects corrupt references', () => {
  const state = initialState(new Date('2026-05-10T12:00:00Z'));
  state.sheets[0].initiatives = fixture().initiatives.map(item => ({ ...item, teamId: state.sheets[0].teams[0].id, sprintId: state.sheets[0].sprints[0].id, stakeholders: [] }));
  assert.equal(state.sheets[0].name, 'Q2 2026');
  assert.deepEqual(validateState(JSON.parse(JSON.stringify(state))), state);
  state.sheets[0].initiatives[0].teamId = 'missing';
  assert.throws(() => validateState(state));
  assert.throws(() => validateState({ schema: 1, sheets: [] }));
  assert.deepEqual(validateState({ schema: 1, activeSheetId: null, sheets: [] }), { schema: 1, activeSheetId: null, sheets: [] });
});

test('backup validation rejects duplicate IDs and unsafe status values', () => {
  const s = fixture();
  const state = { schema: 1, activeSheetId: s.id, sheets: [s] };
  state.sheets[0].initiatives[0].status = '__proto__';
  assert.throws(() => validateState(state));
  state.sheets[0].initiatives[0].status = 'todo';
  state.sheets[0].teams[0].id = state.sheets[0].id;
  assert.throws(() => validateState(state));
});

test('every fresh quarter starts with no initiatives and Sprint 1', () => {
  for (const month of ['01', '04', '07', '10']) {
    const state = initialState(new Date(`2026-${month}-15T12:00:00Z`));
    assert.equal(state.sheets[0].initiatives.length, 0);
    assert.equal(state.sheets[0].sprints[0].number, 1);
    assert.equal(validateState(state), state);
  }
});

test('quarterly sheet numbering stays sequential across calendar quarters or years', () => {
  for (const start of ['2026-09-30', '2026-03-18', '2026-12-23']) {
    const s = createSheet({ name: 'Plan', start, count: 4 });
    assert.deepEqual(s.sprints.map(sprint => sprint.number), [1, 2, 3, 4]);
  }
});

test('repairs saved calendar-reset duplicates without changing sprint IDs or work', () => {
  const s = createSheet({ name: 'Q4 2026', start: '2026-09-30', count: 4 });
  const state = { schema: 1, activeSheetId: s.id, sheets: [s] };
  s.sprints.forEach((sprint, index) => { sprint.number = [1, 1, 2, 3][index]; });
  const ids = s.sprints.map(sprint => sprint.id);
  const work = s.initiatives;
  repairSprintNumbering(state);
  assert.deepEqual(s.sprints.map(sprint => sprint.number), [1, 2, 3, 4]);
  assert.deepEqual(s.sprints.map(sprint => sprint.id), ids);
  assert.equal(s.initiatives, work);
  s.sprints[1].number = 10;
  repairSprintNumbering(state);
  assert.equal(s.sprints[1].number, 10);
});

test('legacy sample cleanup preserves edited and genuine initiatives', () => {
  const state = initialState(new Date('2026-10-15T12:00:00Z'));
  const s = state.sheets[0];
  const sample = { id: idForTest(), title: 'A smoother first five minutes', status: 'progress', owner: 'Alex Chen', description: 'Simplify onboarding and reduce the steps to first value.', teamId: s.teams[0].id, sprintId: s.sprints[0].id, stakeholders: [s.teams[3].id, s.teams[5].id] };
  const edited = { ...sample, id: idForTest(), description: 'Our updated plan' };
  const genuine = { ...sample, id: idForTest(), title: 'My own initiative' };
  s.initiatives.push(sample, edited, genuine);
  s.sprints.forEach((sprint, index) => { sprint.number = 19 + index; });
  removeLegacySamples(state);
  assert.deepEqual(s.initiatives.map(item => item.id), [edited.id, genuine.id]);
  assert.equal(s.sprints[0].number, 1);
  assert.equal(validateState(state), state);
});

const idForTest = () => crypto.randomUUID();
