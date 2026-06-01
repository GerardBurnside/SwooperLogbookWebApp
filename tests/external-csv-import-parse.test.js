const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

/** Loads js/external-csv-import.js (no app.js). */
function loadExternalCsvImport() {
    const extPath = path.join(__dirname, '..', 'js', 'external-csv-import.js');
    const code = fs.readFileSync(extPath, 'utf8');
    const sandbox = { console };
    sandbox.globalThis = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(code, sandbox);
    return sandbox.ExternalCsvImport;
}

const E = loadExternalCsvImport();

test('isExternalSkydivingLogbookCsv detects sample export header', () => {
    const header = 'Saut #,Date,Zone de saut,Aéronef,Équipement,Type de saut,Libération,Notes\n';
    assert.equal(E.isExternalSkydivingLogbookCsv(header), true);
    assert.equal(E.isExternalSkydivingLogbookCsv('{"jumps":[]}'), false);
});

test('parseExternalEquipmentString: harness + canopy + lineset', () => {
    const p = E.parseExternalEquipmentString('Mutant+Petra65-lineset1');
    assert.equal(p.harness, 'Mutant');
    assert.equal(p.canopySegment, 'Petra65');
    assert.equal(p.sizeDigits, '65');
    assert.equal(p.lineset, 1);
    assert.match(p.suggestedCanopyName, /65/);
});

test('parseExternalEquipmentString: harness + PI-71 without lineset suffix', () => {
    const p = E.parseExternalEquipmentString('Odyssey+PI-71');
    assert.equal(p.harness, 'Odyssey');
    assert.equal(p.sizeDigits, '71');
    assert.equal(p.lineset, 1);
});

test('parseExternalEquipmentString: -lineN suffix (not the word lineset)', () => {
    const p = E.parseExternalEquipmentString('Odyssey+Petra68-line1');
    assert.equal(p.harness, 'Odyssey');
    assert.equal(p.canopySegment, 'Petra68');
    assert.equal(p.sizeDigits, '68');
    assert.equal(p.lineset, 1);
    assert.equal(p.suggestedCanopyName, 'Petra 68');
});

test('collectExternalCsvCanopyCandidates: Petra68-line1 maps to existing Petra68', () => {
    const logbook = {
        canopies: [{ id: 'canopy_abc', name: 'Petra68', archived: false }]
    };
    const p = E.parseExternalEquipmentString('Odyssey+Petra68-line1');
    const c = E.collectExternalCsvCanopyCandidates(logbook, p);
    assert.equal(c.length, 1);
    assert.equal(c[0].id, 'canopy_abc');
});

test('collectExternalCsvCanopyCandidates: match by id substring when name differs', () => {
    const logbook = {
        canopies: [{ id: 'rig_main_petra68', name: 'Main sport', archived: false }]
    };
    const p = E.parseExternalEquipmentString('Odyssey+Petra68-line1');
    const c = E.collectExternalCsvCanopyCandidates(logbook, p);
    assert.equal(c.length, 1);
    assert.equal(c[0].name, 'Main sport');
});

test('extractTrailingTwoDigitSize matches canopy names', () => {
    assert.equal(E.extractTrailingTwoDigitSize('Petra 65'), '65');
    assert.equal(E.extractTrailingTwoDigitSize('PI-71'), '71');
    assert.equal(E.extractTrailingTwoDigitSize('Solo'), '');
});

test('parseExternalSkydivingCsv reads sample file rows', () => {
    const csvPath = path.join(__dirname, '..', 'skydiving_logbook.csv');
    const text = fs.readFileSync(csvPath, 'utf8');
    const { ok, rows } = E.parseExternalSkydivingCsv(text);
    assert.equal(ok, true);
    assert.ok(rows.length >= 10);
    assert.equal(rows[0].jumpNumber, 4378);
    assert.equal(rows[0].date, '2020-08-16');
    assert.equal(rows[0].equipmentRaw, 'Mutant+Petra65-lineset1');
});

test('Libération Oui adds cutaway note', () => {
    const csv = [
        'Saut #,Date,Zone de saut,Aéronef,Équipement,Libération,Notes',
        '1,2020-01-01,,,X,Oui,hello'
    ].join('\n');
    const { rows } = E.parseExternalSkydivingCsv(csv);
    assert.equal(rows[0].notesComposed, 'cutaway\nhello');
});

test('mergeExternalCsvRowIntoJump appends notes and fills blanks only', () => {
    const logbook = {
        jumps: [],
        locations: [],
        geocodeLocation: () => {}
    };
    logbook.jumps = [{
        id: 1,
        jumpNumber: 5,
        date: '',
        location: '',
        equipment: '',
        linesetNumber: 1,
        notes: 'local',
        timestamp: '2020-01-01T12:00:00.000Z'
    }];
    const row = {
        jumpNumber: 5,
        date: '2020-06-01',
        location: 'DZ-A',
        equipmentRaw: '',
        notesComposed: 'from csv'
    };
    E.mergeExternalCsvRowIntoJump(logbook, logbook.jumps[0], row, new Map());
    assert.equal(logbook.jumps[0].date, '2020-06-01');
    assert.equal(logbook.jumps[0].location, 'DZ-A');
    assert.match(logbook.jumps[0].notes, /local/);
    assert.match(logbook.jumps[0].notes, /from csv/);
});

test('computeJumpNumberGapInfo: missing numbers strictly below csvMin', () => {
    const logbook = {
        jumps: [{ jumpNumber: 5 }, { jumpNumber: 50 }],
        settings: { startingJumpNumber: 1 }
    };
    const gap = E.computeJumpNumberGapInfo(logbook, 100);
    assert.equal(gap.gapStart, 51);
    assert.equal(gap.gapEnd, 99);
    assert.equal(gap.gapCount, 49);
});

test('computeJumpNumberGapInfo: no gap when csvMin follows last lower jump', () => {
    const logbook = {
        jumps: [{ jumpNumber: 99 }],
        settings: { startingJumpNumber: 1 }
    };
    const gap = E.computeJumpNumberGapInfo(logbook, 100);
    assert.equal(gap, null);
});
