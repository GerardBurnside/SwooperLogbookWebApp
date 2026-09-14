const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createLocalStorageStub() {
    const store = new Map();
    return {
        getItem(key) {
            return store.has(key) ? store.get(key) : null;
        },
        setItem(key, value) {
            store.set(key, String(value));
        },
        removeItem(key) {
            store.delete(key);
        },
        clear() {
            store.clear();
        }
    };
}

function loadSkydivingLogbookClass() {
    const appJsPath = path.join(__dirname, '..', 'js', 'app.js');
    const appJs = fs.readFileSync(appJsPath, 'utf8');
    const localStorage = createLocalStorageStub();
    const sandbox = {
        console,
        setTimeout,
        clearTimeout,
        setInterval,
        clearInterval,
        localStorage,
        confirm: () => true,
        navigator: { onLine: true },
        document: {
            addEventListener: () => {},
            createElement: () => ({ style: {}, click: () => {} }),
            body: { appendChild: () => {}, removeChild: () => {} },
            querySelector: () => null,
            getElementById: () => ({
                addEventListener: () => {},
                style: {},
                classList: { add: () => {}, remove: () => {} },
                value: '',
                textContent: ''
            })
        }
    };
    sandbox.window = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(`${appJs}\nthis.__SkydivingLogbook = SkydivingLogbook;`, sandbox);
    return sandbox.__SkydivingLogbook;
}

function createLogbook(overrides = {}) {
    const SkydivingLogbook = loadSkydivingLogbookClass();
    const logbook = Object.create(SkydivingLogbook.prototype);
    logbook.jumps = [];
    logbook.harnesses = [];
    logbook.canopies = [];
    logbook.locations = [];
    logbook.settings = {
        standardRedThreshold: 160,
        standardOrangeThreshold: 140,
        hybridRedThreshold: 80,
        hybridOrangeThreshold: 60
    };
    Object.assign(logbook, overrides);
    return logbook;
}

function makeCanopy(partial = {}) {
    return {
        id: 'c1',
        name: 'Petra 62',
        archived: false,
        previousJumps: 0,
        linesets: [{ number: 1, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false }],
        ...partial
    };
}

test('canopy previous jumps appear in totals even with no logged jumps', () => {
    const logbook = createLogbook({
        canopies: [makeCanopy({ previousJumps: 120 })]
    });
    const totals = logbook._buildCanopyTotalsStats();
    assert.equal(totals.length, 1);
    assert.equal(totals[0].name, 'Petra 62');
    assert.equal(totals[0].logged, 0);
    assert.equal(totals[0].preApp, 120);
    assert.equal(totals[0].count, 120);
    assert.equal(totals[0].archived, false);
    assert.equal(logbook._hasEquipmentPreviousJumps(), true);
});

test('canopy totals add canopy previous jumps, lineset previous jumps, and logged jumps', () => {
    const canopy = makeCanopy({
        previousJumps: 20,
        linesets: [
            { number: 1, hybrid: false, previousJumps: 5, jumpCount: 2, archived: false },
            { number: 2, hybrid: false, previousJumps: 3, jumpCount: 0, archived: false }
        ]
    });
    const logbook = createLogbook({
        canopies: [canopy],
        jumps: [
            { equipment: 'c1', linesetNumber: 1 },
            { equipment: 'c1', linesetNumber: 1 }
        ]
    });
    const totals = logbook._buildCanopyTotalsStats();
    assert.equal(totals.length, 1);
    assert.equal(totals[0].logged, 2);
    assert.equal(totals[0].preApp, 28);
    assert.equal(totals[0].count, 30);
});

test('archived canopy with previous jumps is included but marked archived', () => {
    const logbook = createLogbook({
        canopies: [makeCanopy({ previousJumps: 40, archived: true })]
    });
    const totals = logbook._buildCanopyTotalsStats();
    assert.equal(totals.length, 1);
    assert.equal(totals[0].archived, true);
    assert.equal(totals[0].preApp, 40);
});

test('canopy with zero previous jumps and no logged jumps is omitted from totals', () => {
    const logbook = createLogbook({
        canopies: [makeCanopy({ previousJumps: 0 })]
    });
    assert.equal(logbook._buildCanopyTotalsStats().length, 0);
    assert.equal(logbook._hasEquipmentPreviousJumps(), false);
});

test('harness previous jumps appear in stats even with no logged jumps', () => {
    const logbook = createLogbook({
        harnesses: [{ id: 'h1', name: 'Mutant', previousJumps: 75, archived: false }]
    });
    const harnessStats = logbook._buildHarnessStats();
    assert.equal(harnessStats.length, 1);
    assert.equal(harnessStats[0].logged, 0);
    assert.equal(harnessStats[0].preApp, 75);
    assert.equal(harnessStats[0].count, 75);
    const active = harnessStats.filter(s => !s.archived && (s.logged > 0 || s.preApp !== 0));
    assert.equal(active.length, 1);
    assert.equal(logbook._hasEquipmentPreviousJumps(), true);
});

test('archived harness with previous jumps is not in the active stats list', () => {
    const logbook = createLogbook({
        harnesses: [{ id: 'h1', name: 'Old Vector', previousJumps: 200, archived: true }]
    });
    const harnessStats = logbook._buildHarnessStats();
    const active = harnessStats.filter(s => !s.archived && (s.logged > 0 || s.preApp !== 0));
    assert.equal(active.length, 0);
    assert.equal(harnessStats[0].archived, true);
    assert.equal(logbook._hasEquipmentPreviousJumps(), true);
});
