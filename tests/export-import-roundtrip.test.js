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
            createElement: () => ({
                style: {},
                click: () => {}
            }),
            body: {
                appendChild: () => {},
                removeChild: () => {}
            },
            querySelector: () => null,
            getElementById: () => ({
                addEventListener: () => {},
                style: {},
                classList: { add: () => {}, remove: () => {} },
                value: '',
                textContent: ''
            })
        },
        URL: {
            createObjectURL: () => 'blob:test',
            revokeObjectURL: () => {}
        },
        Blob: class Blob {
            constructor(parts, opts) {
                this.parts = parts;
                this.opts = opts;
            }
        }
    };

    sandbox.window = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(`${appJs}\nthis.__SkydivingLogbook = SkydivingLogbook;`, sandbox);

    return {
        SkydivingLogbook: sandbox.__SkydivingLogbook,
        localStorage
    };
}

function createHeadlessLogbook() {
    const { SkydivingLogbook, localStorage } = loadSkydivingLogbookClass();
    const logbook = Object.create(SkydivingLogbook.prototype);

    logbook.jumps = [];
    logbook.harnesses = [];
    logbook.canopies = [];
    logbook.locations = [];
    logbook.settings = {
        startingJumpNumber: 1,
        resequenceJumpsFromStartingNumber: true,
        recentJumpsDays: 3,
        recentJumpsGroupByMonth: false,
        standardRedThreshold: 160,
        standardOrangeThreshold: 140,
        hybridRedThreshold: 80,
        hybridOrangeThreshold: 60
    };

    // Disable UI-heavy side effects for isolated unit testing.
    logbook.initializeCanopyLinesetJumpCounts = () => {};
    logbook.ensureJumpIds = () => {};
    logbook.saveToLocalStorage = () => {};
    logbook.saveComponentsToLocalStorage = () => {};
    logbook.markEquipmentModified = () => {};
    logbook.updateEquipmentOptions = () => {};
    logbook.renderJumpsList = () => {};
    logbook.updateStats = () => {};
    logbook.renderEquipmentView = () => {};
    logbook.renderStats = () => {};

    return { logbook, localStorage, SkydivingLogbook };
}

test('export/import round trip restores jump and canopy data', async () => {
    const { logbook, localStorage } = createHeadlessLogbook();

    const canopy = {
        id: 'c_test_1',
        name: 'Canopy Test',
        harnessId: 'h_test_1',
        linesets: [{ number: 1, hybrid: false, previousJumps: 0, jumpCount: 1, archived: false }]
    };
    const jump = {
        id: 123,
        jumpNumber: 1,
        date: '2026-03-06',
        location: 'Test DZ',
        equipment: canopy.id,
        linesetNumber: 1,
        notes: 'roundtrip test',
        timestamp: '2026-03-06T00:00:00.000Z',
        harnessId: 'h_test_1'
    };

    logbook.harnesses = [{ id: 'h_test_1', name: 'Harness Test', previousJumps: 42 }];
    logbook.canopies = [canopy];
    logbook.locations = [{ id: 'loc_test_1', name: 'Test DZ', lat: null, lng: null }];
    logbook.jumps = [jump];

    // Export data.
    const exportedPayload = logbook.buildExportPayload();
    const exportedJson = JSON.stringify(exportedPayload);
    assert.equal(exportedPayload.version, 2);
    assert.equal(exportedPayload.data.jumps.length, 1);
    assert.equal(exportedPayload.data.canopies.length, 1);
    assert.equal(exportedPayload.data.canopies[0].linesets.length, 1);
    assert.equal(exportedPayload.data.jumps[0].harnessId, 'h_test_1');
    assert.equal(exportedPayload.data.canopies[0].harnessId, 'h_test_1');

    // Delete data in app state.
    logbook.jumps = [];
    logbook.harnesses = [];
    logbook.canopies = [];
    logbook.locations = [];

    assert.equal(logbook.jumps.length, 0);
    assert.equal(logbook.canopies.length, 0);

    // Import the previously exported JSON.
    const messages = [];
    logbook.showMessage = (message, type) => {
        messages.push({ message, type });
    };

    const event = {
        target: {
            files: [
                {
                    text: async () => exportedJson
                }
            ],
            value: 'backup.json'
        }
    };

    await logbook.importData(event);

    const pending = logbook._pendingImportPayload;
    assert.ok(pending, 'import should stage JSON payload for merge/replace choice');
    logbook.applyImportMerge(pending);
    logbook.closeImportChoiceModal();

    // Jump and canopy should be restored.
    assert.equal(logbook.jumps.length, 1);
    assert.equal(logbook.jumps[0].id, jump.id);
    assert.equal(logbook.canopies.length, 1);
    assert.equal(logbook.canopies[0].id, canopy.id);
    assert.equal(logbook.canopies[0].linesets.length, 1);
    assert.equal(logbook.canopies[0].harnessId, 'h_test_1');
    assert.equal(logbook.jumps[0].harnessId, 'h_test_1');
    assert.equal(logbook.harnesses[0].previousJumps, 42);

    // Import flow should reset input value and report success.
    assert.equal(event.target.value, '');
    assert.ok(messages.some(m => m.message === 'Data merged successfully!' && m.type === 'success'));

    // Explicit jump #s in JSON turn off chronological renumbering from starting jump.
    assert.equal(logbook.settings.resequenceJumpsFromStartingNumber, false);

    // Settings are persisted during import.
    assert.ok(localStorage.getItem('skydiving-settings'));
});

test('import with invalid JSON keeps existing data unchanged', async () => {
    const { logbook } = createHeadlessLogbook();

    const existingCanopy = {
        id: 'c_existing_1',
        name: 'Canopy Existing',
        linesets: [{ number: 1, hybrid: false, previousJumps: 0, jumpCount: 1, archived: false }]
    };
    const existingJump = {
        id: 456,
        jumpNumber: 7,
        date: '2026-03-05',
        location: 'Existing DZ',
        equipment: existingCanopy.id,
        linesetNumber: 1,
        notes: 'existing data',
        timestamp: '2026-03-05T00:00:00.000Z'
    };

    logbook.jumps = [existingJump];
    logbook.harnesses = [{ id: 'h_existing_1', name: 'Harness Existing' }];
    logbook.canopies = [existingCanopy];
    logbook.locations = [{ id: 'loc_existing_1', name: 'Existing DZ', lat: null, lng: null }];

    const messages = [];
    logbook.showMessage = (message, type) => {
        messages.push({ message, type });
    };

    const event = {
        target: {
            files: [
                {
                    text: async () => '{ not valid json }'
                }
            ],
            value: 'broken-backup.json'
        }
    };

    await logbook.importData(event);

    // Existing data must remain unchanged when import payload is invalid.
    assert.equal(logbook.jumps.length, 1);
    assert.equal(logbook.jumps[0].id, existingJump.id);
    assert.equal(logbook.canopies.length, 1);
    assert.equal(logbook.canopies[0].id, existingCanopy.id);

    // Import flow should still reset input and show parse error feedback.
    assert.equal(event.target.value, '');
    assert.ok(messages.some(m => m.message === 'Import failed: invalid JSON file' && m.type === 'error'));
});

test('JSON replace import with explicit jump numbers disables resequencing; renumberJumps keeps stored #s', () => {
    const { logbook } = createHeadlessLogbook();
    logbook.showMessage = () => {};
    const payload = {
        jumps: [
            { id: 10, jumpNumber: 5020, date: '2020-01-01', location: 'A', equipment: '', linesetNumber: 1, notes: '', timestamp: '2020-01-01T00:00:00.000Z', jumpId: 'j5020' },
            { id: 11, jumpNumber: 5021, date: '2020-02-01', location: 'B', equipment: '', linesetNumber: 1, notes: '', timestamp: '2020-02-01T00:00:00.000Z', jumpId: 'j5021' }
        ],
        harnesses: [],
        canopies: [],
        locations: [],
        settings: { startingJumpNumber: 99, resequenceJumpsFromStartingNumber: true }
    };
    logbook.applyImportReplace(payload);
    assert.equal(logbook.settings.resequenceJumpsFromStartingNumber, false);
    assert.equal(logbook.settings.startingJumpNumber, 99);
    logbook.renumberJumps();
    assert.deepEqual(
        logbook.jumps.map(j => j.jumpNumber).sort((a, b) => a - b),
        [5020, 5021]
    );
});

test('renumberJumps assigns from starting jump when resequencing is enabled', () => {
    const { logbook } = createHeadlessLogbook();
    logbook.settings.resequenceJumpsFromStartingNumber = true;
    logbook.settings.startingJumpNumber = 10;
    logbook.jumps = [
        { id: 1, jumpNumber: 99, date: '2021-02-01', location: 'X', equipment: '', linesetNumber: 1, notes: '', timestamp: '2021-02-01T12:00:00.000Z', jumpId: 'x1' },
        { id: 2, jumpNumber: 1, date: '2021-01-01', location: 'Y', equipment: '', linesetNumber: 1, notes: '', timestamp: '2021-01-01T12:00:00.000Z', jumpId: 'x2' }
    ];
    logbook.renumberJumps();
    assert.deepEqual(logbook.jumps.map(j => j.jumpNumber), [10, 11]);
});
