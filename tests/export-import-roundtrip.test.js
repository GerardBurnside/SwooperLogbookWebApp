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
        recentJumpsDays: 3,
        standardRedThreshold: 160,
        standardOrangeThreshold: 140,
        hybridRedThreshold: 80,
        hybridOrangeThreshold: 60
    };

    // Disable UI-heavy side effects for isolated unit testing.
    logbook.initializeCanopyLinesetJumpCounts = () => {};
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
        timestamp: '2026-03-06T00:00:00.000Z'
    };

    logbook.harnesses = [{ id: 'h_test_1', name: 'Harness Test' }];
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

    // Jump and canopy should be restored.
    assert.equal(logbook.jumps.length, 1);
    assert.equal(logbook.jumps[0].id, jump.id);
    assert.equal(logbook.canopies.length, 1);
    assert.equal(logbook.canopies[0].id, canopy.id);
    assert.equal(logbook.canopies[0].linesets.length, 1);

    // Import flow should reset input value and report success.
    assert.equal(event.target.value, '');
    assert.ok(messages.some(m => m.message === 'Data imported successfully!' && m.type === 'success'));

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
