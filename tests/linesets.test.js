const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('path');
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

function loadSkydivingLogbookClass(equipmentList) {
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
        navigator: { onLine: false },
        document: {
            addEventListener: () => {},
            createElement: () => ({ style: {}, click: () => {} }),
            body: { appendChild: () => {}, removeChild: () => {} },
            querySelector: () => null,
            getElementById: (id) => {
                if (id === 'equipmentList') return equipmentList;
                return {
                    addEventListener: () => {},
                    style: {},
                    classList: { add: () => {}, remove: () => {} },
                    value: '',
                    textContent: ''
                };
            }
        }
    };
    sandbox.window = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(`${appJs}\nthis.__SkydivingLogbook = SkydivingLogbook;`, sandbox);
    return sandbox.__SkydivingLogbook;
}

function createLogbook(overrides = {}, equipmentList = { innerHTML: '' }) {
    const SkydivingLogbook = loadSkydivingLogbookClass(equipmentList);
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
    logbook.showOlderLinesetsByCanopyId = Object.create(null);
    logbook.showMessage = () => {};
    logbook.saveComponentsToLocalStorage = () => {};
    logbook.updateEquipmentOptions = () => {};
    logbook.renderEquipmentView = () => {};
    logbook.updateLinesetHint = () => {};
    logbook._initCanopyDragAndDrop = () => {};
    Object.assign(logbook, overrides);
    return logbook;
}

function ls(number, extra = {}) {
    return { number, hybrid: false, previousJumps: 0, jumpCount: 0, archived: false, ...extra };
}

test('_getLatestAndOlderLinesets: single lineset is latest, none older', () => {
    const logbook = createLogbook();
    const { latest, older } = logbook._getLatestAndOlderLinesets({
        linesets: [ls(1)]
    });
    assert.equal(latest.number, 1);
    assert.equal(older.length, 0);
});

test('_getLatestAndOlderLinesets: highest active is latest; archived are older', () => {
    const logbook = createLogbook();
    const { latest, older } = logbook._getLatestAndOlderLinesets({
        linesets: [ls(1, { archived: true }), ls(2, { archived: true }), ls(3)]
    });
    assert.equal(latest.number, 3);
    assert.equal(older.length, 2);
    assert.equal(older[0].number, 2);
    assert.equal(older[1].number, 1);
});

test('_getLatestAndOlderLinesets: only the highest active is latest when several are unarchived', () => {
    const logbook = createLogbook();
    const { latest, older } = logbook._getLatestAndOlderLinesets({
        linesets: [ls(1), ls(2)]
    });
    assert.equal(latest.number, 2);
    assert.equal(older.length, 1);
    assert.equal(older[0].number, 1);
});

test('_getLatestAndOlderLinesets: falls back to highest number when all archived', () => {
    const logbook = createLogbook();
    const { latest, older } = logbook._getLatestAndOlderLinesets({
        linesets: [ls(1, { archived: true }), ls(2, { archived: true })]
    });
    assert.equal(latest.number, 2);
    assert.equal(older.length, 1);
    assert.equal(older[0].number, 1);
});

test('renderCanopiesWithLinesets hides older linesets and omits archive buttons', () => {
    const equipmentList = { innerHTML: '' };
    const logbook = createLogbook({
        canopies: [{
            id: 'c1',
            name: 'Petra 62',
            archived: false,
            linesets: [ls(1, { archived: true, jumpCount: 10 }), ls(2, { jumpCount: 3 })]
        }]
    }, equipmentList);

    logbook.renderCanopiesWithLinesets();
    const html = equipmentList.innerHTML;

    assert.match(html, /Lineset #2/);
    assert.match(html, /Show older linesets \(1\)/);
    assert.match(html, /older-linesets-c1/);
    assert.match(html, /display:none/);
    assert.doesNotMatch(html, /toggleArchiveLineset/);
    assert.match(html, /toggleArchiveComponent/);
    assert.match(html, /Lineset #1/);
    const olderBlock = html.slice(html.indexOf('older-linesets-group'));
    assert.match(olderBlock, /Lineset #1/);
});

test('renderCanopiesWithLinesets keeps older linesets visible when checkbox was already on', () => {
    const equipmentList = { innerHTML: '' };
    const logbook = createLogbook({
        canopies: [{
            id: 'c1',
            name: 'Petra 62',
            archived: false,
            linesets: [ls(1, { archived: true }), ls(2)]
        }],
        showOlderLinesetsByCanopyId: { c1: true }
    }, equipmentList);

    logbook.renderCanopiesWithLinesets();
    assert.match(equipmentList.innerHTML, /display:block/);
    assert.match(equipmentList.innerHTML, /checked/);
});

test('deleteLineset reactivates the previous lineset when the current one is removed', () => {
    const logbook = createLogbook({
        canopies: [{
            id: 'c1',
            name: 'Petra 62',
            linesets: [ls(1, { archived: true }), ls(2)]
        }]
    });

    logbook.deleteLineset('c1', 2);
    assert.equal(logbook.canopies[0].linesets.length, 1);
    assert.equal(logbook.canopies[0].linesets[0].number, 1);
    assert.equal(logbook.canopies[0].linesets[0].archived, false);
});
