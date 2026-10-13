const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const vm = require('node:vm');

function loadUtil() {
    const filePath = path.join(__dirname, '..', 'js', 'location-gates-util.js');
    const code = require('node:fs').readFileSync(filePath, 'utf8');
    const sandbox = { console };
    sandbox.globalThis = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(code, sandbox);
    return sandbox.LocationGatesUtil;
}

const U = loadUtil();

test('haversineM: same point is zero', () => {
    assert.equal(U.haversineM(32, -111, 32, -111), 0);
});

test('formatDistance: meters vs km', () => {
    assert.equal(U.formatDistance(142), '142 m');
    assert.equal(U.formatDistance(999), '999 m');
    assert.equal(U.formatDistance(1000), '1000 m');
    assert.equal(U.formatDistance(1001), '1.0 km');
    assert.equal(U.formatDistance(1500), '1.5 km');
});

test('formatAccuracy: GPS radius in meters', () => {
    assert.equal(U.formatAccuracy(5), '±5 m');
    assert.equal(U.formatAccuracy(2.4), '±2 m');
    assert.equal(U.formatAccuracy(0.4), '±1 m');
    assert.equal(U.formatAccuracy(85.6), '±86 m');
    assert.equal(U.formatAccuracy(NaN), '');
    assert.equal(U.formatAccuracy(-1), '');
});

test('ensureLocationGatesSeeded: applies once when empty', () => {
    const loc = { id: 'loc_eloy', name: 'Eloy', gates: [] };
    const seed = {
        loc_eloy: [{ label: 'Eloy 020', lat: 32.808, lng: -111.581, source: 'seed:Eloy_020.CSV' }]
    };
    assert.equal(U.ensureLocationGatesSeeded(loc, 'loc_eloy', seed), true);
    assert.equal(loc.gates.length, 1);
    assert.equal(loc.gates[0].label, 'Eloy 020');
    assert.equal(U.ensureLocationGatesSeeded(loc, 'loc_eloy', seed), false);
});

test('ensureLocationGatesSeeded: does not overwrite existing gates', () => {
    const loc = { id: 'loc_eloy', gates: [{ id: 'g1', label: 'Mine', lat: 1, lng: 2 }] };
    const seed = { loc_eloy: [{ label: 'Seed', lat: 3, lng: 4 }] };
    assert.equal(U.ensureLocationGatesSeeded(loc, 'loc_eloy', seed), false);
    assert.equal(loc.gates[0].label, 'Mine');
});

test('buildGatesSeedFromLocations: strips ids and omits empty', () => {
    const locations = [
        { id: 'loc_b', name: 'B', gates: [{ id: 'g1', label: 'Gate B', lat: 1.23456789, lng: 2.5, source: 'seed:x.CSV' }] },
        { id: 'loc_a', name: 'A', gates: [] },
        { id: 'loc_c', name: 'C' }
    ];
    const seed = U.buildGatesSeedFromLocations(locations);
    assert.deepEqual(Object.keys(seed), ['loc_b']);
    assert.equal(seed.loc_b[0].label, 'Gate B');
    assert.equal(seed.loc_b[0].lat, 1.2345679);
    assert.equal(seed.loc_b[0].source, 'seed:x.CSV');
    assert.equal(seed.loc_b[0].id, undefined);
});

test('formatLocationGatesSeedJs: valid assignable snippet', () => {
    const seed = { loc_test: [{ label: 'T', lat: 10, lng: 20 }] };
    const js = U.formatLocationGatesSeedJs(seed);
    assert.match(js, /^\/\*\*/);
    assert.match(js, /window\.LOCATION_GATES_SEED = \{/);
    const sandbox = {};
    vm.runInContext(js, vm.createContext(sandbox));
    assert.deepEqual(sandbox.LOCATION_GATES_SEED.loc_test[0].label, 'T');
});
