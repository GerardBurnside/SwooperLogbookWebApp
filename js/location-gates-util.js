/**
 * Location gate helpers (browser + Node tests via vm).
 */
(function (global) {
    'use strict';

    function haversineM(lat1, lng1, lat2, lng2) {
        const R = 6371000;
        const toRad = (d) => (d * Math.PI) / 180;
        const dLat = toRad(lat2 - lat1);
        const dLng = toRad(lng2 - lng1);
        const a = Math.sin(dLat / 2) ** 2
            + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
        return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    }

    function formatDistance(meters) {
        if (!Number.isFinite(meters)) return '—';
        if (meters <= 1000) return `${Math.round(meters)} m`;
        return `${(meters / 1000).toFixed(1)} km`;
    }

    function formatAccuracy(meters) {
        if (!Number.isFinite(meters) || meters < 0) return '';
        return `±${Math.max(1, Math.round(meters))} m`;
    }

    function cloneGate(gate) {
        const g = {
            id: gate.id,
            label: gate.label,
            lat: gate.lat,
            lng: gate.lng
        };
        if (gate.source) g.source = gate.source;
        return g;
    }

    function cloneGates(gates) {
        if (!Array.isArray(gates)) return [];
        return gates.map(cloneGate);
    }

    function newGateId() {
        if (typeof crypto !== 'undefined' && crypto.randomUUID) {
            return `gate_${crypto.randomUUID()}`;
        }
        return `gate_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
    }

    function gatesFromSeedEntries(seedEntries) {
        if (!Array.isArray(seedEntries)) return [];
        return seedEntries.map((entry) => ({
            id: newGateId(),
            label: entry.label,
            lat: entry.lat,
            lng: entry.lng,
            source: entry.source
        }));
    }

    function ensureLocationGatesSeeded(location, locationId, seedByLocationId) {
        if (!location || !locationId || !seedByLocationId) return false;
        if (Array.isArray(location.gates) && location.gates.length > 0) return false;
        const seedEntries = seedByLocationId[locationId];
        if (!seedEntries || !seedEntries.length) return false;
        location.gates = gatesFromSeedEntries(seedEntries);
        return true;
    }

    /** True when running on local dev (not production GitHub Pages). */
    function isDevHost() {
        if (typeof location === 'undefined') return false;
        const h = location.hostname;
        return h === 'localhost' || h === '127.0.0.1' || h === '[::1]';
    }

    function roundCoord(n) {
        const v = Number(n);
        if (!Number.isFinite(v)) return null;
        return Math.round(v * 1e7) / 1e7;
    }

    /**
     * One saved gate → seed file entry (no id).
     * @param {object} gate
     * @returns {object|null}
     */
    function gateToSeedEntry(gate) {
        if (!gate) return null;
        const lat = roundCoord(gate.lat);
        const lng = roundCoord(gate.lng);
        if (lat == null || lng == null) return null;
        const entry = {
            label: String(gate.label || 'Gate').trim() || 'Gate',
            lat,
            lng
        };
        if (gate.source) entry.source = gate.source;
        return entry;
    }

    /**
     * Build LOCATION_GATES_SEED object from location records (saved gates only).
     * @param {object[]} locations
     * @returns {Record<string, object[]>}
     */
    function buildGatesSeedFromLocations(locations) {
        const seed = {};
        if (!Array.isArray(locations)) return seed;
        const withGates = locations.filter(
            (loc) => loc && loc.id && Array.isArray(loc.gates) && loc.gates.length > 0
        );
        withGates.sort((a, b) => String(a.id).localeCompare(String(b.id)));
        for (const loc of withGates) {
            const entries = loc.gates.map(gateToSeedEntry).filter(Boolean);
            if (entries.length) seed[loc.id] = entries;
        }
        return seed;
    }

    /**
     * Full contents for js/location-gates-seed.js
     * @param {Record<string, object[]>} seed
     * @returns {string}
     */
    function formatLocationGatesSeedJs(seed) {
        const header = '/** Auto-generated or dev-exported — see scripts/build-location-gates-seed.py for CSV regen. */\n';
        return header + 'window.LOCATION_GATES_SEED = ' + JSON.stringify(seed, null, 2) + ';\n';
    }

    const api = {
        haversineM,
        formatDistance,
        formatAccuracy,
        cloneGate,
        cloneGates,
        newGateId,
        gatesFromSeedEntries,
        ensureLocationGatesSeeded,
        isDevHost,
        gateToSeedEntry,
        buildGatesSeedFromLocations,
        formatLocationGatesSeedJs
    };

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = api;
    } else {
        global.LocationGatesUtil = api;
    }
}(typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : global));
