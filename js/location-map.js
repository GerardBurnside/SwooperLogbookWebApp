/**
 * Location map modal — Google Maps, gates, live user distance.
 */
(function (global) {
    'use strict';

    const U = global.LocationGatesUtil;
    let mapsScriptPromise = null;
    let state = null;

    const PERSON_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="32" height="32"><circle cx="12" cy="5" r="3" fill="#1565C0"/><path d="M6 21v-1c0-3.3 2.7-6 6-6s6 2.7 6 6v1H6z" fill="#1565C0"/></svg>';

    function getApiKey() {
        const k = global.GOOGLE_MAPS_API_KEY;
        return typeof k === 'string' ? k.trim() : '';
    }

    function loadGoogleMaps() {
        const key = getApiKey();
        if (!key) {
            return Promise.reject(new Error('NO_KEY'));
        }
        if (global.google && global.google.maps) {
            return Promise.resolve(global.google.maps);
        }
        if (mapsScriptPromise) return mapsScriptPromise;
        mapsScriptPromise = new Promise((resolve, reject) => {
            const cb = '__swooperMapsInit_' + Date.now();
            global[cb] = () => {
                delete global[cb];
                if (global.google && global.google.maps) resolve(global.google.maps);
                else reject(new Error('Maps failed to load'));
            };
            const s = document.createElement('script');
            s.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(key)}&callback=${cb}`;
            s.async = true;
            s.onerror = () => reject(new Error('Maps script error'));
            document.head.appendChild(s);
        });
        return mapsScriptPromise;
    }

    function parseLatLng(lat, lng) {
        const la = parseFloat(lat);
        const ln = parseFloat(lng);
        if (!Number.isFinite(la) || !Number.isFinite(ln)) return null;
        if (la < -90 || la > 90 || ln < -180 || ln > 180) return null;
        return { lat: la, lng: ln };
    }

    function setHud(text) {
        const el = document.getElementById('locationMapHud');
        if (el) el.textContent = text;
    }

    function updateDistanceHud() {
        if (!state || !U) return;
        const sel = state.gates.find((g) => g.id === state.selectedGateId);
        if (!sel) {
            setHud(state.selectedGateId === null && state.gates.length
                ? 'Tap a gate to see live distance from your position'
                : state.gates.length ? 'Tap a gate to see live distance' : 'No gates — use Add gate');
            return;
        }
        if (!state.userPos) {
            setHud(`Selected: ${sel.label} — waiting for GPS…`);
            return;
        }
        const m = U.haversineM(state.userPos.lat, state.userPos.lng, sel.lat, sel.lng);
        setHud(`${U.formatDistance(m)} to ${sel.label}`);
    }

    function syncGateMarkers() {
        if (!state || !state.map) return;
        const maps = global.google.maps;
        const existing = new Map(state.gateMarkers.map((m) => [m.gateId, m]));
        state.gateMarkers = [];
        for (const gate of state.gates) {
            let rec = existing.get(gate.id);
            if (rec) {
                rec.marker.setPosition({ lat: gate.lat, lng: gate.lng });
                rec.marker.setTitle(gate.label);
                rec.marker.setIcon(rec.marker.getIcon());
            } else {
                const marker = new maps.Marker({
                    map: state.map,
                    position: { lat: gate.lat, lng: gate.lng },
                    title: gate.label,
                    draggable: true,
                    label: gate.label.slice(0, 1),
                });
                marker.addListener('click', () => selectGate(gate.id));
                marker.addListener('dragend', () => {
                    const p = marker.getPosition();
                    gate.lat = p.lat();
                    gate.lng = p.lng();
                    notifyChange();
                    updateDistanceHud();
                });
                rec = { gateId: gate.id, marker };
            }
            applyGateSelectionStyle(rec, gate.id === state.selectedGateId);
            state.gateMarkers.push(rec);
        }
        for (const [id, rec] of existing) {
            if (!state.gates.some((g) => g.id === id)) {
                rec.marker.setMap(null);
            }
        }
    }

    function applyGateSelectionStyle(rec, selected) {
        const marker = rec.marker;
        if (selected) {
            marker.setAnimation(global.google.maps.Animation.BOUNCE);
            setTimeout(() => marker.setAnimation(null), 700);
            marker.setZIndex(1000);
        } else {
            marker.setZIndex(undefined);
        }
    }

    function selectGate(gateId) {
        if (!state) return;
        state.selectedGateId = gateId;
        const panel = document.getElementById('locationMapGatePanel');
        const labelInp = document.getElementById('locationMapGateLabel');
        const gate = state.gates.find((g) => g.id === gateId);
        if (panel) panel.style.display = gate ? 'block' : 'none';
        if (labelInp && gate) labelInp.value = gate.label;
        for (const rec of state.gateMarkers) {
            applyGateSelectionStyle(rec, rec.gateId === gateId);
        }
        updateDistanceHud();
    }

    function notifyChange() {
        if (state && typeof state.onGatesChange === 'function') {
            state.onGatesChange(U.cloneGates(state.gates));
        }
    }

    function notifyLatLngChange(lat, lng) {
        if (state && typeof state.onLatLngChange === 'function') {
            state.onLatLngChange(lat, lng);
        }
    }

    function fitBoundsOnce() {
        if (!state || !state.map || state.didFitBounds) return;
        const maps = global.google.maps;
        const bounds = new maps.LatLngBounds();
        let n = 0;
        if (state.locationPos) {
            bounds.extend(state.locationPos);
            n++;
        }
        for (const g of state.gates) {
            bounds.extend({ lat: g.lat, lng: g.lng });
            n++;
        }
        if (state.userPos) {
            bounds.extend(state.userPos);
            n++;
        }
        if (n === 0) return;
        state.didFitBounds = true;
        if (n === 1) {
            state.map.setCenter(bounds.getCenter());
            state.map.setZoom(16);
        } else {
            state.map.fitBounds(bounds, 48);
        }
    }

    function stopWatch() {
        if (state && state.watchId != null && navigator.geolocation) {
            navigator.geolocation.clearWatch(state.watchId);
            state.watchId = null;
        }
    }

    function close() {
        stopWatch();
        const modal = document.getElementById('locationMapModal');
        if (modal) modal.style.display = 'none';
        if (state && state.mapClickListener) {
            global.google.maps.event.removeListener(state.mapClickListener);
        }
        if (state) {
            for (const rec of state.gateMarkers) rec.marker.setMap(null);
            if (state.locationMarker) state.locationMarker.setMap(null);
            if (state.userMarker) state.userMarker.setMap(null);
        }
        state = null;
    }

    function startGeolocation(maps) {
        if (!navigator.geolocation) {
            setHud('Geolocation not supported');
            return;
        }
        if (!state.userMarker) {
            state.userMarker = new maps.Marker({
                map: state.map,
                icon: {
                    url: 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(PERSON_SVG),
                    scaledSize: new maps.Size(32, 32),
                    anchor: new maps.Point(16, 32),
                },
                title: 'You',
                zIndex: 2000,
            });
        }
        state.watchId = navigator.geolocation.watchPosition(
            (pos) => {
                state.userPos = { lat: pos.coords.latitude, lng: pos.coords.longitude };
                state.userMarker.setPosition(state.userPos);
                if (!state.didFitBounds) fitBoundsOnce();
                updateDistanceHud();
            },
            () => {
                if (!state.userPos) setHud('Location unavailable — enable GPS to see live distance');
            },
            { enableHighAccuracy: true, maximumAge: 2000, timeout: 15000 }
        );
    }

    function open(options) {
        const modal = document.getElementById('locationMapModal');
        const mapEl = document.getElementById('locationMapCanvas');
        const errEl = document.getElementById('locationMapError');
        if (!modal || !mapEl) return;

        options = options || {};
        const loc = parseLatLng(options.lat, options.lng);
        const gates = U ? U.cloneGates(options.gates || []) : [];

        close();
        modal.style.display = 'block';
        if (errEl) {
            errEl.style.display = 'none';
            errEl.textContent = '';
        }
        mapEl.innerHTML = '';

        state = {
            map: null,
            gates,
            selectedGateId: gates.length ? gates[0].id : null,
            locationPos: loc,
            userPos: null,
            gateMarkers: [],
            locationMarker: null,
            userMarker: null,
            watchId: null,
            didFitBounds: false,
            addGateMode: false,
            onGatesChange: options.onGatesChange,
            onLatLngChange: options.onLatLngChange,
            mapClickListener: null,
        };

        loadGoogleMaps()
            .then((maps) => {
                const center = loc || (gates[0] ? { lat: gates[0].lat, lng: gates[0].lng } : { lat: 0, lng: 0 });
                const zoom = loc || gates.length ? 16 : 2;
                state.map = new maps.Map(mapEl, {
                    center,
                    zoom,
                    mapTypeId: 'hybrid',
                    gestureHandling: 'greedy',
                });

                if (loc) {
                    state.locationMarker = new maps.Marker({
                        map: state.map,
                        position: loc,
                        title: 'Location',
                        draggable: true,
                        icon: {
                            path: maps.SymbolPath.CIRCLE,
                            scale: 10,
                            fillColor: '#2e7d32',
                            fillOpacity: 1,
                            strokeColor: '#fff',
                            strokeWeight: 2,
                        },
                    });
                    state.locationMarker.addListener('dragend', () => {
                        const p = state.locationMarker.getPosition();
                        state.locationPos = { lat: p.lat(), lng: p.lng() };
                        notifyLatLngChange(p.lat(), p.lng());
                    });
                }

                syncGateMarkers();
                selectGate(state.selectedGateId);
                fitBoundsOnce();
                startGeolocation(maps);

                state.mapClickListener = state.map.addListener('click', (e) => {
                    if (!state.addGateMode) return;
                    state.addGateMode = false;
                    const btn = document.getElementById('locationMapAddGateBtn');
                    if (btn) btn.classList.remove('active');
                    const n = state.gates.length + 1;
                    const gate = {
                        id: U.newGateId(),
                        label: `Gate ${n}`,
                        lat: e.latLng.lat(),
                        lng: e.latLng.lng(),
                    };
                    state.gates.push(gate);
                    syncGateMarkers();
                    selectGate(gate.id);
                    notifyChange();
                    updateDistanceHud();
                });
            })
            .catch((err) => {
                mapEl.innerHTML = '';
                if (errEl) {
                    errEl.style.display = 'block';
                    errEl.textContent = err.message === 'NO_KEY'
                        ? 'Maps API key is not configured. For local dev, copy config/google-maps.example.js to config/google-maps.js.'
                        : 'Could not load Google Maps.';
                }
            });
    }

    function wireUi() {
        document.getElementById('locationMapClose')?.addEventListener('click', () => close());
        document.getElementById('locationMapDone')?.addEventListener('click', () => close());
        document.getElementById('locationMapModal')?.addEventListener('click', (e) => {
            if (e.target.id === 'locationMapModal') close();
        });
        document.getElementById('locationMapAddGateBtn')?.addEventListener('click', () => {
            if (!state || !state.map) return;
            state.addGateMode = true;
            const btn = document.getElementById('locationMapAddGateBtn');
            if (btn) btn.classList.add('active');
            setHud('Tap the map to place a new gate');
        });
        document.getElementById('locationMapDeleteGateBtn')?.addEventListener('click', () => {
            if (!state || !state.selectedGateId) return;
            state.gates = state.gates.filter((g) => g.id !== state.selectedGateId);
            state.selectedGateId = state.gates.length ? state.gates[0].id : null;
            syncGateMarkers();
            selectGate(state.selectedGateId);
            notifyChange();
            updateDistanceHud();
        });
        document.getElementById('locationMapGateLabel')?.addEventListener('change', (e) => {
            if (!state || !state.selectedGateId) return;
            const gate = state.gates.find((g) => g.id === state.selectedGateId);
            if (!gate) return;
            gate.label = (e.target.value || gate.label).trim() || gate.label;
            syncGateMarkers();
            notifyChange();
            updateDistanceHud();
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wireUi);
    } else {
        wireUi();
    }

    global.LocationMap = { open, close };
}(typeof window !== 'undefined' ? window : global));
