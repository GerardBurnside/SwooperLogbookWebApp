/**
 * Location map modal — Google Maps, gates, live user distance.
 */
(function (global) {
    'use strict';

    const U = global.LocationGatesUtil;
    let mapsScriptPromise = null;
    let state = null;
    let onViewportChange = null;

    const PERSON_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 40"><g fill="#1565C0"><circle cx="16" cy="7.6" r="4.4"/><rect x="12.4" y="12" width="7.2" height="10.2" rx="2"/><path d="M12.6 14.2 L4.2 18.8 L5.6 21.2 L13.4 17 Z"/><path d="M19.4 14.2 L27.8 18.8 L26.4 21.2 L18.6 17 Z"/><path d="M13.2 21.8 L10.2 36.8 L13.6 36.8 L15.4 21.8 Z"/><path d="M18.8 21.8 L21.8 36.8 L18.4 36.8 L16.6 21.8 Z"/></g></svg>';
    const PERSON_ICON_SIZE = { w: 32, h: 40 };
    const PERSON_ICON_CENTER = { x: 16, y: 20 };
    const COARSE_ACCURACY_M = 25;
    const OPENING_DIAGONAL_M = 100;
    const MERCATOR_MPP_Z0 = 156543.03392804097;

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

    function syncCenterUserBtn() {
        const btn = document.getElementById('locationMapCenterUserBtn');
        if (!btn) return;
        const ready = !!(state && state.userPos);
        btn.disabled = !ready;
        btn.title = ready
            ? 'Center map on your location'
            : 'Waiting for GPS — center on you when ready';
        btn.setAttribute('aria-label', btn.title);
    }

    function centerMapOnUser() {
        if (!state?.map || !state.userPos) return;
        const size = mapPixelSize();
        state.map.setCenter(state.userPos);
        if (size) {
            state.map.setZoom(
                zoomForViewportDiagonal(state.userPos.lat, OPENING_DIAGONAL_M, size.w, size.h)
            );
        } else {
            state.map.setZoom(19);
        }
    }

    function gpsHudSuffix() {
        if (!state || !state.userPos) return '';
        const acc = U ? U.formatAccuracy(state.userAccuracy) : '';
        if (!acc) return 'GPS';
        if (state.userAccuracy > COARSE_ACCURACY_M) return `GPS ${acc} — waiting for a better fix`;
        return `GPS ${acc}`;
    }

    function geoErrorMessage(err) {
        const code = err && err.code;
        if (code === 1) return 'Location permission denied — enable GPS in browser settings';
        if (code === 2) return 'Location unavailable — check GPS and try again';
        if (code === 3) return 'GPS timed out — wait outdoors for a better signal';
        return 'Location unavailable — enable GPS to see live distance';
    }

    function gateEditingEnabled() {
        return !!(state && state.gateEditingEnabled);
    }

    /** Gates created in the app (not from location-gates-seed.js). */
    function userAddedGates(gates) {
        if (!gates?.length) return [];
        return gates.filter((g) => {
            const src = g.source;
            return !src || !String(src).startsWith('seed:');
        });
    }

    function resetGatesToDefault() {
        if (!state || !state.map || !U || !state.seedEntries?.length) return;
        if (state.gateEditToggle && !gateEditingEnabled()) return;

        const added = userAddedGates(state.gates);
        let msg = 'Reset all gates to the defaults from the app seed? Any changes to gate positions and names will be lost.';
        if (added.length) {
            const names = added.map((g) => g.label || 'Gate').join(', ');
            msg += `\n\nThese gates you added will be removed: ${names}.`;
        }
        if (!global.confirm(msg)) return;

        state.gates = U.cloneGates(U.gatesFromSeedEntries(state.seedEntries));
        state.selectedGateId = state.gates.length ? state.gates[0].id : null;
        state.addGateMode = false;
        const addBtn = document.getElementById('locationMapAddGateBtn');
        if (addBtn) addBtn.classList.remove('active');
        syncGateMarkers();
        selectGate(state.selectedGateId);
        notifyChange();
        updateDistanceHud();
        syncGateEditUi();
    }

    function syncGateEditUi() {
        const editWrap = document.getElementById('locationMapEditWrap');
        const editChk = document.getElementById('locationMapEditCheckbox');
        const addBtn = document.getElementById('locationMapAddGateBtn');
        const resetBtn = document.getElementById('locationMapResetGatesBtn');
        const delBtn = document.getElementById('locationMapDeleteGateBtn');
        const labelInp = document.getElementById('locationMapGateLabel');
        const showToggle = !!(state && state.gateEditToggle);
        const hasSeed = !!(state && state.seedEntries && state.seedEntries.length);
        if (editWrap) editWrap.hidden = !showToggle;
        if (editChk && showToggle) editChk.checked = gateEditingEnabled();
        const canEdit = gateEditingEnabled();
        if (addBtn) addBtn.disabled = !canEdit;
        if (resetBtn) {
            resetBtn.hidden = !hasSeed;
            resetBtn.disabled = !hasSeed || (state.gateEditToggle && !canEdit);
        }
        if (delBtn) delBtn.disabled = !canEdit;
        if (labelInp) labelInp.readOnly = !canEdit;
        if (state && !canEdit && state.addGateMode) {
            state.addGateMode = false;
            if (addBtn) addBtn.classList.remove('active');
        }
        if (state) {
            for (const rec of state.gateMarkers) {
                rec.marker.setDraggable(canEdit);
            }
        }
    }

    function updateDistanceHud() {
        if (!state || !U) return;
        if (state.addGateMode) return;
        const gps = gpsHudSuffix();
        const sel = state.gates.find((g) => g.id === state.selectedGateId);
        if (!sel) {
            const prompt = state.selectedGateId === null && state.gates.length
                ? 'Tap a gate to see live distance from your position'
                : state.gates.length ? 'Tap a gate to see live distance' : (gateEditingEnabled() ? 'No gates — use Add gate' : 'No gates');
            setHud(gps ? `${prompt}  ·  ${gps}` : prompt);
            syncCenterUserBtn();
            return;
        }
        if (!state.userPos) {
            setHud(`Selected: ${sel.label} — waiting for GPS…`);
            syncCenterUserBtn();
            return;
        }
        const m = U.haversineM(state.userPos.lat, state.userPos.lng, sel.lat, sel.lng);
        setHud(`${U.formatDistance(m)} to ${sel.label}  ·  ${gps}`);
        syncCenterUserBtn();
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
                    draggable: gateEditingEnabled(),
                    label: gate.label.slice(0, 1),
                });
                marker.addListener('click', () => selectGate(gate.id));
                marker.addListener('dragend', () => {
                    if (!gateEditingEnabled()) return;
                    const p = marker.getPosition();
                    gate.lat = p.lat();
                    gate.lng = p.lng();
                    notifyChange();
                    updateDistanceHud();
                });
                rec = { gateId: gate.id, marker };
            }
            applyGateSelectionStyle(rec, gate.id === state.selectedGateId);
            rec.marker.setDraggable(gateEditingEnabled());
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

    function gatesCenter() {
        if (!state || !state.gates.length) return null;
        let minLat = Infinity;
        let maxLat = -Infinity;
        let minLng = Infinity;
        let maxLng = -Infinity;
        for (const g of state.gates) {
            if (g.lat < minLat) minLat = g.lat;
            if (g.lat > maxLat) maxLat = g.lat;
            if (g.lng < minLng) minLng = g.lng;
            if (g.lng > maxLng) maxLng = g.lng;
        }
        return { lat: (minLat + maxLat) / 2, lng: (minLng + maxLng) / 2 };
    }

    function openingCenter() {
        return gatesCenter() || (state && (state.locationPos || state.userPos)) || null;
    }

    function zoomForViewportDiagonal(lat, diagonalM, widthPx, heightPx) {
        const diagPx = Math.hypot(widthPx, heightPx);
        if (!(diagPx > 0) || !Number.isFinite(lat) || !(diagonalM > 0)) return 18;
        const mpp = diagonalM / diagPx;
        const cosLat = Math.max(0.01, Math.abs(Math.cos((lat * Math.PI) / 180)));
        const zoom = Math.log2((MERCATOR_MPP_Z0 * cosLat) / mpp);
        if (!Number.isFinite(zoom)) return 18;
        return Math.min(21, Math.max(2, zoom));
    }

    function mapPixelSize() {
        const div = state && state.map ? state.map.getDiv() : null;
        const w = div ? div.clientWidth : 0;
        const h = div ? div.clientHeight : 0;
        if (w > 0 && h > 0) return { w, h };
        const el = document.getElementById('locationMapCanvas');
        const cw = el ? el.clientWidth : 0;
        const ch = el ? el.clientHeight : 0;
        return cw > 0 && ch > 0 ? { w: cw, h: ch } : null;
    }

    function applyOpeningView() {
        if (!state || !state.map || state.didFitBounds) return;
        const center = openingCenter();
        const size = mapPixelSize();
        if (!center || !size) return;
        state.didFitBounds = true;
        state.map.setCenter(center);
        state.map.setZoom(zoomForViewportDiagonal(center.lat, OPENING_DIAGONAL_M, size.w, size.h));
    }

    function stopWatch() {
        if (state && state.watchId != null && navigator.geolocation) {
            navigator.geolocation.clearWatch(state.watchId);
            state.watchId = null;
        }
    }

    function clearModalViewportStyles(modal) {
        if (!modal) return;
        modal.style.top = '';
        modal.style.left = '';
        modal.style.right = '';
        modal.style.bottom = '';
        modal.style.width = '';
        modal.style.height = '';
    }

    function syncModalViewport() {
        const modal = document.getElementById('locationMapModal');
        if (!modal || modal.style.display === 'none') return;
        const vv = global.visualViewport;
        if (vv) {
            modal.style.top = `${vv.offsetTop}px`;
            modal.style.left = `${vv.offsetLeft}px`;
            modal.style.right = 'auto';
            modal.style.bottom = 'auto';
            modal.style.width = `${vv.width}px`;
            modal.style.height = `${vv.height}px`;
        } else {
            modal.style.top = '0';
            modal.style.left = '0';
            modal.style.right = '0';
            modal.style.bottom = '0';
            modal.style.width = 'auto';
            modal.style.height = 'auto';
        }
    }

    function triggerMapResize() {
        if (!state?.map || !global.google?.maps) return;
        global.google.maps.event.trigger(state.map, 'resize');
    }

    function bindModalViewportListeners() {
        unbindModalViewportListeners();
        onViewportChange = () => {
            syncModalViewport();
            triggerMapResize();
        };
        syncModalViewport();
        const vv = global.visualViewport;
        if (vv) {
            vv.addEventListener('resize', onViewportChange);
            vv.addEventListener('scroll', onViewportChange);
        }
        global.addEventListener('resize', onViewportChange);
    }

    function unbindModalViewportListeners() {
        if (!onViewportChange) return;
        const vv = global.visualViewport;
        if (vv) {
            vv.removeEventListener('resize', onViewportChange);
            vv.removeEventListener('scroll', onViewportChange);
        }
        global.removeEventListener('resize', onViewportChange);
        onViewportChange = null;
    }

    function close() {
        stopWatch();
        unbindModalViewportListeners();
        const modal = document.getElementById('locationMapModal');
        if (modal) {
            modal.style.display = 'none';
            clearModalViewportStyles(modal);
        }
        if (global.document?.body) {
            global.document.body.style.overflow = '';
        }
        if (state && state.mapClickListener) {
            global.google.maps.event.removeListener(state.mapClickListener);
        }
        if (state) {
            for (const rec of state.gateMarkers) rec.marker.setMap(null);
            if (state.locationMarker) state.locationMarker.setMap(null);
            if (state.userMarker) state.userMarker.setMap(null);
            if (state.userAccuracyCircle) state.userAccuracyCircle.setMap(null);
        }
        state = null;
        const editWrap = document.getElementById('locationMapEditWrap');
        const addBtn = document.getElementById('locationMapAddGateBtn');
        const resetBtn = document.getElementById('locationMapResetGatesBtn');
        if (editWrap) editWrap.hidden = true;
        if (resetBtn) resetBtn.hidden = true;
        if (addBtn) {
            addBtn.disabled = false;
            addBtn.classList.remove('active');
        }
    }

    function ensureUserOverlay(maps) {
        if (!state.userMarker) {
            state.userMarker = new maps.Marker({
                map: state.map,
                position: state.userPos,
                icon: {
                    url: 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(PERSON_SVG),
                    scaledSize: new maps.Size(PERSON_ICON_SIZE.w, PERSON_ICON_SIZE.h),
                    anchor: new maps.Point(PERSON_ICON_CENTER.x, PERSON_ICON_CENTER.y),
                },
                title: 'You',
                zIndex: 2000,
            });
        }
        if (!state.userAccuracyCircle) {
            state.userAccuracyCircle = new maps.Circle({
                map: state.map,
                center: state.userPos,
                radius: state.userAccuracy > 0 ? state.userAccuracy : 1,
                fillColor: '#1565C0',
                fillOpacity: 0.15,
                strokeColor: '#1565C0',
                strokeOpacity: 0.5,
                strokeWeight: 1,
                clickable: false,
                zIndex: 1999,
            });
        }
    }

    function applyUserPosition(maps, coords) {
        const acc = coords.accuracy;
        state.userPos = { lat: coords.latitude, lng: coords.longitude };
        state.userAccuracy = Number.isFinite(acc) && acc >= 0 ? acc : null;
        ensureUserOverlay(maps);
        state.userMarker.setPosition(state.userPos);
        const accLabel = U ? U.formatAccuracy(state.userAccuracy) : '';
        state.userMarker.setTitle(accLabel ? `You (${accLabel})` : 'You');
        state.userAccuracyCircle.setCenter(state.userPos);
        if (state.userAccuracy != null && state.userAccuracy > 0) {
            state.userAccuracyCircle.setRadius(state.userAccuracy);
            state.userAccuracyCircle.setMap(state.map);
        } else {
            state.userAccuracyCircle.setMap(null);
        }
        syncCenterUserBtn();
    }

    function startGeolocation(maps) {
        if (!navigator.geolocation) {
            setHud('Geolocation not supported');
            syncCenterUserBtn();
            return;
        }
        syncCenterUserBtn();
        state.watchId = navigator.geolocation.watchPosition(
            (pos) => {
                if (!state) return;
                applyUserPosition(maps, pos.coords);
                if (!state.didFitBounds) applyOpeningView();
                updateDistanceHud();
            },
            (err) => {
                if (!state || state.userPos) return;
                setHud(geoErrorMessage(err));
                syncCenterUserBtn();
            },
            { enableHighAccuracy: true, maximumAge: 0, timeout: 30000 }
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
        modal.style.display = 'flex';
        bindModalViewportListeners();
        if (global.document?.body) {
            global.document.body.style.overflow = 'hidden';
        }
        if (errEl) {
            errEl.style.display = 'none';
            errEl.textContent = '';
        }
        mapEl.innerHTML = '';

        const gateEditToggle = !!options.gateEditToggle;
        const seedEntries = Array.isArray(options.seedEntries) ? options.seedEntries : [];
        state = {
            map: null,
            gates,
            seedEntries,
            selectedGateId: gates.length ? gates[0].id : null,
            locationPos: loc,
            userPos: null,
            userAccuracy: null,
            gateMarkers: [],
            locationMarker: null,
            userMarker: null,
            userAccuracyCircle: null,
            watchId: null,
            didFitBounds: false,
            addGateMode: false,
            gateEditToggle,
            gateEditingEnabled: !gateEditToggle,
            onGatesChange: options.onGatesChange,
            onLatLngChange: options.onLatLngChange,
            mapClickListener: null,
        };

        const editChk = document.getElementById('locationMapEditCheckbox');
        if (editChk) editChk.checked = state.gateEditingEnabled;
        syncGateEditUi();

        loadGoogleMaps()
            .then((maps) => {
                const focus = openingCenter();
                const center = focus || { lat: 0, lng: 0 };
                const size = mapEl.clientWidth > 0 && mapEl.clientHeight > 0
                    ? { w: mapEl.clientWidth, h: mapEl.clientHeight }
                    : null;
                const zoom = focus && size
                    ? zoomForViewportDiagonal(center.lat, OPENING_DIAGONAL_M, size.w, size.h)
                    : (focus ? 19 : 2);
                state.map = new maps.Map(mapEl, {
                    center,
                    zoom,
                    mapTypeId: 'hybrid',
                    gestureHandling: 'greedy',
                    isFractionalZoomEnabled: true,
                    maxZoom: 21,
                });
                if (focus && size) state.didFitBounds = true;

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
                syncCenterUserBtn();
                applyOpeningView();
                const refitMap = () => {
                    triggerMapResize();
                    applyOpeningView();
                };
                maps.event.addListenerOnce(state.map, 'idle', refitMap);
                global.requestAnimationFrame(refitMap);
                startGeolocation(maps);

                state.mapClickListener = state.map.addListener('click', (e) => {
                    if (!state.addGateMode || !gateEditingEnabled()) return;
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
        document.getElementById('locationMapEditCheckbox')?.addEventListener('change', (e) => {
            if (!state || !state.gateEditToggle) return;
            state.gateEditingEnabled = !!e.target.checked;
            syncGateEditUi();
            updateDistanceHud();
        });
        document.getElementById('locationMapResetGatesBtn')?.addEventListener('click', () => {
            resetGatesToDefault();
        });
        document.getElementById('locationMapAddGateBtn')?.addEventListener('click', () => {
            if (!state || !state.map || !gateEditingEnabled()) return;
            state.addGateMode = true;
            const btn = document.getElementById('locationMapAddGateBtn');
            if (btn) btn.classList.add('active');
            setHud('Tap the map to place a new gate');
            syncCenterUserBtn();
        });
        document.getElementById('locationMapCenterUserBtn')?.addEventListener('click', () => {
            centerMapOnUser();
        });
        document.getElementById('locationMapDeleteGateBtn')?.addEventListener('click', () => {
            if (!state || !state.selectedGateId || !gateEditingEnabled()) return;
            state.gates = state.gates.filter((g) => g.id !== state.selectedGateId);
            state.selectedGateId = state.gates.length ? state.gates[0].id : null;
            syncGateMarkers();
            selectGate(state.selectedGateId);
            notifyChange();
            updateDistanceHud();
        });
        document.getElementById('locationMapGateLabel')?.addEventListener('change', (e) => {
            if (!state || !state.selectedGateId || !gateEditingEnabled()) return;
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
