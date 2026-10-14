/**
 * Flysight GNSS track CSV parsing and max vertical speed analysis.
 */
(function (global) {
    'use strict';

    const REQUIRED_COLUMNS = ['time', 'hMSL', 'velD'];

    /**
     * @param {string} line
     * @returns {string[]}
     */
    function splitCsvLine(line) {
        return line.split(',');
    }

    /**
     * @param {string} headerLine
     * @returns {Map<string, number>}
     */
    function parseHeaderIndices(headerLine) {
        const cols = splitCsvLine(headerLine.trim());
        const map = new Map();
        cols.forEach((name, i) => {
            const key = name.trim();
            if (key) map.set(key, i);
        });
        return map;
    }

    /**
     * @param {string} text
     * @returns {{ points: { time: string, hMSL: number, velD: number, velN?: number, velE?: number, lat?: number, lon?: number, sAcc?: number }[], error?: string }}
     */
    function parseFlysightCsv(text) {
        const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim() !== '');
        if (lines.length < 2) {
            return { points: [], error: 'File is empty or too short.' };
        }

        let headerIdx = -1;
        let colMap = null;
        for (let i = 0; i < lines.length; i++) {
            const map = parseHeaderIndices(lines[i]);
            if (REQUIRED_COLUMNS.every(c => map.has(c))) {
                headerIdx = i;
                colMap = map;
                break;
            }
        }

        if (!colMap) {
            return { points: [], error: 'Missing required columns (time, hMSL, velD).' };
        }

        const timeCol = colMap.get('time');
        const hmslCol = colMap.get('hMSL');
        const velDCol = colMap.get('velD');
        const velNCol = colMap.has('velN') ? colMap.get('velN') : -1;
        const velECol = colMap.has('velE') ? colMap.get('velE') : -1;
        const latCol = colMap.has('lat') ? colMap.get('lat') : -1;
        const lonCol = colMap.has('lon') ? colMap.get('lon') : -1;
        const sAccCol = colMap.has('sAcc') ? colMap.get('sAcc') : -1;
        const points = [];

        for (let i = headerIdx + 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            const first = line.split(',')[0].trim();
            if (!first || first.startsWith('(')) continue;

            const cols = splitCsvLine(line);
            const time = cols[timeCol]?.trim();
            const hMSL = parseFloat(cols[hmslCol]);
            const velD = parseFloat(cols[velDCol]);
            if (!time || !Number.isFinite(hMSL) || !Number.isFinite(velD)) continue;

            const point = { time, hMSL, velD };
            if (velNCol >= 0 && velECol >= 0) {
                const velN = parseFloat(cols[velNCol]);
                const velE = parseFloat(cols[velECol]);
                if (Number.isFinite(velN) && Number.isFinite(velE)) {
                    point.velN = velN;
                    point.velE = velE;
                }
            }
            if (latCol >= 0 && lonCol >= 0) {
                const lat = parseFloat(cols[latCol]);
                const lon = parseFloat(cols[lonCol]);
                if (Number.isFinite(lat) && Number.isFinite(lon)) {
                    point.lat = lat;
                    point.lon = lon;
                }
            }
            if (sAccCol >= 0) {
                const sAcc = parseFloat(cols[sAccCol]);
                if (Number.isFinite(sAcc)) point.sAcc = sAcc;
            }
            points.push(point);
        }

        if (points.length === 0) {
            return { points: [], error: 'No valid track points found.' };
        }

        return { points };
    }

    const MONTH_NAMES = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];

    /**
     * First track timestamp in the device timezone as "March 9 - 16:53".
     * When the track year is not the current year, includes it: "March 9, 2025 - 16:53".
     * Flysight `time` values are UTC (ISO-8601 with Z); Date local getters apply the phone offset.
     *
     * @param {string} [timeStr]
     * @returns {string}
     */
    function formatTrackStartTitle(timeStr) {
        const ms = Date.parse(timeStr);
        if (!Number.isFinite(ms)) return '';
        const d = new Date(ms);
        const hh = String(d.getHours()).padStart(2, '0');
        const mm = String(d.getMinutes()).padStart(2, '0');
        let datePart = `${MONTH_NAMES[d.getMonth()]} ${d.getDate()}`;
        if (d.getFullYear() !== new Date().getFullYear()) {
            datePart += `, ${d.getFullYear()}`;
        }
        return `${datePart} - ${hh}:${mm}`;
    }

    /** Typical Flysight log interval (10 Hz) when no track is loaded. */
    const DEFAULT_SAMPLE_INTERVAL_SEC = 0.1;

    /**
     * @param {{ time: string }[]} points
     * @returns {number}
     */
    function medianSampleIntervalSec(points) {
        if (points.length < 2) return DEFAULT_SAMPLE_INTERVAL_SEC;

        const deltas = [];
        for (let i = 1; i < points.length; i++) {
            const t0 = Date.parse(points[i - 1].time);
            const t1 = Date.parse(points[i].time);
            if (!Number.isFinite(t0) || !Number.isFinite(t1)) continue;
            const dt = (t1 - t0) / 1000;
            if (dt > 0) deltas.push(dt);
        }

        if (!deltas.length) return DEFAULT_SAMPLE_INTERVAL_SEC;

        deltas.sort((a, b) => a - b);
        const mid = Math.floor(deltas.length / 2);
        return deltas.length % 2 ? deltas[mid] : (deltas[mid - 1] + deltas[mid]) / 2;
    }

    /**
     * @param {number} avgPoints
     * @param {number} sampleIntervalSec
     * @returns {number|null}
     */
    function averagingWindowDurationSec(avgPoints, sampleIntervalSec) {
        const n = Math.max(1, Math.min(20, Math.floor(avgPoints) || 1));
        if (n <= 1 || !Number.isFinite(sampleIntervalSec) || sampleIntervalSec <= 0) return null;
        return (n - 1) * sampleIntervalSec;
    }

    const SWOOP_WINDOW_SEC = 25;
    const STATIONARY_SEC = 2;
    const STATIONARY_RADIUS_M = 2;
    const STATIONARY_SPEED_MS = 1;
    const CURSOR_B_VELD_MS = 1;
    /** Dive angle from horizontal (deg) that marks cursor B after A. */
    const CURSOR_B_DIVE_ANGLE_DEG = 5.5;
    /** AGL (m) that cursor B must stay below after A. Matches whole-metre "2 m" rounding. */
    const CURSOR_B_AGL_M = 2.49;
    /** Place near-ground cursor B at this 1-based sample of a below-CURSOR_B_AGL_M run after A. 0 disables the near-ground candidate. */
    const CURSOR_B_ALT_TICKS = 2;
    /** Real-time dive-angle flattening (deg/s) that marks the start of the recovery arc. */
    const CURSOR_A_FLATTENING_DEG_S = 15;
    /** Later local velD peak may be this fraction of the window max and still count as the dive peak for cursor A. */
    const CURSOR_A_PEAK_FRACTION = 0.9;

    /**
     * Great-circle distance in metres.
     * @param {number} lat1
     * @param {number} lon1
     * @param {number} lat2
     * @param {number} lon2
     * @returns {number}
     */
    function haversineMeters(lat1, lon1, lat2, lon2) {
        const R = 6371000;
        const rad = Math.PI / 180;
        const dLat = (lat2 - lat1) * rad;
        const dLon = (lon2 - lon1) * rad;
        const sLat = Math.sin(dLat / 2);
        const sLon = Math.sin(dLon / 2);
        const h = sLat * sLat + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * sLon * sLon;
        return 2 * R * Math.asin(Math.min(1, Math.sqrt(h)));
    }

    /**
     * Horizontal distance (m) from a gate anchor to the projection of a point onto
     * the gate's local horizontal plane (ENU at the gate, vertical component dropped).
     *
     * @param {{ lat: *, lng: * }} gate
     * @param {number} lat
     * @param {number} lon
     * @returns {number}
     */
    function horizontalDistanceOnGatePlane(gate, lat, lon) {
        const gateLat = coordNum(gate.lat);
        const gateLng = coordNum(gate.lng);
        if (!Number.isFinite(gateLat) || !Number.isFinite(gateLng)) return NaN;
        if (!Number.isFinite(lat) || !Number.isFinite(lon)) return NaN;
        const rad = Math.PI / 180;
        const dLat = (lat - gateLat) * rad;
        const dLon = (lon - gateLng) * rad;
        const R = 6371000;
        const north = dLat * R;
        const east = dLon * R * Math.cos(gateLat * rad);
        return Math.hypot(east, north);
    }

    /**
     * 3D displacement in metres (horizontal haversine + altitude).
     * @param {{ lat?: number, lon?: number, hMSL?: number }} a
     * @param {{ lat?: number, lon?: number, hMSL?: number }} b
     * @returns {number}
     */
    function displacementMeters(a, b) {
        const dAlt = Math.abs((Number(a.hMSL) || 0) - (Number(b.hMSL) || 0));
        if (Number.isFinite(a.lat) && Number.isFinite(a.lon) && Number.isFinite(b.lat) && Number.isFinite(b.lon)) {
            return Math.hypot(haversineMeters(a.lat, a.lon, b.lat, b.lon), dAlt);
        }
        return NaN;
    }

    /**
     * @param {{ lat?: number, lon?: number, hMSL?: number, velN?: number, velE?: number, velD?: number }} origin
     * @param {{ lat?: number, lon?: number, hMSL?: number, velN?: number, velE?: number, velD?: number }} p
     * @returns {boolean}
     */
    function isStationaryRelativeTo(origin, p) {
        const disp = displacementMeters(origin, p);
        if (Number.isFinite(disp)) return disp <= STATIONARY_RADIUS_M;
        const speed = Math.hypot(Number(p.velN) || 0, Number(p.velE) || 0, Number(p.velD) || 0);
        return speed <= STATIONARY_SPEED_MS;
    }

    /**
     * Earliest timestamp at which position stays roughly still for STATIONARY_SEC.
     * Used as landing: analysis and the swoop window drop everything after this.
     *
     * @param {{ time: string, lat?: number, lon?: number, hMSL?: number, velN?: number, velE?: number, velD?: number }[]} points
     * @returns {number} epoch ms, or NaN if no 2 s still period exists
     */
    function findStationaryCutoffMs(points) {
        if (!points || points.length < 2) return NaN;
        const times = points.map(p => Date.parse(p.time));
        for (let i = 0; i < points.length; i++) {
            const t0 = times[i];
            if (!Number.isFinite(t0)) continue;
            let lastStill = t0;
            let still = true;
            for (let j = i; j < points.length; j++) {
                const tj = times[j];
                if (!Number.isFinite(tj)) continue;
                if (!isStationaryRelativeTo(points[i], points[j])) {
                    still = false;
                    break;
                }
                lastStill = tj;
                if ((lastStill - t0) / 1000 >= STATIONARY_SEC) break;
            }
            if (still && (lastStill - t0) / 1000 >= STATIONARY_SEC) return t0;
        }
        return NaN;
    }

    /**
     * Keep samples through the first 2 s still stretch (landing).
     * Drops the post-landing tail so GNSS wander or a walk cannot become
     * ground (0 m AGL) when the device is left on.
     * If no landing is found, returns the original array.
     *
     * @param {{ time: string }[]} points
     * @returns {{ time: string }[]}
     */
    function pointsUpToFirstLanding(points) {
        if (!points || points.length === 0) return points || [];
        const cutoff = findStationaryCutoffMs(points);
        if (!Number.isFinite(cutoff)) return points;
        const kept = points.filter(p => {
            const t = Date.parse(p.time);
            return Number.isFinite(t) && t <= cutoff;
        });
        return kept.length ? kept : points;
    }

    /**
     * Last valid timestamp in the track, or NaN.
     * @param {{ time: string }[]} points
     * @returns {number}
     */
    function lastValidTimeMs(points) {
        for (let i = points.length - 1; i >= 0; i--) {
            const t = Date.parse(points[i].time);
            if (Number.isFinite(t)) return t;
        }
        return NaN;
    }

    /**
     * Flight-path angle from the velocity vector (rad). Positive velD is downward.
     * @param {{ velN?: number, velE?: number, velD: number }} point
     * @returns {number}
     */
    function pathAngleRad(point) {
        const h = Math.hypot(Number(point.velN) || 0, Number(point.velE) || 0);
        return Math.atan2(point.velD, h);
    }

    /**
     * Dive angle from horizontal (deg). Positive is downward; 0 is level.
     * @param {{ velN?: number, velE?: number, velD: number, diveAngleDeg?: number }} point
     * @returns {number}
     */
    function diveAngleDeg(point) {
        if (Number.isFinite(point?.diveAngleDeg)) return point.diveAngleDeg;
        return pathAngleRad(point) * (180 / Math.PI);
    }

    /**
     * Signed path-angle rate in deg/s from `earlier` to `later` in real time.
     * Negative = flattening (dive angle decreasing).
     *
     * @param {{ velN?: number, velE?: number, velD: number }} earlier
     * @param {{ velN?: number, velE?: number, velD: number }} later
     * @param {number} dtSec
     * @returns {number}
     */
    function pathAngleFlatteningDegS(earlier, later, dtSec) {
        if (!Number.isFinite(dtSec) || dtSec <= 0) return 0;
        return (pathAngleRad(later) - pathAngleRad(earlier)) / dtSec * (180 / Math.PI);
    }

    /**
     * @param {{ velN?: number, velE?: number, velD: number }} a
     * @param {{ velN?: number, velE?: number, velD: number }} b
     * @param {number} dtSec
     * @returns {number} absolute path-angle rate in deg/s
     */
    function pathAngleRateDegS(a, b, dtSec) {
        return Math.abs(pathAngleFlatteningDegS(a, b, dtSec));
    }

    /**
     * Last `windowSec` before landing on the trimmed flight track.
     *
     * @param {{ time: string }[]} points
     * @param {number} [windowSec]
     * @returns {{
     *   windowed: { time: string, hMSL: number, velD: number, velN?: number, velE?: number }[],
     *   flight: { time: string }[],
     *   tEnd: number,
     *   error?: string
     * }}
     */
    function flightSwoopWindow(points, windowSec = SWOOP_WINDOW_SEC) {
        if (!points || points.length < 2) {
            return { windowed: [], flight: [], tEnd: NaN, error: 'Not enough track points.' };
        }
        const quality = filterPointsBySpeedAccuracy(points);
        const flight = pointsUpToFirstLanding(quality);
        if (flight.length < 2) {
            return { windowed: [], flight: [], tEnd: NaN, error: 'Not enough track points.' };
        }

        const tEnd = lastValidTimeMs(flight);
        if (!Number.isFinite(tEnd)) {
            return { windowed: [], flight, tEnd: NaN, error: 'Track times are invalid.' };
        }

        const windowed = [];
        for (let i = 0; i < flight.length; i++) {
            const t = Date.parse(flight[i].time);
            if (!Number.isFinite(t)) continue;
            const dt = (tEnd - t) / 1000;
            if (dt >= 0 && dt <= windowSec) windowed.push(flight[i]);
        }
        if (windowed.length < 2) {
            return { windowed: [], flight, tEnd, error: 'Not enough points in the swoop window.' };
        }

        return { windowed, flight, tEnd };
    }

    /**
     * Last `windowSec` before the first 2 s of roughly-still position (landing),
     * in chronological order (index 0 is earliest; the last sample is landing).
     * Points after the still period are dropped. `tRev` is seconds before landing.
     * `velD` is smoothed; `flatteningDegS` is signed dγ/dt from raw velocities
     * (negative = dive angle decreasing in real time).
     *
     * @param {{ time: string, hMSL: number, velD: number, velN?: number, velE?: number, lat?: number, lon?: number }[]} points
     * @param {number} [avgPoints]
     * @param {number} [windowSec]
     * @returns {{
     *   samples: { tRev: number, velD: number, velDRaw: number, velN: number, velE: number, hMSL: number, time: string, pitchRateDegS: number, flatteningDegS: number, diveAngleDeg: number }[],
     *   error?: string
     * }}
     */
    function buildSwoopCursorSeries(
        points,
        avgPoints = 3,
        windowSec = SWOOP_WINDOW_SEC
    ) {
        const { windowed, tEnd, error } = flightSwoopWindow(points, windowSec);
        if (error) {
            return { samples: [], error };
        }

        const velSmooth = movingAverage(windowed.map(p => p.velD), avgPoints);
        const samples = windowed.map((p, i) => {
            const t = Date.parse(p.time);
            const tRev = Number.isFinite(t) ? (tEnd - t) / 1000 : 0;
            const next = windowed[i + 1];
            let flatteningDegS = 0;
            if (next) {
                const tNext = Date.parse(next.time);
                const dt = (Number.isFinite(tNext) && Number.isFinite(t)) ? Math.abs(tNext - t) / 1000 : 0;
                flatteningDegS = pathAngleFlatteningDegS(p, next, dt);
            }
            const sample = {
                tRev,
                velD: velSmooth[i],
                velDRaw: p.velD,
                velN: Number.isFinite(p.velN) ? p.velN : 0,
                velE: Number.isFinite(p.velE) ? p.velE : 0,
                hMSL: Number.isFinite(p.hMSL) ? p.hMSL : NaN,
                time: p.time,
                flatteningDegS,
                pitchRateDegS: Math.abs(flatteningDegS),
                diveAngleDeg: diveAngleDeg(p)
            };
            if (Number.isFinite(p.lat) && Number.isFinite(p.lon)) {
                sample.lat = p.lat;
                sample.lon = p.lon;
            }
            return sample;
        });

        return { samples };
    }

    /**
     * Lowest finite hMSL in the series (ground reference for AGL).
     * @param {{ hMSL?: number }[]} samples
     * @returns {number}
     */
    function minHmsl(samples) {
        let min = Infinity;
        for (let i = 0; i < samples.length; i++) {
            const h = samples[i]?.hMSL;
            if (Number.isFinite(h) && h < min) min = h;
        }
        return min;
    }

    /**
     * @param {{ hMSL?: number } | null | undefined} sample
     * @param {number} groundHmsl
     * @returns {number}
     */
    function aglMeters(sample, groundHmsl) {
        if (!Number.isFinite(sample?.hMSL) || !Number.isFinite(groundHmsl)) return NaN;
        return sample.hMSL - groundHmsl;
    }

    /**
     * @param {{ diveAngleDeg?: number, velN?: number, velE?: number, velD: number }[]} samples
     * @param {number} idxA
     * @param {number} angleThresh
     * @returns {number} index, or -1 if none
     */
    function cursorBFromDiveAngle(samples, idxA, angleThresh) {
        for (let i = idxA + 1; i < samples.length; i++) {
            if (diveAngleDeg(samples[i]) < angleThresh) return i;
        }
        return -1;
    }

    /**
     * Near-ground cursor B: the `ticks`-th consecutive sample below 2.49 m AGL
     * after A toward landing in chronological order, not the first sample that
     * drops below 2.49 m. A run shorter than `ticks` (brief dip) is ignored.
     * `ticks` < 1 disables this candidate.
     *
     * @param {{ hMSL?: number }[]} samples
     * @param {number} idxA
     * @param {number} ticks
     * @param {number} ground
     * @returns {number} index, or -1 if none
     */
    function cursorBFromAltTicks(samples, idxA, ticks, ground) {
        if (!(ticks >= 1)) return -1;
        let belowRun = 0;
        for (let i = idxA + 1; i < samples.length; i++) {
            const agl = aglMeters(samples[i], ground);
            if (Number.isFinite(agl) && agl < CURSOR_B_AGL_M) {
                belowRun++;
                if (belowRun >= ticks) return i;
            } else {
                belowRun = 0;
            }
        }
        return -1;
    }

    /**
     * Index of the highest velD in the series (first max if tied).
     *
     * @param {{ velD: number }[]} samples
     * @returns {number}
     */
    function maxVelDIdx(samples) {
        if (!samples || samples.length === 0) return 0;
        let idx = 0;
        for (let i = 1; i < samples.length; i++) {
            if (samples[i].velD > samples[idx].velD) idx = i;
        }
        return idx;
    }

    /**
     * Last (closest-to-landing) local velD maximum that is still within
     * CURSOR_A_PEAK_FRACTION of the window max. Chronological series: walking
     * from landing backward, the first such peak is last in real time.
     *
     * @param {{ velD: number }[]} samples
     * @returns {number}
     */
    function lastSignificantVelDPeakIdx(samples) {
        if (!samples || samples.length === 0) return 0;
        const globalIdx = maxVelDIdx(samples);
        const globalMax = samples[globalIdx]?.velD;
        if (!Number.isFinite(globalMax) || samples.length < 3) return globalIdx;
        const minPeak = globalMax * CURSOR_A_PEAK_FRACTION;
        for (let i = samples.length - 2; i >= 1; i--) {
            const v = samples[i].velD;
            if (v >= samples[i - 1].velD && v >= samples[i + 1].velD && v >= minPeak) {
                return i;
            }
        }
        return globalIdx;
    }

    /**
     * Entry gates for swoop analysis: distance-marked gates when present,
     * otherwise gates whose label contains "entry", otherwise all gates.
     *
     * @param {{ label?: string }[]} gates
     * @returns {{ label?: string, lat: number, lng: number }[]}
     */
    function entryGatesForSwoop(gates) {
        if (!gates?.length) return [];
        const distance = gates.filter((g) => /\bdistance\b/i.test(String(g.label || '')));
        if (distance.length) return distance;
        const entry = gates.filter((g) => /\bentry\b/i.test(String(g.label || '')));
        if (entry.length) return entry;
        return gates;
    }

    /**
     * @param {{ lat?: number, lon?: number }[]} samples
     * @param {{ lat: number, lng: number }} gate
     * @param {number} fromIdx
     * @param {number} toIdx
     * @returns {{ idx: number, distM: number }}
     */
    function closestSampleIdxToGate(samples, gate, fromIdx, toIdx) {
        const gateLat = coordNum(gate.lat);
        const gateLng = coordNum(gate.lng);
        if (!Number.isFinite(gateLat) || !Number.isFinite(gateLng)) {
            return { idx: -1, distM: Infinity };
        }
        let bestIdx = -1;
        let bestDist = Infinity;
        const start = Math.max(0, fromIdx);
        const end = Math.min(samples.length - 1, toIdx);
        for (let i = start; i <= end; i++) {
            const s = samples[i];
            if (!pointHasGps(s)) continue;
            const d = haversineMeters(s.lat, s.lon, gateLat, gateLng);
            if (d < bestDist || (d === bestDist && i > bestIdx)) {
                bestDist = d;
                bestIdx = i;
            }
        }
        return { idx: bestIdx, distM: bestDist };
    }

    /**
     * Among entry gates, the one whose closest approach to `samples` is shortest.
     *
     * @param {{ lat?: number, lon?: number }[]} samples
     * @param {{ label?: string, lat: number, lng: number }[]} gates
     * @param {number} fromIdx
     * @param {number} toIdx
     * @returns {{ gate?: { label?: string, lat: number, lng: number }, idx: number, distM: number }}
     */
    function closestEntryGatePassOnSeries(samples, gates, fromIdx, toIdx) {
        const entryGates = entryGatesForSwoop(gates);
        let bestGate = null;
        let bestIdx = -1;
        let bestDist = Infinity;
        for (const gate of entryGates) {
            const pass = closestSampleIdxToGate(samples, gate, fromIdx, toIdx);
            if (pass.idx < 0 || !Number.isFinite(pass.distM)) continue;
            if (pass.distM < bestDist) {
                bestDist = pass.distM;
                bestGate = gate;
                bestIdx = pass.idx;
            }
        }
        if (!bestGate || bestIdx < 0) return { idx: -1, distM: Infinity };
        return { gate: bestGate, idx: bestIdx, distM: bestDist };
    }

    /**
     * Cursor B for graphs: the sample immediately before the closest approach
     * to the nearest entry gate (from A toward landing).
     *
     * @param {{ lat?: number, lon?: number }[]} samples
     * @param {number} idxA
     * @param {{ label?: string, lat: number, lng: number }[]} gates
     * @returns {number} index, or -1 if it cannot be computed
     */
    function cursorBBeforeClosestEntryGate(samples, idxA, gates) {
        if (!samples?.length || !gates?.length) return -1;
        const last = samples.length - 1;
        const fromIdx = Math.min(last, Math.max(0, idxA) + 1);
        const pass = closestEntryGatePassOnSeries(samples, gates, fromIdx, last);
        if (pass.idx < fromIdx) return -1;
        const idxB = pass.idx - 1;
        if (idxB <= idxA) return Math.min(last, idxA + 1);
        return idxB;
    }

    /**
     * Default cursor A on a chronological swoop series (last = landing).
     *
     * @param {{ velD: number, flatteningDegS?: number }[]} samples
     * @returns {{ idxA: number, peakVelD: number }}
     */
    function defaultSwoopCursorIdxA(samples) {
        if (!samples || samples.length === 0) {
            return { idxA: 0, peakVelD: 0 };
        }
        const last = samples.length - 1;
        const flattenThresh = -CURSOR_A_FLATTENING_DEG_S;

        const peakIdx = lastSignificantVelDPeakIdx(samples);
        let peakVelD = samples[peakIdx]?.velD;
        if (!Number.isFinite(peakVelD) || peakVelD < 0) peakVelD = 0;

        const flatteningAt = (i) => {
            const v = samples[i]?.flatteningDegS;
            return Number.isFinite(v) ? v : 0;
        };

        let idxA = peakIdx;
        for (let i = peakIdx + 1; i <= last - 1; i++) {
            if (flatteningAt(i) > flattenThresh) continue;
            const later = i + 1 <= last ? flatteningAt(i + 1) : flatteningAt(i);
            if (later > flattenThresh) continue;
            idxA = i;
            break;
        }
        if (idxA >= last) idxA = Math.max(0, last - 1);
        if (idxA < 0) idxA = 0;
        return { idxA, peakVelD };
    }

    /**
     * @param {{ velD: number, flatteningDegS?: number, diveAngleDeg?: number, hMSL?: number }[]} samples
     * @param {number} idxA
     * @param {number} [diveAngleDegThresh]
     * @param {number} [altTicks]
     * @returns {number}
     */
    function defaultSwoopCursorIdxB(
        samples,
        idxA,
        diveAngleDegThresh = CURSOR_B_DIVE_ANGLE_DEG,
        altTicks = CURSOR_B_ALT_TICKS
    ) {
        if (!samples?.length) return 0;
        const last = samples.length - 1;
        const angleThresh = Number.isFinite(diveAngleDegThresh)
            ? diveAngleDegThresh
            : CURSOR_B_DIVE_ANGLE_DEG;
        const ticks = Number.isFinite(altTicks)
            ? Math.max(0, Math.floor(altTicks))
            : CURSOR_B_ALT_TICKS;
        const ground = minHmsl(samples);

        const idxBAngle = angleThresh > 0
            ? cursorBFromDiveAngle(samples, idxA, angleThresh)
            : -1;
        const idxBAlt = ticks > 0
            ? cursorBFromAltTicks(samples, idxA, ticks, ground)
            : -1;
        let idxB;
        if (idxBAngle >= 0 && idxBAlt >= 0) idxB = Math.max(idxBAngle, idxBAlt);
        else if (idxBAngle >= 0) idxB = idxBAngle;
        else if (idxBAlt >= 0) idxB = idxBAlt;
        else idxB = Math.min(last, idxA + 1);
        if (idxB <= idxA) idxB = Math.min(last, idxA + 1);
        return idxB;
    }

    /**
     * Default flare-window cursors on a chronological series (last = landing).
     * A: after the last near-max velD peak, the first sample where dive angle
     *    is flattening strongly (and the next sample toward landing confirms it).
     *    A later peak that is still within CURSOR_A_PEAK_FRACTION of the window
     *    max wins over an earlier, slightly taller one.
     * B: walking from A toward landing in real time. When both the dive-angle
     *    threshold and `altTicks` are non-zero, B is the later of:
     *    - the first sample whose dive angle is below `diveAngleDegThresh`
     *      (default 5.5° from horizontal), and
     *    - the `altTicks`-th consecutive sample below 2.49 m AGL.
     *    Setting either value to 0 disables that condition so B uses only the other.
     *
     * @param {{ velD: number, flatteningDegS?: number, diveAngleDeg?: number, hMSL?: number, velN?: number, velE?: number }[]} samples
     * @param {number} [diveAngleDegThresh]
     * @param {number} [altTicks]
     * @returns {{ idxA: number, idxB: number, peakVelD: number }}
     */
    function defaultSwoopCursorIndices(
        samples,
        diveAngleDegThresh = CURSOR_B_DIVE_ANGLE_DEG,
        altTicks = CURSOR_B_ALT_TICKS
    ) {
        if (!samples || samples.length === 0) {
            return { idxA: 0, idxB: 0, peakVelD: 0 };
        }
        const { idxA, peakVelD } = defaultSwoopCursorIdxA(samples);
        const idxB = defaultSwoopCursorIdxB(samples, idxA, diveAngleDegThresh, altTicks);
        return { idxA, idxB, peakVelD };
    }

    /**
     * Graph modal cursors: same A as default; B is the sample before the
     * closest pass of the nearest entry gate when gates are available.
     *
     * @param {{ velD: number, flatteningDegS?: number, diveAngleDeg?: number, hMSL?: number, lat?: number, lon?: number }[]} samples
     * @param {{ label?: string, lat: number, lng: number }[]} [gates]
     * @param {number} [diveAngleDegThresh]
     * @param {number} [altTicks]
     * @returns {{ idxA: number, idxB: number, peakVelD: number }}
     */
    function graphSwoopCursorIndices(
        samples,
        gates,
        diveAngleDegThresh = CURSOR_B_DIVE_ANGLE_DEG,
        altTicks = CURSOR_B_ALT_TICKS
    ) {
        if (!samples || samples.length === 0) {
            return { idxA: 0, idxB: 0, peakVelD: 0 };
        }
        const { idxA, peakVelD } = defaultSwoopCursorIdxA(samples);
        let idxB = cursorBBeforeClosestEntryGate(samples, idxA, gates);
        if (idxB < 0) {
            idxB = defaultSwoopCursorIdxB(samples, idxA, diveAngleDegThresh, altTicks);
        }
        return { idxA, idxB, peakVelD };
    }

    /**
     * Total trajectory speed (km/h) and AGL (m) at graph cursor B — same B placement as the graph modal.
     *
     * @param {{ time: string, hMSL: number, velD: number, velN?: number, velE?: number, lat?: number, lon?: number }[]} points
     * @param {number} [avgPoints]
     * @param {{ label?: string, lat: number, lng: number }[]} [gates]
     * @param {number} [diveAngleDegThresh]
     * @param {number} [altTicks]
     * @returns {{ speedKmh: number, altitudeM: number }}
     */
    function entrySpeedAtGraphCursorB(
        points,
        avgPoints = 3,
        gates,
        diveAngleDegThresh = CURSOR_B_DIVE_ANGLE_DEG,
        altTicks = CURSOR_B_ALT_TICKS
    ) {
        const series = buildSwoopCursorSeries(points, avgPoints);
        if (series.error || !series.samples.length) {
            return { speedKmh: NaN, altitudeM: NaN };
        }
        const { idxB } = graphSwoopCursorIndices(
            series.samples,
            gates,
            diveAngleDegThresh,
            altTicks
        );
        const sampleB = series.samples[idxB];
        if (!sampleB) return { speedKmh: NaN, altitudeM: NaN };
        const ground = minHmsl(series.samples);
        const ms = trajectorySpeedMs(sampleB);
        const speedKmh = Number.isFinite(ms) ? ms * 3.6 : NaN;
        const agl = aglMeters(sampleB, ground);
        const altitudeM = Number.isFinite(agl) ? Math.max(0, agl) : NaN;
        return { speedKmh, altitudeM };
    }

    /**
     * Seconds from cursor B to the start of the
     * 2 s stationary window (tRev = 0).
     *
     * @param {{ tRev?: number } | null | undefined} sampleB
     * @returns {number}
     */
    function timeAloftSec(sampleB) {
        const t = sampleB?.tRev;
        if (!Number.isFinite(t) || t <= 0) return 0;
        return t;
    }

    /**
     * Seconds between default swoop cursors A and B (recovery arc).
     * Same value as A-B Time on the graph with default cursor placement.
     *
     * @param {{ time: string, hMSL: number, velD: number, velN?: number, velE?: number, lat?: number, lon?: number }[]} points
     * @param {number} [avgPoints]
     * @param {number} [diveAngleDegThresh]
     * @param {number} [altTicks]
     * @returns {number} duration in seconds, or NaN if it cannot be computed
     */
    function recoveryArcSec(
        points,
        avgPoints = 3,
        diveAngleDegThresh = CURSOR_B_DIVE_ANGLE_DEG,
        altTicks = CURSOR_B_ALT_TICKS
    ) {
        const series = buildSwoopCursorSeries(points, avgPoints);
        if (series.error || !series.samples.length) return NaN;
        const { idxA, idxB } = defaultSwoopCursorIndices(
            series.samples,
            diveAngleDegThresh,
            altTicks
        );
        const a = series.samples[idxA];
        const b = series.samples[idxB];
        if (!a || !b) return NaN;
        const dt = Math.abs(b.tRev - a.tRev);
        return Number.isFinite(dt) ? dt : NaN;
    }

    /**
     * @param {number} sec
     * @returns {string}
     */
    function formatDurationSec(sec) {
        if (!Number.isFinite(sec) || sec <= 0) return '';
        if (sec < 0.01) return `${sec.toFixed(3)}s`;
        if (sec < 1) return `${sec.toFixed(2)}s`;
        if (sec < 10) return `${sec.toFixed(1)}s`;
        return `${Math.round(sec)}s`;
    }

    /**
     * Centered moving average; window shrinks near edges.
     * @param {number[]} values
     * @param {number} windowSize
     * @returns {number[]}
     */
    function movingAverage(values, windowSize) {
        const n = Math.max(1, Math.min(20, Math.floor(windowSize) || 1));
        if (values.length === 0) return [];
        const out = new Array(values.length);
        const half = Math.floor(n / 2);

        for (let i = 0; i < values.length; i++) {
            let start = Math.max(0, i - half);
            let end = Math.min(values.length, start + n);
            start = Math.max(0, end - n);
            let sum = 0;
            for (let j = start; j < end; j++) sum += values[j];
            out[i] = sum / (end - start);
        }
        return out;
    }

    const DEFAULT_SPEED_METRIC = 'vertical';
    /** Drop GNSS samples whose speed-accuracy estimate exceeds this (m/s). Missing sAcc is kept. */
    const MAX_SPEED_ACCURACY_MS = 2;

    /**
     * @param {{ velN?: number, velE?: number, velD: number }} point
     * @returns {number}
     */
    function trajectorySpeedMs(point) {
        return Math.hypot(point.velN, point.velE, point.velD);
    }

    /**
     * @param {string} speedMetric
     * @returns {'vertical' | 'total' | 'both'}
     */
    function normalizeSpeedMetric(speedMetric) {
        if (speedMetric === 'total') return 'total';
        if (speedMetric === 'both') return 'both';
        return 'vertical';
    }

    /**
     * @param {{ velN?: number, velE?: number, velD: number }[]} eligible
     * @param {'vertical' | 'total'} mode
     * @returns {number[]}
     */
    function rawSpeedsForMode(eligible, mode) {
        if (mode === 'total') {
            return eligible.map(p => {
                if (!Number.isFinite(p.velN) || !Number.isFinite(p.velE)) return NaN;
                return trajectorySpeedMs(p);
            });
        }
        return eligible.map(p => Math.abs(p.velD));
    }

    /**
     * @param {{ hMSL: number, time: string }[]} eligible
     * @param {number} minHmsl
     * @param {number} avgPoints
     * @param {number[]} rawSpeeds
     * @returns {{ maxSpeedKmh: number, altitudeM: number, time: string }}
     */
    function peakFromRawSpeeds(eligible, minHmsl, avgPoints, rawSpeeds) {
        const windowSize = Math.max(1, Math.min(20, Math.floor(avgPoints) || 1));
        const smoothedVel = movingAverage(rawSpeeds, windowSize);
        const smoothedAlt = movingAverage(eligible.map(p => p.hMSL), windowSize);

        let peakIdx = 0;
        for (let i = 1; i < smoothedVel.length; i++) {
            if (smoothedVel[i] > smoothedVel[peakIdx]) peakIdx = i;
        }

        return {
            maxSpeedKmh: smoothedVel[peakIdx] * 3.6,
            altitudeM: Math.max(0, smoothedAlt[peakIdx] - minHmsl),
            time: eligible[peakIdx].time
        };
    }

    /**
     * Drop points with a reported speed accuracy worse than `maxSAccMs`.
     * Points with missing/non-finite `sAcc` are kept. If every point would be
     * dropped, return the original array so analysis can still run.
     *
     * @param {{ sAcc?: number }[]} points
     * @param {number} [maxSAccMs]
     * @returns {{ sAcc?: number }[]}
     */
    function filterPointsBySpeedAccuracy(points, maxSAccMs = MAX_SPEED_ACCURACY_MS) {
        if (!points || points.length === 0) return [];
        const limit = Number.isFinite(maxSAccMs) && maxSAccMs > 0 ? maxSAccMs : MAX_SPEED_ACCURACY_MS;
        const kept = points.filter(p => !Number.isFinite(p.sAcc) || p.sAcc <= limit);
        return kept.length ? kept : points;
    }

    /**
     * Max vertical speed from graph series (smoothed signed velD, same as the vertical-speed graph).
     *
     * @param {{ velD: number, hMSL?: number, time: string }[]} samples
     * @returns {{ maxSpeedKmh: number, altitudeM: number, time: string }}
     */
    function peakVerticalFromGraphSamples(samples) {
        const peakIdx = maxVelDIdx(samples);
        const ground = minHmsl(samples);
        const s = samples[peakIdx];
        return {
            maxSpeedKmh: s.velD * 3.6,
            altitudeM: Math.max(0, aglMeters(s, ground)),
            time: s.time
        };
    }

    /**
     * Post-landing samples (after the first 2 s still stretch) are ignored so
     * a device left on cannot pull ground / AGL down.
     *
     * @param {{ hMSL: number, velD: number }[]} points
     * @param {number} avgPoints
     * @param {'vertical' | 'total' | 'both'} speedMetric
     * @returns {{
     *   maxVerticalSpeedKmh: number,
     *   maxTotalSpeedKmh?: number,
     *   altitudeM: number,
     *   time: string,
     *   totalPeakTime?: string,
     *   totalPeakAltitudeM?: number,
     *   pointCount: number,
     *   minHmsl: number,
     *   speedMetric: 'vertical' | 'total' | 'both',
     *   error?: string
     * }}
     */
    function analyzeFlysightTrack(
        points,
        avgPoints = 1,
        speedMetric = DEFAULT_SPEED_METRIC
    ) {
        const metric = normalizeSpeedMetric(speedMetric);
        const emptyResult = (overrides = {}) => ({
            maxVerticalSpeedKmh: 0,
            altitudeM: 0,
            time: '',
            pointCount: 0,
            minHmsl: 0,
            speedMetric: metric,
            ...overrides
        });

        if (!points.length) {
            return emptyResult({ error: 'No track points.' });
        }

        const { windowed, error: windowError } = flightSwoopWindow(points);
        if (windowError) {
            return emptyResult({ error: windowError });
        }

        const { samples, error: seriesError } = buildSwoopCursorSeries(points, avgPoints);
        if (seriesError || !samples.length) {
            return emptyResult({ error: seriesError || 'Could not build swoop series.' });
        }

        const groundHmsl = minHmsl(samples);

        if (metric === 'both') {
            const totalRaw = rawSpeedsForMode(windowed, 'total');
            if (totalRaw.some(s => !Number.isFinite(s))) {
                return emptyResult({
                    minHmsl: groundHmsl,
                    error: 'Missing velN/velE columns required for total speed.'
                });
            }

            const verticalPeak = peakVerticalFromGraphSamples(samples);
            const totalPeak = peakFromRawSpeeds(windowed, groundHmsl, avgPoints, totalRaw);

            return {
                maxVerticalSpeedKmh: verticalPeak.maxSpeedKmh,
                maxTotalSpeedKmh: totalPeak.maxSpeedKmh,
                altitudeM: verticalPeak.altitudeM,
                time: verticalPeak.time,
                totalPeakTime: totalPeak.time,
                totalPeakAltitudeM: totalPeak.altitudeM,
                pointCount: samples.length,
                minHmsl: groundHmsl,
                speedMetric: 'both'
            };
        }

        if (metric === 'total') {
            const totalRaw = rawSpeedsForMode(windowed, 'total');
            if (totalRaw.some(s => !Number.isFinite(s))) {
                return emptyResult({
                    minHmsl: groundHmsl,
                    error: 'Missing velN/velE columns required for total speed.'
                });
            }
            const totalPeak = peakFromRawSpeeds(windowed, groundHmsl, avgPoints, totalRaw);
            return {
                maxVerticalSpeedKmh: totalPeak.maxSpeedKmh,
                altitudeM: totalPeak.altitudeM,
                time: totalPeak.time,
                pointCount: samples.length,
                minHmsl: groundHmsl,
                speedMetric: metric
            };
        }

        const verticalPeak = peakVerticalFromGraphSamples(samples);

        return {
            maxVerticalSpeedKmh: verticalPeak.maxSpeedKmh,
            altitudeM: verticalPeak.altitudeM,
            time: verticalPeak.time,
            pointCount: samples.length,
            minHmsl: groundHmsl,
            speedMetric: metric
        };
    }

    /**
     * @param {string} text
     * @param {number} avgPoints
     * @param {'vertical' | 'total' | 'both'} speedMetric
     * @returns {ReturnType<typeof analyzeFlysightTrack> & { points: typeof points }}
     */
    function analyzeFlysightCsv(
        text,
        avgPoints = 1,
        speedMetric = DEFAULT_SPEED_METRIC
    ) {
        const metric = normalizeSpeedMetric(speedMetric);
        const parsed = parseFlysightCsv(text);
        if (parsed.error) {
            return {
                points: [],
                maxVerticalSpeedKmh: 0,
                altitudeM: 0,
                time: '',
                pointCount: 0,
                minHmsl: 0,
                speedMetric: metric,
                error: parsed.error
            };
        }
        const result = analyzeFlysightTrack(parsed.points, avgPoints, metric);
        return { ...result, points: parsed.points };
    }

    const LOCATION_MATCH_MAX_KM = 15;

    /**
     * Local maximum of dive angle when walking backward from cursor A on the graph series.
     *
     * @param {{ diveAngleDeg?: number }[]} samples
     * @param {number} idxA
     * @returns {number}
     */
    function diveAngleLocalPeakIdxBeforeA(samples, idxA) {
        if (!samples?.length) return 0;
        const last = samples.length - 1;
        let a = Math.floor(idxA);
        if (!Number.isFinite(a)) a = last;
        a = Math.max(0, Math.min(a, last));

        const angleAt = (i) => {
            const v = diveAngleDeg(samples[i]);
            return Number.isFinite(v) ? v : -Infinity;
        };

        for (let i = a; i >= 1; i--) {
            const cur = angleAt(i);
            const prev = angleAt(i - 1);
            const next = angleAt(i + 1);
            if (cur >= prev && cur >= next && (cur > prev || cur > next)) return i;
        }
        if (a >= 0 && a <= last) return a;
        return 0;
    }

    /**
     * Closest flight point to any gate on the list.
     *
     * @param {{ lat?: number, lon?: number }[]} flightPoints
     * @param {{ label?: string, lat: number, lng: number }[]} gates
     * @returns {{ point?: { lat: number, lon: number }, gate?: { label?: string, lat: number, lng: number }, gateLabel?: string, minDistM?: number }}
     */
    function closestGatePassOnTrack(flightPoints, gates) {
        if (!flightPoints?.length || !gates?.length) return {};

        let minDist = Infinity;
        let nearestGate = null;
        let point = null;

        for (const gate of gates) {
            const gateLat = coordNum(gate.lat);
            const gateLng = coordNum(gate.lng);
            if (!Number.isFinite(gateLat) || !Number.isFinite(gateLng)) continue;
            for (const p of flightPoints) {
                if (!pointHasGps(p)) continue;
                const d = haversineMeters(p.lat, p.lon, gateLat, gateLng);
                if (d < minDist) {
                    minDist = d;
                    nearestGate = gate;
                    point = p;
                }
            }
        }

        if (!point || !nearestGate || !Number.isFinite(minDist)) return {};
        return {
            point: { lat: point.lat, lon: point.lon },
            gate: nearestGate,
            gateLabel: nearestGate.label,
            minDistM: minDist
        };
    }

    /**
     * Horizontal distance from one point to the closest gate.
     *
     * @param {number} lat
     * @param {number} lon
     * @param {{ label?: string, lat: number, lng: number }[]} gates
     * @returns {{ gateLabel?: string, distM?: number }}
     */
    function nearestGateToPoint(lat, lon, gates) {
        if (!Number.isFinite(lat) || !Number.isFinite(lon) || !gates?.length) return {};

        let minDist = Infinity;
        let nearestGate = null;
        for (const gate of gates) {
            const gateLat = coordNum(gate.lat);
            const gateLng = coordNum(gate.lng);
            if (!Number.isFinite(gateLat) || !Number.isFinite(gateLng)) continue;
            const d = haversineMeters(lat, lon, gateLat, gateLng);
            if (d < minDist) {
                minDist = d;
                nearestGate = gate;
            }
        }

        if (!nearestGate || !Number.isFinite(minDist)) return {};
        return {
            gateLabel: nearestGate.label,
            distM: minDist
        };
    }

    /**
     * Gates for the dropzone matched to this track (saved or seed).
     *
     * @param {{ lat?: number, lon?: number }[]} points
     * @param {{ id: string, gates?: object[] }[]} locations
     * @param {Record<string, object[]>} [seedByLocationId]
     * @param {number} [maxKm]
     * @returns {{ label?: string, lat: number, lng: number }[]}
     */
    function gatesForMatchedTrack(points, locations, seedByLocationId, maxKm) {
        const quality = filterPointsBySpeedAccuracy(points || []);
        const flight = pointsUpToFirstLanding(quality).filter(pointHasGps);
        if (flight.length < 2) return [];

        const match = matchLocationForTrack(flight, locations, maxKm, seedByLocationId);
        if (!match) return [];

        let location = (locations || []).find((l) => l.id === match.locationId);
        if (!location && match.locationId) {
            location = {
                id: match.locationId,
                name: match.name || match.locationId,
                gates: [],
                lat: null,
                lng: null
            };
        }
        return effectiveGatesForLocation(location, seedByLocationId);
    }

    /**
     * @param {{ lat?: number, lon?: number }} p
     * @returns {boolean}
     */
    function pointHasGps(p) {
        return Number.isFinite(p?.lat) && Number.isFinite(p?.lon);
    }

    /**
     * @param {*} n
     * @returns {number}
     */
    function coordNum(n) {
        const v = Number(n);
        return Number.isFinite(v) ? v : NaN;
    }

    /**
     * @param {object[]} seedEntries
     * @returns {{ label?: string, lat: number, lng: number }[]}
     */
    function gatesFromSeedEntriesInline(seedEntries) {
        const U = global.LocationGatesUtil;
        if (U && typeof U.gatesFromSeedEntries === 'function') {
            return U.gatesFromSeedEntries(seedEntries);
        }
        return seedEntries.map((entry) => ({
            label: entry.label,
            lat: entry.lat,
            lng: entry.lng
        }));
    }

    /**
     * @param {object} location
     * @param {Record<string, object[]>} [seedByLocationId]
     * @returns {{ label?: string, lat: number, lng: number }[]}
     */
    function effectiveGatesForLocation(location, seedByLocationId) {
        if (!location) return [];
        if (Array.isArray(location.gates) && location.gates.length) {
            return location.gates;
        }
        let seedEntries = location.id && seedByLocationId
            ? seedByLocationId[location.id]
            : null;
        if (!Array.isArray(seedEntries) || !seedEntries.length) return [];
        return gatesFromSeedEntriesInline(seedEntries);
    }

    /**
     * Min horizontal distance from any flight point to any gate on the location (saved or seed).
     *
     * @param {{ lat?: number, lon?: number }[]} points
     * @param {{ id: string, lat?: *, lng?: *, gates?: object[] }} loc
     * @param {Record<string, object[]>} [seedByLocationId]
     * @returns {{ minDistM: number, via: 'gate' | null }}
     */
    function minTrackDistToLocation(points, loc, seedByLocationId) {
        let minDist = Infinity;
        if (!loc?.id) return { minDistM: Infinity, via: null };

        const gates = effectiveGatesForLocation(loc, seedByLocationId);
        for (const gate of gates) {
            const gateLat = coordNum(gate.lat);
            const gateLng = coordNum(gate.lng);
            if (!Number.isFinite(gateLat) || !Number.isFinite(gateLng)) continue;
            for (const p of points) {
                if (!pointHasGps(p)) continue;
                const d = haversineMeters(p.lat, p.lon, gateLat, gateLng);
                if (d < minDist) minDist = d;
            }
        }

        return {
            minDistM: minDist,
            via: Number.isFinite(minDist) && minDist < Infinity ? 'gate' : null
        };
    }

    /**
     * Nearest location by min track distance to any gate (saved or seed).
     *
     * @param {{ lat?: number, lon?: number }[]} points
     * @param {{ id: string, name?: string, lat?: number, lng?: number, gates?: object[] }[]} locations
     * @param {number} [maxKm]
     * @param {Record<string, object[]>} [seedByLocationId]
     * @returns {{ locationId: string, name: string, minDistM: number, matchVia?: 'gate' } | null}
     */
    function matchLocationForTrack(points, locations, maxKm = LOCATION_MATCH_MAX_KM, seedByLocationId) {
        if (!points?.length) return null;
        const maxM = Number.isFinite(maxKm) && maxKm > 0 ? maxKm * 1000 : LOCATION_MATCH_MAX_KM * 1000;

        const byId = new Map();
        for (const loc of locations || []) {
            if (loc?.id) byId.set(loc.id, loc);
        }
        if (seedByLocationId) {
            for (const seedId of Object.keys(seedByLocationId)) {
                if (!byId.has(seedId)) {
                    byId.set(seedId, { id: seedId, name: seedId, gates: [], lat: null, lng: null });
                }
            }
        }
        if (!byId.size) return null;

        let best = null;
        let bestDist = Infinity;
        let bestVia = null;

        for (const loc of byId.values()) {
            const { minDistM, via } = minTrackDistToLocation(points, loc, seedByLocationId);
            if (!Number.isFinite(minDistM)) continue;
            if (minDistM < bestDist) {
                bestDist = minDistM;
                best = loc;
                bestVia = via;
            }
        }

        if (!best || bestDist > maxM) return null;
        return {
            locationId: best.id,
            name: String(best.name || best.id),
            minDistM: bestDist,
            matchVia: bestVia || undefined
        };
    }

    /**
     * Horizontal distance from the matched gate to the dive-angle peak (snap), measured
     * on the gate's local horizontal plane (snap position projected vertically onto that plane).
     *
     * @param {{ lat?: number, lon?: number, time?: string }[]} points
     * @param {{ id: string, name?: string, lat?: number, lng?: number, gates?: object[] }[]} locations
     * @param {Record<string, object[]>} [seedByLocationId]
     * @param {{
     *   maxKm?: number,
     *   avgPoints?: number,
     *   diveAngleDegThresh?: number,
     *   altTicks?: number
     * }} [options]
     * @returns {{
     *   snapDistanceM?: number,
     *   gateLabel?: string,
     *   gateMinDistM?: number,
     *   divePeakAngleDeg?: number,
     *   idxA?: number,
     *   divePeakIdx?: number,
     *   locationId?: string,
     *   locationName?: string,
     *   reason?: 'no-gps' | 'no-location' | 'no-gates' | 'no-swoop-window' | 'no-gate-pass' | 'no-dive-peak-gps'
     * }}
     */
    function computeGateSnapDistance(points, locations, seedByLocationId, options = {}) {
        const quality = filterPointsBySpeedAccuracy(points || []);
        const flight = pointsUpToFirstLanding(quality).filter(pointHasGps);

        if (flight.length < 2) {
            return { reason: 'no-gps' };
        }

        const match = matchLocationForTrack(
            flight,
            locations || [],
            options.maxKm,
            seedByLocationId
        );
        if (!match) {
            return { reason: 'no-location' };
        }

        let location = (locations || []).find((l) => l.id === match.locationId);
        if (!location && match.locationId) {
            location = {
                id: match.locationId,
                name: match.name || match.locationId,
                gates: [],
                lat: null,
                lng: null
            };
        }
        const gates = effectiveGatesForLocation(location, seedByLocationId);
        if (!gates.length) {
            return {
                reason: 'no-gates',
                locationId: match.locationId,
                locationName: match.name
            };
        }

        const avgPoints = Number.isFinite(options.avgPoints) ? options.avgPoints : 3;
        const series = buildSwoopCursorSeries(quality, avgPoints);
        if (series.error || !series.samples.length) {
            return {
                reason: 'no-swoop-window',
                locationId: match.locationId,
                locationName: match.name
            };
        }

        const { idxA } = defaultSwoopCursorIndices(
            series.samples,
            options.diveAngleDegThresh,
            options.altTicks
        );
        const divePeakIdx = diveAngleLocalPeakIdxBeforeA(series.samples, idxA);
        const peakSample = series.samples[divePeakIdx];
        if (!peakSample || !pointHasGps(peakSample)) {
            return {
                reason: 'no-dive-peak-gps',
                locationId: match.locationId,
                locationName: match.name,
                idxA,
                divePeakIdx
            };
        }

        const gatePass = closestGatePassOnTrack(flight, gates);
        if (!gatePass.point || !Number.isFinite(gatePass.minDistM)) {
            return {
                reason: 'no-gate-pass',
                locationId: match.locationId,
                locationName: match.name
            };
        }

        const snapDistanceM = horizontalDistanceOnGatePlane(
            gatePass.gate,
            peakSample.lat,
            peakSample.lon
        );
        if (!Number.isFinite(snapDistanceM)) {
            return {
                reason: 'no-gate-pass',
                locationId: match.locationId,
                locationName: match.name
            };
        }

        return {
            snapDistanceM,
            gateLabel: gatePass.gateLabel,
            gateMinDistM: gatePass.minDistM,
            divePeakAngleDeg: diveAngleDeg(peakSample),
            idxA,
            divePeakIdx,
            locationId: match.locationId,
            locationName: match.name
        };
    }

    /**
     * @param {{ name?: string, type?: string } | null | undefined} file
     * @returns {boolean}
     */
    function isCsvFile(file) {
        if (!file) return false;
        if (/\.csv$/i.test(String(file.name || ''))) return true;
        return file.type === 'text/csv';
    }

    /**
     * @param {string} a
     * @param {string} b
     * @returns {number}
     */
    function compareNames(a, b) {
        return String(a || '').localeCompare(String(b || ''), undefined, {
            numeric: true,
            sensitivity: 'base'
        });
    }

    /**
     * @param {{ webkitRelativePath?: string }} file
     * @returns {string}
     */
    function relativePathOf(file) {
        return String(file?.webkitRelativePath || '').replace(/\\/g, '/');
    }

    /**
     * Top-level folder from a dropped/selected relative path (`folder/file.csv`).
     * @param {{ webkitRelativePath?: string }} file
     * @returns {string}
     */
    function droppedFolderName(file) {
        const rel = relativePathOf(file);
        if (!rel) return '';
        const i = rel.indexOf('/');
        return i < 0 ? '' : rel.slice(0, i);
    }

    /**
     * Directories first, in increasing name order; files keep their relative order.
     * @param {ArrayLike<{ name?: string, isFile?: boolean, isDirectory?: boolean }> | null | undefined} entries
     * @returns {object[]}
     */
    function sortEntriesForCollection(entries) {
        const list = Array.from(entries || []).filter(Boolean);
        const dirs = list.filter(e => e.isDirectory);
        const files = list.filter(e => !e.isDirectory);
        dirs.sort((a, b) => compareNames(a.name, b.name));
        return [...dirs, ...files];
    }

    /**
     * @param {ArrayLike<{ name?: string, type?: string, webkitRelativePath?: string }> | null | undefined} fileList
     * @returns {{ name?: string, type?: string, webkitRelativePath?: string }[]}
     */
    function collectCsvFilesFromFileList(fileList) {
        if (!fileList || typeof fileList.length !== 'number') return [];
        return Array.from(fileList).filter(isCsvFile).sort((a, b) => {
            const folderCmp = compareNames(droppedFolderName(a), droppedFolderName(b));
            if (folderCmp) return folderCmp;
            const relCmp = compareNames(relativePathOf(a), relativePathOf(b));
            if (relCmp) return relCmp;
            return compareNames(a.name, b.name);
        });
    }

    /**
     * @param {{ file: Function }} entry
     * @returns {Promise<File>}
     */
    function fileFromEntry(entry) {
        return new Promise((resolve, reject) => {
            entry.file(resolve, reject);
        });
    }

    /**
     * `readEntries` may return a partial batch; keep calling until it is empty.
     * @param {{ createReader?: Function }} dirEntry
     * @returns {Promise<object[]>}
     */
    function readAllDirectoryEntries(dirEntry) {
        if (typeof dirEntry.createReader !== 'function') return Promise.resolve([]);
        const reader = dirEntry.createReader();
        const all = [];
        const readBatch = () => new Promise((resolve, reject) => {
            reader.readEntries(resolve, reject);
        }).then((batch) => {
            if (!batch || !batch.length) return all;
            all.push(...batch);
            return readBatch();
        });
        return readBatch();
    }

    /**
     * @param {{ isFile?: boolean, isDirectory?: boolean, file?: Function, createReader?: Function } | null | undefined} entry
     * @param {File[]} out
     * @returns {Promise<void>}
     */
    async function collectCsvFilesFromEntry(entry, out) {
        if (!entry) return;
        try {
            if (entry.isFile) {
                const file = await fileFromEntry(entry);
                if (isCsvFile(file)) out.push(file);
                return;
            }
            if (entry.isDirectory) {
                const children = sortEntriesForCollection(await readAllDirectoryEntries(entry));
                for (const child of children) {
                    await collectCsvFilesFromEntry(child, out);
                }
            }
        } catch (err) {
            console.error('[Flysight] Failed to read entry:', entry.name, err);
        }
    }

    /**
     * Recursively collect CSV File objects from FileSystemEntry values
     * (`webkitGetAsEntry` / dropped folders).
     *
     * @param {ArrayLike<object> | null | undefined} entries
     * @returns {Promise<File[]>}
     */
    async function collectCsvFilesFromEntries(entries) {
        const files = [];
        if (!entries || typeof entries.length !== 'number') return files;
        const ordered = sortEntriesForCollection(entries);
        for (let i = 0; i < ordered.length; i++) {
            await collectCsvFilesFromEntry(ordered[i], files);
        }
        return files;
    }

    /**
     * Shadow `webkitRelativePath` so File System Access picks sort like
     * `<input webkitdirectory>` files (`selectedFolder/.../file.csv`).
     * @param {File} file
     * @param {string} relativePath
     * @returns {File}
     */
    function withRelativePath(file, relativePath) {
        const path = String(relativePath || '').replace(/\\/g, '/');
        if (!file) return file;
        try {
            Object.defineProperty(file, 'webkitRelativePath', {
                configurable: true,
                enumerable: true,
                value: path
            });
        } catch {
            file.webkitRelativePath = path;
        }
        return file;
    }

    /**
     * @param {{ values?: Function, entries?: Function } | null | undefined} dirHandle
     * @returns {Promise<object[]>}
     */
    async function listDirectoryHandleEntries(dirHandle) {
        const entries = [];
        if (!dirHandle) return entries;
        try {
            if (typeof dirHandle.values === 'function') {
                for await (const entry of dirHandle.values()) {
                    if (entry) entries.push(entry);
                }
                return entries;
            }
            if (typeof dirHandle.entries === 'function') {
                for await (const item of dirHandle.entries()) {
                    const entry = Array.isArray(item) ? item[1] : item;
                    if (entry) entries.push(entry);
                }
            }
        } catch (err) {
            console.error('[Flysight] Failed to list directory:', dirHandle.name, err);
        }
        return entries;
    }

    /**
     * @param {ArrayLike<{ name?: string, kind?: string }> | null | undefined} entries
     * @returns {object[]}
     */
    function sortDirectoryHandleEntries(entries) {
        const list = Array.from(entries || []).filter(Boolean);
        const dirs = list.filter(e => e.kind === 'directory');
        const files = list.filter(e => e.kind !== 'directory');
        dirs.sort((a, b) => compareNames(a.name, b.name));
        return [...dirs, ...files];
    }

    /**
     * Recursively collect CSV File objects from a FileSystemDirectoryHandle.
     * Non-CSV entries are skipped by name before they are read.
     *
     * @param {{ name?: string, kind?: string, values?: Function, entries?: Function, getFile?: Function } | null | undefined} dirHandle
     * @returns {Promise<File[]>}
     */
    async function collectCsvFilesFromDirectoryHandle(dirHandle) {
        const files = [];
        if (!dirHandle) return files;

        async function walk(handle, relDir) {
            const children = sortDirectoryHandleEntries(await listDirectoryHandleEntries(handle));
            for (const entry of children) {
                const rel = relDir ? `${relDir}/${entry.name}` : String(entry.name || '');
                if (entry.kind === 'directory') {
                    await walk(entry, rel);
                    continue;
                }
                if (entry.kind !== 'file' || typeof entry.getFile !== 'function') continue;
                if (!isCsvFile({ name: entry.name })) continue;
                try {
                    const file = await entry.getFile();
                    files.push(withRelativePath(file, rel));
                } catch (err) {
                    console.error('[Flysight] Failed to read file:', entry.name, err);
                }
            }
        }

        await walk(dirHandle, dirHandle.name || '');
        return collectCsvFilesFromFileList(files);
    }

    /**
     * True when a directory input (or similar) produced a folder tree
     * (`selectedFolder/.../file.csv`). A flat file picker leaves this empty.
     * @param {ArrayLike<{ webkitRelativePath?: string }> | null | undefined} fileList
     * @returns {boolean}
     */
    function fileListHasRelativePaths(fileList) {
        if (!fileList || typeof fileList.length !== 'number') return false;
        return Array.from(fileList).some(file => relativePathOf(file).includes('/'));
    }

    /**
     * @param {{ webkitRelativePath?: string, name?: string }} file
     * @returns {string[]}
     */
    function relativePathParts(file) {
        const rel = relativePathOf(file);
        if (rel) return rel.split('/').filter(Boolean);
        const name = String(file?.name || '');
        return name ? [name] : [];
    }

    /**
     * @param {string} name
     * @param {string} path
     * @returns {{ name: string, path: string, dirs: Record<string, object>, files: object[] }}
     */
    function emptyVirtualNode(name, path) {
        return { name: String(name || ''), path: String(path || ''), dirs: Object.create(null), files: [] };
    }

    /**
     * @param {{ name: string, path: string, dirs: Record<string, object> | object[], files: object[] }} node
     * @returns {{ name: string, path: string, dirs: object[], files: object[] }}
     */
    function freezeVirtualNode(node) {
        const dirList = Array.isArray(node.dirs)
            ? node.dirs.slice()
            : Object.keys(node.dirs || {}).map(key => node.dirs[key]);
        dirList.sort((a, b) => compareNames(a.name, b.name));
        const files = (node.files || []).slice().sort((a, b) => compareNames(a.name, b.name));
        return {
            name: node.name,
            path: node.path,
            dirs: dirList.map(freezeVirtualNode),
            files
        };
    }

    /**
     * Build a navigable folder tree from `<input webkitdirectory>` files.
     * CSVs only; the selected-folder prefix becomes the root when present.
     *
     * @param {ArrayLike<{ name?: string, type?: string, webkitRelativePath?: string }> | null | undefined} fileList
     * @returns {{ name: string, path: string, dirs: object[], files: object[] }}
     */
    function buildVirtualTreeFromFileList(fileList) {
        const csvs = collectCsvFilesFromFileList(fileList);
        const root = emptyVirtualNode('', '');
        for (const file of csvs) {
            const parts = relativePathParts(file);
            if (!parts.length) continue;
            let node = root;
            const dirParts = parts.slice(0, -1);
            const pathParts = [];
            for (const part of dirParts) {
                pathParts.push(part);
                if (!node.dirs[part]) {
                    node.dirs[part] = emptyVirtualNode(part, pathParts.join('/'));
                }
                node = node.dirs[part];
            }
            node.files.push(file);
        }
        const topNames = Object.keys(root.dirs);
        const start = (!root.files.length && topNames.length === 1)
            ? root.dirs[topNames[0]]
            : root;
        if (!start.name && topNames.length) start.name = topNames[0];
        return freezeVirtualNode(start);
    }

    /**
     * One-level listing for the in-app folder browser: directories first, then CSVs.
     * @param {{ dirs?: object[], files?: object[] } | null | undefined} node
     * @returns {{ dirs: object[], files: object[] }}
     */
    function listVirtualTreeBrowserEntries(node) {
        if (!node) return { dirs: [], files: [] };
        return {
            dirs: Array.from(node.dirs || []),
            files: Array.from(node.files || [])
        };
    }

    /**
     * All CSV File objects under this virtual node, nested folders included.
     * @param {{ dirs?: object[], files?: object[] } | null | undefined} node
     * @returns {object[]}
     */
    function collectCsvFilesFromVirtualNode(node) {
        const files = [];
        function walk(n) {
            if (!n) return;
            for (const dir of n.dirs || []) walk(dir);
            files.push(...(n.files || []));
        }
        walk(node);
        return collectCsvFilesFromFileList(files);
    }

    /**
     * @param {{ dirs?: object[], files?: object[] } | null | undefined} node
     * @returns {number}
     */
    function countCsvFilesFromVirtualNode(node) {
        if (!node) return 0;
        let n = (node.files || []).length;
        for (const dir of node.dirs || []) n += countCsvFilesFromVirtualNode(dir);
        return n;
    }

    /**
     * One-level File System Access listing: directories first, CSV handles only.
     * Does not call `getFile()`.
     *
     * @param {{ values?: Function, entries?: Function } | null | undefined} dirHandle
     * @returns {Promise<{ dirs: object[], files: object[] }>}
     */
    async function listDirectoryHandleBrowserEntries(dirHandle) {
        const children = sortDirectoryHandleEntries(await listDirectoryHandleEntries(dirHandle));
        const dirs = children.filter(e => e.kind === 'directory');
        const files = children.filter(e => e.kind === 'file' && isCsvFile({ name: e.name }));
        return { dirs, files };
    }

    /**
     * Count CSV names under a directory handle without reading file bytes.
     * @param {{ values?: Function, entries?: Function } | null | undefined} dirHandle
     * @returns {Promise<number>}
     */
    async function countCsvFilesFromDirectoryHandle(dirHandle) {
        let count = 0;
        if (!dirHandle) return count;
        async function walk(handle) {
            const children = sortDirectoryHandleEntries(await listDirectoryHandleEntries(handle));
            for (const entry of children) {
                if (entry.kind === 'directory') {
                    await walk(entry);
                    continue;
                }
                if (entry.kind === 'file' && isCsvFile({ name: entry.name })) count += 1;
            }
        }
        await walk(dirHandle);
        return count;
    }

    /**
     * Read chosen CSV handles (used when the browser modal has files checked).
     * @param {ArrayLike<{ name?: string, getFile?: Function }> | null | undefined} fileHandles
     * @param {string} [relDir]
     * @returns {Promise<File[]>}
     */
    async function readCsvFilesFromFileHandles(fileHandles, relDir) {
        const files = [];
        if (!fileHandles || typeof fileHandles.length !== 'number') return files;
        for (const handle of fileHandles) {
            if (!handle || typeof handle.getFile !== 'function') continue;
            if (!isCsvFile({ name: handle.name })) continue;
            try {
                const file = await handle.getFile();
                const rel = relDir ? `${relDir}/${handle.name}` : String(handle.name || '');
                files.push(withRelativePath(file, rel));
            } catch (err) {
                console.error('[Flysight] Failed to read file:', handle.name, err);
            }
        }
        return collectCsvFilesFromFileList(files);
    }

    const Flysight = {
        parseFlysightCsv,
        formatTrackStartTitle,
        isCsvFile,
        collectCsvFilesFromFileList,
        collectCsvFilesFromEntries,
        collectCsvFilesFromDirectoryHandle,
        listDirectoryHandleEntries,
        sortDirectoryHandleEntries,
        listDirectoryHandleBrowserEntries,
        countCsvFilesFromDirectoryHandle,
        readCsvFilesFromFileHandles,
        fileListHasRelativePaths,
        buildVirtualTreeFromFileList,
        listVirtualTreeBrowserEntries,
        collectCsvFilesFromVirtualNode,
        countCsvFilesFromVirtualNode,
        analyzeFlysightTrack,
        analyzeFlysightCsv,
        flightSwoopWindow,
        peakVerticalFromGraphSamples,
        filterPointsBySpeedAccuracy,
        trajectorySpeedMs,
        normalizeSpeedMetric,
        movingAverage,
        medianSampleIntervalSec,
        averagingWindowDurationSec,
        formatDurationSec,
        timeAloftSec,
        recoveryArcSec,
        pathAngleRad,
        diveAngleDeg,
        pathAngleFlatteningDegS,
        pathAngleRateDegS,
        haversineMeters,
        displacementMeters,
        findStationaryCutoffMs,
        pointsUpToFirstLanding,
        lastSignificantVelDPeakIdx,
        maxVelDIdx,
        buildSwoopCursorSeries,
        defaultSwoopCursorIndices,
        defaultSwoopCursorIdxA,
        defaultSwoopCursorIdxB,
        graphSwoopCursorIndices,
        entrySpeedAtGraphCursorB,
        entryGatesForSwoop,
        cursorBBeforeClosestEntryGate,
        closestEntryGatePassOnSeries,
        SWOOP_WINDOW_SEC,
        STATIONARY_SEC,
        STATIONARY_RADIUS_M,
        CURSOR_B_VELD_MS,
        CURSOR_B_DIVE_ANGLE_DEG,
        CURSOR_B_AGL_M,
        CURSOR_B_ALT_TICKS,
        CURSOR_A_FLATTENING_DEG_S,
        CURSOR_A_PEAK_FRACTION,
        DEFAULT_SAMPLE_INTERVAL_SEC,
        MAX_SPEED_ACCURACY_MS,
        DEFAULT_SPEED_METRIC,
        LOCATION_MATCH_MAX_KM,
        diveAngleLocalPeakIdxBeforeA,
        closestGatePassOnTrack,
        nearestGateToPoint,
        gatesForMatchedTrack,
        effectiveGatesForLocation,
        minTrackDistToLocation,
        matchLocationForTrack,
        computeGateSnapDistance
    };

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = Flysight;
    } else {
        global.Flysight = Flysight;
    }
})(typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : global);
