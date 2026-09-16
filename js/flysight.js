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
        return `${MONTH_NAMES[d.getMonth()]} ${d.getDate()} - ${hh}:${mm}`;
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
     * Last `windowSec` before the first 2 s of roughly-still position (landing),
     * in chronological order (index 0 is earliest; the last sample is landing).
     * Points after the still period are dropped. Not limited by the analysis
     * max-height slider. `tRev` is seconds before landing. `velD` is smoothed;
     * `flatteningDegS` is signed dγ/dt from raw velocities (negative = dive
     * angle decreasing in real time).
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
        if (!points || points.length < 2) {
            return { samples: [], error: 'Not enough track points.' };
        }
        const quality = filterPointsBySpeedAccuracy(points);
        const flight = pointsUpToFirstLanding(quality);
        if (flight.length < 2) {
            return { samples: [], error: 'Not enough track points.' };
        }

        const tEnd = lastValidTimeMs(flight);
        if (!Number.isFinite(tEnd)) {
            return { samples: [], error: 'Track times are invalid.' };
        }

        const windowed = [];
        for (let i = 0; i < flight.length; i++) {
            const t = Date.parse(flight[i].time);
            if (!Number.isFinite(t)) continue;
            const dt = (tEnd - t) / 1000;
            if (dt >= 0 && dt <= windowSec) windowed.push(flight[i]);
        }
        if (windowed.length < 2) {
            return { samples: [], error: 'Not enough points in the swoop window.' };
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
            return {
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
        const last = samples.length - 1;
        const flattenThresh = -CURSOR_A_FLATTENING_DEG_S;
        const angleThresh = Number.isFinite(diveAngleDegThresh)
            ? diveAngleDegThresh
            : CURSOR_B_DIVE_ANGLE_DEG;
        const ticks = Number.isFinite(altTicks)
            ? Math.max(0, Math.floor(altTicks))
            : CURSOR_B_ALT_TICKS;
        const ground = minHmsl(samples);

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
        return { idxA, idxB, peakVelD };
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

    const DEFAULT_MAX_HEIGHT_M = 500;
    const MIN_MAX_HEIGHT_M = 1;
    /** Absolute analysis / slider-range cap. The UI slider max defaults to 500 m and is user-configurable. */
    const MAX_MAX_HEIGHT_M = 6000;
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
     * @param {{ hMSL: number }[]} points
     * @param {number} maxHeightM
     * @returns {{ hMSL: number, velD: number }[]}
     */
    function filterPointsByMaxHeight(points, maxHeightM) {
        const minHmsl = points.reduce((min, p) => (p.hMSL < min ? p.hMSL : min), points[0].hMSL);
        const ceiling = Math.max(MIN_MAX_HEIGHT_M, Math.min(MAX_MAX_HEIGHT_M, maxHeightM));
        return points.filter(p => (p.hMSL - minHmsl) <= ceiling);
    }

    /**
     * Post-landing samples (after the first 2 s still stretch) are ignored so
     * a device left on cannot pull ground / AGL down.
     *
     * @param {{ hMSL: number, velD: number }[]} points
     * @param {number} avgPoints
     * @param {number} maxHeightM — ignore points more than this many metres above landing (AGL)
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
        maxHeightM = DEFAULT_MAX_HEIGHT_M,
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

        const quality = filterPointsBySpeedAccuracy(points);
        const flight = pointsUpToFirstLanding(quality);
        if (!flight.length) {
            return emptyResult({ error: 'No track points.' });
        }
        const minHmsl = flight.reduce((min, p) => (p.hMSL < min ? p.hMSL : min), flight[0].hMSL);
        const eligible = filterPointsByMaxHeight(flight, maxHeightM);
        if (!eligible.length) {
            return emptyResult({ minHmsl, error: 'No track points within the max height limit.' });
        }

        if (metric === 'both') {
            const verticalRaw = rawSpeedsForMode(eligible, 'vertical');
            const totalRaw = rawSpeedsForMode(eligible, 'total');
            if (totalRaw.some(s => !Number.isFinite(s))) {
                return emptyResult({
                    minHmsl,
                    error: 'Missing velN/velE columns required for total speed.'
                });
            }

            const verticalPeak = peakFromRawSpeeds(eligible, minHmsl, avgPoints, verticalRaw);
            const totalPeak = peakFromRawSpeeds(eligible, minHmsl, avgPoints, totalRaw);

            return {
                maxVerticalSpeedKmh: verticalPeak.maxSpeedKmh,
                maxTotalSpeedKmh: totalPeak.maxSpeedKmh,
                altitudeM: verticalPeak.altitudeM,
                time: verticalPeak.time,
                totalPeakTime: totalPeak.time,
                totalPeakAltitudeM: totalPeak.altitudeM,
                pointCount: eligible.length,
                minHmsl,
                speedMetric: 'both'
            };
        }

        const mode = metric === 'total' ? 'total' : 'vertical';
        const rawSpeeds = rawSpeedsForMode(eligible, mode);
        if (rawSpeeds.some(s => !Number.isFinite(s))) {
            return emptyResult({
                minHmsl,
                error: 'Missing velN/velE columns required for total speed.'
            });
        }

        const peak = peakFromRawSpeeds(eligible, minHmsl, avgPoints, rawSpeeds);

        return {
            maxVerticalSpeedKmh: peak.maxSpeedKmh,
            altitudeM: peak.altitudeM,
            time: peak.time,
            pointCount: eligible.length,
            minHmsl,
            speedMetric: metric
        };
    }

    /**
     * @param {string} text
     * @param {number} avgPoints
     * @param {number} maxHeightM
     * @param {'vertical' | 'total' | 'both'} speedMetric
     * @returns {ReturnType<typeof analyzeFlysightTrack> & { points: typeof points }}
     */
    function analyzeFlysightCsv(
        text,
        avgPoints = 1,
        maxHeightM = DEFAULT_MAX_HEIGHT_M,
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
        const result = analyzeFlysightTrack(parsed.points, avgPoints, maxHeightM, metric);
        return { ...result, points: parsed.points };
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
     * Top-level folder from a dropped/selected relative path (`folder/file.csv`).
     * @param {{ webkitRelativePath?: string }} file
     * @returns {string}
     */
    function droppedFolderName(file) {
        const rel = String(file?.webkitRelativePath || '').replace(/\\/g, '/');
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
            const relCmp = compareNames(
                String(a.webkitRelativePath || ''),
                String(b.webkitRelativePath || '')
            );
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

    const Flysight = {
        parseFlysightCsv,
        formatTrackStartTitle,
        isCsvFile,
        collectCsvFilesFromFileList,
        collectCsvFilesFromEntries,
        analyzeFlysightTrack,
        analyzeFlysightCsv,
        filterPointsByMaxHeight,
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
        DEFAULT_MAX_HEIGHT_M,
        MIN_MAX_HEIGHT_M,
        MAX_MAX_HEIGHT_M,
        MAX_SPEED_ACCURACY_MS,
        DEFAULT_SPEED_METRIC
    };

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = Flysight;
    } else {
        global.Flysight = Flysight;
    }
})(typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : global);
