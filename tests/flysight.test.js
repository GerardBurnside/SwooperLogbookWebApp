const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadFlysight() {
    const filePath = path.join(__dirname, '..', 'js', 'flysight.js');
    const code = fs.readFileSync(filePath, 'utf8');
    const sandbox = { console };
    sandbox.globalThis = sandbox;
    vm.createContext(sandbox);
    vm.runInContext(code, sandbox);
    return sandbox.Flysight;
}

const F = loadFlysight();

const SAMPLE_CSV = `time,lat,lon,hMSL,velN,velE,velD,hAcc,vAcc,sAcc,heading,cAcc,gpsFix,numSV
,(deg),(deg),(m),(m/s),(m/s),(m/s),(m),(m),(m/s),(deg),(deg),,
2026-07-25T06:58:00.30Z,47.9038276,2.1750580,1434.659,-11.57,-14.79,10.51,10.956,13.298,0.52,231.95744,1.22441,3,9
2026-07-25T06:58:00.50Z,47.9038061,2.1750169,1432.578,-10.92,-15.44,10.54,7.202,7.508,0.38,234.73104,1.02747,3,9
2026-07-25T06:58:00.80Z,47.9037755,2.1749543,1429.179,-9.87,-16.39,10.66,6.002,6.693,0.28,238.93489,0.89907,3,9
2026-07-25T06:58:00.90Z,47.9037662,2.1749327,1428.035,-9.41,-16.64,10.64,5.537,6.225,0.26,240.51135,0.88636,3,9
2026-07-25T06:58:01.00Z,47.9037574,2.1749108,1426.880,-9.05,-16.80,10.69,5.174,5.843,0.25,241.68641,0.87946,3,9
2026-07-25T06:58:01.10Z,47.9037490,2.1748885,1425.761,-8.73,-17.13,10.75,4.874,5.521,0.24,242.99552,0.86933,3,9
2026-07-25T06:58:01.20Z,47.9037409,2.1748658,1424.643,-8.42,-17.38,10.55,4.623,5.254,0.24,244.15600,0.86734,3,9
2026-07-25T06:58:01.30Z,47.9037331,2.1748427,1423.542,-8.13,-17.53,10.57,4.407,5.020,0.23,245.12732,0.86637,3,9
2026-07-25T06:58:01.40Z,47.9037256,2.1748194,1422.449,-7.81,-17.75,10.48,4.219,4.815,0.23,246.25626,0.86483,3,9
2026-07-25T06:58:01.50Z,47.9037183,2.1747956,1421.357,-7.62,-18.09,10.50,4.054,4.634,0.22,247.15819,0.85530,3,9
2026-07-25T06:58:01.60Z,47.9037112,2.1747715,1420.272,-7.29,-18.25,10.50,3.904,4.470,0.22,248.22259,0.85316,3,9`;

test('formatTrackStartTitle is DATE/HOUR in the local timezone', () => {
    const iso = '2026-03-09T15:53:12.10Z';
    const d = new Date(iso);
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    const hh = String(d.getHours()).padStart(2, '0');
    const mm = String(d.getMinutes()).padStart(2, '0');
    assert.equal(F.formatTrackStartTitle(iso), `${months[d.getMonth()]} ${d.getDate()} - ${hh}:${mm}`);
    assert.match(F.formatTrackStartTitle(iso), /^[A-Z][a-z]+ \d{1,2} - \d{2}:\d{2}$/);
});

test('formatTrackStartTitle uses the first parsed CSV timestamp', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    assert.equal(F.formatTrackStartTitle(points[0].time), F.formatTrackStartTitle('2026-07-25T06:58:00.30Z'));
});

test('formatTrackStartTitle returns empty for invalid time', () => {
    assert.equal(F.formatTrackStartTitle(''), '');
    assert.equal(F.formatTrackStartTitle('not-a-date'), '');
    assert.equal(F.formatTrackStartTitle(undefined), '');
});

test('parseFlysightCsv reads sample track', () => {
    const { points, error } = F.parseFlysightCsv(SAMPLE_CSV);
    assert.equal(error, undefined);
    assert.equal(points.length, 11);
    assert.equal(points[0].velD, 10.51);
});

test('analyzeFlysightTrack: max speed and altitude above min hMSL', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const r = F.analyzeFlysightTrack(points, 1);
    assert.equal(r.pointCount, 11);
    assert.ok(r.maxVerticalSpeedKmh > 38 && r.maxVerticalSpeedKmh < 39);
    assert.equal(r.minHmsl, 1420.272);
    assert.ok(r.altitudeM > 4 && r.altitudeM < 8);
});

test('analyzeFlysightTrack: averaging reduces peak vs raw max', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const raw = F.analyzeFlysightTrack(points, 1);
    const smooth = F.analyzeFlysightTrack(points, 5);
    assert.ok(smooth.maxVerticalSpeedKmh <= raw.maxVerticalSpeedKmh);
});

test('parseFlysightCsv rejects missing columns', () => {
    const r = F.parseFlysightCsv('a,b,c\n1,2,3\n');
    assert.ok(r.error);
    assert.equal(r.points.length, 0);
});

test('medianSampleIntervalSec uses median delta between timestamps', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const interval = F.medianSampleIntervalSec(points);
    assert.ok(interval > 0.09 && interval < 0.11);
});

test('averagingWindowDurationSec spans n-1 sample intervals', () => {
    assert.equal(F.averagingWindowDurationSec(1, 0.1), null);
    assert.equal(F.averagingWindowDurationSec(2, 0.05), 0.05);
    assert.equal(F.averagingWindowDurationSec(5, 0.1), 0.4);
});

test('formatDurationSec renders compact seconds suffix', () => {
    assert.equal(F.formatDurationSec(0.05), '0.05s');
    assert.equal(F.formatDurationSec(0.4), '0.40s');
});

test('filterPointsByMaxHeight ignores points above AGL ceiling', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const filtered = F.filterPointsByMaxHeight(points, 10);
    assert.ok(filtered.length < points.length);
    assert.ok(filtered.every(p => p.hMSL <= 1420.272 + 10));
});

test('filterPointsByMaxHeight allows ceilings above the default 500 m', () => {
    const points = [
        { hMSL: 100, velD: 1 },
        { hMSL: 700, velD: 1 },
        { hMSL: 1200, velD: 1 }
    ];
    const at500 = F.filterPointsByMaxHeight(points, 500);
    const at800 = F.filterPointsByMaxHeight(points, 800);
    assert.equal(at500.length, 1);
    assert.equal(at800.length, 2);
    assert.equal(at800[1].hMSL, 700);
    assert.ok(F.MAX_MAX_HEIGHT_M > 500);
});

test('analyzeFlysightTrack: max height limits eligible points', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const full = F.analyzeFlysightTrack(points, 1, 500);
    const limited = F.analyzeFlysightTrack(points, 1, 5);
    assert.equal(full.pointCount, points.length);
    assert.ok(limited.pointCount < full.pointCount);
});

test('analyzeFlysightTrack: speed metric selects vertical vs total trajectory speed', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const vertical = F.analyzeFlysightTrack(points, 1, 500, 'vertical');
    const total = F.analyzeFlysightTrack(points, 1, 500, 'total');
    assert.equal(vertical.speedMetric, 'vertical');
    assert.equal(total.speedMetric, 'total');
    assert.ok(total.maxVerticalSpeedKmh > vertical.maxVerticalSpeedKmh);
    const maxRawTotalKmh = Math.max(...points.map(p => F.trajectorySpeedMs(p))) * 3.6;
    assert.ok(total.maxVerticalSpeedKmh <= maxRawTotalKmh + 0.1);
});

test('analyzeFlysightTrack: both mode returns vertical and total peaks', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    const both = F.analyzeFlysightTrack(points, 1, 500, 'both');
    const vertical = F.analyzeFlysightTrack(points, 1, 500, 'vertical');
    const total = F.analyzeFlysightTrack(points, 1, 500, 'total');
    assert.equal(both.speedMetric, 'both');
    assert.equal(both.maxVerticalSpeedKmh, vertical.maxVerticalSpeedKmh);
    assert.equal(both.maxTotalSpeedKmh, total.maxVerticalSpeedKmh);
    assert.equal(both.time, vertical.time);
    assert.equal(both.totalPeakTime, total.time);
    assert.equal(both.altitudeM, vertical.altitudeM);
    assert.equal(both.totalPeakAltitudeM, total.altitudeM);
});

test('trajectorySpeedMs is 3D velocity magnitude', () => {
    assert.equal(F.trajectorySpeedMs({ velN: 3, velE: 4, velD: 0 }), 5);
});

test('parseFlysightCsv includes lat/lon when present', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    assert.equal(points[0].lat, 47.9038276);
    assert.equal(points[0].lon, 2.1750580);
});

test('parseFlysightCsv includes sAcc when present', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    assert.equal(points[0].sAcc, 0.52);
});

test('filterPointsBySpeedAccuracy drops sAcc above 2 m/s and keeps missing sAcc', () => {
    const points = [
        { hMSL: -586, velD: 220.51, sAcc: 428 },
        { hMSL: 1667, velD: 10.85, sAcc: 22.08 },
        { hMSL: 1763, velD: 1.2, sAcc: 2 },
        { hMSL: 1750, velD: 50, sAcc: 0.3 },
        { hMSL: 120, velD: 0.1 }
    ];
    const kept = F.filterPointsBySpeedAccuracy(points);
    assert.equal(kept.length, 3);
    assert.equal(kept[0].hMSL, 1763);
    assert.equal(kept[1].velD, 50);
    assert.equal(kept[2].hMSL, 120);
});

test('filterPointsBySpeedAccuracy returns original points if all would be dropped', () => {
    const points = [
        { hMSL: 100, velD: 1, sAcc: 5 },
        { hMSL: 101, velD: 2, sAcc: 8 }
    ];
    const kept = F.filterPointsBySpeedAccuracy(points);
    assert.equal(kept, points);
    assert.equal(kept.length, 2);
});

test('pointsUpToFirstLanding keeps a track with no still stretch', () => {
    const { points } = F.parseFlysightCsv(SAMPLE_CSV);
    assert.equal(F.pointsUpToFirstLanding(points), points);
});

test('analyzeFlysightTrack ignores post-landing GNSS wander so 0 m AGL is landing', () => {
    const t0 = Date.parse('2026-01-01T00:00:00.00Z');
    const landHmsl = 115;
    const nMove = 50;
    const points = [];
    for (let i = 0; i < nMove; i++) {
        const frac = i / (nMove - 1);
        const speedScale = frac < 0.7 ? 1 : (1 - (frac - 0.7) / 0.3);
        points.push({
            time: new Date(t0 + i * 100).toISOString(),
            hMSL: 160 - frac * (160 - landHmsl),
            velD: 12 * speedScale,
            velN: 30 * speedScale,
            velE: 20 * speedScale,
            lat: 47.9 + i * 0.00008,
            lon: 2.17,
            sAcc: 0.2
        });
    }
    const land = points[nMove - 1];
    for (let i = 0; i < 40; i++) {
        points.push({
            time: new Date(t0 + (nMove + i) * 100).toISOString(),
            hMSL: land.hMSL,
            velD: 0.05,
            velN: 0.1,
            velE: 0,
            lat: land.lat,
            lon: land.lon,
            sAcc: 0.2
        });
    }
    const after = nMove + 40;
    for (let i = 1; i <= 80; i++) {
        points.push({
            time: new Date(t0 + (after + i) * 100).toISOString(),
            hMSL: land.hMSL - i * 0.35,
            velD: 0.02,
            velN: 0.05,
            velE: 0,
            lat: land.lat,
            lon: land.lon,
            sAcc: 0.9
        });
    }

    const flight = F.pointsUpToFirstLanding(points);
    assert.ok(flight.length < points.length);
    assert.ok(flight.every(p => Date.parse(p.time) <= F.findStationaryCutoffMs(points)));
    assert.ok(Math.min(...flight.map(p => p.hMSL)) > 110);

    const r = F.analyzeFlysightTrack(points, 1, 30, 'total');
    assert.ok(r.minHmsl > 110 && r.minHmsl < 120);
    assert.ok(r.maxVerticalSpeedKmh > 100, `expected swoop speed, got ${r.maxVerticalSpeedKmh}`);
    assert.ok(r.altitudeM > 20 && r.altitudeM <= 30);
    assert.ok(r.pointCount < 80);
});

test('analyzeFlysightTrack ignores lock-on glitch so ground and peak come from the real jump', () => {
    const csv = `time,lat,lon,hMSL,velN,velE,velD,hAcc,vAcc,sAcc,heading,cAcc,gpsFix,numSV
,(deg),(deg),(m),(m/s),(m/s),(m/s),(m),(m),(m/s),(deg),(deg),,
2026-09-12T13:03:14.00Z,47.9043890,2.1625255,-586.421,-6.13,-71.47,220.51,636.209,5254.726,427.99,265.10087,21.88398,3,4
2026-09-12T13:03:14.20Z,47.9037940,2.1657076,1667.132,-13.63,-47.26,10.85,200.972,654.677,22.08,253.90975,17.69536,3,5
2026-09-12T13:07:20.00Z,47.9039000,2.1710000,160.000,20.00,30.00,25.00,0.40,0.60,0.20,240.00000,0.50,3,16
2026-09-12T13:07:20.10Z,47.9039020,2.1710100,157.500,19.50,29.50,24.80,0.40,0.60,0.18,240.00000,0.50,3,16
2026-09-12T13:07:20.20Z,47.9039040,2.1710200,155.000,19.00,29.00,24.50,0.40,0.60,0.17,240.00000,0.50,3,16
2026-09-12T13:07:25.00Z,47.9039600,2.1716100,118.560,0.20,0.10,0.05,0.44,0.62,0.15,40.00000,10.00,3,16`;
    const { points } = F.parseFlysightCsv(csv);
    const r = F.analyzeFlysightTrack(points, 1, 500, 'vertical');
    assert.ok(r.minHmsl > 100 && r.minHmsl < 120);
    assert.ok(r.maxVerticalSpeedKmh > 88 && r.maxVerticalSpeedKmh < 91);
    assert.ok(r.altitudeM > 30 && r.altitudeM < 50);
    assert.ok(r.pointCount >= 4);
});

test('findStationaryCutoffMs is the start of the first 2s still stretch', () => {
    const t0 = Date.parse('2026-01-01T00:00:00.00Z');
    const points = [];
    for (let i = 0; i < 50; i++) {
        points.push({
            time: new Date(t0 + i * 100).toISOString(),
            hMSL: 120 - i * 0.4,
            velD: 4,
            velN: 10,
            velE: 0,
            lat: 47.9 + i * 0.00008,
            lon: 2.17
        });
    }
    const stillStart = points.length;
    for (let i = 0; i < 40; i++) {
        points.push({
            time: new Date(t0 + (stillStart + i) * 100).toISOString(),
            hMSL: points[stillStart - 1].hMSL,
            velD: 0.05,
            velN: 0.1,
            velE: 0,
            lat: points[stillStart - 1].lat,
            lon: points[stillStart - 1].lon
        });
    }
    const cutoff = F.findStationaryCutoffMs(points);
    // Last moving sample is already at the still position, so cutoff is that point.
    assert.equal(cutoff, Date.parse(points[stillStart - 1].time));
});

test('buildSwoopCursorSeries ends at first 2s of still position, then 25s before that', () => {
    const t0 = Date.parse('2026-01-01T00:00:00.00Z');
    const points = [];
    for (let i = 0; i < 400; i++) {
        points.push({
            time: new Date(t0 + i * 100).toISOString(),
            hMSL: 150 - i * 0.1,
            velD: 5,
            velN: 8,
            velE: 0,
            lat: 47.9 + i * 0.00001,
            lon: 2.17
        });
    }
    const lastMove = points[points.length - 1];
    for (let i = 1; i <= 100; i++) {
        points.push({
            time: new Date(t0 + (400 + i) * 100).toISOString(),
            hMSL: lastMove.hMSL,
            velD: 0.02,
            velN: 0.05,
            velE: 0,
            lat: lastMove.lat,
            lon: lastMove.lon
        });
    }
    const { samples, error } = F.buildSwoopCursorSeries(points, 1, 25);
    assert.equal(error, undefined);
    const endTime = Date.parse(samples[0].time);
    const stillStart = Date.parse(points[400].time);
    assert.ok(Math.abs(endTime - stillStart) < 1000);
    assert.ok(samples[samples.length - 1].tRev > 24);
    assert.ok(samples[samples.length - 1].tRev <= 25.05);
    const lastTrackTime = Date.parse(points[points.length - 1].time);
    assert.ok(endTime < lastTrackTime - 5000);
});

test('buildSwoopCursorSeries reverses last 25s so index 0 is landing', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '12-21-30.CSV'), 'utf8');
    const { points, error } = F.parseFlysightCsv(csv);
    assert.equal(error, undefined);
    const { samples } = F.buildSwoopCursorSeries(points, 5, 25);
    assert.ok(samples.length > 50);
    assert.ok(samples[0].tRev < samples[samples.length - 1].tRev);
    assert.ok(samples[0].tRev < 0.2);
    assert.ok(samples[0].velD < 1);
    assert.ok(Number.isFinite(samples[0].hMSL));
    const lastTrackTime = Date.parse(points[points.length - 1].time);
    const landingTime = Date.parse(samples[0].time);
    assert.ok(landingTime < lastTrackTime - 1500);
    assert.ok(samples[samples.length - 1].tRev <= 25.05);
});

test('buildSwoopCursorSeries is not clipped by the analysis max-height ceiling', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '08-55-08.CSV'), 'utf8');
    const { points } = F.parseFlysightCsv(csv);
    const quality = F.filterPointsBySpeedAccuracy(points);
    const below500 = F.filterPointsByMaxHeight(quality, 500);
    const cutoff = F.findStationaryCutoffMs(quality);
    const firstBelow500 = below500[0];
    const clippedSpan = (cutoff - Date.parse(firstBelow500.time)) / 1000;
    assert.ok(clippedSpan < 24, `height-filtered lookback should be under 25s (${clippedSpan})`);
    const { samples } = F.buildSwoopCursorSeries(points, 5);
    assert.ok(samples[samples.length - 1].tRev > 24);
    assert.ok(samples[samples.length - 1].tRev <= 25.05);
});

test('defaultSwoopCursorIndices: A is first strong flattening after peak velD', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '12-21-30.CSV'), 'utf8');
    const { points } = F.parseFlysightCsv(csv);
    const { samples } = F.buildSwoopCursorSeries(points, 5);
    const { idxA, idxB, peakVelD } = F.defaultSwoopCursorIndices(samples);
    assert.ok(idxB < idxA);
    const ground = samples.reduce((min, s) => (Number.isFinite(s.hMSL) && s.hMSL < min ? s.hMSL : min), Infinity);
    assert.ok(
        samples[idxB].diveAngleDeg < F.CURSOR_B_DIVE_ANGLE_DEG
        || samples[idxB].hMSL - ground < F.CURSOR_B_AGL_M
    );
    const peakIdx = F.lastSignificantVelDPeakIdx(samples);
    assert.ok(idxA <= peakIdx);
    if (idxA < peakIdx) {
        assert.ok(samples[idxA].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S);
        if (idxA >= 2) assert.ok(samples[idxA - 1].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S);
        for (let i = peakIdx - 1; i > idxA; i--) {
            const strong = samples[i].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S
                && (i < 2 || samples[i - 1].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S);
            assert.equal(strong, false);
        }
    }
    const dt = samples[idxA].tRev - samples[idxB].tRev;
    assert.ok(dt > 2 && dt < 20);
    assert.ok(peakVelD > 0);
});

test('maxVelDIdx is the first sample with the highest velD', () => {
    assert.equal(F.maxVelDIdx([]), 0);
    assert.equal(F.maxVelDIdx([{ velD: 3 }, { velD: 9 }, { velD: 9 }, { velD: 4 }]), 1);
});

test('lastSignificantVelDPeakIdx prefers a later near-max peak over an earlier taller one', () => {
    const samples = [
        { velD: 1, flatteningDegS: 0, hMSL: 100, diveAngleDeg: 1 },
        { velD: 8, flatteningDegS: -22, hMSL: 101, diveAngleDeg: 8 },
        { velD: 20, flatteningDegS: -22, hMSL: 110, diveAngleDeg: 20 },
        { velD: 41, flatteningDegS: -5, hMSL: 140, diveAngleDeg: 70 },
        { velD: 35, flatteningDegS: 8, hMSL: 160, diveAngleDeg: 60 },
        { velD: 40, flatteningDegS: -18, hMSL: 180, diveAngleDeg: 75 },
        { velD: 42, flatteningDegS: -16, hMSL: 200, diveAngleDeg: 78 },
        { velD: 44, flatteningDegS: 2, hMSL: 220, diveAngleDeg: 80 },
        { velD: 38, flatteningDegS: 5, hMSL: 240, diveAngleDeg: 70 }
    ];
    assert.equal(F.lastSignificantVelDPeakIdx(samples), 3);
    const { idxA } = F.defaultSwoopCursorIndices(samples);
    assert.equal(idxA, 2);
});

test('defaultSwoopCursorIndices A uses the later velD peak on 10-33-00', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '10-33-00.CSV'), 'utf8');
    const { points } = F.parseFlysightCsv(csv);
    const { samples } = F.buildSwoopCursorSeries(points, 3);
    const { idxA } = F.defaultSwoopCursorIndices(samples);
    const peakIdx = F.lastSignificantVelDPeakIdx(samples);
    const ground = samples.reduce((min, s) => (Number.isFinite(s.hMSL) && s.hMSL < min ? s.hMSL : min), Infinity);
    let globalIdx = 0;
    for (let i = 1; i < samples.length; i++) {
        if (samples[i].velD > samples[globalIdx].velD) globalIdx = i;
    }
    assert.ok(samples[peakIdx].tRev < samples[globalIdx].tRev - 1.5);
    assert.ok(samples[peakIdx].velD >= samples[globalIdx].velD * F.CURSOR_A_PEAK_FRACTION);
    assert.ok(idxA <= peakIdx);
    const aglA = samples[idxA].hMSL - ground;
    assert.ok(aglA > 50 && aglA < 90, `A AGL ${aglA}`);
    assert.ok(samples[idxA].tRev > 9.5 && samples[idxA].tRev < 11);
});

test('recoveryArcSec is the default A-B cursor time difference', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '12-21-30.CSV'), 'utf8');
    const { points } = F.parseFlysightCsv(csv);
    const { samples } = F.buildSwoopCursorSeries(points, 5);
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples);
    const expected = Math.abs(samples[idxB].tRev - samples[idxA].tRev);
    assert.equal(F.recoveryArcSec(points, 5), expected);
    assert.ok(expected > 2 && expected < 20);
    assert.ok(Number.isNaN(F.recoveryArcSec([], 5)));
});

test('timeAloftSec is seconds from B to the stationary cutoff', () => {
    assert.equal(F.timeAloftSec({ tRev: 3.4 }), 3.4);
    assert.equal(F.timeAloftSec({ tRev: 0 }), 0);
    assert.equal(F.timeAloftSec(null), 0);
    const csv = fs.readFileSync(path.join(__dirname, '..', '12-21-30.CSV'), 'utf8');
    const { points } = F.parseFlysightCsv(csv);
    const { samples } = F.buildSwoopCursorSeries(points, 5);
    const { idxB } = F.defaultSwoopCursorIndices(samples);
    const aloft = F.timeAloftSec(samples[idxB]);
    assert.equal(aloft, samples[idxB].tRev);
    assert.ok(aloft > 0);
});

test('defaultSwoopCursorIndices places B after A when AGL stays below 2m', () => {
    const samples = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 90 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 100.4, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 101.5, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 110, diveAngleDeg: 90 },
        { velD: 24, flatteningDegS: -4, hMSL: 130, diveAngleDeg: 90 },
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 90 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, 5, 1);
    assert.equal(idxA, 3);
    assert.equal(idxB, 2);
    assert.ok(idxB < idxA);
});

test('defaultSwoopCursorIndices B altitude-tick threshold is configurable', () => {
    const samples = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 90 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 100.4, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 101.5, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 110, diveAngleDeg: 90 },
        { velD: 24, flatteningDegS: -4, hMSL: 130, diveAngleDeg: 90 },
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 90 }
    ];
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 1).idxB, 2);
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 2).idxB, 1);
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 3).idxB, 0);
});

test('defaultSwoopCursorIndices B ignores brief AGL dips shorter than x ticks', () => {
    const samples = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 90 },
        { velD: 0.3, flatteningDegS: 0, hMSL: 100.2, diveAngleDeg: 90 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.5, diveAngleDeg: 90 },
        { velD: 0.5, flatteningDegS: 0, hMSL: 103.0, diveAngleDeg: 90 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.2, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 108, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 120, diveAngleDeg: 90 },
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 90 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, 5, 2);
    assert.equal(idxA, 6);
    assert.equal(idxB, 1);
});

test('defaultSwoopCursorIndices places B after A when dive angle drops below 5°', () => {
    const samples = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100, diveAngleDeg: 0.5 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 150, diveAngleDeg: 4 },
        { velD: 12, flatteningDegS: -22, hMSL: 160, diveAngleDeg: 15 },
        { velD: 20, flatteningDegS: -18, hMSL: 170, diveAngleDeg: 40 },
        { velD: 24, flatteningDegS: -4, hMSL: 180, diveAngleDeg: 55 },
        { velD: 25, flatteningDegS: 1, hMSL: 190, diveAngleDeg: 60 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, 5, 20);
    assert.equal(idxA, 3);
    assert.equal(idxB, 1);
});

test('defaultSwoopCursorIndices B dive-angle threshold is configurable', () => {
    const samples = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100, diveAngleDeg: 0.5 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 150, diveAngleDeg: 4 },
        { velD: 12, flatteningDegS: -22, hMSL: 160, diveAngleDeg: 15 },
        { velD: 20, flatteningDegS: -18, hMSL: 170, diveAngleDeg: 40 },
        { velD: 24, flatteningDegS: -4, hMSL: 180, diveAngleDeg: 55 },
        { velD: 25, flatteningDegS: 1, hMSL: 190, diveAngleDeg: 60 }
    ];
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 20).idxB, 1);
    assert.equal(F.defaultSwoopCursorIndices(samples, 16, 20).idxB, 2);
    assert.equal(F.defaultSwoopCursorIndices(samples, 1, 20).idxB, 0);
});

test('defaultSwoopCursorIndices B is the later of dive-angle and below-2m', () => {
    const angleLater = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 1 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.4, diveAngleDeg: 4 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.0, diveAngleDeg: 8 },
        { velD: 12, flatteningDegS: -22, hMSL: 101.5, diveAngleDeg: 20 },
        { velD: 20, flatteningDegS: -18, hMSL: 110, diveAngleDeg: 40 },
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 60 }
    ];
    const angleLaterCursors = F.defaultSwoopCursorIndices(angleLater, 5, 2);
    assert.equal(angleLaterCursors.idxA, 4);
    assert.equal(angleLaterCursors.idxB, 1);

    const altLater = [
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 1 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.4, diveAngleDeg: 2 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.0, diveAngleDeg: 3 },
        { velD: 12, flatteningDegS: -22, hMSL: 110, diveAngleDeg: 20 },
        { velD: 20, flatteningDegS: -18, hMSL: 120, diveAngleDeg: 40 },
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 60 }
    ];
    const altLaterCursors = F.defaultSwoopCursorIndices(altLater, 5, 2);
    assert.equal(altLaterCursors.idxA, 4);
    assert.equal(altLaterCursors.idxB, 1);
});
