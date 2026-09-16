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

const TITLE_MONTHS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
];

function isoOnMarch9(year) {
    return `${year}-03-09T15:53:12.10Z`;
}

function expectedTrackStartTitle(iso, { includeYear = false } = {}) {
    const d = new Date(iso);
    const hh = String(d.getHours()).padStart(2, '0');
    const mm = String(d.getMinutes()).padStart(2, '0');
    const datePart = includeYear
        ? `${TITLE_MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`
        : `${TITLE_MONTHS[d.getMonth()]} ${d.getDate()}`;
    return `${datePart} - ${hh}:${mm}`;
}

test('formatTrackStartTitle is DATE/HOUR in the local timezone', () => {
    const iso = isoOnMarch9(new Date().getFullYear());
    assert.equal(F.formatTrackStartTitle(iso), expectedTrackStartTitle(iso));
    assert.match(F.formatTrackStartTitle(iso), /^[A-Z][a-z]+ \d{1,2} - \d{2}:\d{2}$/);
});

test('formatTrackStartTitle includes the year when it is not the current year', () => {
    const iso = isoOnMarch9(new Date().getFullYear() - 1);
    assert.equal(F.formatTrackStartTitle(iso), expectedTrackStartTitle(iso, { includeYear: true }));
    assert.match(F.formatTrackStartTitle(iso), /^[A-Z][a-z]+ \d{1,2}, \d{4} - \d{2}:\d{2}$/);
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
    const endTime = Date.parse(samples[samples.length - 1].time);
    const stillStart = Date.parse(points[400].time);
    assert.ok(Math.abs(endTime - stillStart) < 1000);
    assert.ok(samples[0].tRev > 24);
    assert.ok(samples[0].tRev <= 25.05);
    const lastTrackTime = Date.parse(points[points.length - 1].time);
    assert.ok(endTime < lastTrackTime - 5000);
});

test('buildSwoopCursorSeries keeps last 25s in chronological order, ending at landing', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '12-21-30.CSV'), 'utf8');
    const { points, error } = F.parseFlysightCsv(csv);
    assert.equal(error, undefined);
    const { samples } = F.buildSwoopCursorSeries(points, 5, 25);
    assert.ok(samples.length > 50);
    const last = samples[samples.length - 1];
    assert.ok(samples[0].tRev > last.tRev);
    assert.ok(last.tRev < 0.2);
    assert.ok(last.velD < 1);
    assert.ok(Number.isFinite(last.hMSL));
    const lastTrackTime = Date.parse(points[points.length - 1].time);
    const landingTime = Date.parse(last.time);
    assert.ok(landingTime < lastTrackTime - 1500);
    assert.ok(samples[0].tRev <= 25.05);
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
    assert.ok(samples[0].tRev > 24);
    assert.ok(samples[0].tRev <= 25.05);
});

test('defaultSwoopCursorIndices: A is first strong flattening after peak velD', () => {
    const csv = fs.readFileSync(path.join(__dirname, '..', '12-21-30.CSV'), 'utf8');
    const { points } = F.parseFlysightCsv(csv);
    const { samples } = F.buildSwoopCursorSeries(points, 5);
    const { idxA, idxB, peakVelD } = F.defaultSwoopCursorIndices(samples);
    assert.ok(idxB > idxA);
    const ground = samples.reduce((min, s) => (Number.isFinite(s.hMSL) && s.hMSL < min ? s.hMSL : min), Infinity);
    assert.ok(
        samples[idxB].diveAngleDeg < F.CURSOR_B_DIVE_ANGLE_DEG
        || samples[idxB].hMSL - ground < F.CURSOR_B_AGL_M
    );
    const peakIdx = F.lastSignificantVelDPeakIdx(samples);
    assert.ok(idxA >= peakIdx);
    if (idxA > peakIdx) {
        assert.ok(samples[idxA].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S);
        if (idxA + 1 < samples.length) {
            assert.ok(samples[idxA + 1].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S);
        }
        for (let i = peakIdx + 1; i < idxA; i++) {
            const strong = samples[i].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S
                && (i + 1 >= samples.length || samples[i + 1].flatteningDegS <= -F.CURSOR_A_FLATTENING_DEG_S);
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
        { velD: 38, flatteningDegS: 5, hMSL: 240, diveAngleDeg: 70 },
        { velD: 44, flatteningDegS: 2, hMSL: 220, diveAngleDeg: 80 },
        { velD: 42, flatteningDegS: -16, hMSL: 200, diveAngleDeg: 78 },
        { velD: 40, flatteningDegS: -18, hMSL: 180, diveAngleDeg: 75 },
        { velD: 35, flatteningDegS: 8, hMSL: 160, diveAngleDeg: 60 },
        { velD: 41, flatteningDegS: -5, hMSL: 140, diveAngleDeg: 70 },
        { velD: 20, flatteningDegS: -22, hMSL: 110, diveAngleDeg: 20 },
        { velD: 8, flatteningDegS: -22, hMSL: 101, diveAngleDeg: 8 },
        { velD: 1, flatteningDegS: 0, hMSL: 100, diveAngleDeg: 1 }
    ];
    assert.equal(F.lastSignificantVelDPeakIdx(samples), 5);
    const { idxA } = F.defaultSwoopCursorIndices(samples);
    assert.equal(idxA, 6);
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
    assert.ok(idxA >= peakIdx);
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
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 90 },
        { velD: 24, flatteningDegS: -4, hMSL: 130, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 110, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 101.5, diveAngleDeg: 90 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 100.4, diveAngleDeg: 90 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 90 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, 5, 1);
    assert.equal(idxA, 2);
    assert.equal(idxB, 3);
    assert.ok(idxB > idxA);
});

test('defaultSwoopCursorIndices B altitude-tick threshold is configurable', () => {
    const samples = [
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 90 },
        { velD: 24, flatteningDegS: -4, hMSL: 130, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 110, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 101.5, diveAngleDeg: 90 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 100.4, diveAngleDeg: 90 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 90 }
    ];
    // First sample below 2.49 m AGL (from A toward landing) is index 3.
    // x=1 is that first sample; x=2 and x=3 are later samples, not index 3.
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 1).idxB, 3);
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 2).idxB, 4);
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 3).idxB, 5);
});

test('defaultSwoopCursorIndices B treats AGL below 2.49 m as near-ground', () => {
    // Flare altitudes from 08-47-08.CSV. 2.494 m is not below 2.49 m;
    // 2.280 m is the first near-ground sample, so ticks=2 places B at 2.142 m.
    const samples = [
        { velD: 25, flatteningDegS: 1, hMSL: 160, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 140, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 130, diveAngleDeg: 90 },
        { velD: 2.42, flatteningDegS: 0, hMSL: 120.638, diveAngleDeg: 90 },
        { velD: 1.69, flatteningDegS: 0, hMSL: 120.424, diveAngleDeg: 90 },
        { velD: 1.06, flatteningDegS: 0, hMSL: 120.286, diveAngleDeg: 90 },
        { velD: 0.59, flatteningDegS: 0, hMSL: 120.204, diveAngleDeg: 90 },
        { velD: 0.45, flatteningDegS: 0, hMSL: 120.182, diveAngleDeg: 90 },
        { velD: 0.34, flatteningDegS: 0, hMSL: 120.150, diveAngleDeg: 90 },
        { velD: 0.20, flatteningDegS: 0, hMSL: 120.126, diveAngleDeg: 90 },
        { velD: 0.12, flatteningDegS: 0, hMSL: 120.120, diveAngleDeg: 90 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 118.144, diveAngleDeg: 90 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, 5, 2);
    assert.equal(idxA, 1);
    assert.equal(idxB, 5);
    const ground = 118.144;
    assert.ok(samples[idxB].hMSL - ground < F.CURSOR_B_AGL_M);
    assert.ok(samples[idxB - 1].hMSL - ground < F.CURSOR_B_AGL_M);
    assert.ok(samples[idxB - 2].hMSL - ground >= F.CURSOR_B_AGL_M);
});

test('defaultSwoopCursorIndices B ignores brief AGL dips shorter than x ticks', () => {
    const samples = [
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 90 },
        { velD: 20, flatteningDegS: -18, hMSL: 120, diveAngleDeg: 90 },
        { velD: 12, flatteningDegS: -22, hMSL: 108, diveAngleDeg: 90 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.2, diveAngleDeg: 90 },
        { velD: 0.5, flatteningDegS: 0, hMSL: 103.0, diveAngleDeg: 90 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.5, diveAngleDeg: 90 },
        { velD: 0.3, flatteningDegS: 0, hMSL: 100.2, diveAngleDeg: 90 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 90 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, 5, 2);
    assert.equal(idxA, 1);
    assert.equal(idxB, 6);
});

test('defaultSwoopCursorIndices places B after A when dive angle drops below 5.5°', () => {
    const samples = [
        { velD: 25, flatteningDegS: 1, hMSL: 190, diveAngleDeg: 60 },
        { velD: 24, flatteningDegS: -4, hMSL: 180, diveAngleDeg: 55 },
        { velD: 20, flatteningDegS: -18, hMSL: 170, diveAngleDeg: 40 },
        { velD: 12, flatteningDegS: -22, hMSL: 160, diveAngleDeg: 15 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 150, diveAngleDeg: 4 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100, diveAngleDeg: 0.5 }
    ];
    const { idxA, idxB } = F.defaultSwoopCursorIndices(samples, F.CURSOR_B_DIVE_ANGLE_DEG, 20);
    assert.equal(idxA, 2);
    assert.equal(idxB, 4);
});

test('defaultSwoopCursorIndices B dive-angle threshold is configurable', () => {
    const samples = [
        { velD: 25, flatteningDegS: 1, hMSL: 190, diveAngleDeg: 60 },
        { velD: 24, flatteningDegS: -4, hMSL: 180, diveAngleDeg: 55 },
        { velD: 20, flatteningDegS: -18, hMSL: 170, diveAngleDeg: 40 },
        { velD: 12, flatteningDegS: -22, hMSL: 160, diveAngleDeg: 15 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 150, diveAngleDeg: 4 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100, diveAngleDeg: 0.5 }
    ];
    assert.equal(F.defaultSwoopCursorIndices(samples, 5.5, 20).idxB, 4);
    assert.equal(F.defaultSwoopCursorIndices(samples, 16, 20).idxB, 3);
    assert.equal(F.defaultSwoopCursorIndices(samples, 1, 20).idxB, 5);
});

test('defaultSwoopCursorIndices B is the later of dive-angle and below-2.49 m', () => {
    const angleLater = [
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 60 },
        { velD: 20, flatteningDegS: -18, hMSL: 110, diveAngleDeg: 40 },
        { velD: 12, flatteningDegS: -22, hMSL: 101.5, diveAngleDeg: 20 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.0, diveAngleDeg: 8 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.4, diveAngleDeg: 4 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 1 }
    ];
    const angleLaterCursors = F.defaultSwoopCursorIndices(angleLater, 5, 2);
    assert.equal(angleLaterCursors.idxA, 1);
    // 2nd below-2.49 m is idx 3; dive angle <5.5° is idx 4; both satisfied at 4.
    assert.equal(angleLaterCursors.idxB, 4);

    const altLater = [
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 60 },
        { velD: 20, flatteningDegS: -18, hMSL: 120, diveAngleDeg: 40 },
        { velD: 12, flatteningDegS: -22, hMSL: 110, diveAngleDeg: 20 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.0, diveAngleDeg: 3 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.4, diveAngleDeg: 2 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 1 }
    ];
    const altLaterCursors = F.defaultSwoopCursorIndices(altLater, 5, 2);
    assert.equal(altLaterCursors.idxA, 1);
    assert.equal(altLaterCursors.idxB, 4);
});

test('defaultSwoopCursorIndices B ignores a condition when its value is 0', () => {
    const samples = [
        { velD: 25, flatteningDegS: 1, hMSL: 140, diveAngleDeg: 60 },
        { velD: 20, flatteningDegS: -18, hMSL: 120, diveAngleDeg: 40 },
        { velD: 12, flatteningDegS: -22, hMSL: 110, diveAngleDeg: 20 },
        { velD: 0.8, flatteningDegS: -5, hMSL: 101.0, diveAngleDeg: 3 },
        { velD: 0.4, flatteningDegS: 0, hMSL: 100.4, diveAngleDeg: 2 },
        { velD: 0.2, flatteningDegS: 0, hMSL: 100.0, diveAngleDeg: 1 }
    ];
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 2).idxB, 4);
    assert.equal(F.defaultSwoopCursorIndices(samples, 5, 0).idxB, 3);
    assert.equal(F.defaultSwoopCursorIndices(samples, 0, 2).idxB, 4);
});

test('isCsvFile matches .csv names and text/csv types', () => {
    assert.equal(F.isCsvFile({ name: '08-47-08.CSV' }), true);
    assert.equal(F.isCsvFile({ name: 'track.csv', type: '' }), true);
    assert.equal(F.isCsvFile({ name: 'CONFIG.TXT' }), false);
    assert.equal(F.isCsvFile({ name: 'notes', type: 'text/csv' }), true);
    assert.equal(F.isCsvFile(null), false);
});

test('collectCsvFilesFromFileList keeps only CSV files', () => {
    const files = F.collectCsvFilesFromFileList([
        { name: 'a.CSV' },
        { name: 'CONFIG.TXT' },
        { name: 'b.csv' },
        { name: 'readme.md', type: 'text/csv' }
    ]);
    assert.equal(files.map(f => f.name).join(','), 'a.CSV,b.csv,readme.md');
    assert.equal(F.collectCsvFilesFromFileList(null).length, 0);
});

function mockFileEntry(name, type = '') {
    return {
        isFile: true,
        isDirectory: false,
        name,
        file(ok) {
            ok({ name, type });
        }
    };
}

function mockDirEntry(name, children, chunkSize = 2) {
    return {
        isFile: false,
        isDirectory: true,
        name,
        createReader() {
            let offset = 0;
            return {
                readEntries(ok) {
                    const batch = children.slice(offset, offset + chunkSize);
                    offset += chunkSize;
                    ok(batch);
                }
            };
        }
    };
}

test('collectCsvFilesFromEntries walks nested folders and skips non-csv', async () => {
    const nested = mockDirEntry('24-03-09', [
        mockFileEntry('16-53-00.CSV'),
        mockFileEntry('CONFIG.TXT'),
        mockFileEntry('17-22-11.csv')
    ]);
    const root = mockDirEntry('TRACK', [
        nested,
        mockFileEntry('loose.CSV'),
        mockFileEntry('notes.txt')
    ]);
    const files = await F.collectCsvFilesFromEntries([root, mockFileEntry('extra.csv')]);
    assert.equal(
        files.map(f => f.name).join(','),
        '16-53-00.CSV,17-22-11.csv,loose.CSV,extra.csv'
    );
});

test('collectCsvFilesFromEntries paginates directory reads', async () => {
    const children = [
        mockFileEntry('a.csv'),
        mockFileEntry('b.csv'),
        mockFileEntry('c.csv'),
        mockFileEntry('d.txt'),
        mockFileEntry('e.csv')
    ];
    const dir = mockDirEntry('day', children, 2);
    const files = await F.collectCsvFilesFromEntries([dir]);
    assert.equal(files.map(f => f.name).join(','), 'a.csv,b.csv,c.csv,e.csv');
});

test('collectCsvFilesFromEntries processes dropped folders in increasing name order', async () => {
    const later = mockDirEntry('24-03-10', [mockFileEntry('later.csv')]);
    const earlier = mockDirEntry('24-03-08', [mockFileEntry('earlier.csv')]);
    const mid = mockDirEntry('24-03-09', [mockFileEntry('mid.csv')]);
    const files = await F.collectCsvFilesFromEntries([later, earlier, mid]);
    assert.equal(files.map(f => f.name).join(','), 'earlier.csv,mid.csv,later.csv');
});

test('collectCsvFilesFromEntries sorts nested folders by name', async () => {
    const root = mockDirEntry('TRACK', [
        mockDirEntry('24-03-10', [mockFileEntry('c.csv')]),
        mockDirEntry('24-03-08', [mockFileEntry('a.csv')]),
        mockFileEntry('loose.csv'),
        mockDirEntry('24-03-09', [mockFileEntry('b.csv')])
    ]);
    const files = await F.collectCsvFilesFromEntries([root]);
    assert.equal(files.map(f => f.name).join(','), 'a.csv,b.csv,c.csv,loose.csv');
});

test('collectCsvFilesFromFileList orders files by dropped folder name', () => {
    const files = F.collectCsvFilesFromFileList([
        { name: 'z.csv', webkitRelativePath: '24-03-10/z.csv' },
        { name: 'a.csv', webkitRelativePath: '24-03-08/a.csv' },
        { name: 'm.csv', webkitRelativePath: '24-03-09/m.csv' }
    ]);
    assert.equal(files.map(f => f.name).join(','), 'a.csv,m.csv,z.csv');
});

function mockFsFileHandle(name) {
    const handle = {
        name,
        kind: 'file',
        getFileCount: 0,
        async getFile() {
            handle.getFileCount += 1;
            return { name, type: '', webkitRelativePath: '' };
        }
    };
    return handle;
}

function mockFsDirHandle(name, children) {
    return {
        name,
        kind: 'directory',
        async *values() {
            for (const child of children) yield child;
        }
    };
}

test('collectCsvFilesFromDirectoryHandle skips non-csv before reading', async () => {
    const csv = mockFsFileHandle('08-47-08.CSV');
    const txt = mockFsFileHandle('CONFIG.TXT');
    const laterCsv = mockFsFileHandle('09-00-00.csv');
    const notes = mockFsFileHandle('notes.txt');
    const root = mockFsDirHandle('FLY', [
        mockFsDirHandle('24-03-09', [laterCsv]),
        notes,
        mockFsDirHandle('24-03-08', [txt, csv])
    ]);
    const files = await F.collectCsvFilesFromDirectoryHandle(root);
    assert.equal(files.map(f => f.name).join(','), '08-47-08.CSV,09-00-00.csv');
    assert.equal(txt.getFileCount, 0);
    assert.equal(notes.getFileCount, 0);
    assert.equal(csv.getFileCount, 1);
    assert.equal(files[0].webkitRelativePath, 'FLY/24-03-08/08-47-08.CSV');
    assert.equal(files[1].webkitRelativePath, 'FLY/24-03-09/09-00-00.csv');
});

test('collectCsvFilesFromDirectoryHandle returns empty for missing handles', async () => {
    assert.equal((await F.collectCsvFilesFromDirectoryHandle(null)).length, 0);
});

test('fileListHasRelativePaths is true only when a folder tree is present', () => {
    assert.equal(F.fileListHasRelativePaths(null), false);
    assert.equal(F.fileListHasRelativePaths([{ name: 'a.csv', webkitRelativePath: '' }]), false);
    assert.equal(F.fileListHasRelativePaths([{ name: 'a.csv', webkitRelativePath: 'a.csv' }]), false);
    assert.equal(F.fileListHasRelativePaths([
        { name: 'a.csv', webkitRelativePath: '' },
        { name: 'b.csv', webkitRelativePath: 'TRACKS/25-09-16/b.csv' }
    ]), true);
});

test('buildVirtualTreeFromFileList nests Flysight TRACKS days and hides non-csv', () => {
    const tree = F.buildVirtualTreeFromFileList([
        { name: 'CONFIG.TXT', webkitRelativePath: 'TRACKS/CONFIG.TXT' },
        { name: '09-12-33.CSV', webkitRelativePath: 'TRACKS/25-09-16/09-12-33.CSV' },
        { name: '08-47-08.CSV', webkitRelativePath: 'TRACKS/25-09-16/08-47-08.CSV' },
        { name: '10-00-00.csv', webkitRelativePath: 'TRACKS/25-09-17/10-00-00.csv' }
    ]);
    assert.equal(tree.name, 'TRACKS');
    assert.equal(tree.path, 'TRACKS');
    assert.equal(tree.files.length, 0);
    assert.equal(tree.dirs.map(d => d.name).join(','), '25-09-16,25-09-17');
    const listed = F.listVirtualTreeBrowserEntries(tree);
    assert.equal(listed.dirs.map(d => d.name).join(','), '25-09-16,25-09-17');
    assert.equal(listed.files.length, 0);
    const day = tree.dirs[0];
    assert.equal(day.path, 'TRACKS/25-09-16');
    assert.equal(day.files.map(f => f.name).join(','), '08-47-08.CSV,09-12-33.CSV');
    assert.equal(F.countCsvFilesFromVirtualNode(tree), 3);
    assert.equal(F.countCsvFilesFromVirtualNode(day), 2);
});

test('collectCsvFilesFromVirtualNode imports the current folder recursively', () => {
    const tree = F.buildVirtualTreeFromFileList([
        { name: '08-47-08.CSV', webkitRelativePath: 'TRACKS/25-09-16/08-47-08.CSV' },
        { name: '09-12-33.CSV', webkitRelativePath: 'TRACKS/25-09-16/09-12-33.CSV' },
        { name: '10-00-00.csv', webkitRelativePath: 'TRACKS/25-09-17/10-00-00.csv' }
    ]);
    const day = tree.dirs.find(d => d.name === '25-09-16');
    const files = F.collectCsvFilesFromVirtualNode(day);
    assert.equal(files.map(f => f.name).join(','), '08-47-08.CSV,09-12-33.CSV');
    const all = F.collectCsvFilesFromVirtualNode(tree);
    assert.equal(all.map(f => f.name).join(','), '08-47-08.CSV,09-12-33.CSV,10-00-00.csv');
});

test('buildVirtualTreeFromFileList keeps CSVs at the selected folder root', () => {
    const tree = F.buildVirtualTreeFromFileList([
        { name: 'loose.csv', webkitRelativePath: 'TRACKS/loose.csv' },
        { name: 'a.csv', webkitRelativePath: 'TRACKS/25-09-16/a.csv' }
    ]);
    assert.equal(tree.name, 'TRACKS');
    assert.equal(tree.files.map(f => f.name).join(','), 'loose.csv');
    assert.equal(tree.dirs.map(d => d.name).join(','), '25-09-16');
});

test('buildVirtualTreeFromFileList returns an empty root for missing lists', () => {
    const tree = F.buildVirtualTreeFromFileList(null);
    assert.equal(tree.dirs.length, 0);
    assert.equal(tree.files.length, 0);
    assert.equal(F.collectCsvFilesFromVirtualNode(tree).length, 0);
});

test('listDirectoryHandleBrowserEntries lists dirs then csv without reading', async () => {
    const csv = mockFsFileHandle('08-47-08.CSV');
    const txt = mockFsFileHandle('CONFIG.TXT');
    const nested = mockFsFileHandle('later.csv');
    const day = mockFsDirHandle('25-09-16', [nested]);
    const root = mockFsDirHandle('TRACKS', [txt, csv, day]);
    const listed = await F.listDirectoryHandleBrowserEntries(root);
    assert.equal(listed.dirs.map(d => d.name).join(','), '25-09-16');
    assert.equal(listed.files.map(f => f.name).join(','), '08-47-08.CSV');
    assert.equal(csv.getFileCount, 0);
    assert.equal(txt.getFileCount, 0);
    assert.equal(nested.getFileCount, 0);
    assert.equal(await F.countCsvFilesFromDirectoryHandle(root), 2);
});

test('readCsvFilesFromFileHandles reads only the chosen csv handles', async () => {
    const csv = mockFsFileHandle('08-47-08.CSV');
    const txt = mockFsFileHandle('CONFIG.TXT');
    const files = await F.readCsvFilesFromFileHandles([txt, csv], 'TRACKS/25-09-16');
    assert.equal(files.map(f => f.name).join(','), '08-47-08.CSV');
    assert.equal(files[0].webkitRelativePath, 'TRACKS/25-09-16/08-47-08.CSV');
    assert.equal(csv.getFileCount, 1);
    assert.equal(txt.getFileCount, 0);
});
