#!/usr/bin/env python3
"""Regenerate js/location-gates-seed.js from Gates/*.CSV (one gate centroid per file)."""
import json
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GATES_DIR = os.path.join(ROOT, 'Gates')
OUT = os.path.join(ROOT, 'js', 'location-gates-seed.js')

GATE_FILE_TO_LOC = {
    'Eloy_020.CSV': 'loc_eloy',
    'Eloy_200.CSV': 'loc_eloy',
    'Eloy_Speed_Droite.CSV': 'loc_eloy',
    'SDH_Est.CSV': 'loc_lfoz',
    'SDH_Ouest.CSV': 'loc_lfoz',
    'SDH_Speed.CSV': 'loc_lfoz',
    'Klatovy_Speed.CSV': 'loc_klatovy',
    'Pretoria_Speed_Distance.CSV': 'loc_pretoria',
}


def label_from_filename(fn):
    base = fn.rsplit('.', 1)[0]
    return re.sub(r'_', ' ', base)


def parse_centroid(text):
    lines = text.replace('\ufeff', '').splitlines()
    lat_idx = lon_idx = header_idx = -1
    for i, line in enumerate(lines):
        cols = line.split(',')
        if 'lat' in cols and 'lon' in cols:
            lat_idx = cols.index('lat')
            lon_idx = cols.index('lon')
            header_idx = i
            break
    lats, lons = [], []
    for line in lines[header_idx + 1:]:
        line = line.strip()
        if not line or line.startswith('('):
            continue
        cols = line.split(',')
        try:
            lat = float(cols[lat_idx])
            lon = float(cols[lon_idx])
        except (ValueError, IndexError):
            continue
        lats.append(lat)
        lons.append(lon)
    if not lats:
        return None
    return sum(lats) / len(lats), sum(lons) / len(lons)


def main():
    seed = {}
    for fn in sorted(os.listdir(GATES_DIR)):
        if not fn.lower().endswith('.csv'):
            continue
        loc_id = GATE_FILE_TO_LOC.get(fn)
        if not loc_id:
            continue
        path = os.path.join(GATES_DIR, fn)
        with open(path, encoding='utf-8', errors='replace') as f:
            c = parse_centroid(f.read())
        if not c:
            continue
        lat, lon = c
        seed.setdefault(loc_id, []).append({
            'label': label_from_filename(fn),
            'lat': round(lat, 7),
            'lng': round(lon, 7),
            'source': f'seed:{fn}',
        })
    body = '/** Auto-generated from Gates/*.CSV — run scripts/build-location-gates-seed.py to refresh. */\n'
    body += 'window.LOCATION_GATES_SEED = ' + json.dumps(seed, indent=2) + ';\n'
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(body)
    print('Wrote', OUT)


if __name__ == '__main__':
    main()
