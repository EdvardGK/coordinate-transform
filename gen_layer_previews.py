"""Generate per-layer geometry preview PNGs.

Reads the merged master DXF and produces one PNG per theme layer,
plus one per source layer. Output goes to docs/layers/ for use in Notion.

Usage:
    python gen_layer_previews.py
"""
import os, math, ezdxf, json, unicodedata, re, io, base64
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

MASTER_DXF = (
    'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/'
    '10016 - Kistefos/03_Ut/08_Tegninger/DXF_NTM/Global/'
    'KNM_BIMK_MASTER_DATA_NTM10_GLOBAL_METERS.dxf'
)

SOURCE_DIR = (
    'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/'
    '10016 - Kistefos/03_Ut/08_Tegninger/DXF_NTM/Global/'
    'KNM_BIMK_MASTER_DATA_NTM10_GLOBAL_METERS'
)

DOCS_DIR = os.path.join(os.path.dirname(__file__), 'docs', 'layers')

import sys
sys.path.insert(0, os.path.dirname(__file__))
from merge_master_lines import (
    THEMES, normalize_layer, layer_to_theme, normalize_short,
    LAYER_SKIP_COMPILED, INCLUDED_FILES,
)

ACI_COLORS = {
    1: '#ff4444', 2: '#ffff44', 3: '#44ff44', 4: '#44ffff',
    5: '#4444ff', 6: '#ff44ff', 7: '#ffffff', 8: '#808080',
    9: '#c0c0c0', 30: '#ff8800', 34: '#00aaff', 40: '#ffaa80',
    42: '#aa8855', 54: '#aa7733', 56: '#bbaa44', 84: '#88ff44',
    92: '#446633', 94: '#66cc88', 140: '#4488aa', 174: '#cc8866',
    200: '#aa44cc', 220: '#ee88cc', 250: '#666666',
}


def extract_segments(entities):
    """Extract line segments from DXF entities."""
    segs = []
    for e in entities:
        t = e.dxftype()
        if t == 'LINE':
            segs.append(((e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)))
        elif t == 'LWPOLYLINE':
            pts = list(e.get_points(format='xy'))
            for i in range(len(pts) - 1):
                segs.append((pts[i], pts[i+1]))
            if e.closed and len(pts) > 1:
                segs.append((pts[-1], pts[0]))
        elif t == 'POLYLINE':
            pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices
                   if v.dxf.location is not None]
            for i in range(len(pts) - 1):
                segs.append((pts[i], pts[i+1]))
        elif t == 'CIRCLE':
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            n = 24
            pts = [(cx + r*math.cos(2*math.pi*i/n), cy + r*math.sin(2*math.pi*i/n)) for i in range(n+1)]
            for i in range(n):
                segs.append((pts[i], pts[i+1]))
        elif t == 'ARC':
            try:
                cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
                sa, ea = math.radians(e.dxf.start_angle), math.radians(e.dxf.end_angle)
                if ea < sa: ea += 2*math.pi
                n = max(8, int((ea-sa)/(2*math.pi)*32))
                pts = [(cx + r*math.cos(sa+(ea-sa)*i/n), cy + r*math.sin(sa+(ea-sa)*i/n)) for i in range(n+1)]
                for i in range(n):
                    segs.append((pts[i], pts[i+1]))
            except:
                pass
        elif t == 'POINT':
            px, py = e.dxf.location.x, e.dxf.location.y
            d = 0.3
            segs.append(((px-d, py), (px+d, py)))
            segs.append(((px, py-d), (px, py+d)))
    return segs


def render_png(segments, color, title, out_path, figsize=(10, 7), dpi=150):
    """Render line segments to a PNG with dark background."""
    if not segments:
        return
    fig, ax = plt.subplots(figsize=figsize, facecolor='#0d1117')
    ax.set_facecolor('#0d1117')
    ax.set_aspect('equal')
    ax.axis('off')
    lc = LineCollection(segments, colors=color, linewidths=0.3, alpha=0.85)
    ax.add_collection(lc)
    ax.autoscale_view()
    ax.set_title(title, color='#c9d1d9', fontsize=10, pad=8)
    plt.tight_layout(pad=0.3)
    fig.savefig(out_path, dpi=dpi, facecolor='#0d1117', bbox_inches='tight', pad_inches=0.15)
    plt.close(fig)


def safe_filename(name):
    name = unicodedata.normalize('NFC', name)
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r'\s+', '_', name)
    return name[:80]


def main():
    os.makedirs(DOCS_DIR, exist_ok=True)

    print('Reading master DXF...')
    master = ezdxf.readfile(MASTER_DXF)
    msp = master.modelspace()

    theme_entities = {}
    for e in msp:
        layer = e.dxf.layer
        if layer == 'KNM_LEGEND':
            continue
        theme_entities.setdefault(layer, []).append(e)

    print(f'Master: {sum(len(v) for v in theme_entities.values())} entities, {len(theme_entities)} layers')

    # Read source DXFs
    print('Reading source DXFs...')
    source_layers = {}  # (norm_layer, source_short, theme) -> [entities]
    if os.path.exists(SOURCE_DIR):
        for f in os.listdir(SOURCE_DIR):
            if not f.endswith('.dxf'):
                continue
            short = normalize_short(unicodedata.normalize('NFC', f.split('_NTM10')[0]))
            if not any(s in f for s in INCLUDED_FILES):
                continue
            path = os.path.join(SOURCE_DIR, f)
            try:
                doc = ezdxf.readfile(path)
            except:
                continue
            for e in doc.modelspace():
                if e.dxftype() not in ('LINE', 'LWPOLYLINE', 'POLYLINE', 'ARC', 'CIRCLE', 'POINT', 'ELLIPSE'):
                    continue
                orig = e.dxf.layer
                norm = normalize_layer(orig)
                skip = any(s in short and p.search(norm) for s, p in LAYER_SKIP_COMPILED)
                if skip:
                    continue
                theme = layer_to_theme(orig, short)
                key = (norm, short, theme)
                source_layers.setdefault(key, []).append(e)

    print(f'Source: {len(source_layers)} layer groups')

    # Group by theme
    theme_sources = {}
    for (norm, short, theme), ents in source_layers.items():
        theme_sources.setdefault(theme, []).append((norm, short, ents))

    manifest = []

    for theme, (aci, desc) in THEMES.items():
        entities = theme_entities.get(theme, [])
        if not entities:
            continue

        color = ACI_COLORS.get(aci, '#ffffff')
        segs = extract_segments(entities)
        if not segs:
            continue

        print(f'  {theme}: {len(entities)} entities -> {len(segs)} segments')

        # Theme PNG
        theme_png = f'{safe_filename(theme)}.png'
        render_png(segs, color, f'{theme}  ({len(entities):,} entities)', os.path.join(DOCS_DIR, theme_png))

        # Source layer PNGs
        theme_dir = os.path.join(DOCS_DIR, safe_filename(theme))
        src_groups = theme_sources.get(theme, [])

        # Group FKB individual trees
        tree_pat = re.compile(r'^(T.rr)?(Gran|Furu|Lauvtre|Bj.rk|Eik|Hegg|Rogn|L.nn|Or)[_\b-]', re.IGNORECASE)
        grouped_trees = []
        regular = []
        for norm, short, ents in src_groups:
            if tree_pat.match(norm):
                grouped_trees.extend(ents)
            else:
                regular.append((norm, short, ents))
        if grouped_trees:
            regular.append(('FKB_Individual_Trees', 'Nairy_FKB', grouped_trees))

        source_pngs = []
        if regular:
            os.makedirs(theme_dir, exist_ok=True)
        for norm, short, ents in sorted(regular, key=lambda x: -len(x[2])):
            s = extract_segments(ents)
            if not s:
                continue
            src_png = f'{safe_filename(norm)}__{safe_filename(short)}.png'
            render_png(s, color, f'{norm} ({short}, {len(ents):,})',
                       os.path.join(theme_dir, src_png), figsize=(8, 5), dpi=120)
            source_pngs.append({
                'name': norm, 'source_file': short,
                'entities': len(ents), 'filename': f'{safe_filename(theme)}/{src_png}',
            })

        manifest.append({
            'theme': theme, 'aci': aci, 'desc': desc,
            'entities': len(entities), 'filename': theme_png,
            'sources': source_pngs,
        })

    with open(os.path.join(DOCS_DIR, 'manifest.json'), 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    print(f'\nDone. {len(manifest)} theme PNGs + {sum(len(p["sources"]) for p in manifest)} source PNGs')


if __name__ == '__main__':
    main()
