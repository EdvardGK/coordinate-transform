"""Generate one geometry preview PNG per source file used in build_master_v2.

For each source file, applies the same name-mapping + resolver logic as
build_master_v2, then renders ALL geometry from that file with one colour
per *target* (new English) layer.

All previews are rendered at the SAME world extent and SAME pixel dimensions
so they're directly comparable. The world extent is the union bounding box
of geometry across every source, padded slightly.

Output: docs/v2_source_previews/<source_label>.png
"""
import os, sys, math, re, unicodedata
import ezdxf
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection
from matplotlib.patches import Patch

sys.path.insert(0, os.path.dirname(__file__))
from build_master_v2 import (
    BYGNING_MAP, KART_MAP, STIER_MAP, LANDSCAPE_MAP, PARKERING_MAP,
    COORDINATION_LAYERS, DROP,
    make_resolver_single, make_resolver_map, resolver_kotelinjer, resolver_stier,
    explode_all_inserts, entity_in_site, color_for_layer, find_in_master,
    FDE_DIR,
)

OUT_DIR = os.path.join(os.path.dirname(__file__), 'docs', 'v2_source_previews')

ACI_COLORS_DARK_BG = {   # bright, for dark backdrop
    1:  '#ff4444', 2:  '#ffd700', 3:  '#44dd44', 4:  '#44dddd',
    5:  '#5577ff', 6:  '#ff44ff', 7:  '#dddddd', 8:  '#888888',
    9:  '#c0c0c0', 30: '#ff8800', 34: '#00aaff', 40: '#cc8855',
    140: '#a050d0',
}

ACI_COLORS_LIGHT_BG = {  # saturated/dark, for white backdrop
    1:  '#b40000', 2:  '#8a6b00', 3:  '#0c7c0c', 4:  '#006a6a',
    5:  '#1f3aa8', 6:  '#8b008b', 7:  '#202020', 8:  '#404040',
    9:  '#707070', 30: '#b55000', 34: '#005580', 40: '#704818',
    140: '#4a148c',
}

# Render settings — every preview uses these exact values.
# Landscape format, all previews share the same world extent and pixel dims.
FIGSIZE = (14, 8)        # landscape
DPI = 120                # → 1680 × 960 px
LEGEND_FONTSIZE = 7
LINEWIDTH_DENSE  = 0.9
LINEWIDTH_SPARSE = 1.8
MARKERSIZE_DENSE  = 4
MARKERSIZE_SPARSE = 14

# Files that get a REPORT-ONLY colour override so their sparse/single-layer
# content is actually visible. All previews use the dark backdrop — for these
# four, the main data layer is rendered in WHITE (with beefier lines), and
# the coordination markers get magenta / teal so they're visible but don't
# compete with the main data. Purely a preview-rendering choice; doesn't
# affect DWG colours.
REPORT_OVERRIDE_LABELS = {
    'Innmålt_Elektro',
    'Innmålt_Kunstinstallasjon',
    'Innmålt_Tre',
    'TreStammer_Diameter',
}
REPORT_MAIN_COLOR  = '#ffffff'   # white
REPORT_COORD_TEAL  = '#00c8c8'   # teal
REPORT_COORD_MAGENTA = '#ff30c0' # magenta

def aci_to_hex(aci):
    return ACI_COLORS_DARK_BG.get(aci, '#ffffff')

def extract_segments(e):
    t = e.dxftype()
    segs = []
    try:
        if t == 'LINE':
            segs.append(((e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)))
        elif t == 'LWPOLYLINE':
            pts = list(e.get_points(format='xy'))
            for i in range(len(pts) - 1):
                segs.append((pts[i], pts[i+1]))
            if e.closed and len(pts) > 1:
                segs.append((pts[-1], pts[0]))
        elif t == 'POLYLINE':
            pts = []
            for v in e.vertices:
                if v.dxf.flags & 128:
                    continue
                pts.append((v.dxf.location.x, v.dxf.location.y))
            for i in range(len(pts) - 1):
                segs.append((pts[i], pts[i+1]))
        elif t == 'CIRCLE':
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            n = 24
            pts = [(cx + r*math.cos(2*math.pi*i/n), cy + r*math.sin(2*math.pi*i/n))
                   for i in range(n+1)]
            for i in range(n):
                segs.append((pts[i], pts[i+1]))
        elif t == 'ARC':
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            sa, ea = math.radians(e.dxf.start_angle), math.radians(e.dxf.end_angle)
            if ea < sa: ea += 2*math.pi
            n = max(8, int((ea-sa)/(2*math.pi)*32))
            pts = [(cx + r*math.cos(sa+(ea-sa)*i/n),
                    cy + r*math.sin(sa+(ea-sa)*i/n)) for i in range(n+1)]
            for i in range(n):
                segs.append((pts[i], pts[i+1]))
        elif t == 'POINT':
            px, py = e.dxf.location.x, e.dxf.location.y
            d = 0.3
            segs.append(((px-d, py), (px+d, py)))
            segs.append(((px, py-d), (px, py+d)))
        elif t == '3DFACE':
            pts = []
            for i in range(4):
                v = getattr(e.dxf, f'vtx{i}')
                pts.append((v.x, v.y))
            for i in range(len(pts)):
                segs.append((pts[i], pts[(i+1) % len(pts)]))
    except Exception:
        pass
    return segs

def text_xy(e):
    t = e.dxftype()
    try:
        if t in ('TEXT', 'MTEXT'):
            return (e.dxf.insert.x, e.dxf.insert.y)
    except Exception:
        pass
    return None

def collect_source(src_path, resolver, label):
    """Read a source file and return:
       (layer_segments, layer_text_pts, total_entities, bbox_or_None)
       bbox is (xmin, xmax, ymin, ymax) of all kept geometry.
    """
    if not src_path or not os.path.exists(src_path):
        return None
    try:
        doc = ezdxf.readfile(src_path)
    except Exception as ex:
        print(f'    READ FAILED: {ex}')
        return None
    msp = doc.modelspace()
    explode_all_inserts(msp)

    layer_segments = {}
    layer_text_pts = {}
    xmin = ymin =  float('inf')
    xmax = ymax = -float('inf')

    for e in msp:
        t = e.dxftype()
        if t in ('HATCH', 'INSERT'):
            continue
        if not entity_in_site(e):
            continue
        try:
            tgt = resolver(e)
        except Exception:
            tgt = None
        if tgt is None or tgt is DROP:
            continue

        segs = extract_segments(e)
        for (x1, y1), (x2, y2) in segs:
            if x1 < xmin: xmin = x1
            if x2 < xmin: xmin = x2
            if x1 > xmax: xmax = x1
            if x2 > xmax: xmax = x2
            if y1 < ymin: ymin = y1
            if y2 < ymin: ymin = y2
            if y1 > ymax: ymax = y1
            if y2 > ymax: ymax = y2
        if segs:
            layer_segments.setdefault(tgt, []).extend(segs)

        tp = text_xy(e)
        if tp is not None:
            x, y = tp
            if x < xmin: xmin = x
            if x > xmax: xmax = x
            if y < ymin: ymin = y
            if y > ymax: ymax = y
            layer_text_pts.setdefault(tgt, []).append(tp)

    total = sum(len(v) for v in layer_segments.values()) + sum(len(v) for v in layer_text_pts.values())
    bbox = None
    if xmin != float('inf'):
        bbox = (xmin, xmax, ymin, ymax)
    return layer_segments, layer_text_pts, total, bbox

def render(label, layer_segments, layer_text_pts, world_extent, out_path):
    """Render one source preview at the fixed FIGSIZE/DPI/world-extent on a
    dark backdrop. Files in REPORT_OVERRIDE_LABELS swap their ACI colours for
    white + magenta + teal so sparse single-category content stands out."""
    total_entities = sum(len(s) for s in layer_segments.values()) + \
                     sum(len(t) for t in layer_text_pts.values())
    override = label in REPORT_OVERRIDE_LABELS

    bg = '#0d1117'
    fg = '#c9d1d9'
    legend_bg = '#0d1117'
    legend_edge = '#333'
    lw = LINEWIDTH_SPARSE if override else LINEWIDTH_DENSE
    ms = MARKERSIZE_SPARSE if override else MARKERSIZE_DENSE
    alpha = 1.0

    fig, ax = plt.subplots(figsize=FIGSIZE, facecolor=bg)
    ax.set_facecolor(bg)
    ax.set_aspect('equal')
    ax.axis('off')

    sorted_layers = sorted(
        layer_segments.keys(),
        key=lambda l: -(len(layer_segments[l]) + len(layer_text_pts.get(l, [])))
    )

    # For override files, assign coordination-marker layers to magenta/teal
    # deterministically (alternating) so each marker is distinguishable.
    coord_cycle = [REPORT_COORD_MAGENTA, REPORT_COORD_TEAL]
    coord_idx = {}

    legend_handles = []
    for tgt_layer in sorted_layers:
        if override:
            if tgt_layer.startswith('COORDINATION_'):
                if tgt_layer not in coord_idx:
                    coord_idx[tgt_layer] = coord_cycle[len(coord_idx) % len(coord_cycle)]
                col = coord_idx[tgt_layer]
            else:
                col = REPORT_MAIN_COLOR
        else:
            col = aci_to_hex(color_for_layer(tgt_layer))
        segs = layer_segments[tgt_layer]
        if segs:
            ax.add_collection(LineCollection(segs, colors=col,
                                             linewidths=lw, alpha=alpha))
        tpts = layer_text_pts.get(tgt_layer, [])
        if tpts:
            xs = [p[0] for p in tpts]
            ys = [p[1] for p in tpts]
            ax.scatter(xs, ys, c=col, s=ms, alpha=alpha, marker='.', linewidths=0)
        n = len(segs) + len(tpts)
        name = tgt_layer if len(tgt_layer) <= 32 else tgt_layer[:30] + '…'
        legend_handles.append(Patch(facecolor=col, edgecolor='none',
                                    label=f'{name} ({n:,})'))

    xmin, xmax, ymin, ymax = world_extent
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)

    title = f'{label}  —  {total_entities:,} drawn entities  •  {len(sorted_layers)} target layers'
    ax.set_title(title, color=fg, fontsize=11, pad=8, loc='left')

    if legend_handles:
        ax.legend(
            handles=legend_handles, loc='upper right',
            fontsize=LEGEND_FONTSIZE, frameon=True, framealpha=0.75,
            facecolor=legend_bg, edgecolor=legend_edge,
            labelcolor=fg,
            ncol=1 if len(legend_handles) <= 18 else 2,
        )

    plt.tight_layout(pad=0.4)
    fig.savefig(out_path, dpi=DPI, facecolor=bg,
                bbox_inches=None, pad_inches=0.15)
    plt.close(fig)

def union_extent(bboxes, pad_frac=0.03, target_aspect=FIGSIZE[0]/FIGSIZE[1]):
    """Union of bboxes, padded, then expanded to match the figure aspect so
    every preview fills the canvas consistently. Pads equally after aspect
    expansion."""
    xmins = [b[0] for b in bboxes]
    xmaxs = [b[1] for b in bboxes]
    ymins = [b[2] for b in bboxes]
    ymaxs = [b[3] for b in bboxes]
    xmin, xmax = min(xmins), max(xmaxs)
    ymin, ymax = min(ymins), max(ymaxs)
    w = xmax - xmin
    h = ymax - ymin
    cx, cy = (xmin + xmax) / 2, (ymin + ymax) / 2
    if w / h < target_aspect:
        w = h * target_aspect
    else:
        h = w / target_aspect
    w *= 1 + 2 * pad_frac
    h *= 1 + 2 * pad_frac
    return (cx - w/2, cx + w/2, cy - h/2, cy + h/2)

def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    # (path, resolver, label, extent_contributor)
    # extent_contributor=False for FKB files so they don't inflate the view —
    # the project area is the consultant footprint (~1 km), not the municipal
    # map area (2 km+). FKB previews are rendered WITHIN the project view,
    # clipping municipal content outside the site.
    sources = [
        (os.path.join(FDE_DIR, '3DE_Bygning_fra fkb.dxf'),
            make_resolver_map(BYGNING_MAP), '3DE_Bygning', False),
        (os.path.join(FDE_DIR, '3DE_Kart fra fkb.dxf'),
            make_resolver_map(KART_MAP), '3DE_Kart', False),
        (find_in_master('Innmålt_Elektro_NTM10_global_meters.dxf'),
            make_resolver_single('Surveyed_Electrical'), 'Innmålt_Elektro', True),
        (find_in_master('Innmålt_Kunstinstallasjon_NTM10_global_meters.dxf'),
            make_resolver_single('Surveyed_ArtInstallation'), 'Innmålt_Kunstinstallasjon', True),
        (find_in_master('Innmålt_Tre_NTM10_global_meters.dxf'),
            make_resolver_single('Tree_Surveyed'), 'Innmålt_Tre', True),
        (find_in_master('TreStammer_Diameter_NTM10_global_meters.dxf'),
            make_resolver_single('Tree_Trunk'), 'TreStammer_Diameter', True),
        (find_in_master('Kotelinjer_MedPåskrift_2D_NTM10_global_meters.dxf'),
            resolver_kotelinjer, 'Kotelinjer_MedPåskrift_2D', True),
        (find_in_master('KNM_Stier_NTM10_global.dxf'),
            resolver_stier, 'KNM_Stier', True),
        (find_in_master('landscape plan visitor centre_NTM10_global_meters.dxf'),
            make_resolver_map(LANDSCAPE_MAP), 'landscape_plan_visitor_centre', True),
        (find_in_master('Parkering_kistefos_LARK 1_NTM10_global_meters.dxf'),
            make_resolver_map(PARKERING_MAP), 'Parkering_kistefos_LARK_1', True),
    ]

    print(f'Pass 1: collecting geometry + bboxes...')
    collected = []
    bboxes = []
    for src, resolver, label, extent_contributor in sources:
        print(f'  {label}')
        result = collect_source(src, resolver, label)
        if result is None:
            print(f'    SKIP')
            continue
        layer_segments, layer_text_pts, total, bbox = result
        print(f'    layers={len(layer_segments)}  entities={total}  '
              f'bbox={None if bbox is None else tuple(round(v,1) for v in bbox)}'
              + ('  [extent contributor]' if extent_contributor else ''))
        if bbox is not None and extent_contributor:
            bboxes.append(bbox)
        collected.append((label, layer_segments, layer_text_pts))

    if not bboxes:
        print('No geometry collected. Aborting.')
        return

    extent = union_extent(bboxes)
    w_px = int(FIGSIZE[0] * DPI)
    h_px = int(FIGSIZE[1] * DPI)
    print(f'\nFixed render extent: '
          f'x [{extent[0]:.1f}..{extent[1]:.1f}] ({extent[1]-extent[0]:.0f} m)  '
          f'y [{extent[2]:.1f}..{extent[3]:.1f}] ({extent[3]-extent[2]:.0f} m)')
    print(f'Figure size: {FIGSIZE[0]}" × {FIGSIZE[1]}"  →  {w_px} × {h_px} px (landscape)')

    print(f'\nPass 2: rendering at fixed extent → {OUT_DIR}')
    for label, layer_segments, layer_text_pts in collected:
        safe_label = re.sub(r'[<>:"/\\|?*\s]', '_', unicodedata.normalize('NFC', label))
        out_path = os.path.join(OUT_DIR, f'{safe_label}.png')
        render(label, layer_segments, layer_text_pts, extent, out_path)
        print(f'  {safe_label}.png')

    print(f'\nDone. {len(collected)} previews written.')

if __name__ == '__main__':
    main()
