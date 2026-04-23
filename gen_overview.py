"""Generate HTML overview with matplotlib geometry previews, layer-colored.
Cards show geometry. Clicking opens popup with legend."""
import json, os, ezdxf, math, base64, io
from collections import Counter
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection

BASE = 'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/10016 - Kistefos'
DXF_CONVERTED = 'c:/Users/edkjo/repos/coordinate-transform/tmp_dxf_all'

with open('c:/Users/edkjo/repos/coordinate-transform/acc_analysis.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

results = data['results']

crs_colors = {
    'NTM10_global_m': '#27ae60', 'NTM10_global_mm': '#2ecc71',
    'UTM32_m': '#2980b9', 'UTM32_mm': '#3498db',
    'LOCAL_m': '#f39c12', 'LOCAL_mm': '#e67e22',
    'UNKNOWN': '#e74c3c', 'NO_COORDS': '#95a5a6',
}

LAYER_PALETTE = [
    '#4a9eff', '#ff6b6b', '#51cf66', '#ffd43b', '#cc5de8',
    '#ff922b', '#20c997', '#f06595', '#7950f2', '#15aabf',
    '#e64980', '#82c91e', '#fab005', '#4c6ef5', '#fd7e14',
    '#38d9a9', '#d6336c', '#748ffc', '#a9e34b', '#f783ac',
]

def fmt_size(b):
    if b > 1_000_000: return f'{b/1_000_000:.1f} MB'
    elif b > 1000: return f'{b/1000:.0f} KB'
    return f'{b} B'


def render_preview(dxf_path, small=True):
    """Render a DXF file to a base64 PNG using matplotlib. Returns (b64_small, b64_large, layer_map)."""
    try:
        doc = ezdxf.readfile(dxf_path)
        msp = doc.modelspace()
    except:
        return None, None, {}

    layer_lines = {}  # {layer: [((x1,y1),(x2,y2)), ...]}
    layer_colors = {}
    max_ent = 8000
    count = 0

    for e in msp:
        count += 1
        if count > max_ent:
            break
        t = e.dxftype()
        lyr = e.dxf.layer
        if lyr not in layer_colors:
            layer_colors[lyr] = LAYER_PALETTE[len(layer_colors) % len(LAYER_PALETTE)]
        if lyr not in layer_lines:
            layer_lines[lyr] = []

        if t == 'LINE':
            layer_lines[lyr].append(((e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)))
        elif t == 'LWPOLYLINE':
            pts = list(e.get_points(format='xy'))
            for i in range(len(pts) - 1):
                layer_lines[lyr].append((pts[i], pts[i + 1]))
            if e.closed and len(pts) > 1:
                layer_lines[lyr].append((pts[-1], pts[0]))
        elif t == 'CIRCLE':
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            segs = 16
            pts = [(cx + r * math.cos(2 * math.pi * i / segs),
                     cy + r * math.sin(2 * math.pi * i / segs)) for i in range(segs + 1)]
            for i in range(len(pts) - 1):
                layer_lines[lyr].append((pts[i], pts[i + 1]))
        elif t == 'ARC':
            try:
                cx, cy = e.dxf.center.x, e.dxf.center.y
                r = e.dxf.radius
                sa, ea = math.radians(e.dxf.start_angle), math.radians(e.dxf.end_angle)
                if ea < sa: ea += 2 * math.pi
                segs = max(8, int((ea - sa) / (2 * math.pi) * 32))
                pts = [(cx + r * math.cos(sa + (ea - sa) * i / segs),
                         cy + r * math.sin(sa + (ea - sa) * i / segs)) for i in range(segs + 1)]
                for i in range(len(pts) - 1):
                    layer_lines[lyr].append((pts[i], pts[i + 1]))
            except:
                pass

    if not any(layer_lines.values()):
        return None, None, {}

    def make_fig(figsize, dpi):
        fig, ax = plt.subplots(figsize=figsize, facecolor='#0d1117')
        ax.set_facecolor('#0d1117')
        ax.set_aspect('equal')
        ax.axis('off')

        for lyr, segs in layer_lines.items():
            if not segs:
                continue
            col = layer_colors[lyr]
            lc = LineCollection(segs, colors=col, linewidths=0.4, alpha=0.8)
            ax.add_collection(lc)

        ax.autoscale_view()
        plt.tight_layout(pad=0.2)

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=dpi, facecolor='#0d1117', bbox_inches='tight', pad_inches=0.1)
        plt.close(fig)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode()

    b64_small = make_fig((2.8, 1.8), 80)
    b64_large = make_fig((8, 5), 120)
    return b64_small, b64_large, layer_colors


def find_dxf_for(r):
    full = os.path.join(BASE, r['path'])
    if full.lower().endswith('.dxf') and os.path.exists(full):
        return full
    name = os.path.splitext(os.path.basename(r['path']))[0]
    converted = os.path.join(DXF_CONVERTED, name + '.dxf')
    if os.path.exists(converted):
        return converted
    orig_dxf = os.path.join(BASE, '02_Arbeid/09_Tegninger/DXF_Original', name + '.dxf')
    if os.path.exists(orig_dxf):
        return orig_dxf
    return None


# Build cards
import html as html_mod
cards = []
for i, r in enumerate(sorted(results, key=lambda x: x['path'])):
    fname = os.path.basename(r['path'])
    folder = os.path.dirname(r['path']).replace('\\', '/')
    color = crs_colors.get(r['crs'], '#bdc3c7')
    fmt = r.get('format', 'DXF')
    card_id = f'card-{i}'

    dxf_path = find_dxf_for(r)
    if dxf_path:
        b64_small, b64_large, layer_map = render_preview(dxf_path)
        if b64_small:
            img_small = f'<img src="data:image/png;base64,{b64_small}" style="width:100%;height:180px;object-fit:contain;">'
            img_large = f'<img src="data:image/png;base64,{b64_large}" style="width:100%;max-height:500px;object-fit:contain;">'
        else:
            img_small = '<div class="no-preview">Empty</div>'
            img_large = img_small
            layer_map = {}
        legend_items = ''.join(
            f'<div class="legend-item"><span class="legend-dot" style="background:{c}"></span>{html_mod.escape(l)}</div>'
            for l, c in sorted(layer_map.items())
        )
    else:
        img_small = '<div class="no-preview">No DXF available</div>'
        img_large = img_small
        legend_items = ''

    img_large_esc = html_mod.escape(img_large)
    legend_esc = html_mod.escape(legend_items)

    cards.append(f'''
    <div class="card" data-crs="{r['crs']}" data-text="{fname.lower()} {folder.lower()}"
         onclick="openPopup('{html_mod.escape(fname)}', `{img_large_esc}`, `{legend_esc}`, '{r['crs'].replace('_',' ')}', '{r['units']}', '{r['entities']}', '{fmt_size(r['size'])}', `{html_mod.escape(folder)}`)">
      <div class="card-header">
        <div class="crs-badge" style="background:{color}">{r['crs'].replace('_', ' ')}</div>
        <span class="fmt-badge {'fmt-dwg' if fmt == 'DWG' else 'fmt-dxf'}">{fmt}</span>
      </div>
      <div class="preview">{img_small}</div>
      <div class="filename">{fname}</div>
      <div class="folder">{folder}</div>
      <div class="meta">
        <span><b>Units:</b> {r['units']}</span>
        <span><b>Ent:</b> {r['entities']}</span>
        <span><b>Size:</b> {fmt_size(r['size'])}</span>
      </div>
    </div>''')

    if (i + 1) % 10 == 0:
        print(f'  Generated {i + 1}/{len(results)} cards...')

crs_counts = Counter(r['crs'] for r in results)
summary_items = ''.join(
    f'<span class="summary-item" style="background:{crs_colors.get(c, "#bdc3c7")}">{c.replace("_", " ")}: {n}</span>'
    for c, n in crs_counts.most_common()
)
crs_options = ''.join(
    f'<option value="{c}">{c.replace("_"," ")} ({n})</option>'
    for c, n in crs_counts.most_common()
)

html_out = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Kistefos ACC - File CRS Analysis</title>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #0d1117; color: #eee; padding: 20px; }}
  h1 {{ margin-bottom: 10px; font-size: 1.4em; color: #fff; }}
  .summary {{ display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 20px; }}
  .summary-item {{ padding: 4px 12px; border-radius: 12px; font-size: 0.8em; color: #fff; font-weight: 600; }}
  .filter-bar {{ margin-bottom: 16px; display: flex; gap: 8px; align-items: center; }}
  .filter-bar input {{ padding: 6px 12px; border-radius: 6px; border: 1px solid #333; background: #161b22; color: #eee; font-size: 0.9em; width: 300px; }}
  .filter-bar select {{ padding: 6px 8px; border-radius: 6px; border: 1px solid #333; background: #161b22; color: #eee; }}
  .grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }}
  .card {{ background: #161b22; border-radius: 8px; overflow: hidden; border: 1px solid #21262d; transition: transform 0.1s, border-color 0.2s; cursor: pointer; }}
  .card:hover {{ transform: scale(1.02); border-color: #4a9eff; }}
  .card-header {{ display: flex; justify-content: space-between; align-items: center; padding: 8px 10px 4px; }}
  .crs-badge {{ display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 0.65em; font-weight: 700; color: #fff; text-transform: uppercase; }}
  .fmt-badge {{ font-size: 0.65em; font-weight: 700; padding: 2px 6px; border-radius: 3px; }}
  .fmt-dwg {{ background: #1f6feb; color: #fff; }}
  .fmt-dxf {{ background: #238636; color: #fff; }}
  .preview {{ width: 100%; height: 180px; display: flex; align-items: center; justify-content: center; background: #0d1117; }}
  .preview img {{ width: 100%; height: 180px; object-fit: contain; }}
  .no-preview {{ color: #484f58; font-size: 0.8em; }}
  .filename {{ font-weight: 600; font-size: 0.82em; padding: 6px 10px 2px; word-break: break-all; }}
  .folder {{ font-size: 0.68em; color: #484f58; padding: 0 10px 6px; word-break: break-all; }}
  .meta {{ display: flex; gap: 10px; font-size: 0.7em; color: #8b949e; padding: 4px 10px 10px; }}
  .count {{ color: #8b949e; font-size: 0.9em; margin-bottom: 12px; }}
  .overlay {{ display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); z-index: 1000; justify-content: center; align-items: center; }}
  .overlay.active {{ display: flex; }}
  .popup {{ background: #161b22; border-radius: 12px; border: 1px solid #30363d; max-width: 90vw; max-height: 90vh; overflow-y: auto; padding: 24px; position: relative; }}
  .popup-close {{ position: absolute; top: 12px; right: 16px; font-size: 1.5em; cursor: pointer; color: #8b949e; }}
  .popup-close:hover {{ color: #fff; }}
  .popup-title {{ font-size: 1.2em; font-weight: 700; margin-bottom: 8px; }}
  .popup-meta {{ display: flex; gap: 16px; font-size: 0.85em; color: #8b949e; margin-bottom: 16px; flex-wrap: wrap; }}
  .popup-preview {{ margin-bottom: 16px; }}
  .popup-preview img {{ width: 100%; max-height: 500px; object-fit: contain; }}
  .popup-legend {{ display: flex; flex-wrap: wrap; gap: 4px 14px; }}
  .legend-item {{ display: flex; align-items: center; gap: 4px; font-size: 0.75em; color: #8b949e; }}
  .legend-dot {{ width: 10px; height: 10px; border-radius: 2px; flex-shrink: 0; }}
  .legend-title {{ font-weight: 600; font-size: 0.85em; margin-bottom: 8px; color: #c9d1d9; }}
</style>
</head>
<body>
<h1>Kistefos ACC - DXF/DWG CRS Analysis ({len(results)} files)</h1>
<div class="summary">{summary_items}</div>
<div class="filter-bar">
  <input type="text" id="search" placeholder="Filter by filename or path..." oninput="filterCards()">
  <select id="crsFilter" onchange="filterCards()">
    <option value="">All CRS</option>
    {crs_options}
  </select>
</div>
<div class="count" id="count">{len(results)} files</div>
<div class="grid" id="grid">
{''.join(cards)}
</div>
<div class="overlay" id="overlay" onclick="if(event.target===this)closePopup()">
  <div class="popup">
    <span class="popup-close" onclick="closePopup()">&times;</span>
    <div class="popup-title" id="popup-title"></div>
    <div class="popup-meta" id="popup-meta"></div>
    <div class="popup-preview" id="popup-preview"></div>
    <div class="legend-title">Layers</div>
    <div class="popup-legend" id="popup-legend"></div>
  </div>
</div>
<script>
function filterCards() {{
  const q = document.getElementById('search').value.toLowerCase();
  const crs = document.getElementById('crsFilter').value;
  const cards = document.querySelectorAll('.card');
  let shown = 0;
  cards.forEach(c => {{
    const text = c.dataset.text;
    const cardCrs = c.dataset.crs;
    const matchQ = !q || text.includes(q);
    const matchCRS = !crs || cardCrs === crs;
    c.style.display = (matchQ && matchCRS) ? '' : 'none';
    if (matchQ && matchCRS) shown++;
  }});
  document.getElementById('count').textContent = shown + ' files';
}}
function openPopup(title, largeSvg, legend, crs, units, entities, size, folder) {{
  document.getElementById('popup-title').textContent = title;
  document.getElementById('popup-meta').innerHTML =
    `<span><b>CRS:</b> ${{crs}}</span><span><b>Units:</b> ${{units}}</span><span><b>Entities:</b> ${{entities}}</span><span><b>Size:</b> ${{size}}</span><span style="color:#484f58">${{folder}}</span>`;
  document.getElementById('popup-preview').innerHTML = largeSvg;
  document.getElementById('popup-legend').innerHTML = legend;
  document.getElementById('overlay').classList.add('active');
}}
function closePopup() {{
  document.getElementById('overlay').classList.remove('active');
}}
document.addEventListener('keydown', e => {{ if (e.key === 'Escape') closePopup(); }});
</script>
</body>
</html>'''

outpath = 'c:/Users/edkjo/repos/coordinate-transform/acc_overview.html'
with open(outpath, 'w', encoding='utf-8') as f:
    f.write(html_out)
print(f'Saved: {outpath} ({len(results)} files)')
