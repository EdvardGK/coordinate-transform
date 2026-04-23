"""Merge selected DXF files into one master, color-coded by source file."""
import ezdxf, os
from ezdxf.addons import Importer

base = 'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/10016 - Kistefos/03_Ut/08_Tegninger/DXF_NTM/Lokal'

# Build file list from actual directory listing to avoid encoding issues
bangs_dir = os.path.join(base, 'BANGS_Innmaaling')
design_dir = os.path.join(base, 'Design')

# Map: substring to match -> ACI color
wanted = {
    'Vegteknisk': 1,        # red
    'Kotelinjer': 2,        # yellow
    'Terrengmodell': 3,     # green
    'TreStammer': 4,        # cyan
    'Veg_TIN': 5,           # blue
    'Elektro': 6,           # magenta
    'Kunstinstallasjon': 30,# orange
    'lt_Tre': 140,          # brown
    'Nairy': 200,           # purple
    'Stier': 150,           # teal
    'Parkering': 40,        # salmon
    'landscape plan': 7,    # white
}

files_to_merge = []

# Scan BANGS
for f in os.listdir(bangs_dir):
    if not f.endswith('.dxf'):
        continue
    if 'BIMK_Master' in f:
        continue  # skip the old master
    for substr, color in wanted.items():
        if substr in f:
            files_to_merge.append((os.path.join(bangs_dir, f), color, f))
            break

# Scan Design
for f in os.listdir(design_dir):
    if not f.endswith('.dxf'):
        continue
    for substr, color in wanted.items():
        if substr in f:
            files_to_merge.append((os.path.join(design_dir, f), color, f))
            break

# ACAD landscape plan at root
for f in os.listdir(base):
    if 'landscape plan' in f and f.endswith('.dxf'):
        for substr, color in wanted.items():
            if substr in f:
                files_to_merge.append((os.path.join(base, f), color, f))
                break

print(f'Files to merge: {len(files_to_merge)}')
for path, color, fname in files_to_merge:
    print(f'  [{color:>3}] {fname}')

# Merge
merged = ezdxf.new('R2018')
merged.header['$INSUNITS'] = 6
merged.header['$MEASUREMENT'] = 1
msp_out = merged.modelspace()

total = 0
for path, aci_color, fname in files_to_merge:
    short = fname.split('_NTM10')[0].split('_NTM10')[0]
    layer_prefix = f'SRC|{short}'

    src = ezdxf.readfile(path)
    src_msp = src.modelspace()

    importer = Importer(src, merged)
    importer.import_modelspace()
    importer.finalize()

    all_entities = list(msp_out)
    new_entities = all_entities[total:]

    seen_layers = set()
    for e in new_entities:
        orig_layer = e.dxf.layer
        new_layer = f'{layer_prefix}|{orig_layer}'
        if new_layer not in seen_layers:
            if new_layer not in merged.layers:
                merged.layers.add(new_layer, color=aci_color)
            seen_layers.add(new_layer)
        e.dxf.layer = new_layer

    count = len(new_entities)
    total += count
    print(f'  Merged {short}: {count} entities, {len(seen_layers)} layers')

out_path = os.path.join(base, 'KNM_MASTER_NTM10_local_meters.dxf')
merged.saveas(out_path)
print(f'\nSaved: {out_path}')
print(f'Total entities: {total}')
print(f'Total layers: {len(merged.layers)}')

# Convert to DWG via ODA
import subprocess
tmp_in = os.path.dirname(out_path)
tmp_out = os.path.join(os.path.dirname(tmp_in), 'DWG_NTM', 'Lokal')
os.makedirs(tmp_out, exist_ok=True)

# ODA needs input dir, so copy to a temp dir with just our file
import shutil, tempfile
with tempfile.TemporaryDirectory() as tmpdir:
    shutil.copy2(out_path, tmpdir)
    oda = r'C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe'
    result = subprocess.run([oda, tmpdir, tmp_out, 'ACAD2018', 'DWG', '0', '1'],
                          capture_output=True, timeout=60)
    dwg_out = os.path.join(tmp_out, 'KNM_MASTER_NTM10_local_meters.dwg')
    if os.path.exists(dwg_out):
        print(f'DWG saved: {dwg_out}')
    else:
        print(f'DWG conversion may have failed. Check {tmp_out}')
