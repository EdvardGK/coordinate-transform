"""Merge selected DXF files into one master (GLOBAL version), color-coded by source file."""
import ezdxf, os, subprocess, shutil, tempfile
from ezdxf.addons import Importer

base = 'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/10016 - Kistefos/03_Ut/08_Tegninger/DXF_NTM/Global'
bangs_dir = os.path.join(base, 'BANGS_Innmaaling')
design_dir = os.path.join(base, 'Design')

wanted = {
    'Vegteknisk': 1, 'Kotelinjer': 2, 'Terrengmodell': 3, 'TreStammer': 4,
    'Veg_TIN': 5, 'Elektro': 6, 'Kunstinstallasjon': 30, 'lt_Tre': 140,
    'Nairy': 200, 'Stier': 150, 'Parkering': 40, 'landscape plan': 7,
}

files_to_merge = []
seen_short = set()

for d in [bangs_dir, design_dir, base]:
    if not os.path.exists(d):
        continue
    for f in os.listdir(d):
        if not f.endswith('.dxf') or 'BIMK_Master' in f or 'KNM_MASTER' in f:
            continue
        for substr, color in wanted.items():
            if substr in f:
                short = f.split('_NTM10')[0]
                if short not in seen_short:
                    files_to_merge.append((os.path.join(d, f), color, f))
                    seen_short.add(short)
                break

print(f'Files to merge: {len(files_to_merge)}')
for path, color, fname in files_to_merge:
    print(f'  [{color:>3}] {fname}')

merged = ezdxf.new('R2018')
merged.header['$INSUNITS'] = 6
merged.header['$MEASUREMENT'] = 1
msp_out = merged.modelspace()

total = 0
for path, aci_color, fname in files_to_merge:
    short = fname.split('_NTM10')[0]
    layer_prefix = f'SRC|{short}'

    src = ezdxf.readfile(path)
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

out_path = os.path.join(base, 'KNM_MASTER_NTM10_global_meters.dxf')
merged.saveas(out_path)
print(f'\nSaved DXF: {out_path}')
print(f'Total entities: {total}, Layers: {len(merged.layers)}')

# Convert to DWG
dwg_out_dir = os.path.join(os.path.dirname(os.path.dirname(base)), 'DWG_NTM', 'Global')
os.makedirs(dwg_out_dir, exist_ok=True)

with tempfile.TemporaryDirectory() as tmpdir:
    shutil.copy2(out_path, tmpdir)
    oda = r'C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe'
    subprocess.run([oda, tmpdir, dwg_out_dir, 'ACAD2018', 'DWG', '0', '1'], capture_output=True, timeout=60)
    dwg_path = os.path.join(dwg_out_dir, 'KNM_MASTER_NTM10_global_meters.dwg')
    if os.path.exists(dwg_path):
        print(f'Saved DWG: {dwg_path}')
    else:
        print(f'DWG conversion may need more time. Check: {dwg_out_dir}')
