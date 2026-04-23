"""Derive local + mm DXF/DWG variants from a global meters DXF.

The global meters DXF is the canonical source of truth. Local meters, global mm
and local mm are all derived from it without re-merging:

    LOCAL meters = GLOBAL meters offset by -BASEPOINT
    GLOBAL mm    = GLOBAL meters x1000
    LOCAL mm     = LOCAL meters x1000

Usage:
    python derive_variants.py path/to/MASTER_NAME_NTM10_GLOBAL_METERS.dxf

Expects the input filename to contain `_NTM10_GLOBAL_METERS`. The output paths
are derived by string substitution and live next to the input (DXFs) and in
the matching DWG_NTM/{Global,Lokal} folders (DWGs).
"""
import os, sys
from merge_master_lines import (
    derive_local_from_global, scale_to_mm, convert_to_dwg, BASEPOINT,
)


def derive_all(global_meters_dxf, dxf_local_dir, dwg_global_dir, dwg_local_dir):
    base = os.path.basename(global_meters_dxf)
    if '_NTM10_GLOBAL_METERS' not in base:
        raise ValueError(f'Input must contain "_NTM10_GLOBAL_METERS": {base}')

    dxf_global_dir = os.path.dirname(global_meters_dxf)
    name_local_m  = base.replace('_NTM10_GLOBAL_METERS', '_NTM10_LOCAL_METERS')
    name_global_mm = base.replace('_NTM10_GLOBAL_METERS', '_NTM10_GLOBAL_MM')
    name_local_mm  = base.replace('_NTM10_GLOBAL_METERS', '_NTM10_LOCAL_MM')

    local_m   = os.path.join(dxf_local_dir,  name_local_m)
    global_mm = os.path.join(dxf_global_dir, name_global_mm)
    local_mm  = os.path.join(dxf_local_dir,  name_local_mm)

    print(f'\n=== Source: {base} ===')
    print(f'  basepoint: {BASEPOINT}')

    print('\n[1/3] Derive LOCAL METERS')
    derive_local_from_global(global_meters_dxf, local_m)

    print('\n[2/3] Derive GLOBAL MM')
    scale_to_mm(global_meters_dxf, global_mm)

    print('\n[3/3] Derive LOCAL MM (from local meters)')
    scale_to_mm(local_m, local_mm)

    print('\n=== DWG conversion ===')
    convert_to_dwg(global_meters_dxf, dwg_global_dir)
    convert_to_dwg(global_mm,         dwg_global_dir)
    convert_to_dwg(local_m,           dwg_local_dir)
    convert_to_dwg(local_mm,          dwg_local_dir)

    print('\nDone.')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    global_meters_dxf = sys.argv[1]

    if len(sys.argv) >= 5:
        dxf_local_dir = sys.argv[2]
        dwg_global_dir = sys.argv[3]
        dwg_local_dir = sys.argv[4]
    else:
        # Default ACC layout: infer Lokal/Global pair from input path.
        dxf_global_dir = os.path.dirname(global_meters_dxf)
        if dxf_global_dir.endswith('Global'):
            base = os.path.dirname(dxf_global_dir)
            dxf_local_dir = os.path.join(base, 'Lokal')
        else:
            print(f'Cannot infer paths from {dxf_global_dir}; pass dxf_local_dir, dwg_global_dir, dwg_local_dir explicitly.')
            sys.exit(1)
        # DWG dirs sit at ../DWG_NTM/{Global,Lokal}
        dxf_root = os.path.dirname(os.path.dirname(dxf_global_dir))  # .../08_Tegninger
        dwg_global_dir = os.path.join(dxf_root, 'DWG_NTM', 'Global')
        dwg_local_dir = os.path.join(dxf_root, 'DWG_NTM', 'Lokal')

    derive_all(global_meters_dxf, dxf_local_dir, dwg_global_dir, dwg_local_dir)
