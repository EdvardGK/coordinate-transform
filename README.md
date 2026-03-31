# Coordinate Transform - UTM32 to NTM Sone 10

Tools for reprojecting DWG/DXF geometry from EUREF89 UTM Zone 32N (EPSG:25832) to EUREF89 NTM Sone 10 (EPSG:5110).

Developed for Kistefos Museum project (10016).

## Scripts

### `transform_dwg.py` - Full DWG pipeline (requires AutoCAD)
Opens DWG in running AutoCAD via COM, reprojects every vertex, saves as DWG.

```
python transform_dwg.py <input.dwg>
```

Outputs:
- `*_NTM10_global_meters.dwg` — world NTM coordinates
- `*_NTM10_local_meters.dwg` — offset to project basepoint (0,0)

Handles: lines, polylines (2D/3D/LW), circles, arcs, ellipses, splines, text, blocks, points.

### `transform_dxf.py` - DXF pipeline (no AutoCAD needed)
Reprojects DXF geometry using ezdxf + pyproj.

```
python transform_dxf.py <input.dxf>
```

### `build_ifc.py` - DXF to IFC conversion
Converts DXF geometry to IFC slabs/walls with color by layer.

### `build_contour_ifc.py` - Contour lines to IFC
Creates contour wall IFC and triangulated terrain mesh IFC from 3D contour DXF.

## Configuration

All scripts share these parameters (edit at top of each file):

| Parameter | Value | Description |
|---|---|---|
| UTM_BP_E | 575200.0 | UTM32 basepoint easting |
| UTM_BP_N | 6676400.0 | UTM32 basepoint northing |
| NEW_BP_E | 92200.0 | NTM10 project basepoint easting |
| NEW_BP_N | 1247000.0 | NTM10 project basepoint northing |
| ROT_E | 92800.0 | Rotation verification point easting |
| ROT_N | 1248100.0 | Rotation verification point northing |

## Coordination markers

All outputs include three identical markers (10m diameter circle + crosshair + label):
- **Old basepoint** (red) — UTM32 origin mapped to NTM10: E=92083.507, N=1247024.149
- **New basepoint** (green) — Project basepoint: E=92200, N=1247000
- **Rotation point** (blue) — Verification point: E=92800, N=1248100

Two markers at known world coordinates allow verification of both position AND rotation when linking files.

## Requirements

```
pip install ezdxf pyproj openpyxl ifcopenshell scipy numpy
```

For `transform_dwg.py` additionally:
```
pip install pywin32
```
And a running AutoCAD/Civil 3D instance.

## Coordinate systems

- **Source:** EUREF89 UTM Zone 32N (EPSG:25832)
- **Target:** EUREF89 Norwegian Transverse Mercator Zone 10 (EPSG:5110)
- **Height:** NN2000 (preserved, not transformed)
- **Datum:** ETRS89 (same for both, zero-error transformation)

## Important notes

- Every vertex is individually reprojected via pyproj — not a simple offset
- The difference between offset and proper reprojection is ~13m at 500m from basepoint
- DWG files store coordinates in model space relative to UTM32 basepoint — the basepoint itself is metadata only and does not transform the geometry on export
