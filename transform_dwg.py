"""
Transform DWG from UTM32 to NTM Sone 10 via AutoCAD COM.
Opens DWG in running AutoCAD, reprojects every vertex, saves as DWG.

Outputs:
  - *_NTM10_global_meters.dwg  (world NTM coordinates)
  - *_NTM10_local_meters.dwg   (offset to basepoint origin)

Usage:
  1. Open AutoCAD / Civil 3D
  2. python transform_dwg.py <input.dwg> [--basepoint-e 92200 --basepoint-n 1247000]

Requires: pywin32, pyproj
"""
import win32com.client
import pythoncom
import os
import sys
import time
import math
from pyproj import Transformer

# ── Config ──
UTM_BP_E = 575200.0
UTM_BP_N = 6676400.0
NEW_BP_E = 92200.0
NEW_BP_N = 1247000.0
ROT_E = 92800.0
ROT_N = 1248100.0

_transformer = Transformer.from_crs("EPSG:25832", "EPSG:5110", always_xy=True)


def to_ntm(model_x, model_y):
    """Model space (m, relative to UTM32 basepoint) -> NTM10 world coords."""
    return _transformer.transform(UTM_BP_E + model_x, UTM_BP_N + model_y)


def vtpnt(x, y, z=0.0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, z])


def vtfloat(lst):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


def transform_entity(e, offset_e=0, offset_n=0):
    """Transform a single AutoCAD entity from model coords to NTM10 + optional offset."""
    etype = e.ObjectName

    try:
        if etype in ("AcDbLine",):
            sp = e.StartPoint
            ep = e.EndPoint
            nx1, ny1 = to_ntm(sp[0], sp[1])
            nx2, ny2 = to_ntm(ep[0], ep[1])
            e.StartPoint = vtpnt(nx1 - offset_e, ny1 - offset_n, sp[2])
            e.EndPoint = vtpnt(nx2 - offset_e, ny2 - offset_n, ep[2])
            return True

        elif etype in ("AcDbPolyline", "AcDbLightWeightPolyline"):
            coords = list(e.Coordinates)
            # LWPolyline: pairs of (x, y)
            new_coords = []
            for i in range(0, len(coords), 2):
                nx, ny = to_ntm(coords[i], coords[i + 1])
                new_coords.extend([nx - offset_e, ny - offset_n])
            e.Coordinates = vtfloat(new_coords)
            return True

        elif etype == "AcDb2dPolyline":
            # 2D polyline with vertex sub-entities
            for i in range(e.Count):
                v = e.Item(i)
                p = v.Coordinate
                nx, ny = to_ntm(p[0], p[1])
                v.Coordinate = vtpnt(nx - offset_e, ny - offset_n, p[2])
            return True

        elif etype == "AcDb3dPolyline":
            coords = list(e.Coordinates)
            new_coords = []
            for i in range(0, len(coords), 3):
                nx, ny = to_ntm(coords[i], coords[i + 1])
                new_coords.extend([nx - offset_e, ny - offset_n, coords[i + 2]])
            e.Coordinates = vtfloat(new_coords)
            return True

        elif etype == "AcDbCircle":
            c = e.Center
            nx, ny = to_ntm(c[0], c[1])
            e.Center = vtpnt(nx - offset_e, ny - offset_n, c[2])
            return True

        elif etype == "AcDbEllipse":
            c = e.Center
            nx, ny = to_ntm(c[0], c[1])
            e.Center = vtpnt(nx - offset_e, ny - offset_n, c[2])
            return True

        elif etype == "AcDbMText":
            ip = e.InsertionPoint
            nx, ny = to_ntm(ip[0], ip[1])
            e.InsertionPoint = vtpnt(nx - offset_e, ny - offset_n, ip[2])
            return True

        elif etype == "AcDbText":
            ip = e.InsertionPoint
            nx, ny = to_ntm(ip[0], ip[1])
            e.InsertionPoint = vtpnt(nx - offset_e, ny - offset_n, ip[2])
            return True

        elif etype == "AcDbHatch":
            # Hatches are complex - skip transform, they'll follow if associative
            return False

        elif etype == "AcDbBlockReference":
            ip = e.InsertionPoint
            nx, ny = to_ntm(ip[0], ip[1])
            e.InsertionPoint = vtpnt(nx - offset_e, ny - offset_n, ip[2])
            return True

        elif etype == "AcDbPoint":
            p = e.Coordinates
            nx, ny = to_ntm(p[0], p[1])
            e.Coordinates = vtpnt(nx - offset_e, ny - offset_n, p[2])
            return True

        elif etype == "AcDbArc":
            c = e.Center
            nx, ny = to_ntm(c[0], c[1])
            e.Center = vtpnt(nx - offset_e, ny - offset_n, c[2])
            return True

        elif etype == "AcDbSpline":
            # Get control points
            pts = list(e.ControlPoints)
            new_pts = []
            for i in range(0, len(pts), 3):
                nx, ny = to_ntm(pts[i], pts[i + 1])
                new_pts.extend([nx - offset_e, ny - offset_n, pts[i + 2]])
            e.ControlPoints = vtfloat(new_pts)
            # Also transform fit points if present
            try:
                fpts = list(e.FitPoints)
                new_fpts = []
                for i in range(0, len(fpts), 3):
                    nx, ny = to_ntm(fpts[i], fpts[i + 1])
                    new_fpts.extend([nx - offset_e, ny - offset_n, fpts[i + 2]])
                e.FitPoints = vtfloat(new_fpts)
            except:
                pass
            return True

        else:
            return False

    except Exception as ex:
        return False


def add_markers(ms, offset_e=0, offset_n=0):
    """Add coordination markers."""
    old_e, old_n = to_ntm(0, 0)
    markers = [
        ("COORDINATION_MARKER_OLD", old_e - offset_e, old_n - offset_n, 1,
         f"OLD BASEPOINT (UTM32: 575200/6676400)\nNTM10: E={old_e:.3f} N={old_n:.3f}"),
        ("COORDINATION_MARKER_NEW", NEW_BP_E - offset_e, NEW_BP_N - offset_n, 3,
         f"NEW BASEPOINT NTM10\nE={NEW_BP_E:.3f} N={NEW_BP_N:.3f}"),
        ("COORDINATION_MARKER_ROTATION", ROT_E - offset_e, ROT_N - offset_n, 5,
         f"ROTATION POINT NTM10\nE={ROT_E:.3f} N={ROT_N:.3f}"),
    ]

    doc = ms.Application.ActiveDocument
    for layer_name, bx, by, color, label in markers:
        # Create layer
        try:
            layer = doc.Layers.Add(layer_name)
            layer.Color = color
        except:
            pass

        # Circle
        c = ms.AddCircle(vtpnt(bx, by, 0), 5.0)
        c.Layer = layer_name
        c.Color = color

        # Crosshair
        l1 = ms.AddLine(vtpnt(bx - 7, by, 0), vtpnt(bx + 7, by, 0))
        l1.Layer = layer_name
        l1.Color = color
        l2 = ms.AddLine(vtpnt(bx, by - 7, 0), vtpnt(bx, by + 7, 0))
        l2.Layer = layer_name
        l2.Color = color

        # Label
        mt = ms.AddMText(vtpnt(bx + 6, by + 3, 0), 0, label)
        mt.Height = 2.0
        mt.Layer = layer_name


def process_dwg(input_path, output_global, output_local):
    """Full pipeline: open DWG, transform, save global + local versions."""

    print(f"Connecting to AutoCAD...")
    acad = win32com.client.Dispatch("AutoCAD.Application")
    acad.Visible = True

    print(f"Opening: {input_path}")
    doc = acad.Documents.Open(input_path)
    ms = doc.ModelSpace

    entity_count = ms.Count
    print(f"Entities: {entity_count}")

    # ── Global version (world NTM coords) ──
    print("\nTransforming to NTM10 (global)...")
    transformed = 0
    skipped = 0
    for i in range(entity_count):
        e = ms.Item(i)
        if transform_entity(e, offset_e=0, offset_n=0):
            transformed += 1
        else:
            skipped += 1
        if (i + 1) % 200 == 0:
            print(f"  {i + 1}/{entity_count}...")

    print(f"Transformed: {transformed}, Skipped: {skipped}")

    # Add markers
    print("Adding markers...")
    add_markers(ms, offset_e=0, offset_n=0)

    # Set units
    doc.SetVariable("INSUNITS", 6)
    doc.SetVariable("MEASUREMENT", 1)
    doc.SetVariable("LUNITS", 2)

    # Save global
    print(f"Saving global: {output_global}")
    doc.SaveAs(output_global)

    # ── Local version (offset to basepoint) ──
    print("\nOffsetting to basepoint for local version...")
    for i in range(ms.Count):
        e = ms.Item(i)
        etype = e.ObjectName
        try:
            if etype in ("AcDbLine",):
                sp = e.StartPoint
                ep = e.EndPoint
                e.StartPoint = vtpnt(sp[0] - NEW_BP_E, sp[1] - NEW_BP_N, sp[2])
                e.EndPoint = vtpnt(ep[0] - NEW_BP_E, ep[1] - NEW_BP_N, ep[2])
            elif etype in ("AcDbPolyline", "AcDbLightWeightPolyline"):
                coords = list(e.Coordinates)
                new_coords = []
                for j in range(0, len(coords), 2):
                    new_coords.extend([coords[j] - NEW_BP_E, coords[j + 1] - NEW_BP_N])
                e.Coordinates = vtfloat(new_coords)
            elif etype == "AcDb3dPolyline":
                coords = list(e.Coordinates)
                new_coords = []
                for j in range(0, len(coords), 3):
                    new_coords.extend([coords[j] - NEW_BP_E, coords[j + 1] - NEW_BP_N, coords[j + 2]])
                e.Coordinates = vtfloat(new_coords)
            elif etype in ("AcDbCircle", "AcDbArc", "AcDbEllipse"):
                c = e.Center
                e.Center = vtpnt(c[0] - NEW_BP_E, c[1] - NEW_BP_N, c[2])
            elif etype in ("AcDbMText", "AcDbText"):
                ip = e.InsertionPoint
                e.InsertionPoint = vtpnt(ip[0] - NEW_BP_E, ip[1] - NEW_BP_N, ip[2])
            elif etype == "AcDbBlockReference":
                ip = e.InsertionPoint
                e.InsertionPoint = vtpnt(ip[0] - NEW_BP_E, ip[1] - NEW_BP_N, ip[2])
            elif etype == "AcDbSpline":
                pts = list(e.ControlPoints)
                new_pts = []
                for j in range(0, len(pts), 3):
                    new_pts.extend([pts[j] - NEW_BP_E, pts[j + 1] - NEW_BP_N, pts[j + 2]])
                e.ControlPoints = vtfloat(new_pts)
            elif etype == "AcDbPoint":
                p = e.Coordinates
                e.Coordinates = vtpnt(p[0] - NEW_BP_E, p[1] - NEW_BP_N, p[2])
        except:
            pass

    # Save local
    print(f"Saving local: {output_local}")
    doc.SaveAs(output_local)

    print("\nDone!")
    print(f"  Global: {output_global}")
    print(f"  Local:  {output_local}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python transform_dwg.py <input.dwg>")
        print("  Outputs: *_NTM10_global_meters.dwg and *_NTM10_local_meters.dwg")
        sys.exit(1)

    input_path = os.path.abspath(sys.argv[1])
    base = os.path.splitext(input_path)[0]
    output_dir = os.path.dirname(input_path)
    name = os.path.splitext(os.path.basename(input_path))[0]

    output_global = os.path.join(output_dir, f"{name}_NTM10_global_meters.dwg")
    output_local = os.path.join(output_dir, f"{name}_NTM10_local_meters.dwg")

    process_dwg(input_path, output_global, output_local)
