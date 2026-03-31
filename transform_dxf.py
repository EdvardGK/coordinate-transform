"""
Transform DXF from UTM32 to NTM Sone 10.
Reprojects every vertex using pyproj. No AutoCAD needed.

Outputs:
  - *_NTM10_global_meters.dxf  (world NTM coordinates)
  - *_NTM10_local_meters.dxf   (offset to basepoint origin)

Usage:
  python transform_dxf.py <input.dxf>

Requires: ezdxf, pyproj
"""
import ezdxf
import os
import sys
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


def to_ntm_world(utm_e, utm_n):
    """UTM32 world coords -> NTM10 world coords."""
    return _transformer.transform(utm_e, utm_n)


def detect_coord_type(doc):
    """Detect if coordinates are model-relative or world UTM32."""
    msp = doc.modelspace()
    xs = []
    for e in msp:
        t = e.dxftype()
        if t == "LINE":
            xs.append(e.dxf.start.x)
        elif t == "LWPOLYLINE":
            for x, y in e.get_points(format="xy"):
                xs.append(x)
                break
        elif t == "POLYLINE":
            for p in e.points():
                xs.append(p[0])
                break
        if len(xs) >= 10:
            break

    if not xs:
        return "unknown"

    avg_x = sum(xs) / len(xs)
    if avg_x > 100000:  # World UTM32 coordinates (575xxx)
        return "world_utm32"
    else:  # Model-relative (0-1000 range)
        return "model_relative"


def transform_entity(e, transform_fn):
    """Transform entity coordinates using the given function."""
    t = e.dxftype()
    if t == "LINE":
        nx, ny = transform_fn(e.dxf.start.x, e.dxf.start.y)
        e.dxf.start = (nx, ny, e.dxf.start.z)
        nx, ny = transform_fn(e.dxf.end.x, e.dxf.end.y)
        e.dxf.end = (nx, ny, e.dxf.end.z)
    elif t == "LWPOLYLINE":
        pts = list(e.get_points(format="xyseb"))
        new_pts = []
        for x, y, s, ew, b in pts:
            nx, ny = transform_fn(x, y)
            new_pts.append((nx, ny, s, ew, b))
        e.set_points(new_pts, format="xyseb")
    elif t == "POLYLINE":
        for v in e.vertices:
            loc = v.dxf.location
            nx, ny = transform_fn(loc[0], loc[1])
            v.dxf.location = (nx, ny, loc[2])
    elif t == "CIRCLE":
        nx, ny = transform_fn(e.dxf.center.x, e.dxf.center.y)
        e.dxf.center = (nx, ny, e.dxf.center.z)
    elif t == "ELLIPSE":
        nx, ny = transform_fn(e.dxf.center.x, e.dxf.center.y)
        e.dxf.center = (nx, ny, e.dxf.center.z)
    elif t == "MTEXT":
        ins = e.dxf.insert
        nx, ny = transform_fn(ins.x, ins.y)
        e.dxf.insert = (nx, ny, ins.z)
    elif t == "HATCH":
        try:
            for path in e.paths:
                if hasattr(path, "vertices"):
                    new_verts = []
                    for v in path.vertices:
                        nx, ny = transform_fn(v[0], v[1])
                        new_verts.append((nx, ny) + v[2:])
                    path.vertices = new_verts
                if hasattr(path, "edges"):
                    for edge in path.edges:
                        for attr in ["start", "end", "center"]:
                            if hasattr(edge, attr):
                                p = getattr(edge, attr)
                                nx, ny = transform_fn(p[0], p[1])
                                setattr(edge, attr, (nx, ny))
        except:
            pass


def add_markers(doc, msp, offset_e=0, offset_n=0):
    """Add coordination markers."""
    old_bp_e, old_bp_n = to_ntm(0, 0)

    markers = [
        ("COORDINATION_MARKER_OLD", old_bp_e - offset_e, old_bp_n - offset_n, 1,
         f"OLD BASEPOINT (UTM32: {UTM_BP_E:.0f}/{UTM_BP_N:.0f})\nNTM10: E={old_bp_e:.3f} N={old_bp_n:.3f}"),
        ("COORDINATION_MARKER_NEW", NEW_BP_E - offset_e, NEW_BP_N - offset_n, 3,
         f"NEW BASEPOINT NTM10\nE={NEW_BP_E:.3f} N={NEW_BP_N:.3f}"),
        ("COORDINATION_MARKER_ROTATION", ROT_E - offset_e, ROT_N - offset_n, 5,
         f"ROTATION POINT NTM10\nE={ROT_E:.3f} N={ROT_N:.3f}"),
    ]

    for layer_name, bx, by, color, label in markers:
        doc.layers.add(layer_name, color=color)
        attribs = {"layer": layer_name, "color": color}
        msp.add_circle((bx, by, 0), radius=5.0, dxfattribs=attribs)
        msp.add_line((bx - 7, by, 0), (bx + 7, by, 0), dxfattribs=attribs)
        msp.add_line((bx, by - 7, 0), (bx, by + 7, 0), dxfattribs=attribs)
        msp.add_mtext(label, dxfattribs={
            **attribs, "insert": (bx + 6, by + 3, 0), "char_height": 2.0})


def process_dxf(input_path):
    """Transform DXF and output global + local versions."""
    print(f"Reading: {input_path}")
    doc = ezdxf.readfile(input_path)
    msp = doc.modelspace()

    coord_type = detect_coord_type(doc)
    print(f"Coordinate type: {coord_type}")

    if coord_type == "world_utm32":
        transform_fn = to_ntm_world
        print("Using world UTM32 -> NTM10 transform")
    else:
        transform_fn = to_ntm
        print("Using model-relative -> NTM10 transform")

    # Transform all entities
    count = 0
    for e in msp:
        transform_entity(e, transform_fn)
        count += 1

    # Fix headers
    doc.header["$DIMSCALE"] = 1.0
    doc.header["$LTSCALE"] = 1.0
    doc.header["$MEASUREMENT"] = 1
    doc.header["$INSUNITS"] = 6

    # Add markers
    add_markers(doc, msp, offset_e=0, offset_n=0)

    # Save global
    base = os.path.splitext(input_path)[0]
    global_path = f"{base}_NTM10_global_meters.dxf"
    doc.saveas(global_path)
    print(f"Saved global: {global_path} ({count} entities)")

    # Create local version by offsetting
    doc2 = ezdxf.readfile(global_path)
    msp2 = doc2.modelspace()

    def offset_fn(x, y):
        return x - NEW_BP_E, y - NEW_BP_N

    for e in msp2:
        transform_entity(e, offset_fn)

    local_path = f"{base}_NTM10_local_meters.dxf"
    doc2.saveas(local_path)
    print(f"Saved local: {local_path} (basepoint at origin)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python transform_dxf.py <input.dxf>")
        sys.exit(1)

    process_dxf(sys.argv[1])
