"""
Build two IFC files from 20cm contour lines:
1. Contour walls: 5cm wide walls with inside edge on the line
2. Terrain model: triangulated surface from contour vertices

Coordinates properly reprojected from UTM32 to NTM Sone 10.
"""
import ezdxf
import ifcopenshell
import ifcopenshell.api
import numpy as np
from scipy.spatial import Delaunay
from pyproj import Transformer
import math

DXF_PATH = r"C:\Users\edkjo\DC\ACCDocs\Skiplum AS\Skiplum Backup\Project Files\10016 - Kistefos\01_Inn\02_Tegninger\ACAD-Kotelinjer_20cm.dxf"
IFC_WALLS_OUT = r"C:\Users\edkjo\DC\ACCDocs\Skiplum AS\Skiplum Backup\Project Files\10016 - Kistefos\03_Ut\08_Tegninger\Kotelinjer_20cm_walls_NTM10.ifc"
IFC_TERRAIN_OUT = r"C:\Users\edkjo\DC\ACCDocs\Skiplum AS\Skiplum Backup\Project Files\10016 - Kistefos\03_Ut\08_Tegninger\Kotelinjer_20cm_terrain_NTM10.ifc"

WALL_WIDTH = 0.05  # 5cm
WALL_HEIGHT = 0.20  # contour interval

_transformer = Transformer.from_crs("EPSG:25832", "EPSG:5110", always_xy=True)


def to_ntm(x, y):
    return _transformer.transform(x, y)


def offset_polyline_outward(pts_2d, width):
    """Offset a polyline outward (to the right) by width.
    Returns the offset points. Combined with original = wall outline."""
    offset_pts = []
    n = len(pts_2d)
    for i in range(n):
        # Get direction vectors
        if i == 0:
            dx = pts_2d[1][0] - pts_2d[0][0]
            dy = pts_2d[1][1] - pts_2d[0][1]
        elif i == n - 1:
            dx = pts_2d[-1][0] - pts_2d[-2][0]
            dy = pts_2d[-1][1] - pts_2d[-2][1]
        else:
            dx = pts_2d[i + 1][0] - pts_2d[i - 1][0]
            dy = pts_2d[i + 1][1] - pts_2d[i - 1][1]

        length = math.sqrt(dx * dx + dy * dy)
        if length < 1e-10:
            offset_pts.append(pts_2d[i])
            continue

        # Normal to the right (outward)
        nx = dy / length * width
        ny = -dx / length * width

        offset_pts.append((pts_2d[i][0] + nx, pts_2d[i][1] + ny))

    return offset_pts


# ── Read DXF and collect contour data ──
print("Reading DXF...")
doc = ezdxf.readfile(DXF_PATH)
msp = doc.modelspace()

contours = []  # list of (z, [(ntm_e, ntm_n), ...])
all_points = []  # for terrain mesh

for e in msp:
    if e.dxftype() == "POLYLINE":
        pts_raw = list(e.points())
        if not pts_raw:
            continue
        z = pts_raw[0][2]

        ntm_pts = []
        for p in pts_raw:
            nx, ny = to_ntm(p[0], p[1])
            ntm_pts.append((nx, ny))
            all_points.append((nx, ny, p[2]))

        contours.append((z, ntm_pts))

print(f"Contours: {len(contours)}")
print(f"Total vertices: {len(all_points)}")
print(f"Z range: {min(p[2] for p in all_points):.2f} - {max(p[2] for p in all_points):.2f}")


# ══════════════════════════════════════════
# IFC 1: Contour walls
# ══════════════════════════════════════════
print("\nBuilding contour walls IFC...")

ifc = ifcopenshell.api.run("project.create_file")
project = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcProject", name="Kistefos Kotelinjer")
ifcopenshell.api.run("unit.assign_unit", ifc, length={"is_metric": True, "raw": "METERS"})

ctx = ifcopenshell.api.run("context.add_context", ifc, context_type="Model")
body = ifcopenshell.api.run(
    "context.add_context", ifc,
    context_type="Model", context_identifier="Body",
    target_view="MODEL_VIEW", parent=ctx,
)

site = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcSite", name="Kistefos")
ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=project, products=[site])
building = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcBuilding", name="Terrain")
ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=site, products=[building])
storey = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcBuildingStorey", name="Ground")
ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=building, products=[storey])

wall_count = 0
for z, pts in contours:
    if len(pts) < 2:
        continue

    # Create wall outline: original line + offset line reversed + close
    offset = offset_polyline_outward(pts, WALL_WIDTH)
    # Wall profile: go along original, come back along offset (reversed)
    outline = list(pts) + list(reversed(offset))

    name = f"Contour_{z:.2f}m"
    wall = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcWall", name=name)

    ifc_pts = [ifc.createIfcCartesianPoint([float(p[0]), float(p[1])]) for p in outline]
    ifc_pts.append(ifc_pts[0])
    polyline = ifc.createIfcPolyline(ifc_pts)
    profile = ifc.createIfcArbitraryClosedProfileDef("AREA", None, polyline)

    solid = ifc.createIfcExtrudedAreaSolid(
        profile,
        ifc.createIfcAxis2Placement3D(ifc.createIfcCartesianPoint([0.0, 0.0, float(z)])),
        ifc.createIfcDirection([0.0, 0.0, 1.0]),
        WALL_HEIGHT,
    )

    rep = ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [solid])
    prod_rep = ifc.createIfcProductDefinitionShape(None, None, [rep])
    wall.Representation = prod_rep
    wall.ObjectPlacement = ifcopenshell.api.run("geometry.edit_object_placement", ifc, product=wall)
    ifcopenshell.api.run("spatial.assign_container", ifc, relating_structure=storey, products=[wall])

    # Color by elevation (blue=low, red=high)
    z_norm = (z - 115.5) / (150.1 - 115.5)
    r, g, b = z_norm, 0.3, 1.0 - z_norm
    color = ifc.createIfcColourRgb(None, r, g, b)
    style_render = ifc.createIfcSurfaceStyleRendering(color, 0.0, None, None, None, None, None, None, "FLAT")
    style = ifc.createIfcSurfaceStyle(name, "BOTH", [style_render])
    ifc.createIfcStyledItem(solid, [ifc.createIfcPresentationStyleAssignment([style])], None)

    # Property set
    pset = ifcopenshell.api.run("pset.add_pset", ifc, product=wall, name="NOSKI_Contour")
    ifcopenshell.api.run("pset.edit_pset", ifc, pset=pset, properties={
        "Elevation_NN2000": str(z),
        "ContourInterval": "0.20m",
        "SourceFile": "Kotelinjer_20cm.DWG",
        "CoordinateSystem": "EUREF89 NTM Sone 10 (EPSG:5110)",
        "SourceCRS": "EUREF89 UTM32 (EPSG:25832)",
        "HeightSystem": "NN2000",
    })

    wall_count += 1
    if wall_count % 200 == 0:
        print(f"  {wall_count} walls...")

# ── Coordination markers ──
OLD_BP_E, OLD_BP_N = to_ntm(575200, 6676400)
NEW_BP_E, NEW_BP_N = 92200.0, 1247000.0
MARKER_Z = 130.0  # mid-range elevation for visibility

for bp_name, bp_e, bp_n, radius, height in [
    ("COORDINATION_MARKER_NEW", NEW_BP_E, NEW_BP_N, 2.0, 5.0),
    ("COORDINATION_MARKER_OLD (UTM32: 575200/6676400)", OLD_BP_E, OLD_BP_N, 1.5, 3.0),
]:
    marker = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcSlab", name=bp_name)
    m_pts = [(bp_e + radius * math.cos(2 * math.pi * i / 36),
              bp_n + radius * math.sin(2 * math.pi * i / 36)) for i in range(36)]
    m_profile = ifc.createIfcArbitraryClosedProfileDef("AREA", None,
        ifc.createIfcPolyline([ifc.createIfcCartesianPoint([float(p[0]), float(p[1])]) for p in m_pts]
                              + [ifc.createIfcCartesianPoint([float(m_pts[0][0]), float(m_pts[0][1])])]))
    m_solid = ifc.createIfcExtrudedAreaSolid(m_profile,
        ifc.createIfcAxis2Placement3D(ifc.createIfcCartesianPoint([0.0, 0.0, MARKER_Z])),
        ifc.createIfcDirection([0.0, 0.0, 1.0]), height)
    m_rep = ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [m_solid])
    marker.Representation = ifc.createIfcProductDefinitionShape(None, None, [m_rep])
    marker.ObjectPlacement = ifcopenshell.api.run("geometry.edit_object_placement", ifc, product=marker)
    ifcopenshell.api.run("spatial.assign_container", ifc, relating_structure=storey, products=[marker])

ifc.write(IFC_WALLS_OUT)
print(f"Walls IFC saved: {IFC_WALLS_OUT}")
print(f"Total walls: {wall_count} + 2 markers")


# ══════════════════════════════════════════
# IFC 2: Terrain mesh
# ══════════════════════════════════════════
print("\nBuilding terrain mesh IFC...")

# Deduplicate and subsample points for triangulation
pts_array = np.array(all_points)
print(f"Raw points: {len(pts_array)}")

# Subsample if too many points (Delaunay can be slow)
if len(pts_array) > 50000:
    idx = np.random.choice(len(pts_array), 50000, replace=False)
    pts_array = pts_array[idx]
    print(f"Subsampled to: {len(pts_array)}")

xy = pts_array[:, :2]
z_vals = pts_array[:, 2]

print("Triangulating...")
tri = Delaunay(xy)
triangles = tri.simplices
print(f"Triangles: {len(triangles)}")

# Filter out very large triangles (gaps between contours)
# Max edge length threshold
MAX_EDGE = 15.0  # meters
good_triangles = []
for t_idx in triangles:
    p0, p1, p2 = xy[t_idx[0]], xy[t_idx[1]], xy[t_idx[2]]
    e1 = np.sqrt((p1[0] - p0[0]) ** 2 + (p1[1] - p0[1]) ** 2)
    e2 = np.sqrt((p2[0] - p1[0]) ** 2 + (p2[1] - p1[1]) ** 2)
    e3 = np.sqrt((p0[0] - p2[0]) ** 2 + (p0[1] - p2[1]) ** 2)
    if e1 < MAX_EDGE and e2 < MAX_EDGE and e3 < MAX_EDGE:
        good_triangles.append(t_idx)

good_triangles = np.array(good_triangles)
print(f"Filtered triangles (edge < {MAX_EDGE}m): {len(good_triangles)}")

# Build IFC
ifc2 = ifcopenshell.api.run("project.create_file")
project2 = ifcopenshell.api.run("root.create_entity", ifc2, ifc_class="IfcProject", name="Kistefos Terreng")
ifcopenshell.api.run("unit.assign_unit", ifc2, length={"is_metric": True, "raw": "METERS"})

ctx2 = ifcopenshell.api.run("context.add_context", ifc2, context_type="Model")
body2 = ifcopenshell.api.run(
    "context.add_context", ifc2,
    context_type="Model", context_identifier="Body",
    target_view="MODEL_VIEW", parent=ctx2,
)

site2 = ifcopenshell.api.run("root.create_entity", ifc2, ifc_class="IfcSite", name="Kistefos")
ifcopenshell.api.run("aggregate.assign_object", ifc2, relating_object=project2, products=[site2])
building2 = ifcopenshell.api.run("root.create_entity", ifc2, ifc_class="IfcBuilding", name="Terrain")
ifcopenshell.api.run("aggregate.assign_object", ifc2, relating_object=site2, products=[building2])
storey2 = ifcopenshell.api.run("root.create_entity", ifc2, ifc_class="IfcBuildingStorey", name="Ground")
ifcopenshell.api.run("aggregate.assign_object", ifc2, relating_object=building2, products=[storey2])

# Create terrain as IfcGeographicElement with IfcTriangulatedFaceSet
terrain = ifcopenshell.api.run(
    "root.create_entity", ifc2, ifc_class="IfcGeographicElement", name="Terrain_20cm_contours"
)

# Point list
coord_list = ifc2.createIfcCartesianPointList3D(
    [tuple(float(v) for v in (xy[i][0], xy[i][1], z_vals[i])) for i in range(len(xy))]
)

# Triangle indices (IFC uses 1-based)
face_indices = [[int(t[0]) + 1, int(t[1]) + 1, int(t[2]) + 1] for t in good_triangles]

face_set = ifc2.createIfcTriangulatedFaceSet(coord_list, None, None, face_indices, None)

# Color the terrain (green-brown)
color2 = ifc2.createIfcColourRgb(None, 0.45, 0.55, 0.30)
style_render2 = ifc2.createIfcSurfaceStyleRendering(color2, 0.0, None, None, None, None, None, None, "FLAT")
style2 = ifc2.createIfcSurfaceStyle("Terrain", "BOTH", [style_render2])
ifc2.createIfcStyledItem(face_set, [ifc2.createIfcPresentationStyleAssignment([style2])], None)

rep2 = ifc2.createIfcShapeRepresentation(body2, "Body", "Tessellation", [face_set])
prod_rep2 = ifc2.createIfcProductDefinitionShape(None, None, [rep2])
terrain.Representation = prod_rep2
terrain.ObjectPlacement = ifcopenshell.api.run("geometry.edit_object_placement", ifc2, product=terrain)
ifcopenshell.api.run("spatial.assign_container", ifc2, relating_structure=storey2, products=[terrain])

# Property set
pset2 = ifcopenshell.api.run("pset.add_pset", ifc2, product=terrain, name="NOSKI_Terrain")
ifcopenshell.api.run("pset.edit_pset", ifc2, pset=pset2, properties={
    "SourceFile": "Kotelinjer_20cm.DWG",
    "ContourInterval": "0.20m",
    "CoordinateSystem": "EUREF89 NTM Sone 10 (EPSG:5110)",
    "SourceCRS": "EUREF89 UTM32 (EPSG:25832)",
    "HeightSystem": "NN2000",
    "Vertices": str(len(xy)),
    "Triangles": str(len(good_triangles)),
    "ElevationRange": f"{min(z_vals):.2f} - {max(z_vals):.2f}m",
})

# ── Coordination markers for terrain ──
for bp_name, bp_e, bp_n, radius, height in [
    ("COORDINATION_MARKER_NEW", NEW_BP_E, NEW_BP_N, 2.0, 5.0),
    ("COORDINATION_MARKER_OLD (UTM32: 575200/6676400)", OLD_BP_E, OLD_BP_N, 1.5, 3.0),
]:
    marker2 = ifcopenshell.api.run("root.create_entity", ifc2, ifc_class="IfcSlab", name=bp_name)
    m_pts2 = [(bp_e + radius * math.cos(2 * math.pi * i / 36),
               bp_n + radius * math.sin(2 * math.pi * i / 36)) for i in range(36)]
    m_profile2 = ifc2.createIfcArbitraryClosedProfileDef("AREA", None,
        ifc2.createIfcPolyline([ifc2.createIfcCartesianPoint([float(p[0]), float(p[1])]) for p in m_pts2]
                               + [ifc2.createIfcCartesianPoint([float(m_pts2[0][0]), float(m_pts2[0][1])])]))
    m_solid2 = ifc2.createIfcExtrudedAreaSolid(m_profile2,
        ifc2.createIfcAxis2Placement3D(ifc2.createIfcCartesianPoint([0.0, 0.0, MARKER_Z])),
        ifc2.createIfcDirection([0.0, 0.0, 1.0]), height)
    m_rep2 = ifc2.createIfcShapeRepresentation(body2, "Body", "SweptSolid", [m_solid2])
    marker2.Representation = ifc2.createIfcProductDefinitionShape(None, None, [m_rep2])
    marker2.ObjectPlacement = ifcopenshell.api.run("geometry.edit_object_placement", ifc2, product=marker2)
    ifcopenshell.api.run("spatial.assign_container", ifc2, relating_structure=storey2, products=[marker2])

ifc2.write(IFC_TERRAIN_OUT)
print(f"\nTerrain IFC saved: {IFC_TERRAIN_OUT}")
print(f"Vertices: {len(xy)}, Triangles: {len(good_triangles)}, + 2 markers")
