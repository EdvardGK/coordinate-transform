"""
Build IFC from DXF geometry. Closed polylines become slab outlines directly.
Open polylines are paired by matching endpoints to form closed loops.
Circles become circular slabs. All slabs 0.5m thick at Z=0.
Coordinates in NTM Sone 10 (georeferenced).
"""
import ezdxf
import ifcopenshell
import ifcopenshell.api
import math
import time
from collections import defaultdict
from pyproj import Transformer

DXF_PATH = "ACAD-KNM_Stier_meter_NTM.dxf"
IFC_OUT = "KNM_Stier_NTM10.ifc"
UTM_BP_E = 575200.0  # UTM32 basepoint - model coords are relative to this
UTM_BP_N = 6676400.0
SLAB_THICKNESS = 0.5
ENDPOINT_TOL = 2.0  # meters tolerance for matching endpoints
GROUND_Z = 122.83  # average Z from survey (NN2000)

# Proper coordinate transform: UTM32 -> NTM Sone 10
_transformer = Transformer.from_crs("EPSG:25832", "EPSG:5110", always_xy=True)


def to_ntm(model_x, model_y):
    """Transform model-space coords (meters, relative to UTM32 basepoint) to NTM10."""
    return _transformer.transform(UTM_BP_E + model_x, UTM_BP_N + model_y)

doc = ezdxf.readfile(DXF_PATH)
msp = doc.modelspace()

# ── Collect geometry by layer ──
layers = defaultdict(lambda: {"closed": [], "open": [], "lines": [], "circles": [], "ellipses": []})

for e in msp:
    t = e.dxftype()
    layer = e.dxf.layer

    if t == "LWPOLYLINE":
        pts = [to_ntm(x, y) for x, y in e.get_points(format="xy")]
        if e.closed:
            layers[layer]["closed"].append(pts)
        else:
            layers[layer]["open"].append(pts)
    elif t == "LINE":
        s = to_ntm(e.dxf.start.x, e.dxf.start.y)
        end = to_ntm(e.dxf.end.x, e.dxf.end.y)
        layers[layer]["lines"].append([s, end])
    elif t == "CIRCLE":
        cx, cy = to_ntm(e.dxf.center.x, e.dxf.center.y)
        r = e.dxf.radius
        # Approximate circle as polygon - transform each vertex
        circle_pts_model = [(e.dxf.center.x + r * math.cos(2 * math.pi * i / 36),
                             e.dxf.center.y + r * math.sin(2 * math.pi * i / 36)) for i in range(36)]
        pts = [to_ntm(x, y) for x, y in circle_pts_model]
        layers[layer]["circles"].append(pts)
    elif t == "ELLIPSE":
        try:
            verts = list(e.vertices(list(range(0, 360, 10))))
            pts = [to_ntm(v.x, v.y) for v in verts]
            if len(pts) >= 3:
                layers[layer]["closed"].append(pts)
        except:
            pass


def dist(p1, p2):
    return math.sqrt((p1[0] - p2[0]) ** 2 + (p1[1] - p2[1]) ** 2)


def pair_open_polylines(open_pls, tol=ENDPOINT_TOL):
    """Try to pair open polylines into closed loops by matching endpoints."""
    closed = []
    used = set()

    for i in range(len(open_pls)):
        if i in used:
            continue
        pts_i = open_pls[i]
        start_i, end_i = pts_i[0], pts_i[-1]

        best_j = None
        best_type = None

        for j in range(i + 1, len(open_pls)):
            if j in used:
                continue
            pts_j = open_pls[j]
            start_j, end_j = pts_j[0], pts_j[-1]

            # Case 1: end_i matches start_j, end_j matches start_i (loop)
            if dist(end_i, start_j) < tol and dist(end_j, start_i) < tol:
                best_j = j
                best_type = "forward"
                break
            # Case 2: end_i matches end_j, start_j matches start_i
            if dist(end_i, end_j) < tol and dist(start_i, start_j) < tol:
                best_j = j
                best_type = "reverse"
                break
            # Case 3: start_i matches start_j (parallel paths going same direction)
            if dist(start_i, start_j) < tol and dist(end_i, end_j) < tol:
                best_j = j
                best_type = "parallel_same"
                break
            # Case 4: start_i matches end_j
            if dist(start_i, end_j) < tol and dist(end_i, start_j) < tol:
                best_j = j
                best_type = "forward"
                break

        if best_j is not None:
            pts_j = open_pls[best_j]
            if best_type == "forward":
                combined = pts_i + pts_j
            elif best_type == "reverse":
                combined = pts_i + list(reversed(pts_j))
            elif best_type == "parallel_same":
                combined = pts_i + list(reversed(pts_j))
            closed.append(combined)
            used.add(i)
            used.add(best_j)

    # Remaining unpaired - close them individually (connect start to end)
    unpaired = []
    for i in range(len(open_pls)):
        if i not in used:
            pts = open_pls[i]
            if len(pts) >= 3:
                # Close it off
                closed.append(pts)
                unpaired.append(i)
            elif len(pts) == 2:
                # Skip 2-point open segments (tree symbol lines etc)
                pass

    return closed, len(unpaired)


# ── Pair open polylines per layer ──
all_outlines = {}  # layer -> list of closed point lists
stats = {}

for layer in sorted(layers.keys()):
    d = layers[layer]
    outlines = []

    # Direct closed polylines
    outlines.extend(d["closed"])

    # Circles
    outlines.extend(d["circles"])

    # Pair open polylines
    if d["open"]:
        paired, n_unpaired = pair_open_polylines(d["open"])
        outlines.extend(paired)

    # Lines - group into chains if possible, otherwise skip (too short for slabs)
    # Lines on besøkssenter are individual segments, not slab outlines

    if outlines:
        all_outlines[layer] = outlines
        stats[layer] = len(outlines)

print("Outlines per layer:")
for layer, count in sorted(stats.items()):
    print(f"  {layer}: {count}")
print(f"\nTotal slab outlines: {sum(stats.values())}")

# ── Build IFC ──
ifc = ifcopenshell.api.run("project.create_file")
project = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcProject", name="Kistefos Stier")

# Units - meters
ifcopenshell.api.run("unit.assign_unit", ifc, length={"is_metric": True, "raw": "METERS"})

# Context
ctx = ifcopenshell.api.run("context.add_context", ifc, context_type="Model")
body = ifcopenshell.api.run(
    "context.add_context",
    ifc,
    context_type="Model",
    context_identifier="Body",
    target_view="MODEL_VIEW",
    parent=ctx,
)

# Site
site = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcSite", name="Kistefos")
ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=project, products=[site])

# Building
building = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcBuilding", name="Landscape")
ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=site, products=[building])

# Storey
storey = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcBuildingStorey", name="Ground")
ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=building, products=[storey])

# Coordination marker - cylinder at basepoint
marker_slab = ifcopenshell.api.run(
    "root.create_entity", ifc, ifc_class="IfcSlab", name="COORDINATION_MARKER_NEW"
)
OLD_BP_E, OLD_BP_N = to_ntm(0, 0)  # model origin = old basepoint in NTM
NEW_BP_E, NEW_BP_N = 92200.0, 1247000.0
marker_pts = [(NEW_BP_E + 2 * math.cos(2 * math.pi * i / 36),
               NEW_BP_N + 2 * math.sin(2 * math.pi * i / 36)) for i in range(36)]
marker_profile = ifc.createIfcArbitraryClosedProfileDef(
    "AREA",
    None,
    ifc.createIfcPolyline([ifc.createIfcCartesianPoint([float(p[0]), float(p[1])]) for p in marker_pts]
                          + [ifc.createIfcCartesianPoint([float(marker_pts[0][0]), float(marker_pts[0][1])])]),
)
marker_solid = ifc.createIfcExtrudedAreaSolid(
    marker_profile,
    ifc.createIfcAxis2Placement3D(ifc.createIfcCartesianPoint([0.0, 0.0, GROUND_Z])),
    ifc.createIfcDirection([0.0, 0.0, 1.0]),
    5.0,  # 5m tall cylinder for visibility
)
marker_rep = ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [marker_solid])
marker_prod_rep = ifc.createIfcProductDefinitionShape(None, None, [marker_rep])
marker_slab.Representation = marker_prod_rep
marker_slab.ObjectPlacement = ifcopenshell.api.run(
    "geometry.edit_object_placement", ifc, product=marker_slab
)
ifcopenshell.api.run("spatial.assign_container", ifc, relating_structure=storey, products=[marker_slab])

# Old basepoint marker
old_e, old_n = OLD_BP_E, OLD_BP_N
old_marker = ifcopenshell.api.run(
    "root.create_entity", ifc, ifc_class="IfcSlab", name="COORDINATION_MARKER_OLD (UTM32: 575200/6676400)"
)
old_pts = [(old_e + 1.5 * math.cos(2 * math.pi * i / 36),
            old_n + 1.5 * math.sin(2 * math.pi * i / 36)) for i in range(36)]
old_profile = ifc.createIfcArbitraryClosedProfileDef(
    "AREA",
    None,
    ifc.createIfcPolyline([ifc.createIfcCartesianPoint([float(p[0]), float(p[1])]) for p in old_pts]
                          + [ifc.createIfcCartesianPoint([float(old_pts[0][0]), float(old_pts[0][1])])]),
)
old_solid = ifc.createIfcExtrudedAreaSolid(
    old_profile,
    ifc.createIfcAxis2Placement3D(ifc.createIfcCartesianPoint([0.0, 0.0, GROUND_Z])),
    ifc.createIfcDirection([0.0, 0.0, 1.0]),
    3.0,  # 3m tall, shorter than new marker
)
old_rep = ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [old_solid])
old_prod_rep = ifc.createIfcProductDefinitionShape(None, None, [old_rep])
old_marker.Representation = old_prod_rep
old_marker.ObjectPlacement = ifcopenshell.api.run(
    "geometry.edit_object_placement", ifc, product=old_marker
)
ifcopenshell.api.run("spatial.assign_container", ifc, relating_structure=storey, products=[old_marker])

# Layer color map (R, G, B 0-1)
LAYER_COLOR_MAP = {
    "00 Stier(Nye konstruksjoner)": (0.82, 0.71, 0.55),          # sand/path
    "00- bes\u00f8kssenter(Nye konstruksjoner)": (0.7, 0.7, 0.7),  # gray/buildings
    "00- nye stier(Nye konstruksjoner)": (0.9, 0.8, 0.6),         # light sand
    "00- ny jord(Nye konstruksjoner)": (0.55, 0.35, 0.17),        # brown/earth
    "00-forflytning jord(Nye konstruksjoner)": (0.6, 0.4, 0.2),   # brown
    "00- Hotlink Bygg(Nye konstruksjoner)": (0.8, 0.8, 0.8),      # light gray
    "761- Veier_ Kj\u00f8reveier_ sykkel- og gangveier mv_(Nye konstruksjoner)": (0.3, 0.3, 0.3),  # dark gray/asphalt
    "790- Utstyr_ m\u00f8bler_ lekeapparater(Nye konstruksjoner)": (0.9, 0.5, 0.1),  # orange
    "837- Planteplan tr\u00e6r pisk BJ\u00d8RK(Nye konstruksjoner)": (0.4, 0.7, 0.3),  # green
    "837- Planteplan tr\u00e6r pisk FURU(Nye konstruksjoner)": (0.2, 0.5, 0.2),        # dark green
    "837- Planteplan tr\u00e6r pisk HEGG(Nye konstruksjoner)": (0.5, 0.8, 0.4),        # light green
    "837- Planteplan tr\u00e6r pisk OR(Nye konstruksjoner)": (0.3, 0.6, 0.3),          # green
    "837- Planteplan tr\u00e6r pisk ROGN(Nye konstruksjoner)": (0.6, 0.8, 0.3),        # yellow-green
    "837- Planteplan tr\u00e6r so 12-14(Nye konstruksjoner)": (0.1, 0.5, 0.1),         # forest green
    "837- Planteplan tr\u00e6r transplantert(Nye konstruksjoner)": (0.3, 0.7, 0.5),    # teal
    "AnnetGjerde(Nye konstruksjoner)": (0.5, 0.5, 0.5),           # gray/fence
    "Veg(Nye konstruksjoner)": (0.25, 0.25, 0.25),                # asphalt
    "Vegdekkekant(Nye konstruksjoner)": (0.35, 0.35, 0.35),       # road edge
    "Veggr\u00f8ft\u00e5pen(Nye konstruksjoner)": (0.3, 0.5, 0.7),  # blue/ditch
    "VegkantAnnetVegareal(Nye konstruksjoner)": (0.4, 0.4, 0.4),
    "VegkantAvkj\u00f8rsel(Nye konstruksjoner)": (0.4, 0.4, 0.4),
    "VegkantFiktiv(Nye konstruksjoner)": (0.5, 0.5, 0.5),
    "Vegrekkverk(Nye konstruksjoner)": (0.6, 0.6, 0.6),           # guardrail
    "Veranda(Nye konstruksjoner)": (0.65, 0.5, 0.35),             # wood
    "f-64000_fortau_FLATE_GANGSYKKELVEG(Nye konstruksjoner)": (0.7, 0.7, 0.7),  # sidewalk
    "f-veg_21000_FLATE_KJ\u00d8REFELT_ASFALT(Nye konstruksjoner)": (0.2, 0.2, 0.2),  # asphalt
    "PLANKART januar(Nye konstruksjoner)": (0.9, 0.9, 0.4),       # yellow/plan
    "0(Nye konstruksjoner)": (0.6, 0.6, 0.6),
}
layer_colors = {}

# Classify layers into IFC types
WALL_PREFIXES = ["AnnetGjerde", "Vegrekkverk", "Vegbom", "Veggr"]  # fences, guardrails, barriers, ditches
SKIP_PREFIXES = ["837- Planteplan", "83-- Tekst", "00- Hotlink"]  # trees, text, hotlinks

WALL_HEIGHT = 1.5  # meters

def get_ifc_class(layer):
    for prefix in SKIP_PREFIXES:
        if prefix in layer:
            return None
    for prefix in WALL_PREFIXES:
        if prefix in layer:
            return "IfcWall"
    return "IfcSlab"

# Create elements
slab_count = 0
wall_count = 0
skipped_layers = set()
for layer, outlines in sorted(all_outlines.items()):
    ifc_class = get_ifc_class(layer)
    if ifc_class is None:
        skipped_layers.add(layer)
        continue

    for i, pts in enumerate(outlines):
        if len(pts) < 3:
            continue

        name = f"{layer}_{i}"
        element = ifcopenshell.api.run("root.create_entity", ifc, ifc_class=ifc_class, name=name)
        is_wall = (ifc_class == "IfcWall")
        extrude_height = WALL_HEIGHT if is_wall else SLAB_THICKNESS
        extrude_z_start = GROUND_Z if is_wall else GROUND_Z - SLAB_THICKNESS

        # Create profile from points
        ifc_pts = [ifc.createIfcCartesianPoint([float(p[0]), float(p[1])]) for p in pts]
        ifc_pts.append(ifc_pts[0])  # close the loop
        polyline = ifc.createIfcPolyline(ifc_pts)
        profile = ifc.createIfcArbitraryClosedProfileDef("AREA", None, polyline)

        # Extrude
        solid = ifc.createIfcExtrudedAreaSolid(
            profile,
            ifc.createIfcAxis2Placement3D(
                ifc.createIfcCartesianPoint([0.0, 0.0, extrude_z_start])
            ),
            ifc.createIfcDirection([0.0, 0.0, 1.0]),
            extrude_height,
        )

        rep = ifc.createIfcShapeRepresentation(body, "Body", "SweptSolid", [solid])
        prod_rep = ifc.createIfcProductDefinitionShape(None, None, [rep])
        element.Representation = prod_rep
        element.ObjectPlacement = ifcopenshell.api.run(
            "geometry.edit_object_placement", ifc, product=element
        )
        ifcopenshell.api.run(
            "spatial.assign_container", ifc, relating_structure=storey, products=[element]
        )

        # Color by layer
        if layer not in layer_colors:
            layer_colors[layer] = LAYER_COLOR_MAP.get(layer, (0.6, 0.6, 0.6))
        r, g, b = layer_colors[layer]
        color = ifc.createIfcColourRgb(None, r, g, b)
        surface_style = ifc.createIfcSurfaceStyleRendering(
            color, 0.0, None, None, None, None, None, None, "FLAT"
        )
        style = ifc.createIfcSurfaceStyle(layer, "BOTH", [surface_style])
        styled_item = ifc.createIfcStyledItem(solid, [ifc.createIfcPresentationStyleAssignment([style])], None)

        # NOSKI_DXF property set
        pset = ifcopenshell.api.run(
            "pset.add_pset", ifc, product=element, name="NOSKI_DXF"
        )
        ifcopenshell.api.run(
            "pset.edit_pset",
            ifc,
            pset=pset,
            properties={
                "Layer": layer,
                "SourceFile": "KNM_Stier.dwg",
                "IfcClass": ifc_class,
                "CoordinateSystem": "EUREF89 NTM Sone 10 (EPSG:5110)",
                "BasePoint_NTM_E": str(OLD_BP_E),
                "BasePoint_NTM_N": str(OLD_BP_N),
                "BasePoint_UTM32_E": "575200",
                "BasePoint_UTM32_N": "6676400",
                "HeightSystem": "NN2000",
                "Thickness": str(extrude_height),
            },
        )

        if is_wall:
            wall_count += 1
        else:
            slab_count += 1

print(f"\nSkipped layers (trees/text): {sorted(skipped_layers)}")

ifc.write(IFC_OUT)
print(f"\nIFC saved: {IFC_OUT}")
print(f"Slabs: {slab_count}, Walls: {wall_count}, Markers: 2")
