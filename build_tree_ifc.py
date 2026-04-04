"""
Build IFC tree models from Kistefos survey DXF data.

Reads two DXF files:
  - Innmålt_Tre: tree blocks with species/height/crown/trunk attributes
  - TreStammer_Diameter: trunk cross-section circles with accurate Z

Creates parametric 3D trees with distinct shapes:
  - Conifer (Gran, Furu): pointed cone crown
  - Deciduous (Lauvtre, Bjørk): rounded/layered dome crown
  - Dead conifer (TørrGran): sparse pointed trunk-only
  - Dead deciduous (Lauvtre-Tørr): bare trunk

Outputs NTM Sone 10 georeferenced IFC (global + local).
"""

import ezdxf
import ifcopenshell
import ifcopenshell.api
import math
import os
import re
import sys
from collections import defaultdict
from pyproj import Transformer

# ── Config ──────────────────────────────────────────────────────────────────
DXF_DIR = os.path.dirname(os.path.abspath(__file__))
# Resolve source DXF paths - check common locations
_PROJECT_BASE = os.path.join(
    os.path.expanduser("~"),
    "DC", "ACCDocs", "Skiplum AS", "Skiplum Backup",
    "Project Files", "10016 - Kistefos",
)
_DXF_SRC = os.path.join(_PROJECT_BASE, "02_Arbeid", "09_Tegninger", "DXF_UTF")
_OUT_DIR = os.path.join(_PROJECT_BASE, "03_Ut", "08_Tegninger")

def _find_dxf(directory, keyword):
    """Find DXF file by keyword, handling unicode normalization."""
    for f in os.listdir(directory):
        if keyword in f and f.endswith(".dxf"):
            return os.path.join(directory, f)
    raise FileNotFoundError(f"No DXF matching '{keyword}' in {directory}")

TREE_DXF = _find_dxf(_DXF_SRC, "Tre")  # matches Innmålt_Tre but not TreStammer
# Resolve: pick the one WITHOUT "Stammer"
_tree_candidates = [f for f in os.listdir(_DXF_SRC)
                    if "Tre" in f and "Stammer" not in f and f.endswith(".dxf")]
TREE_DXF = os.path.join(_DXF_SRC, _tree_candidates[0])
TRUNK_DXF = _find_dxf(_DXF_SRC, "TreStammer")
IFC_OUT_GLOBAL = os.path.join(_OUT_DIR, "Kistefos_Trees_NTM10_global_meters.ifc")
IFC_OUT_LOCAL = os.path.join(_OUT_DIR, "Kistefos_Trees_NTM10_local_meters.ifc")

UTM_BP_E = 575200.0
UTM_BP_N = 6676400.0
NTM_BP_E = 92200.0
NTM_BP_N = 1247000.0

NSEG_CONIFER = 7   # fewer sides = more organic, less obviously "round"
NSEG_DECIDUOUS = 12  # wider crowns need more segments to read as cylindrical

# ── Coordinate transform ───────────────────────────────────────────────────
_transformer = Transformer.from_crs("EPSG:25832", "EPSG:5110", always_xy=True)


def to_ntm(utm_e, utm_n):
    """Transform UTM32 world coords to NTM Sone 10."""
    return _transformer.transform(utm_e, utm_n)


# ── Tree shape categories ──────────────────────────────────────────────────
# Maps species name (from DXF) to shape category
SPECIES_CATEGORY = {
    "Gran":         "conifer",
    "Furu":         "conifer",
    "Bjørk":        "deciduous",
    "Lauvtre":      "deciduous",
    "TørrGran":     "dead_conifer",
    "Lauvtre-Tørr": "dead_deciduous",
}

# Colors per species (R, G, B  0–1)
SPECIES_CROWN_COLOR = {
    "Gran":         (0.20, 0.50, 0.18),   # medium green
    "Furu":         (0.10, 0.35, 0.12),   # dark green (darker than Gran)
    "Bjørk":        (0.40, 0.75, 0.25),   # bright green
    "Lauvtre":      (0.40, 0.75, 0.25),   # bright green
    "TørrGran":     (0.40, 0.50, 0.35),   # dull green/grey
    "Lauvtre-Tørr": (0.55, 0.65, 0.20),   # green going on yellow
}

SPECIES_TRUNK_COLOR = {
    "Bjørk":        (0.90, 0.90, 0.88),   # white birch bark
}
DEFAULT_TRUNK_COLOR = (0.45, 0.30, 0.15)   # brown


# ── Parse DXF data ─────────────────────────────────────────────────────────
def parse_tree_labels(dxf_path):
    """Parse Innmålt_Tre.dxf — returns list of tree dicts."""
    doc = ezdxf.readfile(dxf_path)
    trees = []
    for ins in doc.modelspace():
        if ins.dxftype() != "INSERT":
            continue
        block = doc.blocks.get(ins.dxf.name)
        if not block:
            continue

        # Find TEXT with attributes and LINE/POINT for position
        text_ent = None
        pos_x, pos_y, pos_z = None, None, None

        for e in block:
            if e.dxftype() == "TEXT" and "TreTyp:" in e.dxf.text:
                text_ent = e
                # Text insert position is near the tree
                pos_x = e.dxf.insert.x
                pos_y = e.dxf.insert.y
                pos_z = e.dxf.insert.z
            elif e.dxftype() == "LINE" and pos_x is None:
                # Lines form a cross at tree position — use midpoint of first line
                pos_x = (e.dxf.start.x + e.dxf.end.x) / 2
                pos_y = (e.dxf.start.y + e.dxf.end.y) / 2
                pos_z = 0.0

        if text_ent is None or pos_x is None:
            continue

        txt = text_ent.dxf.text.replace("^J", "\n")
        attrs = {}
        for line in txt.split("\n"):
            line = line.strip()
            if ":" in line:
                key, val = line.split(":", 1)
                attrs[key.strip()] = val.strip()

        species = attrs.get("TreTyp", "Unknown")
        try:
            trunk_dia = float(attrs.get("StammeDia", "0.2"))
        except ValueError:
            trunk_dia = 0.2
        try:
            height = float(attrs.get("TreHøyde", "10"))
        except ValueError:
            height = 10.0
        try:
            crown_dia = float(attrs.get("KronDia", "4"))
        except ValueError:
            crown_dia = 4.0

        trees.append({
            "species": species,
            "trunk_dia": trunk_dia,
            "height": height,
            "crown_dia": crown_dia,
            "utm_e": pos_x,
            "utm_n": pos_y,
            "z": pos_z,
            "source": "label",
        })

    return trees


def parse_trunk_circles(dxf_path):
    """Parse TreStammer_Diameter.dxf — returns list of trunk dicts."""
    doc = ezdxf.readfile(dxf_path)
    trunks = []
    for ins in doc.modelspace():
        if ins.dxftype() != "INSERT":
            continue
        block = doc.blocks.get(ins.dxf.name)
        if not block:
            continue
        for e in block:
            if e.dxftype() == "CIRCLE":
                trunks.append({
                    "utm_e": e.dxf.center.x,
                    "utm_n": e.dxf.center.y,
                    "z": e.dxf.center.z,
                    "trunk_dia": e.dxf.radius * 2,
                })
    return trunks


def match_trees(label_trees, trunk_circles, tol=2.0):
    """Match label trees to trunk circles by proximity. Use trunk Z when matched."""
    # Build spatial index (simple — 1500 trees is fine with brute force)
    matched = []
    used_trunks = set()

    for tree in label_trees:
        best_dist = tol
        best_idx = None
        for i, trunk in enumerate(trunk_circles):
            if i in used_trunks:
                continue
            d = math.sqrt((tree["utm_e"] - trunk["utm_e"]) ** 2 +
                          (tree["utm_n"] - trunk["utm_n"]) ** 2)
            if d < best_dist:
                best_dist = d
                best_idx = i

        if best_idx is not None:
            trunk = trunk_circles[best_idx]
            used_trunks.add(best_idx)
            tree["z"] = trunk["z"]  # use trunk Z (more accurate)
            tree["trunk_dia"] = trunk["trunk_dia"]  # use measured diameter
            tree["utm_e"] = trunk["utm_e"]  # use trunk center position
            tree["utm_n"] = trunk["utm_n"]
        matched.append(tree)

    print(f"  Matched {len(used_trunks)}/{len(label_trees)} trees to trunk circles")
    return matched


# ── Geometry builders ──────────────────────────────────────────────────────
def ring_points(cx, cy, cz, radius, n=NSEG_DECIDUOUS):
    """Generate n points around a circle at (cx, cy, cz)."""
    return [(cx + radius * math.cos(2 * math.pi * i / n),
             cy + radius * math.sin(2 * math.pi * i / n),
             cz) for i in range(n)]


def make_frustum_faces(bottom_ring, top_ring):
    """Create triangulated faces connecting two rings of equal length."""
    n = len(bottom_ring)
    faces = []
    for i in range(n):
        j = (i + 1) % n
        # Two triangles per quad
        faces.append([bottom_ring[i], bottom_ring[j], top_ring[j]])
        faces.append([bottom_ring[i], top_ring[j], top_ring[i]])
    return faces


def make_cap_faces(ring, center, flip=False):
    """Create triangulated fan cap."""
    n = len(ring)
    faces = []
    for i in range(n):
        j = (i + 1) % n
        if flip:
            faces.append([center, ring[j], ring[i]])
        else:
            faces.append([center, ring[i], ring[j]])
    return faces


def build_conifer_geometry(height, crown_dia, trunk_dia):
    """
    Conifer (Gran/Furu): short trunk + single large cone tapering to a point.
    7-sided polygon — fewer sides reads more organic than a smooth cone.
    """
    parts = []
    n = NSEG_CONIFER
    trunk_r = trunk_dia / 2
    trunk_h = height * 0.12  # conifers: branches almost reach the ground
    crown_h = height - trunk_h
    crown_r = crown_dia / 2

    # Trunk — short cylinder
    bottom = ring_points(0, 0, 0, trunk_r, n)
    top = ring_points(0, 0, trunk_h, trunk_r, n)
    faces = make_frustum_faces(bottom, top)
    faces += make_cap_faces(bottom, (0, 0, 0), flip=True)
    faces += make_cap_faces(top, (0, 0, trunk_h))
    parts.append(("trunk", faces))

    # Crown — single cone from wide base to pointed tip
    base_ring = ring_points(0, 0, trunk_h, crown_r, n)
    tip = (0, 0, height)
    faces = make_cap_faces(base_ring, tip)
    faces += make_cap_faces(base_ring, (0, 0, trunk_h), flip=True)
    parts.append(("crown", faces))

    return parts


def build_deciduous_geometry(height, crown_dia, trunk_dia):
    """
    Deciduous (Lauvtre/Bjørk): tall bare trunk, branches start high up,
    then a wide cylindrical/slightly domed crown that doesn't taper much.
    """
    parts = []
    trunk_r = trunk_dia / 2
    # Branches start at ~40% of height (at least 2m clearance)
    trunk_h = max(height * 0.40, 2.5)
    crown_h = height - trunk_h
    crown_r = crown_dia / 2

    # Trunk — tall cylinder, slight taper
    bottom = ring_points(0, 0, 0, trunk_r)
    top = ring_points(0, 0, trunk_h, trunk_r * 0.85)
    faces = make_frustum_faces(bottom, top)
    faces += make_cap_faces(bottom, (0, 0, 0), flip=True)
    faces += make_cap_faces(top, (0, 0, trunk_h))
    parts.append(("trunk", faces))

    # Crown — wide cylinder with a domed top
    # Lower 70%: straight cylinder at full crown radius
    cyl_h = crown_h * 0.7
    dome_h = crown_h * 0.3

    # Cylinder portion
    cyl_bottom = ring_points(0, 0, trunk_h, crown_r)
    cyl_top = ring_points(0, 0, trunk_h + cyl_h, crown_r)
    faces = make_frustum_faces(cyl_bottom, cyl_top)
    faces += make_cap_faces(cyl_bottom, (0, 0, trunk_h), flip=True)
    parts.append(("crown", faces))

    # Dome top — 2 frustums tapering to a rounded top
    mid_z = trunk_h + cyl_h + dome_h * 0.5
    top_z = trunk_h + cyl_h + dome_h
    mid_ring = ring_points(0, 0, mid_z, crown_r * 0.75)
    top_ring = ring_points(0, 0, top_z, crown_r * 0.35)

    faces = make_frustum_faces(cyl_top, mid_ring)
    parts.append(("crown", faces))

    faces = make_frustum_faces(mid_ring, top_ring)
    faces += make_cap_faces(top_ring, (0, 0, top_z))
    parts.append(("crown", faces))

    return parts


def build_dead_conifer_geometry(height, crown_dia, trunk_dia):
    """Dead conifer: trunk + thinner/sparser cone (still has old branches). 7-sided."""
    parts = []
    n = NSEG_CONIFER
    trunk_r = trunk_dia / 2
    trunk_h = height * 0.15
    crown_r = crown_dia / 2 * 0.55  # narrower than live conifer
    crown_h = height - trunk_h

    # Trunk
    bottom = ring_points(0, 0, 0, trunk_r, n)
    top = ring_points(0, 0, trunk_h, trunk_r, n)
    faces = make_frustum_faces(bottom, top)
    faces += make_cap_faces(bottom, (0, 0, 0), flip=True)
    faces += make_cap_faces(top, (0, 0, trunk_h))
    parts.append(("trunk", faces))

    # Sparse cone crown — same shape as conifer but smaller radius
    base_ring = ring_points(0, 0, trunk_h, crown_r, n)
    tip = (0, 0, height)
    faces = make_cap_faces(base_ring, tip)
    faces += make_cap_faces(base_ring, (0, 0, trunk_h), flip=True)
    parts.append(("crown", faces))

    return parts


def build_dead_deciduous_geometry(height, crown_dia, trunk_dia):
    """Dead deciduous: tall trunk + small sparse crown blob."""
    parts = []
    trunk_r = trunk_dia / 2
    trunk_h = max(height * 0.45, 2.5)
    crown_r = crown_dia / 2 * 0.5  # smaller than live deciduous
    crown_h = height - trunk_h

    # Trunk — tapered
    bottom = ring_points(0, 0, 0, trunk_r)
    top = ring_points(0, 0, trunk_h, trunk_r * 0.8)
    faces = make_frustum_faces(bottom, top)
    faces += make_cap_faces(bottom, (0, 0, 0), flip=True)
    faces += make_cap_faces(top, (0, 0, trunk_h))
    parts.append(("trunk", faces))

    # Small cylindrical crown with dome
    cyl_top_z = trunk_h + crown_h * 0.6
    dome_top_z = trunk_h + crown_h
    cyl_bottom = ring_points(0, 0, trunk_h, crown_r)
    cyl_top = ring_points(0, 0, cyl_top_z, crown_r)
    faces = make_frustum_faces(cyl_bottom, cyl_top)
    faces += make_cap_faces(cyl_bottom, (0, 0, trunk_h), flip=True)
    parts.append(("crown", faces))

    # Dome cap
    dome_ring = ring_points(0, 0, dome_top_z, crown_r * 0.3)
    faces = make_frustum_faces(cyl_top, dome_ring)
    faces += make_cap_faces(dome_ring, (0, 0, dome_top_z))
    parts.append(("crown", faces))

    return parts


def build_conifer_layered_geometry(height, crown_dia, trunk_dia):
    """
    DEMO SHAPE — stacked rotated cones for a more detailed silhouette.
    Each tier is a cone, slightly smaller and rotated from the one below.
    Looks great up close. Murders your file size at 1500 trees.
    """
    parts = []
    n = NSEG_CONIFER
    trunk_r = trunk_dia / 2
    trunk_h = height * 0.15
    crown_h = height - trunk_h
    crown_r = crown_dia / 2

    # Trunk
    bottom = ring_points(0, 0, 0, trunk_r, n)
    top = ring_points(0, 0, trunk_h, trunk_r, n)
    faces = make_frustum_faces(bottom, top)
    faces += make_cap_faces(bottom, (0, 0, 0), flip=True)
    faces += make_cap_faces(top, (0, 0, trunk_h))
    parts.append(("trunk", faces))

    # Stacked cones — each tier overlaps the one above, rotated
    n_tiers = 5
    tier_h = crown_h / n_tiers
    for t in range(n_tiers):
        # Each tier starts slightly below the previous tier's midpoint (overlap)
        base_z = trunk_h + t * tier_h * 0.75
        tip_z = base_z + tier_h * 1.3
        # Radius shrinks toward the top
        frac = 1.0 - (t / n_tiers) * 0.7
        r = crown_r * frac
        # Rotate each tier by an offset angle for organic irregularity
        angle_offset = t * (2 * math.pi / n_tiers) * 0.6
        ring = [(r * math.cos(2 * math.pi * i / n + angle_offset),
                 r * math.sin(2 * math.pi * i / n + angle_offset),
                 base_z) for i in range(n)]
        tip = (0, 0, min(tip_z, height))
        faces = make_cap_faces(ring, tip)
        faces += make_cap_faces(ring, (0, 0, base_z), flip=True)
        parts.append(("crown", faces))

    return parts


GEOMETRY_BUILDERS = {
    "conifer":          build_conifer_geometry,
    "conifer_layered":  build_conifer_layered_geometry,
    "deciduous":        build_deciduous_geometry,
    "dead_conifer":     build_dead_conifer_geometry,
    "dead_deciduous":   build_dead_deciduous_geometry,
}


# ── IFC helpers ────────────────────────────────────────────────────────────
def faces_to_ifc_brep(ifc, faces):
    """Convert triangle face list to IfcFacetedBrep."""
    ifc_faces = []
    for tri in faces:
        pts = [ifc.createIfcCartesianPoint([float(p[0]), float(p[1]), float(p[2])])
               for p in tri]
        # Close the loop
        pts.append(pts[0])
        loop = ifc.createIfcPolyLoop(pts[:-1])
        bound = ifc.createIfcFaceOuterBound(loop, True)
        ifc_faces.append(ifc.createIfcFace([bound]))

    shell = ifc.createIfcClosedShell(ifc_faces)
    return ifc.createIfcFacetedBrep(shell)


def create_ifc_color(ifc, r, g, b):
    """Create surface style for coloring."""
    color = ifc.createIfcColourRgb(None, r, g, b)
    rendering = ifc.createIfcSurfaceStyleRendering(
        color, 0.0, None, None, None, None, None, None, "FLAT"
    )
    style = ifc.createIfcSurfaceStyle(None, "BOTH", [rendering])
    return style


# ── Main build ─────────────────────────────────────────────────────────────
def build_ifc(trees, local=False):
    """Build IFC file from tree list. If local=True, subtract NTM basepoint."""
    ifc = ifcopenshell.api.run("project.create_file")
    project = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcProject", name="Kistefos Trees"
    )
    ifcopenshell.api.run("unit.assign_unit", ifc, length={"is_metric": True, "raw": "METERS"})

    ctx = ifcopenshell.api.run("context.add_context", ifc, context_type="Model")
    body = ifcopenshell.api.run(
        "context.add_context", ifc,
        context_type="Model",
        context_identifier="Body",
        target_view="MODEL_VIEW",
        parent=ctx,
    )

    site = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcSite", name="Kistefos")
    ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=project, products=[site])

    building = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcBuilding", name="Landscape"
    )
    ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=site, products=[building])

    storey = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcBuildingStorey", name="Ground"
    )
    ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=building, products=[storey])

    # Pre-create color styles per species
    trunk_styles = {}
    for sp, (r, g, b) in SPECIES_TRUNK_COLOR.items():
        trunk_styles[sp] = create_ifc_color(ifc, r, g, b)
    default_trunk_style = create_ifc_color(ifc, *DEFAULT_TRUNK_COLOR)
    crown_styles = {}
    for sp, (r, g, b) in SPECIES_CROWN_COLOR.items():
        crown_styles[sp] = create_ifc_color(ifc, r, g, b)

    offset_e = NTM_BP_E if local else 0.0
    offset_n = NTM_BP_N if local else 0.0

    counts = defaultdict(int)
    for i, tree in enumerate(trees):
        # Transform position
        ntm_e, ntm_n = to_ntm(tree["utm_e"], tree["utm_n"])
        x = ntm_e - offset_e
        y = ntm_n - offset_n
        z = tree["z"]

        category = SPECIES_CATEGORY.get(tree["species"], "deciduous")
        builder = GEOMETRY_BUILDERS[category]

        # Clamp values to reasonable ranges
        height = max(tree["height"], 2.0)
        crown_dia = max(tree["crown_dia"], 1.0)
        trunk_dia = max(tree["trunk_dia"], 0.05)

        parts = builder(height, crown_dia, trunk_dia)

        # Combine all faces into one brep, offset to tree position
        all_faces_trunk = []
        all_faces_crown = []
        for part_name, faces in parts:
            offset_faces = [
                [(p[0] + x, p[1] + y, p[2] + z) for p in tri]
                for tri in faces
            ]
            if part_name == "trunk":
                all_faces_trunk.extend(offset_faces)
            else:
                all_faces_crown.extend(offset_faces)

        # Create IFC elements — separate trunk and crown for coloring
        items = []
        species = tree["species"]
        t_style = trunk_styles.get(species, default_trunk_style)
        c_style = crown_styles.get(species, crown_styles.get("Lauvtre"))

        if all_faces_trunk:
            brep_trunk = faces_to_ifc_brep(ifc, all_faces_trunk)
            ifc.createIfcStyledItem(
                brep_trunk,
                [ifc.createIfcPresentationStyleAssignment([t_style])],
                None,
            )
            items.append(brep_trunk)
        if all_faces_crown:
            brep_crown = faces_to_ifc_brep(ifc, all_faces_crown)
            ifc.createIfcStyledItem(
                brep_crown,
                [ifc.createIfcPresentationStyleAssignment([c_style])],
                None,
            )
            items.append(brep_crown)

        if not items:
            continue

        name = f"Tree_{i+1:04d}_{tree['species']}"
        element = ifcopenshell.api.run(
            "root.create_entity", ifc,
            ifc_class="IfcBuildingElementProxy",
            name=name,
        )

        rep = ifc.createIfcShapeRepresentation(body, "Body", "Brep", items)
        prod_rep = ifc.createIfcProductDefinitionShape(None, None, [rep])
        element.Representation = prod_rep
        element.ObjectPlacement = ifcopenshell.api.run(
            "geometry.edit_object_placement", ifc, product=element
        )
        ifcopenshell.api.run(
            "spatial.assign_container", ifc,
            relating_structure=storey, products=[element],
        )

        # Property set
        pset = ifcopenshell.api.run("pset.add_pset", ifc, product=element, name="NOSKI_Trees")
        ifcopenshell.api.run(
            "pset.edit_pset", ifc, pset=pset,
            properties={
                "Species": tree["species"],
                "Category": category,
                "TrunkDiameter": str(round(tree["trunk_dia"], 3)),
                "TreeHeight": str(round(tree["height"], 1)),
                "CrownDiameter": str(round(tree["crown_dia"], 1)),
                "Position_E": str(round(ntm_e, 3)),
                "Position_N": str(round(ntm_n, 3)),
                "Elevation_Z": str(round(z, 2)),
                "CoordinateSystem": "EUREF89 NTM Sone 10 (EPSG:5110)",
                "HeightSystem": "NN2000",
                "SourceFile": "ACAD-Innmålt_Tre.dxf + ACAD-TreStammer_Diameter.dxf",
            },
        )

        counts[category] += 1

        if (i + 1) % 200 == 0:
            print(f"  {i+1}/{len(trees)} trees...")

    # Coordination markers
    for marker_name, me, mn in [
        ("COORD_MARKER_NTM_BP (E=92200 N=1247000)", NTM_BP_E, NTM_BP_N),
        ("COORD_MARKER_UTM32_BP (E=575200 N=6676400)", *to_ntm(UTM_BP_E, UTM_BP_N)),
    ]:
        mx = me - offset_e
        my = mn - offset_n
        marker_ring = ring_points(mx, my, 120.0, 2.0, 24)
        top_ring = ring_points(mx, my, 125.0, 2.0, 24)
        faces = make_frustum_faces(marker_ring, top_ring)
        faces += make_cap_faces(marker_ring, (mx, my, 120.0), flip=True)
        faces += make_cap_faces(top_ring, (mx, my, 125.0))
        brep = faces_to_ifc_brep(ifc, faces)

        marker = ifcopenshell.api.run(
            "root.create_entity", ifc, ifc_class="IfcBuildingElementProxy", name=marker_name
        )
        rep = ifc.createIfcShapeRepresentation(body, "Body", "Brep", [brep])
        prod_rep = ifc.createIfcProductDefinitionShape(None, None, [rep])
        marker.Representation = prod_rep
        marker.ObjectPlacement = ifcopenshell.api.run(
            "geometry.edit_object_placement", ifc, product=marker
        )
        ifcopenshell.api.run(
            "spatial.assign_container", ifc, relating_structure=storey, products=[marker]
        )

    return ifc, counts


def main():
    print("=" * 60)
    print("Kistefos Tree IFC Generator")
    print("=" * 60)

    # Parse DXF files
    print(f"\n1. Reading tree labels: {os.path.basename(TREE_DXF)}")
    label_trees = parse_tree_labels(TREE_DXF)
    print(f"   Found {len(label_trees)} labelled trees")

    print(f"\n2. Reading trunk circles: {os.path.basename(TRUNK_DXF)}")
    trunk_circles = parse_trunk_circles(TRUNK_DXF)
    print(f"   Found {len(trunk_circles)} trunk circles")

    print(f"\n3. Matching trees to trunks...")
    trees = match_trees(label_trees, trunk_circles)

    # Species stats
    from collections import Counter
    sp_counts = Counter(t["species"] for t in trees)
    print(f"\n   Species distribution:")
    for sp, cnt in sp_counts.most_common():
        cat = SPECIES_CATEGORY.get(sp, "deciduous")
        print(f"     {sp}: {cnt} ({cat})")

    # Build global IFC
    print(f"\n4. Building global IFC (NTM10 world coordinates)...")
    ifc_global, counts = build_ifc(trees, local=False)
    ifc_global.write(IFC_OUT_GLOBAL)
    print(f"   Saved: {IFC_OUT_GLOBAL}")
    for cat, cnt in sorted(counts.items()):
        print(f"     {cat}: {cnt}")

    # Build local IFC
    print(f"\n5. Building local IFC (basepoint at origin)...")
    ifc_local, counts = build_ifc(trees, local=True)
    ifc_local.write(IFC_OUT_LOCAL)
    print(f"   Saved: {IFC_OUT_LOCAL}")

    print(f"\nDone! {len(trees)} trees in {len(counts)} categories.")


if __name__ == "__main__":
    main()
