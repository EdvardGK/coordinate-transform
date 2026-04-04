"""
Generate a single hyper-detailed birch leaf (Betula pendula) in IFC.
Video demo prop: "this is what ONE leaf costs in geometry."

One unified mesh — petiole tube morphs into flat blade, veins are
Z-displacement ridges in the surface, not separate objects.

Hierarchy (as displacement layers, not separate geometry):
  1. Spine arc — main droop curve
  2. Cross-section morph — tube (petiole) → flat (blade)
  3. Midrib ridge — Gaussian bump at s=0
  4. Lateral vein ridges — paths on the surface, displaced up
  5. Inter-vein billow — membrane sags between veins
  6. Edge — doubly serrate (primary teeth at vein tips, secondary between)
  7. Organic noise — asymmetry, twist, low-freq undulation
"""
import ifcopenshell
import ifcopenshell.api
import math

# ── Parameters ───────────────────────────────────────────────────────────────
LEAF_LENGTH = 0.06          # 6cm blade
LEAF_WIDTH = 0.042          # ~4cm at widest
PETIOLE_LENGTH = 0.025      # 2.5cm stem
PETIOLE_RADIUS = 0.0012     # 1.2mm stem radius

N_T = 55                    # grid: along length (petiole + blade)
N_S = 18                    # grid: across each half (midrib to edge)
T_MIN = -0.42               # t domain start (petiole base)
T_MAX = 1.0                 # t domain end (blade tip)

VEIN_PAIRS = 7              # lateral vein pairs
DROOP = 0.10                # tip droop as fraction of leaf length
TWIST_DEG = 5               # gentle twist base-to-tip
CURL = 0.03                 # edge curl (subtle for birch)
ASYMMETRY = 0.04            # one side 4% wider

SCALE_DEMO = 10.0


# ── Utility ──────────────────────────────────────────────────────────────────

def smoothstep(x, edge0, edge1):
    """Hermite interpolation, 0 at edge0, 1 at edge1."""
    t = max(0.0, min(1.0, (x - edge0) / (edge1 - edge0)))
    return t * t * (3 - 2 * t)


def gauss(x, sigma):
    """Unnormalized Gaussian."""
    return math.exp(-(x * x) / (2 * sigma * sigma))


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 1: SPINE — the centerline arc
# ══════════════════════════════════════════════════════════════════════════════

def spine_x(t):
    """X position along the spine. Petiole at negative t, blade at positive."""
    if t <= 0:
        return t * PETIOLE_LENGTH / abs(T_MIN)  # maps T_MIN..0 to -PETIOLE_LENGTH..0
    else:
        return t * LEAF_LENGTH

def spine_z(t):
    """Main droop arc. Petiole slightly upward, blade droops."""
    if t <= 0:
        # Petiole: slight upward curve
        return 0.015 * LEAF_LENGTH * (-t / abs(T_MIN)) ** 1.5
    else:
        # Blade: gravity droop
        return -DROOP * LEAF_LENGTH * t ** 2.5


# ══════════════════════════════════════════════════════════════════════════════
# LAYER 2: CROSS-SECTION MORPH — tube to flat blade
# ══════════════════════════════════════════════════════════════════════════════

def blade_alpha(t):
    """0 = pure tube (petiole), 1 = pure flat blade."""
    return smoothstep(t, -0.12, 0.08)


def petiole_radius(t):
    """Petiole tube radius — tapers toward blade."""
    if t >= 0.08:
        return 0
    r = PETIOLE_RADIUS * (1.0 - smoothstep(t, -0.25, 0.05) * 0.6)
    return r


def leaf_half_width(t, side_sign=1.0):
    """Half-width of the blade at parameter t (0=base, 1=tip)."""
    if t <= 0:
        return 0
    # Triangular-ovate profile: widest at ~30% from base
    w = math.sin(t * math.pi) ** 0.55
    w *= (1.0 - t ** 2.8)            # taper to acuminate tip
    w *= min(1.0, (t / 0.15) ** 1.2)  # gradual cuneate base
    hw = w * (LEAF_WIDTH / 2)
    # Asymmetry
    if side_sign < 0:
        hw *= (1.0 + ASYMMETRY)
    return hw


# ══════════════════════════════════════════════════════════════════════════════
# DOUBLY SERRATE EDGE
# ══════════════════════════════════════════════════════════════════════════════

def vein_t_positions():
    """T-positions where lateral veins leave the midrib."""
    return [0.08 + 0.78 * i / (VEIN_PAIRS - 1) for i in range(VEIN_PAIRS)]


def serration_offset(t):
    """Doubly serrate edge: primary teeth at vein positions, secondary between."""
    if t < 0.05 or t > 0.96:
        return 0.0

    vein_ts = vein_t_positions()
    # Primary teeth: one per vein, tied to where each vein meets the edge
    # The vein reaches the edge at approximately t_branch + delta
    amp_base = 0.0018 * math.sin(t * math.pi) ** 0.4

    # Primary serrations — sharper, larger
    primary_freq = VEIN_PAIRS * 2  # teeth on both sides of each vein
    p_phase = t * primary_freq * math.pi
    primary = amp_base * (0.5 * math.sin(p_phase) + 0.5 * abs(math.sin(p_phase)))

    # Secondary serrations — smaller, higher frequency, ride on top of primary
    secondary_freq = primary_freq * 3.3
    s_phase = t * secondary_freq * math.pi
    secondary = amp_base * 0.30 * abs(math.sin(s_phase))

    return primary + secondary


# ══════════════════════════════════════════════════════════════════════════════
# VEIN PATHS — defined in (t, s) space
# ══════════════════════════════════════════════════════════════════════════════

class VeinSystem:
    """Lateral vein paths and their influence on the surface."""

    def __init__(self):
        self.vein_ts = vein_t_positions()
        # Each vein has a t-extent: how far along t it takes to reach the edge
        self.vein_deltas = []
        for i, vt in enumerate(self.vein_ts):
            # Lower veins extend further forward
            delta = 0.10 + 0.05 * (1 - i / VEIN_PAIRS)
            self.vein_deltas.append(delta)

    def vein_s_at_t(self, vein_idx, t):
        """Where is vein `vein_idx` in s-space at parameter t? None if outside vein."""
        vt = self.vein_ts[vein_idx]
        delta = self.vein_deltas[vein_idx]
        if t < vt or t > vt + delta:
            return None
        progress = (t - vt) / delta
        return progress ** 0.65 * 0.93  # concave curve, almost reaches edge

    def ridge_z(self, t, s):
        """Z displacement from vein ridges at (t, s). Gaussian cross-profile."""
        if t <= 0.02:
            return 0
        z = 0
        sigma = 0.055  # vein width in s-space
        for vi in range(VEIN_PAIRS):
            sv = self.vein_s_at_t(vi, t)
            if sv is None:
                continue
            dist = abs(s - sv)
            # Height tapers toward edge
            taper = max(0, 1.0 - sv * 0.7)
            height = 0.0055 * LEAF_WIDTH * taper
            z += height * gauss(dist, sigma)
        return z

    def billow_z(self, t, s):
        """Inter-vein sag — membrane droops between veins like tent fabric."""
        if t <= 0.05 or s < 0.05 or s > 0.88:
            return 0
        # Find nearest vein s-position
        min_dist = 1.0
        for vi in range(VEIN_PAIRS):
            sv = self.vein_s_at_t(vi, t)
            if sv is not None:
                min_dist = min(min_dist, abs(s - sv))
        # Also consider midrib (s=0) and leaf edge (s=1)
        min_dist = min(min_dist, s, 1.0 - s)
        # Sag increases with distance from nearest vein
        sag = -0.003 * LEAF_WIDTH * smoothstep(min_dist, 0.02, 0.08)
        return sag


# ══════════════════════════════════════════════════════════════════════════════
# Z-DISPLACEMENT STACK
# ══════════════════════════════════════════════════════════════════════════════

def midrib_ridge(t, s):
    """Gaussian ridge along the midrib (s=0). Continuous with petiole."""
    if t < -0.15:
        return 0
    # Midrib height: prominent at base, tapers to tip
    height = LEAF_WIDTH * 0.022 * (1 - max(0, t) * 0.65)
    # Smooth onset from petiole transition
    onset = smoothstep(t, -0.15, 0.05)
    sigma = 0.07  # midrib width in s-space
    return height * onset * gauss(s, sigma)


def edge_curl(t, s):
    """Subtle upward curl at the edges."""
    if t <= 0:
        return 0
    return CURL * LEAF_WIDTH * s ** 3 * math.sin(t * math.pi) ** 0.5


def organic_noise(t, s, side_sign):
    """Low-frequency undulation + twist for irregularity."""
    z = 0
    if t > 0:
        # Twist
        twist_rad = math.radians(TWIST_DEG) * t
        hw = leaf_half_width(t, side_sign)
        z += side_sign * s * hw * math.sin(twist_rad) * 0.15

        # Low-freq waviness
        z += 0.0006 * LEAF_WIDTH * math.sin(t * 7.3 + s * 4.1) * math.sin(t * 3.7 - s * 5.9)

        # Gravity on wider sections — slight extra sag where the leaf is widest
        z -= 0.002 * LEAF_WIDTH * math.sin(t * math.pi) ** 2 * s * 0.3
    return z


# ══════════════════════════════════════════════════════════════════════════════
# UNIFIED GRID
# ══════════════════════════════════════════════════════════════════════════════

class UnifiedLeafGrid:
    """
    One parametric grid for the entire leaf. Petiole tube morphs into flat blade.
    Veins, billow, serrations are all Z-displacements on this grid.
    """

    def __init__(self, scale=1.0):
        self.scale = scale
        self.veins = VeinSystem()

        # Vertex storage: indexed by (side_idx, i, j)
        self.vertices = []
        self.idx = {}  # (side, i, j) -> vertex index

        for side_idx, side_sign in enumerate([1.0, -1.0]):
            for i in range(N_T + 1):
                t = T_MIN + (T_MAX - T_MIN) * i / N_T

                for j in range(N_S + 1):
                    s = j / N_S  # 0=midrib, 1=edge

                    # --- Spine position ---
                    cx = spine_x(t)
                    cz = spine_z(t)

                    # --- Cross-section morph ---
                    alpha = blade_alpha(t)
                    pet_r = petiole_radius(t)
                    hw = leaf_half_width(t, side_sign)

                    # Tube cross-section: s maps around the tube
                    # s=0 is top (midrib), s=1 is side/bottom
                    theta = s * math.pi
                    tube_y = side_sign * pet_r * math.sin(theta)
                    tube_z = pet_r * math.cos(theta)

                    # Blade cross-section: s maps from midrib to edge
                    serr = serration_offset(t) if t > 0 else 0
                    blade_y = side_sign * s * (hw + serr)
                    blade_z = 0.0

                    # Blend
                    y = tube_y * (1 - alpha) + blade_y * alpha
                    local_z = tube_z * (1 - alpha) + blade_z * alpha

                    # --- Z displacement stack (blade region only) ---
                    dz = 0
                    dz += midrib_ridge(t, s)
                    if t > 0.02:
                        dz += self.veins.ridge_z(t, s)
                        dz += self.veins.billow_z(t, s)
                    dz += edge_curl(t, s)
                    dz += organic_noise(t, s, side_sign)

                    # Slight forward sweep at edges
                    x_sweep = 0
                    if t > 0:
                        x_sweep = s * 0.012 * LEAF_LENGTH

                    # Final vertex
                    vx = (cx + x_sweep) * scale
                    vy = y * scale
                    vz = (cz + local_z + dz) * scale

                    # Share midrib vertices between sides
                    if side_idx == 1 and j == 0:
                        self.idx[(side_idx, i, j)] = self.idx[(0, i, 0)]
                    else:
                        self.idx[(side_idx, i, j)] = len(self.vertices)
                        self.vertices.append((vx, vy, vz))

    def build_faces(self):
        faces = []
        v = self.vertices

        for side_idx in range(2):
            for i in range(N_T):
                for j in range(N_S):
                    i00 = self.idx[(side_idx, i, j)]
                    i10 = self.idx[(side_idx, i+1, j)]
                    i01 = self.idx[(side_idx, i, j+1)]
                    i11 = self.idx[(side_idx, i+1, j+1)]

                    # Skip degenerate quads (collapsed at tip or base)
                    if i00 == i01 and i10 == i11:
                        continue

                    # Top face
                    faces.append([v[i00], v[i10], v[i11]])
                    faces.append([v[i00], v[i11], v[i01]])
                    # Bottom face (flipped winding)
                    faces.append([v[i00], v[i11], v[i10]])
                    faces.append([v[i00], v[i01], v[i11]])

        # Petiole tube closure: stitch side 0's s=N_S to side 1's s=N_S
        for i in range(N_T):
            t = T_MIN + (T_MAX - T_MIN) * i / N_T
            t_next = T_MIN + (T_MAX - T_MIN) * (i + 1) / N_T
            if blade_alpha(t) > 0.97:
                break  # blade is open, no more tube closure

            a = self.idx[(0, i, N_S)]
            b = self.idx[(0, i+1, N_S)]
            c = self.idx[(1, i+1, N_S)]
            d = self.idx[(1, i, N_S)]

            faces.append([v[a], v[b], v[c]])
            faces.append([v[a], v[c], v[d]])
            faces.append([v[a], v[c], v[b]])
            faces.append([v[a], v[d], v[c]])

        # Cap the petiole base (i=0)
        base_verts_0 = [self.idx[(0, 0, j)] for j in range(N_S + 1)]
        base_verts_1 = [self.idx[(1, 0, j)] for j in range(N_S + 1)]
        # Combine into a ring: side 0 from j=0..N_S, then side 1 from j=N_S..0
        ring = base_verts_0 + list(reversed(base_verts_1[1:]))  # skip j=0 duplicate
        center = v[self.idx[(0, 0, 0)]]  # midrib vertex at base
        # Fan triangulation
        for k in range(len(ring) - 1):
            faces.append([center, v[ring[k]], v[ring[k+1]]])
            faces.append([center, v[ring[k+1]], v[ring[k]]])

        return faces


# ══════════════════════════════════════════════════════════════════════════════
# IFC OUTPUT
# ══════════════════════════════════════════════════════════════════════════════

def faces_to_ifc_brep(ifc, faces):
    ifc_faces = []
    for tri in faces:
        pts = [ifc.createIfcCartesianPoint([float(p[0]), float(p[1]), float(p[2])])
               for p in tri]
        loop = ifc.createIfcPolyLoop(pts)
        bound = ifc.createIfcFaceOuterBound(loop, True)
        ifc_faces.append(ifc.createIfcFace([bound]))
    shell = ifc.createIfcClosedShell(ifc_faces)
    return ifc.createIfcFacetedBrep(shell)


def create_ifc_color(ifc, r, g, b):
    color = ifc.createIfcColourRgb(None, r, g, b)
    rendering = ifc.createIfcSurfaceStyleRendering(
        color, 0.0, None, None, None, None, None, None, "FLAT"
    )
    return ifc.createIfcSurfaceStyle(None, "BOTH", [rendering])


def build_leaf_ifc(out_path, scale=1.0, label=""):
    ifc = ifcopenshell.api.run("project.create_file")
    project = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcProject", name="Leaf Detail Demo"
    )
    ifcopenshell.api.run("unit.assign_unit", ifc, length={"is_metric": True, "raw": "METERS"})
    ctx = ifcopenshell.api.run("context.add_context", ifc, context_type="Model")
    body = ifcopenshell.api.run(
        "context.add_context", ifc, context_type="Model",
        context_identifier="Body", target_view="MODEL_VIEW", parent=ctx,
    )
    site = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcSite", name="Demo")
    ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=project, products=[site])
    building = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcBuilding", name="Demo")
    ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=site, products=[building])
    storey = ifcopenshell.api.run("root.create_entity", ifc, ifc_class="IfcBuildingStorey", name="Ground")
    ifcopenshell.api.run("aggregate.assign_object", ifc, relating_object=building, products=[storey])

    leaf_color = create_ifc_color(ifc, 0.18, 0.50, 0.10)

    print("  Building unified leaf mesh...")
    leaf = UnifiedLeafGrid(scale)
    faces = leaf.build_faces()

    brep = faces_to_ifc_brep(ifc, faces)
    ifc.createIfcStyledItem(
        brep, [ifc.createIfcPresentationStyleAssignment([leaf_color])], None
    )

    element = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcBuildingElementProxy",
        name=f"Birch_Leaf{label}",
    )
    rep = ifc.createIfcShapeRepresentation(body, "Body", "Brep", [brep])
    element.Representation = ifc.createIfcProductDefinitionShape(None, None, [rep])
    element.ObjectPlacement = ifcopenshell.api.run(
        "geometry.edit_object_placement", ifc, product=element
    )
    ifcopenshell.api.run(
        "spatial.assign_container", ifc, relating_structure=storey, products=[element]
    )

    n_tri = len(faces)
    pset = ifcopenshell.api.run("pset.add_pset", ifc, product=element, name="Demo_GeometryStats")
    ifcopenshell.api.run("pset.edit_pset", ifc, pset=pset, properties={
        "Triangles": str(n_tri),
        "Vertices": str(len(leaf.vertices)),
        "VeinPairs": str(VEIN_PAIRS),
        "Purpose": "One leaf. Now imagine 200,000 of these per tree.",
    })

    ifc.write(out_path)
    return n_tri


def main():
    import os
    out_dir = os.path.dirname(os.path.abspath(__file__))

    real_path = os.path.join(out_dir, "demo_leaf_realsize.ifc")
    n1 = build_leaf_ifc(real_path, scale=1.0, label="_RealSize_6cm")
    size1 = os.path.getsize(real_path)

    demo_path = os.path.join(out_dir, "demo_leaf_10x.ifc")
    n2 = build_leaf_ifc(demo_path, scale=SCALE_DEMO, label="_10x_Demo")
    size2 = os.path.getsize(demo_path)

    print("=" * 60)
    print("Birch Leaf (Betula pendula) — Unified Mesh")
    print("=" * 60)
    print(f"\nOne continuous surface: petiole -> blade -> veins -> teeth")
    print(f"\nReal size (6cm):  {n1:,} tri -- {size1/1024:.0f} KB")
    print(f"10x scale (60cm): {n2:,} tri -- {size2/1024:.0f} KB")
    print(f"Vertices: {len(UnifiedLeafGrid(1.0).vertices):,}")
    print(f"\n--- The punchline ---")
    print(f"One leaf:             {n1:>10,} triangles   {size1/1024:>6.0f} KB")
    print(f"One birch tree:       {n1*200_000:>10,} triangles  (200k leaves)")
    print(f"190 deciduous trees:  {n1*200_000*190:>10,} triangles")
    print(f"")
    print(f"Our 7-sided cone:               42 triangles")
    print(f"All 1,544 trees:           ~150,000 triangles   42 MB")


if __name__ == "__main__":
    main()
