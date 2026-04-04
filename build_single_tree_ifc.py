"""
Procedural birch tree (Betula pendula) grown by simulating 30 years of
phototropic growth. Each year, every active bud pushes a new shoot toward
the brightest available light, constrained by gravity, self-shading, and
the tree's own structural limits.

The sun sweeps east-to-west daily, north-to-south seasonally (Norway, 60N).
A bud in shade produces a weak shoot or dies. A bud in light produces
a strong shoot that becomes next year's branch.

Output: IFC via IfcOpenShell. Four visual layers:
  1. Trunk (white bark)
  2. Main branches (pale bark)
  3. Delivery branches (grey-brown)
  4. Leaves (green, at active tips)
"""
import ifcopenshell
import ifcopenshell.api
import math
import numpy as np
from dataclasses import dataclass, field

# ── Growth parameters ────────────────────────────────────────────────────────
SEED = 42
YEARS = 100                   # max years — soil will run out before this
LATITUDE = 60.0               # Jevnaker, Norway — degrees north

# Soil: file size IS the nutrition in the ground
SOIL_MB = 300.0               # megabytes of nutrition available
BYTES_PER_TRIANGLE = 380      # approximate IFC bytes per triangle (measured)
SOIL_TRIANGLES = int(SOIL_MB * 1024 * 1024 / BYTES_PER_TRIANGLE)

# Shoot growth
SHOOT_LENGTH_BASE = 0.9       # max annual shoot length (meters) in full sun
SHOOT_LENGTH_MIN = 0.12       # minimum viable shoot
TRUNK_SHOOT_BOOST_YOUNG = 2.2  # strong apical dominance when young
TRUNK_SHOOT_BOOST_OLD = 1.1    # leader slows with age, crown fills out
GRAVITY_DROOP = 0.035          # less droop — lets branches reach outward more
MAX_DROOP_RATE = 0.12          # max droop per growth step

# Branching
BRANCH_PROB = 0.70            # probability of lateral bud activating per year
LEADER_BRANCH_PROB = 0.90     # leader almost always branches
MAX_ACTIVE_TIPS = 3000        # performance cap
BUD_DEATH_SHADE = 0.03        # light below this = bud dies (very tolerant)

# Da Vinci's rule — radius accumulation
RING_THICKNESS = 0.005        # meters per growth ring (thicker rings)

# Sun sampling for light calculation
SUN_SAMPLES = 10              # directions to sample

# Crown self-shading — simplified ray test
SHADE_RADIUS = 0.35           # tighter shade radius — less aggressive culling

# Geometry
TUBE_SIDES_TRUNK = 10
TUBE_SIDES_BRANCH = 6
TUBE_SIDES_TWIG = 4
LEAF_SIZE = 0.05
LEAVES_PER_TIP = 5


# ── Sun model (Norway, 60N) ─────────────────────────────────────────────────

def sun_directions(month=6):
    """
    Sample sun directions for a given month at 60N latitude.
    Returns list of unit vectors pointing toward the sun at different times.
    Summer: high arc, long days. Winter: low, short days.
    """
    # Solar declination (approximate)
    decl = 23.45 * math.sin(math.radians(360/365 * (284 + month * 30)))
    lat = math.radians(LATITUDE)
    decl_rad = math.radians(decl)

    directions = []
    # Sample the sun's arc through the day
    for hour_angle in np.linspace(-90, 90, SUN_SAMPLES):
        ha = math.radians(hour_angle)
        # Solar altitude
        sin_alt = (math.sin(lat) * math.sin(decl_rad) +
                   math.cos(lat) * math.cos(decl_rad) * math.cos(ha))
        if sin_alt <= 0.05:
            continue  # sun below horizon
        alt = math.asin(sin_alt)

        # Solar azimuth (simplified)
        cos_az = (math.sin(decl_rad) - math.sin(lat) * sin_alt) / (math.cos(lat) * math.cos(alt) + 1e-10)
        cos_az = max(-1, min(1, cos_az))
        az = math.acos(cos_az)
        if hour_angle > 0:
            az = 2 * math.pi - az

        # Direction vector (toward the sun)
        dx = math.sin(az) * math.cos(alt)
        dy = math.cos(az) * math.cos(alt)
        dz = math.sin(alt)
        directions.append(np.array([dx, dy, dz]))

    if not directions:
        directions.append(np.array([0, 0.5, 0.87]))  # fallback

    return directions


# ── Tree data structure ──────────────────────────────────────────────────────

@dataclass
class TreeNode:
    pos: np.ndarray              # 3D position
    parent_idx: int              # index of parent node (-1 for root)
    birth_year: int              # year this node was created
    is_tip: bool                 # active growing tip
    is_leader: bool              # apical leader (trunk continuation)
    depth: int                   # generations from trunk (0=trunk)
    direction: np.ndarray        # growth direction when this node was created
    n_downstream: int = 1        # number of tips below this node (for Da Vinci)


class GrowingTree:
    """Simulates a birch tree growing year by year."""

    def __init__(self, rng):
        self.rng = rng
        self.nodes = []
        # Seed: root at ground level
        root = TreeNode(
            pos=np.array([0.0, 0.0, 0.0]),
            parent_idx=-1, birth_year=0,
            is_tip=True, is_leader=True, depth=0,
            direction=np.array([0.0, 0.0, 1.0])
        )
        self.nodes.append(root)

    def _rebuild_shade_grid(self):
        """Build a 3D voxel grid for O(1) shade lookups."""
        self._voxel_size = 0.8  # meters per voxel
        self._shade_voxels = set()
        for node in self.nodes:
            if node.birth_year <= 0:
                continue
            vx = int(node.pos[0] / self._voxel_size)
            vy = int(node.pos[1] / self._voxel_size)
            vz = int(node.pos[2] / self._voxel_size)
            self._shade_voxels.add((vx, vy, vz))

    def light_at(self, pos, sun_dirs):
        """
        How much light reaches this position?
        Voxel-based: march along each sun ray, check if any voxel is occupied.
        O(ray_length / voxel_size) per direction instead of O(n_nodes).
        Returns 0..1.
        """
        if not self._shade_voxels:
            return 1.0

        vs = self._voxel_size
        hits = 0
        max_steps = 30  # max ray march steps

        for sun_dir in sun_dirs:
            blocked = False
            # March from pos toward the sun
            for step in range(1, max_steps):
                sample = pos + sun_dir * (step * vs * 0.7)
                if sample[2] < 0 or sample[2] > 30:
                    break
                vx = int(sample[0] / vs)
                vy = int(sample[1] / vs)
                vz = int(sample[2] / vs)
                if (vx, vy, vz) in self._shade_voxels:
                    blocked = True
                    break
            if not blocked:
                hits += 1

        return hits / len(sun_dirs) if sun_dirs else 0.5

    def best_growth_direction(self, pos, parent_dir, sun_dirs, light):
        """
        Grow toward the brightest open sky, biased by parent direction.
        Average unblocked sun directions to find the light pull.
        """
        vs = self._voxel_size
        sun_pull = np.zeros(3)
        for sd in sun_dirs:
            # Quick 3-step ray check
            blocked = False
            for step in range(1, 4):
                sample = pos + sd * (step * vs)
                vx = int(sample[0] / vs)
                vy = int(sample[1] / vs)
                vz = int(sample[2] / vs)
                if (vx, vy, vz) in self._shade_voxels:
                    blocked = True
                    break
            if not blocked:
                sun_pull += sd

        n = np.linalg.norm(sun_pull)
        if n > 0.01:
            sun_pull /= n
        else:
            sun_pull = np.array([0, 0, 0.5])

        grow_dir = normalize(
            parent_dir * 0.5 +
            sun_pull * 0.4 +
            self.rng.randn(3) * 0.1
        )
        return grow_dir

    def apply_gravity(self, direction, pos):
        """Droop based on horizontal distance from trunk axis."""
        horiz_dist = math.sqrt(pos[0]**2 + pos[1]**2)
        droop = min(GRAVITY_DROOP * horiz_dist, MAX_DROOP_RATE)
        if droop < 0.001 or direction[2] < -0.9:
            return direction

        horiz = np.array([direction[0], direction[1], 0.0])
        hl = np.linalg.norm(horiz)
        if hl < 1e-6:
            return direction

        # Rotate down
        axis = normalize(np.cross(np.array([0, 0, 1.0]), horiz))
        c, s = math.cos(-droop), math.sin(-droop)
        d = direction
        result = d * c + np.cross(axis, d) * s + axis * np.dot(axis, d) * (1 - c)
        return normalize(result)

    def estimate_triangles(self):
        """Estimate how many triangles the current tree would produce."""
        # Each node = one tube segment (~2*n_sides triangles) + tips get leaves
        n_tips = sum(1 for n in self.nodes if n.is_tip)
        # Average tube sides ~5, so ~10 tri per node, + 4 tri per leaf * LEAVES_PER_TIP
        return len(self.nodes) * 10 + n_tips * LEAVES_PER_TIP * 4

    def soil_remaining(self):
        """Fraction of soil nutrition remaining (0..1)."""
        used = self.estimate_triangles()
        return max(0, 1.0 - used / SOIL_TRIANGLES)

    def grow_one_year(self, year):
        """Simulate one year of growth."""
        # Check soil
        soil = self.soil_remaining()
        if soil <= 0:
            # No nutrition left — tree can't grow
            return False

        # Rebuild shade grid for this year
        self._rebuild_shade_grid()

        # Sun directions: combine the actual solar arc with diffuse sky light
        # An open-grown tree gets light from all directions, not just the sun arc
        sun_dirs = sun_directions(month=6)
        # Add diffuse sky light from all compass directions (overhead hemisphere)
        for az_deg in range(0, 360, 45):
            for alt_deg in [30, 60]:
                az = math.radians(az_deg)
                alt = math.radians(alt_deg)
                sun_dirs.append(np.array([
                    math.cos(az) * math.cos(alt),
                    math.sin(az) * math.cos(alt),
                    math.sin(alt)
                ]))

        tips = [(i, n) for i, n in enumerate(self.nodes) if n.is_tip]

        # Performance cap: if too many tips, kill the weakest
        if len(tips) > MAX_ACTIVE_TIPS:
            # Sort by light, keep the best
            tip_light = []
            for idx, node in tips:
                light = self.light_at(node.pos, sun_dirs)
                tip_light.append((idx, node, light))
            tip_light.sort(key=lambda x: x[2], reverse=True)
            for idx, node, light in tip_light[MAX_ACTIVE_TIPS:]:
                node.is_tip = False
            tips = [(idx, node) for idx, node, light in tip_light[:MAX_ACTIVE_TIPS]]

        new_nodes = []

        for tip_idx, tip_node in tips:
            tip_node.is_tip = False  # this tip will produce a new node

            # How much light does this tip get?
            light = self.light_at(tip_node.pos, sun_dirs)

            if light < BUD_DEATH_SHADE and not tip_node.is_leader:
                continue  # bud dies — too much shade (leader never gives up)

            # Growth direction
            if tip_node.is_leader:
                # Leader: strongly vertical with slight wobble
                # Trunk doesn't chase the sun — it just grows UP
                wobble = self.rng.randn(3) * 0.04
                wobble[2] = 0  # no vertical wobble
                grow_dir = normalize(np.array([0, 0, 1.0]) + wobble)
                # No gravity droop on the leader — trunk stays upright
            else:
                # Branches: chase the light
                grow_dir = self.best_growth_direction(
                    tip_node.pos, tip_node.direction, sun_dirs, light)
                # Apply gravity droop
                grow_dir = self.apply_gravity(grow_dir, tip_node.pos)

            # Shoot length depends on light AND available soil nutrition
            shoot_len = SHOOT_LENGTH_BASE * light * max(0.2, soil)
            if tip_node.is_leader:
                # Apical dominance weakens with age — crown fills out
                boost = TRUNK_SHOOT_BOOST_YOUNG + (TRUNK_SHOOT_BOOST_OLD - TRUNK_SHOOT_BOOST_YOUNG) * min(1, year / 25)
                shoot_len *= boost
            shoot_len = max(shoot_len, SHOOT_LENGTH_MIN * soil)
            # Decrease shoot length with age (growth slows, but not too much)
            shoot_len *= max(0.55, 1.0 - year / (YEARS * 2.0))

            new_pos = tip_node.pos + grow_dir * shoot_len

            # Create the new apical node (continuation of this branch)
            new_node = TreeNode(
                pos=new_pos,
                parent_idx=tip_idx,
                birth_year=year,
                is_tip=True,
                is_leader=tip_node.is_leader,
                depth=tip_node.depth,
                direction=grow_dir,
            )
            new_nodes.append(new_node)

            # Lateral branching? Probability scales with light AND nutrition
            if self.rng.random() < BRANCH_PROB * light * max(0.1, soil):
                # Direction: outward from trunk + random + slightly up
                trunk_xy = np.array([tip_node.pos[0], tip_node.pos[1], 0.0])
                if np.linalg.norm(trunk_xy) > 0.05:
                    outward = normalize(trunk_xy)
                else:
                    outward = normalize(np.array([self.rng.randn(), self.rng.randn(), 0.0]))
                lat_dir = normalize(outward * 0.5 + self.rng.randn(3) * 0.3 + np.array([0, 0, 0.2]))
                lat_len = shoot_len * 0.85

                lat_node = TreeNode(
                    pos=tip_node.pos + lat_dir * lat_len,
                    parent_idx=tip_idx,
                    birth_year=year,
                    is_tip=True,
                    is_leader=False,
                    depth=tip_node.depth + 1,
                    direction=lat_dir,
                )
                new_nodes.append(lat_node)

            # Leader also spawns lateral branches
            if tip_node.is_leader and year > 2 and self.rng.random() < LEADER_BRANCH_PROB:
                perp = self.rng.randn(3)
                perp -= np.dot(perp, grow_dir) * grow_dir
                n = np.linalg.norm(perp)
                if n > 0.01:
                    lat_dir = normalize(perp)
                    lat_dir = normalize(lat_dir + np.array([0, 0, 0.15]))
                    lat_len = shoot_len * 0.5

                    lat_node = TreeNode(
                        pos=tip_node.pos + lat_dir * lat_len,
                        parent_idx=tip_idx,
                        birth_year=year,
                        is_tip=True,
                        is_leader=False,
                        depth=1,
                        direction=lat_dir,
                    )
                    new_nodes.append(lat_node)

        # Dormant bud activation: the trunk/leader can sprout new branches
        # from any point along its length in the crown zone, not just the tip.
        # This fills the crown from bottom to top over the years.
        if year > 3:
            trunk_nodes = [(i, n) for i, n in enumerate(self.nodes)
                           if n.is_leader and not n.is_tip and n.pos[2] > 2.0]
            for idx, node in trunk_nodes:
                # Dormant buds along the trunk activate based on age and light.
                # Younger nodes (higher up) sprout more readily, but older nodes
                # also get a chance — this fills the lower crown over time.
                age = year - node.birth_year
                # Fresh nodes sprout eagerly, old nodes occasionally
                age_factor = max(0.3, 1.0 - age / 40.0)
                sprout_prob = 0.18 * age_factor * soil
                if self.rng.random() < sprout_prob:
                    # Check light at this position
                    node_light = self.light_at(node.pos, sun_dirs)
                    if node_light < 0.2:
                        continue
                    # Push outward from trunk axis — open-grown tree
                    outward = normalize(np.array([
                        node.pos[0] + self.rng.randn() * 0.3,
                        node.pos[1] + self.rng.randn() * 0.3,
                        0]))
                    if np.linalg.norm(outward) < 0.01:
                        outward = normalize(np.array([self.rng.randn(), self.rng.randn(), 0]))
                    lat_dir = normalize(outward * 0.6 + np.array([0, 0, 0.3]) + self.rng.randn(3) * 0.1)
                    lat_len = SHOOT_LENGTH_BASE * 0.7 * node_light * soil

                    lat_node = TreeNode(
                        pos=node.pos + lat_dir * lat_len,
                        parent_idx=idx,
                        birth_year=year,
                        is_tip=True,
                        is_leader=False,
                        depth=1,
                        direction=lat_dir,
                    )
                    new_nodes.append(lat_node)

        # Add all new nodes
        for node in new_nodes:
            self.nodes.append(node)

        # Update downstream counts (for Da Vinci radius)
        self._update_downstream()
        return True  # growth occurred

    def _update_downstream(self):
        """Count how many tips are downstream of each node."""
        for node in self.nodes:
            node.n_downstream = 0

        # Walk up from each tip to root, incrementing
        for i, node in enumerate(self.nodes):
            if not node.is_tip:
                continue
            idx = i
            while idx >= 0:
                self.nodes[idx].n_downstream += 1
                idx = self.nodes[idx].parent_idx

    def compute_radius(self, node):
        """
        Radius from downstream tip count + age accumulation.
        Pure Da Vinci (sqrt) makes branches too thin.
        Power of 0.4 keeps proportions realistic.
        Leader nodes get a minimum radius — the trunk is always visible.
        """
        r = RING_THICKNESS * max(node.n_downstream, 1) ** 0.4
        # Leader nodes: minimum trunk radius that grows with age
        if node.is_leader or node.parent_idx < 0:
            # Trunk sections accumulate girth over years
            # Use n_downstream as proxy for maturity
            min_r = 0.015 + 0.004 * max(node.n_downstream, 1) ** 0.3
            r = max(r, min_r)
        return r


# ── Geometry ─────────────────────────────────────────────────────────────────

def normalize(v):
    n = np.linalg.norm(v)
    return v / n if n > 1e-10 else np.array([0, 0, 1.0])


def oriented_ring(center, direction, radius, n_sides):
    d = normalize(direction)
    if abs(d[2]) < 0.9:
        up = np.array([0.0, 0.0, 1.0])
    else:
        up = np.array([1.0, 0.0, 0.0])
    u = normalize(np.cross(d, up))
    v = np.cross(d, u)
    pts = []
    for i in range(n_sides):
        theta = 2 * math.pi * i / n_sides
        p = center + radius * (math.cos(theta) * u + math.sin(theta) * v)
        pts.append(tuple(p))
    return pts


def frustum_faces(ring_bot, ring_top):
    n = len(ring_bot)
    faces = []
    for i in range(n):
        j = (i + 1) % n
        faces.append([ring_bot[i], ring_bot[j], ring_top[j]])
        faces.append([ring_bot[i], ring_top[j], ring_top[i]])
    return faces


def cap_faces(ring, center, flip=False):
    n = len(ring)
    faces = []
    for i in range(n):
        j = (i + 1) % n
        if flip:
            faces.append([center, ring[j], ring[i]])
        else:
            faces.append([center, ring[i], ring[j]])
    return faces


def tree_to_faces(tree):
    """Convert the grown tree into triangle faces, separated by layer."""
    trunk_faces = []
    main_faces = []
    delivery_faces = []
    leaf_positions = []

    for i, node in enumerate(tree.nodes):
        if node.parent_idx < 0:
            continue

        parent = tree.nodes[node.parent_idx]
        r_start = tree.compute_radius(parent)
        r_end = tree.compute_radius(node)

        # Knuckle: barely perceptible bump at growth nodes
        r_start *= 1.012

        direction = node.pos - parent.pos
        length = np.linalg.norm(direction)
        if length < 0.01:
            continue

        # Choose layer and tube sides based on radius
        if r_start > 0.04:
            n_sides = TUBE_SIDES_TRUNK
            target = trunk_faces
        elif r_start > 0.012:
            n_sides = TUBE_SIDES_BRANCH
            target = main_faces
        else:
            n_sides = TUBE_SIDES_TWIG
            target = delivery_faces

        d = direction / length
        ring0 = oriented_ring(parent.pos, d, r_start, n_sides)
        ring1 = oriented_ring(node.pos, d, r_end, n_sides)
        target.extend(frustum_faces(ring0, ring1))

        # Collect leaf positions at tips
        if node.is_tip:
            leaf_positions.append((node.pos.copy(), node.direction.copy()))

    # Ground cap
    if tree.nodes:
        root = tree.nodes[0]
        r = tree.compute_radius(root)
        ring = oriented_ring(root.pos, np.array([0, 0, 1.0]), r, TUBE_SIDES_TRUNK)
        trunk_faces.extend(cap_faces(ring, tuple(root.pos), flip=True))

    return trunk_faces, main_faces, delivery_faces, leaf_positions


def make_leaf(position, branch_dir, rng):
    """Simple double-sided diamond leaf."""
    up = np.array([0.0, 0.0, 1.0])
    normal = normalize(branch_dir * 0.3 + up * 0.7 + rng.randn(3) * 0.15)
    if abs(normal[2]) < 0.9:
        right = normalize(np.cross(normal, up))
    else:
        right = normalize(np.cross(normal, np.array([1, 0, 0])))
    fwd = normalize(np.cross(right, normal))

    half_w = LEAF_SIZE * 0.35
    half_h = LEAF_SIZE * 0.5
    c = np.array(position)
    tip = tuple(c + fwd * half_h)
    base = tuple(c - fwd * half_h * 0.5)
    left = tuple(c - right * half_w)
    right_pt = tuple(c + right * half_w)

    return [
        [base, right_pt, tip], [base, tip, left],
        [base, tip, right_pt], [base, left, tip],
    ]


def place_leaves(leaf_positions, rng):
    faces = []
    for pos, direction in leaf_positions:
        for _ in range(LEAVES_PER_TIP):
            offset = rng.randn(3) * 0.04
            offset[2] = abs(offset[2]) * 0.5
            faces.extend(make_leaf(pos + offset, direction, rng))
    return faces


# ── IFC output ───────────────────────────────────────────────────────────────

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


def build_tree_ifc(out_path):
    rng = np.random.RandomState(SEED)

    # Grow the tree — soil nutrition is the constraint
    tree = GrowingTree(rng)
    final_year = 0
    print(f"  Soil: {SOIL_MB:.0f} MB = ~{SOIL_TRIANGLES:,} triangles of nutrition\n")
    for year in range(1, YEARS + 1):
        grew = tree.grow_one_year(year)
        if not grew:
            print(f"  Year {year:>2}: SOIL EXHAUSTED -- tree has reached mature size")
            break
        final_year = year
        n_tips = sum(1 for n in tree.nodes if n.is_tip)
        soil_pct = tree.soil_remaining() * 100
        est_tri = tree.estimate_triangles()
        if year % 5 == 0 or year == YEARS or soil_pct < 5:
            print(f"  Year {year:>2}: {len(tree.nodes):>6} nodes, {n_tips:>4} tips, "
                  f"~{est_tri:>8,} tri, soil: {soil_pct:>5.1f}%")

    # Convert to geometry
    print("\n  Converting to geometry...")
    trunk_faces, main_faces, delivery_faces, leaf_positions = tree_to_faces(tree)

    print(f"  Placing leaves at {len(leaf_positions)} tips...")
    leaf_faces = place_leaves(leaf_positions, rng)

    n_trunk = len(trunk_faces)
    n_main = len(main_faces)
    n_delivery = len(delivery_faces)
    n_leaf = len(leaf_faces)
    n_total = n_trunk + n_main + n_delivery + n_leaf
    n_leaves = len(leaf_positions) * LEAVES_PER_TIP

    # Build IFC
    print("  Writing IFC...")
    ifc = ifcopenshell.api.run("project.create_file")
    project = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcProject", name="Single Tree Demo"
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

    trunk_bark = create_ifc_color(ifc, 0.92, 0.90, 0.85)
    main_bark = create_ifc_color(ifc, 0.80, 0.75, 0.68)
    delivery_color = create_ifc_color(ifc, 0.55, 0.50, 0.40)
    leaf_color = create_ifc_color(ifc, 0.28, 0.62, 0.16)

    items = []
    for faces, color in [(trunk_faces, trunk_bark), (main_faces, main_bark),
                          (delivery_faces, delivery_color), (leaf_faces, leaf_color)]:
        if faces:
            brep = faces_to_ifc_brep(ifc, faces)
            ifc.createIfcStyledItem(
                brep, [ifc.createIfcPresentationStyleAssignment([color])], None)
            items.append(brep)

    element = ifcopenshell.api.run(
        "root.create_entity", ifc, ifc_class="IfcBuildingElementProxy",
        name=f"Betula_Pendula_{YEARS}yr",
    )
    rep = ifc.createIfcShapeRepresentation(body, "Body", "Brep", items)
    element.Representation = ifc.createIfcProductDefinitionShape(None, None, [rep])
    element.ObjectPlacement = ifcopenshell.api.run(
        "geometry.edit_object_placement", ifc, product=element
    )
    ifcopenshell.api.run(
        "spatial.assign_container", ifc, relating_structure=storey, products=[element]
    )

    pset = ifcopenshell.api.run("pset.add_pset", ifc, product=element, name="Demo_TreeStats")
    ifcopenshell.api.run("pset.edit_pset", ifc, pset=pset, properties={
        "Years": str(YEARS),
        "Nodes": str(len(tree.nodes)),
        "Triangles_Trunk": str(n_trunk),
        "Triangles_MainBranches": str(n_main),
        "Triangles_DeliveryBranches": str(n_delivery),
        "Triangles_Leaves": str(n_leaf),
        "Triangles_Total": str(n_total),
        "LeafCount": str(n_leaves),
        "Species": "Betula pendula (silver birch)",
    })

    ifc.write(out_path)
    return n_trunk, n_main, n_delivery, n_leaf, n_leaves, len(tree.nodes)


def main():
    import os
    import subprocess
    from datetime import datetime
    out_dir = os.path.dirname(os.path.abspath(__file__))
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    try:
        git_hash = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=out_dir, stderr=subprocess.DEVNULL
        ).decode().strip()
    except Exception:
        git_hash = "nogit"
    out_path = os.path.join(out_dir, f"demo_birch_tree_{stamp}_{git_hash}.ifc")

    n_trunk, n_main, n_delivery, n_leaf, n_leaves, n_nodes = build_tree_ifc(out_path)
    n_total = n_trunk + n_main + n_delivery + n_leaf
    size = os.path.getsize(out_path)

    print("=" * 64)
    print(f"Birch Tree (Betula pendula) -- Growth Simulation")
    print("=" * 64)
    print(f"\n  Soil: {SOIL_MB:.0f} MB nutrition")
    print(f"  Growth: {n_nodes:,} nodes")
    print(f"\n  1. Trunk:              {n_trunk:>8,} triangles")
    print(f"  2. Main branches:      {n_main:>8,} triangles")
    print(f"  3. Delivery branches:  {n_delivery:>8,} triangles")
    print(f"  4. Leaves:             {n_leaf:>8,} triangles  ({n_leaves:,} leaves)")
    print(f"  {'':->52}")
    print(f"  TOTAL:                 {n_total:>8,} triangles   {size/1024:.0f} KB")
    print(f"\n  File: {out_path}")
    print(f"\n--- The comparison ---")
    print(f"  One detailed leaf:          8,068 triangles")
    print(f"  One {YEARS}-year tree:      {n_total:>9,} triangles")
    print(f"  1,544 cone-trees:        ~150,000 triangles")
    print(f"  One REAL tree (200k x 8k): 1,600,000,000 triangles")


if __name__ == "__main__":
    main()
