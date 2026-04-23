"""Build KNM master DATA DXF v2 from FKB baseline + 8 consultant files.

Applies the mapping from 02_Arbeid/00_Docs/layer_mapping_draft.md:
- Re-layers every entity to an English target vocabulary (FKB/SOSI + NS 3451 fallback)
- Drops explicitly excluded layers
- Strips (Nye konstruksjoner) suffix and trailing underscores
- Collapses tree-metadata layers to single targets
- Flattens Z to 0 throughout
- 3DE_Bru is skipped (3DSOLID + EXTRUDEDSURFACE can't be flattened cleanly)

Output: temp DXF + DWG for review before promoting into ACC.
"""
import ezdxf, os, sys, unicodedata, shutil, subprocess
from ezdxf.addons import importer

sys.stdout.reconfigure(encoding='utf-8')

MASTER_DIR = r'C:\Users\edkjo\DC\ACCDocs\Skiplum AS\Skiplum Backup\Project Files\10016 - Kistefos\03_Ut\08_Tegninger\DWG_NTM\Global\KNM_BIMK_MASTER_DATA_NTM10_GLOBAL_METERS'
FDE_DIR    = r'C:\Users\edkjo\AppData\Local\Temp\3de_out'
DXF_OUT_DIR = r'C:\Users\edkjo\DC\ACCDocs\Skiplum AS\Skiplum Backup\Project Files\10016 - Kistefos\03_Ut\08_Tegninger\DXF_NTM\Global'
DWG_OUT_DIR = r'C:\Users\edkjo\DC\ACCDocs\Skiplum AS\Skiplum Backup\Project Files\10016 - Kistefos\03_Ut\08_Tegninger\DWG_NTM\Global'
OUT_BASE   = 'KNM_BIMK_MASTER_DATA_NTM10_GLOBAL_METERS_v2'
ODA        = r'C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe'

def nfc(s): return unicodedata.normalize('NFC', s)

DROP = object()
COORDINATION_LAYERS = {
    'COORDINATION_MARKER_OLD',
    'COORDINATION_MARKER_NEW',
    'COORDINATION_MARKER_ROTATION',
}

# Entity types dropped wholesale. Master rule (CLAUDE.md):
# "lines/points/text only — no hatches, no block inserts".
# INSERTs are handled by explosion (world-positioned constituents), so
# they don't appear here. Any residual INSERT that refuses to explode is
# dropped via the fallback below.
DROP_ENTITY_TYPES = {'HATCH'}
RESIDUAL_INSERT_DROP = True

# Geographic filter. Anything outside the KNM site corridor is dropped —
# catches strays left in block definitions (e.g. stray LINEs that live at
# block-local coords far from the insert point and escape to outer space
# after explode).
KNM_E_MIN, KNM_E_MAX = 91000, 94000
KNM_N_MIN, KNM_N_MAX = 1246000, 1249000

def entity_points_xy(e):
    t = e.dxftype()
    pts = []
    try:
        if t == 'LINE':
            pts.append((e.dxf.start.x, e.dxf.start.y))
            pts.append((e.dxf.end.x, e.dxf.end.y))
        elif t == 'LWPOLYLINE':
            for pt in e.get_points():
                pts.append((pt[0], pt[1]))
        elif t == 'POLYLINE':
            for v in e.vertices:
                if v.dxf.flags & 128:
                    continue  # skip polyface face records
                pts.append((v.dxf.location.x, v.dxf.location.y))
        elif t in ('TEXT', 'MTEXT'):
            p = e.dxf.insert
            pts.append((p.x, p.y))
        elif t in ('CIRCLE', 'ARC', 'ELLIPSE'):
            c = e.dxf.center
            pts.append((c.x, c.y))
        elif t == '3DFACE':
            for i in range(4):
                v = getattr(e.dxf, f'vtx{i}')
                pts.append((v.x, v.y))
        elif t == 'POINT':
            p = e.dxf.location
            pts.append((p.x, p.y))
    except Exception:
        pass
    return pts

def entity_in_site(e):
    pts = entity_points_xy(e)
    if not pts:
        return True  # can't determine, keep
    cx = sum(x for x, _ in pts) / len(pts)
    cy = sum(y for _, y in pts) / len(pts)
    return KNM_E_MIN <= cx <= KNM_E_MAX and KNM_N_MIN <= cy <= KNM_N_MAX

def strip_layer(lay):
    lay = lay.strip()
    if lay.endswith('(Nye konstruksjoner)'):
        lay = lay[:-len('(Nye konstruksjoner)')].rstrip()
    # NOTE: do not rstrip underscores — some legit layer names end with `_`
    # (e.g. '761- Veier_ Kjøreveier_ sykkel- og gangveier mv_'). Trailing-_
    # layers from 3DE are handled explicitly in BYGNING_MAP.
    return lay

# ─── Mappings ────────────────────────────────────────────────────────────
BYGNING_MAP = {
    'BygningsavgrensningTakkant':      'Building_RoofEdge',
    'BygningslinjerMønelinje':         'Building_RoofRidge',
    'BygningslinjerBygningslinje':     'Building_Line',
    'BygningGrunnflateTak':            'Building_RoofFootprint',
    'BygningslinjerTaksprangVedTopp':  'Building_EaveTop',
    'BygningslinjerTaksprangVedBunn':  'Building_EaveBottom',
    'BygningsvedhengVeranda':          'Building_Veranda',
    'BygningBygningsdelelinje':        'Building_PartLine',
    'BygninglinjerHjelpelinje3D':      'Building_Helper3D',
    'BygningsvedhengTrappBygg':        'Building_ExteriorStair',
    'BygningsvedhengLåvebru':          'Building_BarnRamp',
    'BygningsavgrensningGrunnmur':     'Building_FoundationWall',
    'Kantutsnitt':                     'Building_EdgeCutout',
    'UklassifisertObjekt':             'Unclassified',
    'Takovegbygg_':                    'Building_RoofedOverpass',
    'Takoverbygg Takkant_':            'Building_RoofedOverpass_Edge',
    'Solid - 2':                       DROP,
    'Linje - 3':                       DROP,
}

KART_MAP = {
    # Colleague-curated drops (see layer panel screenshot 2026-04-14).
    # These are layers represented primarily as surface/mesh geometry where
    # the flattened 2D projection is ugly/unhelpful. Their edge-layer siblings
    # (Vegdekkeoverflatekant, Innsjøkant, KaiBryggeKant, MurGrunnrisskant)
    # are kept.
    'Høydekurve':                      DROP,  # superseded by Kotelinjer
    'Forsenkningskurve':               DROP,  # depression contours
    'Lekeplass':                       DROP,  # playground
    'BruGrunnflate':                   DROP,  # bridge base surface
    'GangSykkelveg':                   DROP,  # pedestrian/cycle surface
    'Innsjø':                          DROP,  # lake interior
    'Vegdekkeoverflate':               DROP,  # road pavement surface
    # Kept
    'Vegdekkeoverflatekant':           'Road_PavementEdge',
    'Gjerde':                          'Fence',
    'AnnetVegarealAvgrensning':        'Road_OtherAreaBoundary',
    'ElvBekkKant':                     'River_StreamEdge',
    'Støttemur':                       'RetainingWall',
    'VegKantFiktiv':                   'Road_EdgeVirtual',
    'KanalGrøftKant':                  'Ditch_Edge',
    'Vegskulderkant':                  'Road_ShoulderEdge',
    'Kjørebanekant':                   'Road_LaneEdge',
    'Bruenhet':                        'Bridge_Unit',
    'Masteomriss':                     'Mast_Outline',
    'FlomløpKant':                     'FloodChannel_Edge',
    'Ledningsnett Framføringsveger':   'Utility_RouteCorridor',
    'FiktivDelelinje':                 'VirtualDivisionLine',
    'Spormidt':                        'Track_Centerline',
    'Overvannsledning':                'StormwaterPipe',
    'Vannledning':                     'WaterPipe',
    'MurGrunnrissflate':               'Wall_Footprint',
    'KaiBryggeKant':                   'QuayPier_Edge',
    'KaiBrygge':                       'QuayPier',
    'Innsjøkant':                      'Lake_Edge',
    'MurGrunnrisskant':                'Wall_FootprintEdge',
    'BygningsvedhengTrappBygg':        'Building_ExteriorStair',
    'Kantutsnitt':                     'Map_EdgeCutout',
}

STIER_MAP = {
    '00- besøkssenter':                            'VisitorCentre_Project',
    '00- nye stier':                               'Path_New',
    '00 Stier':                                    'Path_New',
    '00- ny jord':                                 'Soil_New',
    '00- Hotlink Bygg':                            DROP,
    '00-forflytning jord':                         'Soil_Reshaping',
    '837- Planteplan trær pisk FURU':              'PlantingPlan_Tree_Pine',
    '837- Planteplan trær pisk BJØRK':             'PlantingPlan_Tree_Birch',
    '837- Planteplan trær so 12-14':               'PlantingPlan_Tree_Standard12-14',
    '837- Planteplan trær pisk OR':                'PlantingPlan_Tree_Alder',
    '837- Planteplan trær pisk HEGG':              'PlantingPlan_Tree_BirdCherry',
    '837- Planteplan trær pisk ROGN':              'PlantingPlan_Tree_Rowan',
    '837- Planteplan trær transplantert':          'PlantingPlan_Tree_Transplanted',
    'AnnetGjerde':                                 'Fence_Other',
    'Vegdekkekant':                                'Road_PavementEdge',
    'VegkantFiktiv':                               'Road_EdgeVirtual',
    'Vegrekkverk':                                 'Road_GuardRail',
    'f-veg_21000_FLATE_KJØREFELT_ASFALT':          'Road_Asphalt_21000',
    'VegkantAnnetVegareal':                        'Road_OtherAreaBoundary',
    'Veg':                                         'Road_PavementSurface',
    'Veranda':                                     'Building_Veranda',
    '790- Utstyr_ møbler_ lekeapparater':          'OutdoorFurniture_PlayEquipment',
    'PLANKART januar':                             'Zoning_PlanImport',
    'VeggrøftÅpen':                                'Road_DitchOpen',
    '761- Veier_ Kjøreveier_ sykkel- og gangveier mv_': 'Route_New',
    '83-- Tekst':                                  'Annotation_Text',
    'f-veg_95000_FLATE_KJØREFELT_ASFALT':          'Road_Asphalt_95000',
    'f-64000_fortau_FLATE_GANGSYKKELVEG':          'Road_PedestrianCycle',
    'Vegbom':                                      'Road_Barrier',
    'VegkantAvkjørsel':                            'Road_Driveway_Edge',
}

LANDSCAPE_MAP = {
    '00- ARK 12 PLAN 1 (250408)':                  'VisitorCentre_Architecture_Plan1',
    '773- Nytt tre':                               'Tree_New',
    '762- Forplass grafikk':                       'Forecourt_Graphics',
    '790- Utstyr_ møbler_ lekeapparater':          'OutdoorFurniture_PlayEquipment',
    '721- Støttemurer og andre murer':             'RetainingWall_New',
    '762- Plasser':                                'Forecourt',
    '900- RiVei koter 10cm':                       'RoadDesign_Contour10cm',
    '72-- Utendørs konstruksjoner':                'OutdoorStructures_New',
    '900- RiVei koter 20cm':                       'RoadDesign_Contour20cm',
    '702- møte eks ny kote':                       'Terrain_ContourMeeting',
    '712- Ny kote 20cm':                           'Terrain_NewContour20cm',
    '716- Ny punkthøyde':                          'Terrain_NewSpotHeight',
    '779- Stauder_ bregner_ prydgress_ løkplanter':'Planting_Perennials',
    '717- Fallpil':                                'Terrain_SlopeArrow',
    '712- Ny kote 10cm':                           'Terrain_NewContour10cm',
    '713- Ny kote tekst':                          'Terrain_NewContourLabel',
    '775- Busker':                                 'Planting_Shrubs',
    '830- Tekst':                                  'Annotation_Text',
    '733- Utendørs brannsslokking':                'Outdoor_FireProtection',
    '761- Veier_ Kjøreveier_ sykkel- og gangveier mv_': 'Route_New',
    '00- Sti 1-250':                               'Path_DetailPlan',
    '771- Plen_ gress_ eng':                       'Planting_Lawn',
    '00- Stein og fjell':                          'Rock_Outcrop',
    '731- Utendørs VA':                            'Outdoor_WaterSewer',
    'Sporing ytterkant':                           'VehicleTracking_OuterEdge',
    'TiltakGrense':                                'Planning_MeasureBoundary',
    'Vegrekkverk':                                 'Road_GuardRail',
    '762- Forplass grafikk B':                     'Forecourt_Graphics_B',
    '738- Utendørs fontener og springvann':        'Outdoor_Fountain',
    'Innmålt elektro':                             DROP,
    '710- Eks koter':                              DROP,
    '772- Eks tre':                                DROP,
    '711- Eks kote tekst':                         DROP,
    '830- tekst fjernet fra situasjonsplan':       DROP,
    '00- Hotlink Bygg':                            DROP,
    '0':                                           'Site_Outline_Misc',
}

PARKERING_MAP = {
    '0':                                           'Parking_Drawing',
    '_p-plass_boks':                               'Parking_Box',
    '_p-plass_dim':                                'Parking_Dimension',
    'f-veg_21000_FLATE_KJØREFELT_ASFALT':          'Road_Asphalt_21000',
    'Kjřrebane':                                   'Road_Lane',
    'f-64000_fortau_FLATE_GANGSYKKELVEG':          'Road_PedestrianCycle',
    '83-- Tekst':                                  'Annotation_Text',
    'f-veg_95000_FLATE_KJØREFELT_ASFALT':          'Road_Asphalt_95000',
    '0 (Skravur)':                                 'Hatch_Misc',
    'Generelt - DWG-_PDF-import':                  DROP,
}

def map_layer(mapping, raw_layer):
    if raw_layer in COORDINATION_LAYERS:
        return raw_layer
    cleaned = strip_layer(raw_layer)
    if cleaned in mapping:
        return mapping[cleaned]
    if raw_layer in mapping:
        return mapping[raw_layer]
    return None

# ─── Z flatten ────────────────────────────────────────────────────────────
def flatten_z(e):
    t = e.dxftype()
    try:
        if t == 'LINE':
            s = e.dxf.start; en = e.dxf.end
            e.dxf.start = (s.x, s.y, 0.0)
            e.dxf.end   = (en.x, en.y, 0.0)
        elif t == 'LWPOLYLINE':
            e.dxf.elevation = 0.0
        elif t == 'POLYLINE':
            for v in e.vertices:
                l = v.dxf.location
                v.dxf.location = (l.x, l.y, 0.0)
        elif t in ('TEXT', 'ATTRIB', 'ATTDEF'):
            p = e.dxf.insert
            e.dxf.insert = (p.x, p.y, 0.0)
            if e.dxf.hasattr('align_point'):
                a = e.dxf.align_point
                e.dxf.align_point = (a.x, a.y, 0.0)
        elif t == 'MTEXT':
            p = e.dxf.insert
            e.dxf.insert = (p.x, p.y, 0.0)
        elif t in ('CIRCLE', 'ARC'):
            c = e.dxf.center
            e.dxf.center = (c.x, c.y, 0.0)
        elif t == '3DFACE':
            for i in range(4):
                v = getattr(e.dxf, f'vtx{i}')
                setattr(e.dxf, f'vtx{i}', (v.x, v.y, 0.0))
        elif t == 'INSERT':
            p = e.dxf.insert
            e.dxf.insert = (p.x, p.y, 0.0)
        elif t == 'HATCH':
            try: e.dxf.elevation = (0.0, 0.0, 0.0)
            except Exception: pass
    except Exception:
        pass

# ─── Resolvers per file ──────────────────────────────────────────────────
def make_resolver_single(target):
    def f(e):
        if e.dxf.layer in COORDINATION_LAYERS:
            return e.dxf.layer
        return target
    return f

def make_resolver_map(mapping):
    def f(e):
        return map_layer(mapping, e.dxf.layer)
    return f

def resolver_kotelinjer(e):
    if e.dxf.layer in COORDINATION_LAYERS:
        return e.dxf.layer
    t = e.dxftype()
    if t in ('LINE', 'LWPOLYLINE', 'POLYLINE', 'ARC'):
        return 'ContourLine_20cm'
    if t in ('TEXT', 'MTEXT'):
        return 'ContourLabel_20cm'
    if t == 'CIRCLE':
        return DROP
    return None

def resolver_stier(e):
    lay = e.dxf.layer
    if lay in COORDINATION_LAYERS:
        return lay
    cleaned = strip_layer(lay)
    if cleaned == '0':
        t = e.dxftype()
        if t in ('MTEXT', 'CIRCLE'):
            return DROP
        return 'Site_Boundary_Misc'
    return map_layer(STIER_MAP, lay)

# ─── Process ─────────────────────────────────────────────────────────────
# Feature-category color palette (ACI color index).
# Longest prefix wins; checked in declaration order.
LAYER_COLOR_RULES = [
    ('VisitorCentre_',        140),  # purple
    ('Building_',               4),  # cyan
    ('Bridge_',                 4),  # cyan (same family)
    ('Road_',                   2),  # yellow
    ('Route_',                  2),
    ('RoadDesign_',             2),
    ('Parking_',                6),  # magenta
    ('Forecourt',               6),
    ('Path_',                   2),
    ('Tree_',                   3),  # green
    ('PlantingPlan_',           3),
    ('Planting_',               3),
    ('Terrain_',                8),  # grey
    ('ContourLine',             8),
    ('ContourLabel',            8),
    ('RetainingWall',          30),  # orange
    ('Wall_',                  30),
    ('Fence',                  30),
    ('River_',                  5),  # blue
    ('Lake',                    5),
    ('QuayPier',                5),
    ('Water',                   5),
    ('Stormwater',              5),
    ('Ditch_',                  5),
    ('FloodChannel',            5),
    ('Outdoor_Water',           5),
    ('Outdoor_Fire',            1),  # red
    ('Surveyed_',               1),
    ('COORDINATION_',           1),
    ('Utility_',               40),  # brown-ish
    ('Mast_',                  40),
    ('Track_',                 40),
    ('Rock_',                  40),
    ('Soil_',                  40),
    ('OutdoorFurniture',        6),
    ('OutdoorStructures',      30),
    ('Site_',                   8),
    ('Annotation_',             7),
    ('Planning_',               1),
    ('Unclassified',            7),
    ('VehicleTracking_',        2),
    ('VirtualDivisionLine',     8),
    ('Map_EdgeCutout',          8),
    ('Zoning_',                 1),
]

def color_for_layer(name):
    for prefix, color in LAYER_COLOR_RULES:
        if name.startswith(prefix):
            return color
    return 7  # white/black default

def ensure_layer(doc, name):
    if name not in doc.layers:
        doc.layers.add(name, dxfattribs={'color': color_for_layer(name)})
    else:
        # Layer already exists (e.g. imported from source) — override color
        # so the feature palette is consistent regardless of source order.
        doc.layers.get(name).color = color_for_layer(name)

def explode_all_inserts(msp, max_passes=5):
    """Recursively explode every INSERT in the modelspace until none remain
    (or max_passes is hit). Returns (exploded_count, residual_count)."""
    exploded = 0
    for _pass in range(max_passes):
        inserts = [e for e in msp if e.dxftype() == 'INSERT']
        if not inserts:
            return exploded, 0
        for ins in inserts:
            try:
                ins.explode()
                exploded += 1
            except Exception:
                pass
    residual = sum(1 for e in msp if e.dxftype() == 'INSERT')
    return exploded, residual

def process_file(target_doc, source_path, resolver, label):
    print(f'  {label}')
    try:
        src = ezdxf.readfile(source_path)
    except Exception as ex:
        print(f'    READ FAILED: {ex}')
        return 0
    src_msp = src.modelspace()
    exploded, residual = explode_all_inserts(src_msp)
    if exploded:
        print(f'    exploded {exploded} INSERTs  (residual: {residual})')
    imp = importer.Importer(src, target_doc)
    kept, dropped, unmapped = 0, 0, 0
    unmapped_layers = {}
    to_import = []
    target_layers_seen = set()
    oob = 0
    for e in src_msp:
        src_lay = e.dxf.layer
        if e.dxftype() in DROP_ENTITY_TYPES:
            dropped += 1
            continue
        if RESIDUAL_INSERT_DROP and e.dxftype() == 'INSERT':
            dropped += 1
            continue
        if not entity_in_site(e):
            oob += 1
            dropped += 1
            continue
        tgt = resolver(e)
        if tgt is DROP:
            dropped += 1
            continue
        if tgt is None:
            unmapped += 1
            unmapped_layers[src_lay] = unmapped_layers.get(src_lay, 0) + 1
            continue
        e.dxf.layer = tgt
        # POLYLINE has sub-entities (vertices + SEQEND) that carry their own
        # layer attribute. If we only rename the POLYLINE, SEQEND/Vertex
        # entities retain the source layer — ODA re-creates those layer
        # definitions during DXF→DWG, re-introducing Norwegian names.
        if e.dxftype() == 'POLYLINE':
            try:
                for v in e.vertices:
                    v.dxf.layer = tgt
                if e.seqend is not None:
                    e.seqend.dxf.layer = tgt
            except Exception:
                pass
        target_layers_seen.add(tgt)
        flatten_z(e)
        to_import.append(e)
        kept += 1
    # Pre-create target layers in the SOURCE doc so the Importer copies them
    # across (Importer walks src.layers for each layer referenced by entities).
    # Also create in target_doc as a belt-and-braces fallback.
    for tgt in target_layers_seen:
        if tgt not in src.layers:
            src.layers.add(tgt)
        ensure_layer(target_doc, tgt)
    try:
        imp.import_entities(to_import)
        imp.import_tables()  # pulls over layers, styles, linetypes, etc.
        imp.finalize()
    except Exception as ex:
        print(f'    IMPORT FAILED: {ex}')
    print(f'    kept={kept} dropped={dropped} unmapped={unmapped} oob={oob}')
    if unmapped_layers:
        for lay, n in sorted(unmapped_layers.items(), key=lambda x: -x[1])[:8]:
            print(f'      UNMAPPED  {n:>5}  {lay}')
    return kept

def find_in_master(name):
    want = nfc(name)
    for f in os.listdir(MASTER_DIR):
        if nfc(f) == want:
            return os.path.join(MASTER_DIR, f)
    return None

def main():
    os.makedirs(DXF_OUT_DIR, exist_ok=True)
    os.makedirs(DWG_OUT_DIR, exist_ok=True)

    target_doc = ezdxf.new('R2018', setup=True)
    target_doc.units = 6
    msp = target_doc.modelspace()

    sources = [
        (os.path.join(FDE_DIR, '3DE_Bygning_fra fkb.dxf'),                      make_resolver_map(BYGNING_MAP),         '3DE_Bygning'),
        (os.path.join(FDE_DIR, '3DE_Kart fra fkb.dxf'),                         make_resolver_map(KART_MAP),            '3DE_Kart'),
        (find_in_master('Innmålt_Elektro_NTM10_global_meters.dxf'),              make_resolver_single('Surveyed_Electrical'),       'Innmålt_Elektro'),
        (find_in_master('Innmålt_Kunstinstallasjon_NTM10_global_meters.dxf'),    make_resolver_single('Surveyed_ArtInstallation'),  'Innmålt_Kunstinstallasjon'),
        (find_in_master('Innmålt_Tre_NTM10_global_meters.dxf'),                  make_resolver_single('Tree_Surveyed'),             'Innmålt_Tre'),
        (find_in_master('TreStammer_Diameter_NTM10_global_meters.dxf'),          make_resolver_single('Tree_Trunk'),                'TreStammer_Diameter'),
        (find_in_master('Kotelinjer_MedPåskrift_2D_NTM10_global_meters.dxf'),    resolver_kotelinjer,                                'Kotelinjer_MedPåskrift_2D'),
        (find_in_master('KNM_Stier_NTM10_global.dxf'),                          resolver_stier,                                     'KNM_Stier'),
        (find_in_master('landscape plan visitor centre_NTM10_global_meters.dxf'), make_resolver_map(LANDSCAPE_MAP),                 'landscape plan visitor centre'),
        (find_in_master('Parkering_kistefos_LARK 1_NTM10_global_meters.dxf'),    make_resolver_map(PARKERING_MAP),                  'Parkering_kistefos_LARK 1'),
    ]

    print('Building master DATA v2...')
    total = 0
    for src, resolver, label in sources:
        if not src or not os.path.exists(src):
            print(f'  SKIP (not found): {label}')
            continue
        total += process_file(target_doc, src, resolver, label)

    # Per-layer summary
    msp_layers = {}
    for e in msp:
        msp_layers[e.dxf.layer] = msp_layers.get(e.dxf.layer, 0) + 1

    # Apply feature-category colors to every target layer, overriding any
    # colors carried over by the Importer's table merge.
    colored = 0
    for lay_name in msp_layers:
        try:
            layer = target_doc.layers.get(lay_name)
            layer.color = color_for_layer(lay_name)
            colored += 1
        except Exception:
            pass
    print(f'\nApplied colors to {colored} layers.')

    # First purge block definitions that nothing references. After explode,
    # all INSERTs should be gone, so every non-layout block is orphaned.
    referenced_blocks = {'*Model_Space', '*Paper_Space', '*Paper_Space0'}
    for e in msp:
        if e.dxftype() == 'INSERT':
            referenced_blocks.add(e.dxf.name)
    orphan_blocks = [b.name for b in target_doc.blocks
                     if b.name not in referenced_blocks
                     and not b.name.startswith('*')]
    for name in orphan_blocks:
        try:
            target_doc.blocks.delete_block(name, safe=False)
        except Exception:
            pass
    print(f'Purged {len(orphan_blocks)} orphaned block definitions.')

    # Purge unused layers from the layer table.
    # A layer is "used" if any entity anywhere in the doc references it —
    # including inside remaining block definitions (paperspace, layouts).
    PRESERVE = {'0', 'Defpoints'}
    used = set(msp_layers.keys())
    for block in target_doc.blocks:
        for e in block:
            try:
                used.add(e.dxf.layer)
            except Exception:
                pass
    to_purge = [l.dxf.name for l in target_doc.layers
                if l.dxf.name not in used and l.dxf.name not in PRESERVE]
    failures = 0
    for name in to_purge:
        try:
            target_doc.layers.remove(name)
        except Exception:
            failures += 1
    print(f'\nPurged {len(to_purge) - failures} unused layers from layer table '
          f'({failures} failed, {len(used)} used incl. block-internal).')

    out_dxf = os.path.join(DXF_OUT_DIR, f'{OUT_BASE}.dxf')
    print(f'\nTotal entities kept: {total}')
    print(f'Unique target layers: {len(msp_layers)}')
    print(f'\nPer-layer counts (top 30):')
    for lay, n in sorted(msp_layers.items(), key=lambda x: -x[1])[:30]:
        print(f'  {n:>7}  {lay}')

    print(f'\nWriting DXF: {out_dxf}')
    target_doc.saveas(out_dxf)

    # ODA convert to DWG, place next to the canonical master DWG
    print(f'\nConverting DXF -> DWG via ODA...')
    staging = r'C:\Users\edkjo\AppData\Local\Temp\_master_oda'
    tmp_in  = os.path.join(staging, 'in')
    tmp_out = os.path.join(staging, 'out')
    os.makedirs(tmp_in, exist_ok=True)
    os.makedirs(tmp_out, exist_ok=True)
    for f in os.listdir(tmp_in):  os.remove(os.path.join(tmp_in, f))
    for f in os.listdir(tmp_out): os.remove(os.path.join(tmp_out, f))
    shutil.copy(out_dxf, os.path.join(tmp_in, f'{OUT_BASE}.dxf'))
    subprocess.run([ODA, tmp_in, tmp_out, 'ACAD2018', 'DWG', '0', '1'], check=False)
    dwg_src = os.path.join(tmp_out, f'{OUT_BASE}.dwg')
    dwg_dst = os.path.join(DWG_OUT_DIR, f'{OUT_BASE}.dwg')
    if os.path.exists(dwg_src):
        shutil.move(dwg_src, dwg_dst)
        print(f'DWG: {dwg_dst}')
    else:
        print('DWG: conversion FAILED')

    print('\nDone.')

if __name__ == '__main__':
    main()
