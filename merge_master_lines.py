"""Merge selected DXF files into KNM_BIMK_MASTER_DATA — theme-grouped layers,
lines/points/text only, source-priority dedup, embedded legend block.
Produces global + local + mm variants + DWG conversion."""
import ezdxf, os, re, subprocess, shutil, tempfile, unicodedata
from ezdxf.addons import Importer


def normalize_short(name):
    """Normalize a source short-name to NFC and strip KNM_BIMK_ prefix.

    The prefix appears on the global Nairy file but not on the local one.
    NFC matters because os.listdir on this Windows install returns NFD-decomposed
    Norwegian characters (a + combining ring instead of composed å)."""
    name = unicodedata.normalize('NFC', name)
    if name.startswith('KNM_BIMK_'):
        name = name[len('KNM_BIMK_'):]
    return name

KEEP_TYPES = {'LINE', 'LWPOLYLINE', 'POLYLINE', 'ARC', 'CIRCLE', 'POINT', 'ELLIPSE', 'TEXT', 'MTEXT'}

# Appended to every output file name. Set via env var or by importer scripts
# that want to produce a parallel taxonomy variant (e.g. `_NS3451`, `_discipline`).
TAXONOMY_SUFFIX = os.environ.get('KNM_TAXONOMY_SUFFIX', '')


def _master_name(variant, ext):
    """Build master filename with optional taxonomy suffix.

    variant: 'GLOBAL_METERS' | 'GLOBAL_MM' | 'LOCAL_METERS' | 'LOCAL_MM'
    ext: 'dxf' | 'md' | 'pdf'
    """
    return f'KNM_BIMK_MASTER_DATA_NTM10_{variant}{TAXONOMY_SUFFIX}.{ext}'

# === THEMES === theme_name -> (ACI color, English description)
# Structure follows NS 3451:2022 Bygningsdelstabell where applicable.
# Non-NS themes (Topography, Regulatory, Internal, Other) cover data that
# falls outside NS 3451's scope (natural terrain, zoning, coordination).
THEMES = {
    # --- NS 3451 - 2 Bygning (Building) ---
    # Navn and Veiledning are quoted from NS 3451:2022 Bygningsdelstabell (Notion).
    'KNM_2_Building':                       (1,   '2 Bygning — omfatter bygningsmessige deler (includes heritage sites folded in)'),
    # --- NS 3451 - 7 Utendørs (2-digit fallbacks for layers without 3-digit detail) ---
    'KNM_70_Outdoor_General':               (8,   '70 Utendørs, generelt — alt utenfor bygningen, men innenfor tomtegrensen'),
    'KNM_71_Terrain':                       (4,   '71 Bearbeidet terreng — treated terrain, earthworks'),
    'KNM_72_Outdoor_Constructions':         (56,  '72 Utendørs konstruksjoner — inkludert graving, sprengning, fundamenter og gjenfylling'),
    'KNM_73_Pipework':                      (34,  '73 Utendørs røranlegg — alle VVS-installasjoner og anlegg utenfor bygget'),
    'KNM_74_Electrical_Outdoor':            (40,  '74 Utendørs elkraft — alle elkraftinstallasjoner og anlegg utenfor bygget'),
    'KNM_75_Telecom_Automation':            (200, '75 Utendørs ekom og automatisering — tele and automation outside building'),
    'KNM_76_Roads_Plazas':                  (6,   '76 Veier og plasser — overbygning inkl. forsterkningslag og kantstein'),
    'KNM_77_Park_Green':                    (3,   '77 Park og grøntanlegg — inkludert bearbeiding av undergrunnen, vekstjord, gjødsling, såing og planting, inkludert hager'),
    'KNM_78_External_Infrastructure':       (54,  '78 Utendørs infrastruktur — fysisk tilknytning til eksterne systemer'),
    'KNM_79_Outdoor_Other':                 (230, '79 Andre utendørs anlegg — catch-all when 2-digit inndeling is insufficient'),
    # --- NS 3451 - 7XX Nivå 3 (quoted Navn/Veiledning from Notion) ---
    'KNM_721_Retaining_Walls':              (16,  '721 Utendørs støttemurer og andre murer — inkludert støyskjermer utført som voll/mur'),
    'KNM_731_Water_Sewer':                  (171, '731 Utendørs VA — anlegg for vannforsyning, spillvann og bortledning av overflatevann'),
    'KNM_733_Fire_Water':                   (10,  '733 Utendørs brannslokking — inkludert hydranter og brannkummer med anlegg for vanntilførsel'),
    'KNM_738_Fountains':                    (150, '738 Utendørs fontener og springvann — inkludert særskilte pumper og renseanlegg'),
    'KNM_761_Routes':                       (12,  '761 Veier — kjøreveier, sykkel- og gangveier, løyper, stier, baner m.m. (all route types, not just roads)'),
    'KNM_762_Forecourt':                    (182, '762 — not in the NS 3451:2022 table directly; cross-referenced by 774 as "mindre områder, løkker o.l." Consultant uses for Forplass (Forecourt — entry/drop-off area in front of buildings)'),
    'KNM_771_Grass':                        (92,  '771 Utendørs gressarealer — plen, blomstereng m.m.'),
    'KNM_772_Planting':                     (62,  '772 Utendørs beplantning — blomsterbed, busker og trær (flowerbeds, shrubs, trees — all planting)'),
    'KNM_773_Outdoor_Equipment':            (44,  '773 Utendørs utstyr — benker, lekeapparater, flaggstenger, utsmykning (skulpturer), inkludert fundamentering'),
    'KNM_775_Bleachers':                    (106, '775 Utendørs tribuner og amfier (NS 3451). NOTE: Asplan Viak uses 775 for "Busker" (shrubs), which per NS actually belongs under 772 Utendørs beplantning. Layers are routed here to respect the consultant prefix.'),
    'KNM_777_Reserved':                     (94,  '777 (Reservert) — koden skal ikke benyttes per NS 3451:2022. Asplan Viak uses "777 Skjøtsel" (maintenance) which is a deviation from the standard.'),
    'KNM_779_Other_Park_Green':             (64,  '779 Andre deler for park og grøntanlegg — for anvendelse når øvrig 77X-inndeling ikke er dekkende. Consultant uses for stauder, bregner, prydgress, løkplanter.'),
    'KNM_790_Furniture_Equipment':          (210, '790 — not in the NS 3451:2022 table. Consultant uses as a child of 79 "Andre utendørs anlegg" for utstyr, møbler, lekeapparater (furniture, equipment, playground). Note: in strict NS 3451, outdoor furniture belongs under 773 Utendørs utstyr.'),
    # --- Non-NS: Topography (survey data outside NS 3451 scope) ---
    'KNM_Topography_Contours':              (2,   'Contour lines, elevation text, spot heights (survey)'),
    'KNM_Topography_Water_Natural':         (140, 'Natural water: rivers, lakes, streams, dams'),
    # --- Non-NS: Regulatory ---
    'KNM_Regulatory':                       (9,   'Plankart, tiltaksgrense, zoning, administrative boundaries'),
    # --- Non-NS: Internal project coordination ---
    'KNM_Internal_Map_Markers_And_Other':   (250, 'Defpoints, coordination markers, helpers, text, annotations'),
    # --- Catch-all ---
    'KNM_Other':                            (7,   'Unmapped — for human review'),
}

# === LAYER RULES === First match wins. Pattern is matched against the
# normalized layer name (with "(Nye konstruksjoner)" suffix stripped).
# Theme = None means "fall through to source fallback".
#
# Rule order:
#   1. Ambiguous fallthroughs (0, Diverse, Layer N)
#   2. Internal / coordination / text (defpoints, markers, annotations)
#   3. Regulatory (tiltak, plankart, regulering)
#   4. Heritage → folds into KNM_2_Building
#   5. Topography (natural contours and water — not NS 3451 scope)
#   6. NS 3451 numeric prefixes (^2-, ^7XX-) — highest-confidence mapping
#   7. Descriptive Norwegian keywords as fallback mapping to NS themes
LAYER_RULES = [
    # --- Ambiguous / fallthrough to source ---
    (r'^0$',                                 None),
    (r'^Diverse$',                           None),
    (r'^Layer\s*\d+$',                       None),

    # === Internal / project coordination / text ===
    (r'^COORDINATION_MARKER',                'KNM_Internal_Map_Markers_And_Other'),
    (r'^Defpoints$',                         'KNM_Internal_Map_Markers_And_Other'),
    (r'^Sporing',                            'KNM_Internal_Map_Markers_And_Other'),
    (r'^00_TRACET',                          'KNM_Internal_Map_Markers_And_Other'),
    (r'^BeskrivendeHjelpe',                  'KNM_Internal_Map_Markers_And_Other'),
    (r'^Hjelpelinje',                        'KNM_Internal_Map_Markers_And_Other'),
    (r'^827-\s*Snittlinjer',                 'KNM_Internal_Map_Markers_And_Other'),
    # 837- Planteplan trær ... = new tree planting plan → 77 Park (must come before generic 83-)
    (r'^837-.*Planteplan.*tr',               'KNM_77_Park_Green'),
    (r'^83[0-9]?-',                          'KNM_Internal_Map_Markers_And_Other'),
    (r'^847-',                               'KNM_Internal_Map_Markers_And_Other'),
    (r'\bTekst\b',                           'KNM_Internal_Map_Markers_And_Other'),

    # === Regulatory ===
    (r'^TiltakGrense',                       'KNM_Regulatory'),
    (r'^PblTiltak',                          'KNM_Regulatory'),
    (r'^Arealbrukgrense',                    'KNM_Regulatory'),
    (r'^Arealgrense',                        'KNM_Regulatory'),
    (r'^Anleg[gs]*omr.de',                   'KNM_Regulatory'),
    (r'^PLANKART',                           'KNM_Regulatory'),
    (r'regulering',                          'KNM_Regulatory'),

    # === Heritage → folds into Building ===
    (r'^Kulturminne',                        'KNM_2_Building'),

    # === Topography - natural contours (not NS 3451 scope) ===
    (r'^70[12]-',                            'KNM_Topography_Contours'),
    (r'^71[0-7]-',                           'KNM_Topography_Contours'),
    (r'^900-\s*RiVei\s*koter',               'KNM_Topography_Contours'),
    (r'Kotelinjer',                          'KNM_Topography_Contours'),
    (r'\bkote\b',                            'KNM_Topography_Contours'),
    (r'^PresH.ydetall',                      'KNM_Topography_Contours'),
    (r'^Forsenkningskurve',                  'KNM_Topography_Contours'),

    # === Topography - natural water ===
    (r'^Dam(kant)?$',                        'KNM_Topography_Water_Natural'),
    (r'^ElvBekk',                            'KNM_Topography_Water_Natural'),
    (r'^Floml.pkant',                        'KNM_Topography_Water_Natural'),
    (r'^ElvelinjeFiktiv',                    'KNM_Topography_Water_Natural'),
    (r'^Innsj.',                             'KNM_Topography_Water_Natural'),
    (r'^Kanal(Gr.ft)?$',                     'KNM_Topography_Water_Natural'),
    (r'^Fisketrapp',                         'KNM_Topography_Water_Natural'),

    # === NS 3451 - 7XX Nivå 3 explicit prefixes (3-digit before 2-digit fallback) ===
    # 72X
    (r'^721-',                               'KNM_721_Retaining_Walls'),
    # 73X
    (r'^731-',                               'KNM_731_Water_Sewer'),
    (r'^733-',                               'KNM_733_Fire_Water'),
    (r'^738-',                               'KNM_738_Fountains'),
    # 76X
    (r'^761-',                               'KNM_761_Routes'),
    (r'^762-',                               'KNM_762_Forecourt'),
    # 77X
    (r'^771-',                               'KNM_771_Grass'),
    (r'^772-',                               'KNM_772_Planting'),
    (r'^773-',                               'KNM_773_Outdoor_Equipment'),
    (r'^775-',                               'KNM_775_Bleachers'),
    (r'^777-?\s*Skj',                        'KNM_777_Reserved'),
    (r'^779-',                               'KNM_779_Other_Park_Green'),
    # 79X
    (r'^790-',                               'KNM_790_Furniture_Equipment'),

    # === NS 3451 - 2-digit fallback (catches remaining 7XX codes) ===
    (r'^70[3-9]-',                           'KNM_70_Outdoor_General'),
    (r'^72[0-9]?-',                          'KNM_72_Outdoor_Constructions'),
    (r'^73[0-9]?-',                          'KNM_73_Pipework'),
    (r'^74[0-9]?-',                          'KNM_74_Electrical_Outdoor'),
    (r'^75[0-9]?-',                          'KNM_75_Telecom_Automation'),
    (r'^76[0-9]?-',                          'KNM_76_Roads_Plazas'),
    (r'^77[0-9]?-',                          'KNM_77_Park_Green'),
    (r'^78[0-9]?-',                          'KNM_78_External_Infrastructure'),
    (r'^79[0-9]?-',                          'KNM_79_Outdoor_Other'),

    # === NS 2 Bygning - descriptive Norwegian building keywords ===
    (r'^Bygning',                            'KNM_2_Building'),
    (r'^AnnenBygning',                       'KNM_2_Building'),
    (r'^Bru\b',                              'KNM_2_Building'),
    (r'^Bruavgrensning',                     'KNM_2_Building'),
    (r'^Veranda',                            'KNM_2_Building'),
    (r'^00-?\s*ARK',                         'KNM_2_Building'),
    (r'^00-?\s*Hotlink\s*Bygg',              'KNM_2_Building'),
    (r'^Bygningsdelelinje',                  'KNM_2_Building'),
    (r'^Bygningslinje',                      'KNM_2_Building'),
    (r'^Takkant',                            'KNM_2_Building'),
    (r'^M.nelinje',                          'KNM_2_Building'),
    (r'^Taksprang',                          'KNM_2_Building'),
    (r'^TakoverbyggKant',                    'KNM_2_Building'),
    (r'^Takoverbygg',                        'KNM_2_Building'),
    (r'^Pipe(kant)?$',                       'KNM_2_Building'),
    (r'^L.vebru',                            'KNM_2_Building'),
    (r'^TrappBygg',                          'KNM_2_Building'),
    (r'^BautaStatue',                        'KNM_2_Building'),

    # === NS 72 Outdoor constructions - descriptive keywords ===
    (r'^Mur\b',                              'KNM_72_Outdoor_Constructions'),
    (r'^MurLoddrett',                        'KNM_72_Outdoor_Constructions'),
    (r'^00-?\s*Stein\s*og\s*fjell',          'KNM_72_Outdoor_Constructions'),
    (r'^Fritt(st.ende)?Trapp',               'KNM_72_Outdoor_Constructions'),
    (r'^Fundament',                          'KNM_72_Outdoor_Constructions'),
    (r'^Skr.Forst.tning',                    'KNM_72_Outdoor_Constructions'),
    (r'^KaiBrygge',                          'KNM_72_Outdoor_Constructions'),
    (r'^Sv.mmebasseng',                      'KNM_72_Outdoor_Constructions'),

    # === NS 731 Water/sewer - descriptive keywords ===
    (r'^Stikkrenne',                         'KNM_731_Water_Sewer'),
    (r'^R.rgate',                            'KNM_731_Water_Sewer'),

    # === NS 74 Electrical - descriptive keywords ===
    (r'^Masteomriss',                        'KNM_74_Electrical_Outdoor'),
    (r'^Innm.lt\s*kunstinstall',             'KNM_79_Outdoor_Other'),  # art installation survey
    (r'^Innm.lt\b',                          'KNM_74_Electrical_Outdoor'),

    # === NS 76 Roads and plazas - descriptive keywords ===
    (r'^Veg(dekkekant|kant|rekkverk|bom|gr.ft)', 'KNM_76_Roads_Plazas'),
    (r'^Veg$',                               'KNM_76_Roads_Plazas'),
    (r'^Vegskulderkant',                     'KNM_76_Roads_Plazas'),
    (r'^AnnetVegareal',                      'KNM_76_Roads_Plazas'),
    (r'^Kj.rebane',                          'KNM_76_Roads_Plazas'),
    (r'^Spormidt',                           'KNM_76_Roads_Plazas'),
    (r'^f-veg_',                             'KNM_76_Roads_Plazas'),
    (r'^AnnetGjerde',                        'KNM_76_Roads_Plazas'),
    (r'^GangSykkelveg',                      'KNM_76_Roads_Plazas'),
    (r'^Gangvegkant',                        'KNM_76_Roads_Plazas'),
    (r'^Fortauskant',                        'KNM_76_Roads_Plazas'),
    (r'^00-?\s*nye?\s*stier',                'KNM_76_Roads_Plazas'),
    (r'^00\s*Stier',                         'KNM_76_Roads_Plazas'),
    (r'^00-?\s*bes',                         'KNM_76_Roads_Plazas'),
    (r'^f-64000',                            'KNM_76_Roads_Plazas'),
    (r'^_p-plass',                           'KNM_76_Roads_Plazas'),
    (r'^Parkering',                          'KNM_76_Roads_Plazas'),
    (r'^00-?\s*(fase\s*\d+\s*)?parkering',   'KNM_76_Roads_Plazas'),

    # === NS 77 Park and vegetation - descriptive keywords ===
    (r'^T.rr?(Lauvtre|Gran|Furu|Bj.rk|Eik|Hegg|Rogn)', 'KNM_77_Park_Green'),
    (r'^(Lauvtre|Bj.rk|Furu|Eik|Gran|Hegg|Rogn|L.nn|Or)[\b_-]', 'KNM_77_Park_Green'),
    (r'^(Lauvtre|Bj.rk|Furu|Eik|Gran|Hegg|Rogn|L.nn|Or)$', 'KNM_77_Park_Green'),
    (r'^Tre(Stamme|Krone|Punkt)',            'KNM_77_Park_Green'),
    (r'^Hekk',                               'KNM_77_Park_Green'),
    (r'^Arealressurs',                       'KNM_77_Park_Green'),
    (r'^00-?\s*ny\s*jord',                   'KNM_77_Park_Green'),
    (r'^00-?forflytning\s*jord',             'KNM_77_Park_Green'),
    (r'transplantert',                       'KNM_77_Park_Green'),

    # === NS 79 Other outdoor - descriptive keywords ===
    (r'^Lekeplass',                          'KNM_79_Outdoor_Other'),

    # === NS 71 Terrain - earthworks fallback ===
    (r'Veg_TIN',                             'KNM_71_Terrain'),
]

LAYER_RULES_COMPILED = [(re.compile(p, re.IGNORECASE), t) for p, t in LAYER_RULES]


def _normalize_dict_keys(d):
    return {unicodedata.normalize('NFC', k): v for k, v in d.items()}


def _normalize_priority_lists(d):
    return {k: [unicodedata.normalize('NFC', s) for s in v] for k, v in d.items()}

# === SOURCE FALLBACK === If no rule matches, theme based on source file.
SOURCE_FALLBACK = {
    '240806_Nairy Baghramian 1 1':    'KNM_Other',
    'landscape plan visitor centre':  'KNM_Other',
    'KNM_Stier':                      'KNM_76_Roads_Plazas',
    'Kotelinjer_20cm':                'KNM_Topography_Contours',
    'Kotelinjer_MedPåskrift_2D':      'KNM_Topography_Contours',
    '230831_Veg_TIN':                 'KNM_71_Terrain',
    'Parkering_kistefos_LARK 1':      'KNM_76_Roads_Plazas',
}
SOURCE_FALLBACK = _normalize_dict_keys(SOURCE_FALLBACK)

# === SOURCE SCOPE === per source file → physical scope on the Kistefos site.
# Values: 'KNM' (new museum), 'BS' (visitor centre / besøkssenter),
#         'TT' (The Twist), 'Site' (site-wide / ownership defaults to Kistefos).
# Used as a prefix in report descriptions so stakeholders can tell which
# contract each layer belongs to. Default = 'Site' when not listed.
SOURCE_SCOPE = {
    'landscape plan visitor centre':  'BS',
    '240806_Nairy Baghramian 1 1':    'Site',
    'KNM_Stier':                      'Site',
    'Kotelinjer_20cm':                'Site',
    'Kotelinjer_MedPåskrift_2D':      'Site',
    '230831_Veg_TIN':                 'Site',
    'Parkering_kistefos_LARK 1':      'Site',
}
SOURCE_SCOPE = _normalize_dict_keys(SOURCE_SCOPE)

def scope_for(source_short):
    return SOURCE_SCOPE.get(unicodedata.normalize('NFC', source_short), 'Site')

# === SOURCE PRIORITY === per theme, highest priority source first.
# Only the top-priority source actually present contributes to that theme.
# Themes not listed: all sources contribute (no dedup).
#
# IMPORTANT: source-priority dedup is only safe when multiple sources draw the
# *same physical features*. For most themes here, the consultant drawings cover
# *different parts of the site* (e.g. landscape plan = visitor centre area,
# Nairy = art installation area, Kotelinjer_20cm vs MedPåskrift = different
# contour products with different content), so dedup would lose real geometry.
# We only dedup KNM_Helpers because every source emits the same 4 coordination
# markers and Defpoints — those are true duplicates.
SOURCE_PRIORITY = {
    # Only one set of coordination markers needed in the master
    'KNM_Internal_Map_Markers_And_Other': [
        'landscape plan visitor centre',
    ],
    # Contours: only MedPåskrift_2D — it's the 20 cm survey with labels.
    # Other sources (Kotelinjer_20cm, landscape plan contours, Nairy contours)
    # are dropped from MASTER_DATA. MASTER_LOCATION is unaffected (built separately).
    'KNM_Topography_Contours': [
        'Kotelinjer_MedPåskrift_2D',
    ],
}
SOURCE_PRIORITY = _normalize_priority_lists(SOURCE_PRIORITY)

# === FILE INCLUSION === substrings — only DXFs matching one of these get merged.
INCLUDED_FILES = [
    'Kotelinjer', 'Veg_TIN', 'Nairy',
    'Stier', 'Parkering', 'landscape plan',
]

# === LAYER SKIP === (source_short_substring, layer_name_regex) pairs.
# Entities matching both are dropped entirely (not mapped to any theme).
# Use this to remove known duplicates between sources.
LAYER_SKIP = [
    # Nairy file carries a copy of BANGS 20cm contour survey — already in Kotelinjer_20cm
    ('Nairy', r'^innm.lt\s*kotelinjer\s*20'),
]
LAYER_SKIP_COMPILED = [(s, re.compile(p, re.IGNORECASE)) for s, p in LAYER_SKIP]

ODA = r'C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe'

# New basepoint (NTM10 → local origin) — see CLAUDE.md
BASEPOINT = (92200.0, 1247000.0, 0.0)


def normalize_layer(name):
    """Strip Norwegian suffix used by Nairy/landscape files."""
    return name.replace('(Nye konstruksjoner)', '').strip()


# Themes whose content is inherently existing (surveyed terrain, heritage).
_EXISTING_BY_THEME = {
    'KNM_Topography_Contours',
    'KNM_Topography_Water_Natural',
}

# Sources where '(Nye konstruksjoner)' is a file-level template artifact,
# not a real per-layer new/existing flag. The Nairy FKB survey file tags
# every layer with the suffix even though its content is existing terrain.
_SUFFIX_NOT_MEANINGFUL = {
    unicodedata.normalize('NFC', '240806_Nairy Baghramian 1 1'),
}

# Sources whose content is predominantly surveyed/existing by nature.
# Default to 'Existing' when no other signal is present.
_SOURCE_DEFAULT_EXISTING = {
    unicodedata.normalize('NFC', '240806_Nairy Baghramian 1 1'),
    unicodedata.normalize('NFC', 'Kotelinjer_20cm'),
    unicodedata.normalize('NFC', 'Kotelinjer_MedPåskrift_2D'),
}

_NEW_SUFFIX = '(Nye konstruksjoner)'

_EXISTING_PAT = re.compile(
    r'(\(Eksisterende\)|\bEks(isterende)?\b|\beks\s|'
    r'\bEks[_ ]?tre|tørr|^T[øo]rr|^Innm\.?lt|innmålt)',
    re.IGNORECASE,
)
_NEW_PAT = re.compile(
    r'(\bNytt?\b|\bny\b|^00-?\s*nye?|transplantert|Planteplan|f-veg_|f-64000)',
    re.IGNORECASE,
)


def infer_status(layer_name, theme, source_short=None):
    """Return 'New', 'Existing', or '' based on layer name, theme, source.

    Rules (first match wins):
      1. Heritage (Kulturminne match) → Existing
      2. '(Nye konstruksjoner)' suffix → New (only if source is not in
         _SUFFIX_NOT_MEANINGFUL, where the suffix is a template artifact)
      3. Other NEW markers (Ny, Planteplan, transplantert…) → New
      4. EXISTING markers (Eks, tørr, Innmålt…) → Existing
      5. Theme default (topography themes) → Existing
      6. Source default (FKB/survey sources) → Existing
      7. Otherwise '' (unknown — lands in the unsuffixed base layer)
    """
    src_norm = unicodedata.normalize('NFC', source_short) if source_short else ''
    suffix_meaningful = src_norm not in _SUFFIX_NOT_MEANINGFUL

    if re.match(r'^Kulturminne', layer_name, re.IGNORECASE):
        return 'Existing'
    if suffix_meaningful and _NEW_SUFFIX in layer_name:
        return 'New'
    if _NEW_PAT.search(layer_name):
        return 'New'
    if _EXISTING_PAT.search(layer_name):
        return 'Existing'
    if theme in _EXISTING_BY_THEME:
        return 'Existing'
    if src_norm in _SOURCE_DEFAULT_EXISTING:
        return 'Existing'
    return ''


def layer_to_theme(layer_name, source_short):
    norm = normalize_layer(layer_name)
    for pat, theme in LAYER_RULES_COMPILED:
        if pat.search(norm):
            if theme is None:
                break  # fall through to source fallback
            return theme
    return SOURCE_FALLBACK.get(source_short, 'KNM_Other')


def collect_files(base_dir):
    """Collect DXFs in base_dir and Design/BANGS subfolders, deduped by short name."""
    bangs = os.path.join(base_dir, 'BANGS_Innmaaling')
    design = os.path.join(base_dir, 'Design')
    files = []
    seen = set()

    for d in [bangs, design, base_dir]:
        if not os.path.exists(d):
            continue
        for f in os.listdir(d):
            if not f.endswith('.dxf'):
                continue
            if 'BIMK_Master' in f or 'BIMK_MASTER' in f:
                continue
            if 'KNM_MASTER' in f:
                continue
            if 'redigert' in f:
                continue
            if f.startswith('ACAD-'):
                continue
            if 'Stier_flere' in f:
                continue
            if not any(s in f for s in INCLUDED_FILES):
                continue
            short = normalize_short(f.split('_NTM10')[0])
            if short in seen:
                continue
            files.append((os.path.join(d, f), f, short))
            seen.add(short)
    return sorted(files, key=lambda x: x[2])


def purge_unused_layers(doc, protect=('0', 'Defpoints', 'KNM_LEGEND')):
    """Remove layer table entries AND orphaned block definitions.

    After a theme-remap merge, the Importer leaves thousands of orphaned source
    layers in doc.layers AND thousands of block definitions that reference those
    layers. If we only remove the layer table entries, ODA File Converter will
    recreate them from the block defs during DXF→DWG conversion — defeating the
    entire purpose of theme grouping.

    Solution: purge unused blocks first (they reference old layers), then purge
    the now-unreferenced layers.
    """
    msp = doc.modelspace()

    # Step 1: find blocks actually referenced in modelspace (there shouldn't be
    # any INSERT entities since we filter by KEEP_TYPES, but be safe).
    used_blocks = set()
    for e in msp:
        if e.dxftype() == 'INSERT':
            used_blocks.add(e.dxf.name)

    # Step 2: collect blocks referenced by dimstyles (arrowheads etc.)
    dimstyle_blocks = set()
    for ds in doc.dimstyles:
        for attr in ('dimblk', 'dimblk1', 'dimblk2', 'dimldrblk'):
            try:
                val = ds.dxf.get(attr)
                if val:
                    dimstyle_blocks.add(val)
            except Exception:
                pass

    # Step 3: purge unreferenced block definitions.
    # Keep special blocks (*Model_Space, *Paper_Space, etc.) and dimstyle blocks.
    block_names = [b.name for b in doc.blocks if not b.name.startswith('*')]
    blocks_removed = 0
    for bname in block_names:
        if bname not in used_blocks and bname not in dimstyle_blocks:
            try:
                doc.blocks.delete_block(bname, safe=False)
                blocks_removed += 1
            except Exception:
                pass

    # Step 3: find layers actually used by remaining entities (modelspace + any
    # surviving block defs).
    used_layers = set()
    for e in msp:
        try:
            used_layers.add(e.dxf.layer)
        except AttributeError:
            pass
    for block in doc.blocks:
        for e in block:
            try:
                used_layers.add(e.dxf.layer)
            except AttributeError:
                pass

    # Step 4: purge unused layers. Only keep layers that have entities,
    # plus a few protected ones (0, Defpoints, KNM_LEGEND). Base theme layers
    # without entities (common when every entity falls into a _New/_Existing
    # suffix variant) are purged so the output layer table stays tight.
    keep = set(protect) | used_layers
    layers_removed = 0
    for layer in list(doc.layers):
        name = layer.dxf.name
        if name not in keep:
            try:
                doc.layers.remove(name)
                layers_removed += 1
            except Exception:
                pass

    print(f'  Purged {blocks_removed} unused blocks, {layers_removed} unused layers')
    return layers_removed


def offset_entity(e, dx, dy, dz=0.0):
    """Offset all positional coordinates of an entity. Sizes/radii are unchanged."""
    t = e.dxftype()
    if t == 'LINE':
        s = e.dxf.start
        e.dxf.start = (s.x + dx, s.y + dy, s.z + dz)
        s = e.dxf.end
        e.dxf.end = (s.x + dx, s.y + dy, s.z + dz)
    elif t == 'LWPOLYLINE':
        pts = list(e.get_points(format='xyseb'))
        e.set_points([(x+dx, y+dy, sv, ew, b) for x, y, sv, ew, b in pts], format='xyseb')
    elif t in ('CIRCLE', 'ARC'):
        c = e.dxf.center
        e.dxf.center = (c.x+dx, c.y+dy, c.z+dz)
    elif t == 'POINT':
        loc = e.dxf.location
        e.dxf.location = (loc.x+dx, loc.y+dy, loc.z+dz)
    elif t == 'ELLIPSE':
        c = e.dxf.center
        e.dxf.center = (c.x+dx, c.y+dy, c.z+dz)
    elif t == 'POLYLINE':
        for v in e.vertices:
            loc = v.dxf.location
            v.dxf.location = (loc[0]+dx, loc[1]+dy, loc[2]+dz)
    elif t in ('TEXT', 'MTEXT'):
        ins = e.dxf.insert
        e.dxf.insert = (ins.x+dx, ins.y+dy, ins.z+dz)


def derive_local_from_global(global_dxf, local_dxf, basepoint=BASEPOINT):
    """Read a global meters DXF, offset every entity by -basepoint, save as local."""
    doc = ezdxf.readfile(global_dxf)
    msp = doc.modelspace()
    dx, dy, dz = -basepoint[0], -basepoint[1], -basepoint[2]
    for e in msp:
        offset_entity(e, dx, dy, dz)
    doc.header['$INSUNITS'] = 6  # meters
    doc.saveas(local_dxf)
    print(f'Derived local: {local_dxf}')
    return local_dxf


def scale_entity(e, scale):
    t = e.dxftype()
    if t == 'LINE':
        s = e.dxf.start
        e.dxf.start = (s.x * scale, s.y * scale, s.z * scale)
        s = e.dxf.end
        e.dxf.end = (s.x * scale, s.y * scale, s.z * scale)
    elif t == 'LWPOLYLINE':
        pts = list(e.get_points(format='xyseb'))
        e.set_points([(x*scale, y*scale, sv*scale, ew*scale, b) for x, y, sv, ew, b in pts], format='xyseb')
    elif t == 'CIRCLE':
        c = e.dxf.center
        e.dxf.center = (c.x*scale, c.y*scale, c.z*scale)
        e.dxf.radius *= scale
    elif t == 'ARC':
        c = e.dxf.center
        e.dxf.center = (c.x*scale, c.y*scale, c.z*scale)
        e.dxf.radius *= scale
    elif t == 'POINT':
        loc = e.dxf.location
        e.dxf.location = (loc.x*scale, loc.y*scale, loc.z*scale)
    elif t == 'ELLIPSE':
        c = e.dxf.center
        e.dxf.center = (c.x*scale, c.y*scale, c.z*scale)
        maj = e.dxf.major_axis
        e.dxf.major_axis = (maj[0]*scale, maj[1]*scale, maj[2]*scale)
    elif t == 'POLYLINE':
        for v in e.vertices:
            loc = v.dxf.location
            v.dxf.location = (loc[0]*scale, loc[1]*scale, loc[2]*scale)
    elif t in ('TEXT', 'MTEXT'):
        ins = e.dxf.insert
        e.dxf.insert = (ins.x*scale, ins.y*scale, ins.z*scale)
        if hasattr(e.dxf, 'char_height'):
            e.dxf.char_height *= scale
        if hasattr(e.dxf, 'height'):
            e.dxf.height *= scale


def merge_files(files_to_merge, out_path):
    merged = ezdxf.new('R2018')
    merged.header['$INSUNITS'] = 6
    merged.header['$MEASUREMENT'] = 1
    msp_out = merged.modelspace()

    sources_present = {short for _, _, short in files_to_merge}

    # Decide winning source per theme that has a priority list
    theme_winner = {}
    for theme, prio in SOURCE_PRIORITY.items():
        for s in prio:
            if s in sources_present:
                theme_winner[theme] = s
                break

    # Pre-create theme layers
    for theme, (color, _) in THEMES.items():
        if theme not in merged.layers:
            merged.layers.add(theme, color=color)

    # Stats
    # theme -> {(orig_layer_norm, source_short): count}
    theme_layer_map = {t: {} for t in THEMES}
    theme_counts = {t: 0 for t in THEMES}
    dropped_dedup = {}  # (theme, source) -> count

    total_imported = 0
    for path, fname, short in files_to_merge:
        print(f'\n--- {short} ---')
        src = ezdxf.readfile(path)
        is_mm = '_mm.' in fname or '_mm_' in fname

        importer = Importer(src, merged)
        importer.import_modelspace()
        importer.finalize()

        all_entities = list(msp_out)
        new_entities = all_entities[total_imported:]

        kept = 0
        dropped_type = 0
        dropped_d = 0
        to_remove = []

        for e in new_entities:
            if e.dxftype() not in KEEP_TYPES:
                to_remove.append(e)
                dropped_type += 1
                continue

            orig_layer = e.dxf.layer
            orig_norm = normalize_layer(orig_layer)

            # Check per-source layer skip rules (known duplicates)
            skip = False
            for src_sub, pat in LAYER_SKIP_COMPILED:
                if src_sub in short and pat.search(orig_norm):
                    skip = True
                    break
            if skip:
                to_remove.append(e)
                dropped_d += 1
                key = ('SKIP', short)
                dropped_dedup[key] = dropped_dedup.get(key, 0) + 1
                continue

            theme = layer_to_theme(orig_layer, short)

            winner = theme_winner.get(theme)
            if winner and winner != short:
                to_remove.append(e)
                dropped_d += 1
                key = (theme, short)
                dropped_dedup[key] = dropped_dedup.get(key, 0) + 1
                continue

            if is_mm:
                scale_entity(e, 0.001)

            # Split theme into up to 3 variants by inferred status (New / Existing / unknown base)
            status = infer_status(orig_layer, theme, short)
            output_layer = f'{theme}_{status}' if status else theme
            if output_layer not in merged.layers:
                color = THEMES[theme][0]
                merged.layers.add(output_layer, color=color)
            e.dxf.layer = output_layer
            theme_counts[theme] += 1

            norm = normalize_layer(orig_layer)
            key = (norm, short)
            theme_layer_map[theme][key] = theme_layer_map[theme].get(key, 0) + 1

            kept += 1

        for e in to_remove:
            msp_out.delete_entity(e)

        total_imported = len(list(msp_out))
        print(f'  kept={kept}, type-dropped={dropped_type}, dedup-dropped={dropped_d}')

    # Theme summary
    print('\n=== Theme summary ===')
    for theme in THEMES:
        if theme_counts[theme]:
            print(f'  {theme}: {theme_counts[theme]} entities, {len(theme_layer_map[theme])} merged source layers')
    if dropped_dedup:
        print('\n=== Dedup drops ===')
        for (theme, src), n in sorted(dropped_dedup.items()):
            print(f'  {theme} <- {src}: dropped {n}')

    # Purge unused layers from the table. After theme-remapping, every source
    # layer imported by the Importer is orphaned (no entities reference it),
    # but it's still in doc.layers — so AutoCAD's Layer Manager shows thousands
    # of empty layers. Drop them all; keep only themes + a few protected names.
    purged = purge_unused_layers(merged)
    print(f'\nPurged {purged} unused layers from the layer table')

    add_legend_block(merged, theme_counts)

    merged.saveas(out_path)
    total_final = len(list(merged.modelspace()))
    used_themes = sum(1 for c in theme_counts.values() if c)
    print(f'\nSaved: {out_path}')
    print(f'Total entities: {total_final}, Theme layers used: {used_themes}')

    return out_path, theme_layer_map, theme_counts, dropped_dedup


def compute_bbox(msp):
    min_x = min_y = float('inf')
    max_x = max_y = float('-inf')
    found = False
    for e in msp:
        t = e.dxftype()
        try:
            if t == 'LINE':
                for p in (e.dxf.start, e.dxf.end):
                    if p.x < min_x: min_x = p.x
                    if p.x > max_x: max_x = p.x
                    if p.y < min_y: min_y = p.y
                    if p.y > max_y: max_y = p.y
                    found = True
            elif t == 'LWPOLYLINE':
                for x, y, *_ in e.get_points():
                    if x < min_x: min_x = x
                    if x > max_x: max_x = x
                    if y < min_y: min_y = y
                    if y > max_y: max_y = y
                    found = True
            elif t in ('CIRCLE', 'ARC'):
                c = e.dxf.center
                r = e.dxf.radius
                if c.x - r < min_x: min_x = c.x - r
                if c.x + r > max_x: max_x = c.x + r
                if c.y - r < min_y: min_y = c.y - r
                if c.y + r > max_y: max_y = c.y + r
                found = True
            elif t == 'POINT':
                loc = e.dxf.location
                if loc.x < min_x: min_x = loc.x
                if loc.x > max_x: max_x = loc.x
                if loc.y < min_y: min_y = loc.y
                if loc.y > max_y: max_y = loc.y
                found = True
        except Exception:
            pass
    if not found:
        return None
    return (min_x, min_y, max_x, max_y)


def add_legend_block(doc, theme_counts):
    """Add an MTEXT legend just to the right of the model extents."""
    msp = doc.modelspace()
    bbox = compute_bbox(msp)
    if bbox is None:
        return
    min_x, min_y, max_x, max_y = bbox
    width = max_x - min_x
    height = max_y - min_y
    margin = max(width, height) * 0.05
    legend_x = max_x + margin
    legend_y = max_y
    text_height = max(0.5, max(width, height) * 0.004)

    lines = ['KNM_BIMK_MASTER_DATA  -  LEGEND', '']
    lines.append('Layers are grouped by theme. Each theme = one layer + one color.')
    lines.append('')
    lines.append('Theme layer                ACI   Description')
    lines.append('-' * 70)
    for theme, (color, desc) in THEMES.items():
        if theme_counts.get(theme, 0) == 0:
            continue
        lines.append(f'{theme:26s} {color:>3}   {desc}')
    lines.append('')
    lines.append('See KNM_BIMK_MASTER_DATA_LEGEND.md for the full source-layer mapping.')

    text = '\\P'.join(lines)

    if 'KNM_LEGEND' not in doc.layers:
        doc.layers.add('KNM_LEGEND', color=7)

    mtext = msp.add_mtext(text, dxfattribs={
        'layer': 'KNM_LEGEND',
        'char_height': text_height,
        'insert': (legend_x, legend_y),
    })
    mtext.dxf.attachment_point = 1  # top-left


def write_legend_md(theme_layer_map, theme_counts, dropped_dedup, sources, variant_label, out_path):
    """Write a full attachable report covering the layer cleanup."""
    from datetime import date

    total_entities = sum(theme_counts.values())
    used_themes = [t for t in THEMES if theme_counts.get(t, 0) > 0]

    lines = [
        f'# KNM_BIMK_MASTER_DATA_NTM10_{variant_label}_METERS{TAXONOMY_SUFFIX}',
        '## Layer cleanup report',
        '',
        f'_Generated {date.today().isoformat()} by `merge_master_lines.py`._',
        '',
        '## Background',
        '',
        'The master file uses **theme-based layer grouping** aligned with',
        '**NS 3451:2022 Bygningsdelstabell** where applicable:',
        '',
        '- Each entity is reassigned to a single English theme layer (e.g. `KNM_2_Building`,',
        '  `KNM_76_Roads_Plazas`, `KNM_77_Park_Green`).',
        '- NS 3451 codes are preserved in the theme name so the classification is traceable',
        '  to the standard. Non-NS themes (`KNM_Topography_*`, `KNM_Regulatory`,',
        '  `KNM_Internal_*`) cover data that falls outside NS 3451\'s scope.',
        '- Layer color is derived from the theme, not from the source file.',
        '- Each source layer is also tagged with its **scope** (`BS` = visitor centre,',
        '  `KNM` = new museum, `TT` = The Twist, `Site` = site-wide / default) and its',
        '  **status** (`New` / `Existing`) where inferable from layer naming.',
        '- A short legend block is embedded as MTEXT in the upper-right corner of the master DXF/DWG.',
        '- This document is the full mapping legend; attach it alongside the deliverable.',
        '',
        '## Summary',
        '',
        f'- **Variant:** {variant_label} (NTM10, meters)',
        f'- **Source drawings merged:** {len(sources)}',
        f'- **Theme layers used:** {len(used_themes)} (down from ~4,500 raw layers)',
        f'- **Total entities in master:** {total_entities:,}',
        f'- **Duplicate marker entities removed:** {sum(dropped_dedup.values()):,}',
        '',
        '## Theme legend',
        '',
        'Each theme has one color and one purpose. Toggle a single layer to show/hide all',
        'features of that kind across the entire site. The **New** / **Existing** columns',
        'split each theme by status inferred from source layer naming ('
        '`(Nye konstruksjoner)`, `Eks`, `Innmålt`, `Kulturminne` = existing, etc.). ',
        '**?** = status not inferable from name.',
        '',
        '| Master layer | ACI | Description | Total | New | Existing | ? |',
        '|---|---|---|---|---|---|---|',
    ]
    # Pre-compute per-theme status counts from theme_layer_map
    status_by_theme = {t: {'New': 0, 'Existing': 0, '': 0} for t in THEMES}
    for theme in THEMES:
        for (orig_layer, src), cnt in theme_layer_map.get(theme, {}).items():
            st = infer_status(orig_layer, theme, src)
            status_by_theme[theme][st] += cnt

    for theme, (color, desc) in THEMES.items():
        n = theme_counts.get(theme, 0)
        if n == 0:
            continue
        sc = status_by_theme[theme]
        lines.append(
            f'| `{theme}` | {color} | {desc} | {n:,} | {sc["New"]:,} | {sc["Existing"]:,} | {sc[""]:,} |'
        )

    lines.extend([
        '',
        '## Source layer mapping',
        '',
        'For each theme, the table below lists the original consultant layer names that were',
        'merged into it, along with the source drawing each came from. Use this to verify the',
        'cleanup put the right content in the right theme.',
        '',
    ])

    for theme, (color, desc) in THEMES.items():
        n = theme_counts.get(theme, 0)
        if n == 0:
            continue
        # group entries by source for readability
        by_source = {}
        for (orig_layer, src), cnt in theme_layer_map[theme].items():
            by_source.setdefault(src, []).append((orig_layer, cnt))

        lines.append(f'### `{theme}` (ACI {color}) — {n:,} entities')
        lines.append('')
        lines.append(desc)
        lines.append('')
        if not by_source:
            lines.append('_No source layers contributed._')
        else:
            for src in sorted(by_source):
                src_layers = sorted(by_source[src])
                src_total = sum(c for _, c in src_layers)
                src_scope = scope_for(src)
                lines.append(f'**[{src_scope}] from `{src}`** ({src_total:,} entities, {len(src_layers)} layer{"s" if len(src_layers)!=1 else ""})')
                lines.append('')
                def _fmt(ol, c, _src=src):
                    status = infer_status(ol, theme, _src)
                    tag = f'[{status}] ' if status else ''
                    return f'- {tag}`{ol}` &nbsp; ({c:,})'
                # Truncate huge layer lists (Nairy art installation has ~1400)
                if len(src_layers) > 25:
                    for ol, c in src_layers[:20]:
                        lines.append(_fmt(ol, c))
                    lines.append(f'- _... +{len(src_layers)-20} more layers (full list omitted for brevity)_')
                else:
                    for ol, c in src_layers:
                        lines.append(_fmt(ol, c))
                lines.append('')
        lines.append('')

    lines.extend([
        '## Source drawings',
        '',
        'The following consultant drawings were merged. Each is the NTM10 transform of the',
        'original DWG/DXF received from the consultant.',
        '',
    ])
    for s in sorted(sources):
        lines.append(f'- [{scope_for(s)}] `{s}`')
    lines.append('')

    lines.extend([
        '## Notes for the architect',
        '',
        '- The master is **lines/points/text only** — no hatches, no block inserts. Hatches and',
        '  blocks from the source drawings are stripped during merge.',
        '- The `KNM_LEGEND` layer in the DXF holds the embedded legend MTEXT block. Toggle it off',
        '  if it gets in the way of model views.',
        '- Layer colors use ACI 1–8 for primary themes (fully visible in Dalux) and higher ACI',
        '  values for secondary themes.',
        '- Where multiple consultants drew features in the same area (e.g. contour lines from',
        '  both `Kotelinjer_20cm` and `landscape plan visitor centre`), some geometric overlap',
        '  may exist on the merged theme layer. We chose to preserve all source content rather',
        '  than risk dropping unique parts of the site.',
        '- The **`KNM_Other`** theme is the unmapped bucket — any layer whose name does not',
        '  match the NS 3451 prefix rules or descriptive keyword rules lands here for human',
        '  review. Ideally this should be small; a large `KNM_Other` means the rules need more',
        '  keyword mappings added.',
        '- If a theme color or grouping needs to change, edit `THEMES` or `LAYER_RULES` in',
        '  `merge_master_lines.py` and re-run the script. Theme names follow NS 3451:2022',
        '  Bygningsdelstabell codes (2, 70–79) where applicable.',
        '',
        '---',
        '',
        f'_Report generated automatically. {total_entities:,} entities, {len(used_themes)} themes._',
    ])

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'Report MD: {out_path}')


def scale_to_mm(meters_path, mm_path):
    """Scale a meters DXF x1000 to produce an mm version."""
    doc = ezdxf.readfile(meters_path)
    msp = doc.modelspace()
    s = 1000.0
    for e in msp:
        scale_entity(e, s)
    doc.header['$INSUNITS'] = 4  # millimeters
    doc.saveas(mm_path)
    print(f'Saved mm: {mm_path}')
    return mm_path


def convert_to_dwg(dxf_path, dwg_dir):
    os.makedirs(dwg_dir, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmpdir:
        shutil.copy2(dxf_path, tmpdir)
        subprocess.run([ODA, tmpdir, dwg_dir, 'ACAD2018', 'DWG', '0', '1'],
                      capture_output=True, timeout=120)
    name = os.path.splitext(os.path.basename(dxf_path))[0] + '.dwg'
    dwg_path = os.path.join(dwg_dir, name)
    if os.path.exists(dwg_path):
        print(f'DWG: {dwg_path}')
    else:
        print(f'DWG conversion check: {dwg_dir}')


def write_report_pair(theme_map, theme_counts, dropped_dedup, sources_short, dxf_dir, dwg_dir, label):
    """Write the .md + .pdf legend report and copy alongside DWG."""
    report_md = os.path.join(dxf_dir, f'KNM_BIMK_MASTER_DATA_NTM10_{label}_METERS{TAXONOMY_SUFFIX}_REPORT.md')
    write_legend_md(theme_map, theme_counts, dropped_dedup, sources_short, label, report_md)
    shutil.copy2(report_md, os.path.join(dwg_dir, os.path.basename(report_md)))
    try:
        from md_to_pdf import md_to_pdf
        pdf_path = report_md.replace('.md', '.pdf')
        md_to_pdf(report_md, pdf_path)
        shutil.copy2(pdf_path, os.path.join(dwg_dir, os.path.basename(pdf_path)))
    except Exception as ex:
        print(f'PDF generation skipped: {ex}')


def run_pipeline(dxf_global, dwg_global, dxf_local, dwg_local):
    """Master pipeline.

    GLOBAL meters is the source of truth (merged from consultant drawings).
    LOCAL meters, GLOBAL mm, LOCAL mm are all derived from it without re-merging:
      - LOCAL meters = GLOBAL meters offset by -BASEPOINT
      - GLOBAL mm    = GLOBAL meters x1000
      - LOCAL mm     = LOCAL meters x1000
    """
    print('=== MERGE GLOBAL METERS (canonical master) ===')
    files = collect_files(dxf_global)
    print(f'Files: {len(files)}')
    for _, _, short in files:
        print(f'  - {short}')

    global_m = os.path.join(dxf_global, _master_name('GLOBAL_METERS', 'dxf'))
    global_m, theme_map, theme_counts, dropped_dedup = merge_files(files, global_m)
    sources_short = [s for _, _, s in files]

    print('\n=== DERIVE LOCAL METERS (offset by basepoint) ===')
    local_m = os.path.join(dxf_local, _master_name('LOCAL_METERS', 'dxf'))
    derive_local_from_global(global_m, local_m)

    print('\n=== DERIVE GLOBAL MM (x1000) ===')
    global_mm = os.path.join(dxf_global, _master_name('GLOBAL_MM', 'dxf'))
    scale_to_mm(global_m, global_mm)

    print('\n=== DERIVE LOCAL MM (x1000 from local meters) ===')
    local_mm = os.path.join(dxf_local, _master_name('LOCAL_MM', 'dxf'))
    scale_to_mm(local_m, local_mm)

    print('\n=== DWG conversion ===')
    convert_to_dwg(global_m,  dwg_global)
    convert_to_dwg(global_mm, dwg_global)
    convert_to_dwg(local_m,   dwg_local)
    convert_to_dwg(local_mm,  dwg_local)

    print('\n=== Reports ===')
    write_report_pair(theme_map, theme_counts, dropped_dedup, sources_short, dxf_global, dwg_global, 'GLOBAL')
    write_report_pair(theme_map, theme_counts, dropped_dedup, sources_short, dxf_local,  dwg_local,  'LOCAL')


if __name__ == '__main__':
    ACC = 'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/10016 - Kistefos'
    run_pipeline(
        dxf_global=f'{ACC}/03_Ut/08_Tegninger/DXF_NTM/Global',
        dwg_global=f'{ACC}/03_Ut/08_Tegninger/DWG_NTM/Global',
        dxf_local=f'{ACC}/03_Ut/08_Tegninger/DXF_NTM/Lokal',
        dwg_local=f'{ACC}/03_Ut/08_Tegninger/DWG_NTM/Lokal',
    )
    print('\nDone.')
