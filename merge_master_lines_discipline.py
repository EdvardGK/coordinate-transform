"""Run the merge_master_lines pipeline with the v1 discipline-based taxonomy.

Kept in parallel with the NS 3451 taxonomy in `merge_master_lines.py` so that
both master deliverables can be produced side-by-side while the team decides
which taxonomy best fits the project. Output files get a `_discipline` suffix.

The discipline-based taxonomy uses Norwegian consulting disciplines as the
top-level grouping (ARK / LARK / RIVeg / RIVA / Survey / Municipality / INT).
It is the "previous" version that was in production before the 2026-04-13
NS 3451 rewrite and is preserved as a point-in-time snapshot.
"""
import os
import re
import unicodedata

os.environ.setdefault('KNM_TAXONOMY_SUFFIX', '_discipline')

import merge_master_lines as mml  # noqa: E402  (env var must be set first)


# === THEMES === discipline-based (v1) ======================================
mml.THEMES = {
    # --- ARK: Architecture ---
    'KNM_ARK_Buildings':            (1,   'Buildings, rooflines, walls, outdoor structures'),
    # --- LARK: Landscape architecture ---
    'KNM_LARK_Trees_Existing':      (3,   'Trees: existing, surveyed, dead/dry'),
    'KNM_LARK_Trees_New':           (84,  'Trees: new, planted, transplanted'),
    'KNM_LARK_Paths_Existing':      (174, 'Paths: existing pedestrian/cycle, fences'),
    'KNM_LARK_Paths_New':           (94,  'Paths: new/designed pedestrian, sidewalks'),
    'KNM_LARK_Vegetation':          (92,  'Non-tree vegetation: lawn, grass, hedges, soil'),
    'KNM_LARK_Forecourt':              (5,   'Forecourts, entry areas, hardscape graphics'),
    'KNM_LARK_Site_Furniture':      (230, 'Outdoor furniture, equipment, playground installations'),
    'KNM_LARK_Parking_New':         (40,  'Parking: new/designed spaces and surfaces'),
    # --- RIVeg: Road engineering ---
    'KNM_RIVeg_Roads_Existing':     (6,   'Roads: existing vehicular (surfaces, edges, ditches)'),
    'KNM_RIVeg_Roads_New':          (56,  'Roads: new/designed vehicular'),
    'KNM_RIVeg_Earthworks':         (4,   'Road engineering terrain (Veg_TIN)'),
    # --- RIVA: Water/sewer infrastructure ---
    'KNM_RIVA_Water':               (34,  'VA pipes, fire fighting, culverts'),
    # --- Survey: Survey/measurement ---
    'KNM_Survey_Electric':          (200, 'Electrical survey (markers, masts)'),
    'KNM_Survey_Art':               (220, 'Art installation survey (Nairy Baghramian)'),
    # --- Municipality: Municipal / kartdata ---
    'KNM_Municipality_Contours':    (2,   'Contour lines (existing + new, 10/20 cm, elevation text)'),
    'KNM_Municipality_Water':       (140, 'Natural water: rivers, lakes, dams, flood lines'),
    'KNM_Municipality_Heritage':    (54,  'Cultural heritage sites (fredning/preservation)'),
    'KNM_Municipality_Borders':     (9,   'Regulatory/administrative borders, zoning, construction zones'),
    # --- INT: Internal / project coordination ---
    'KNM_INT_Markers':              (250, 'Coordination markers, defpoints, helpers'),
    'KNM_INT_Text':                 (7,   'Text, dimensions, annotations'),
    # --- Catch-all ---
    'KNM_OTHER':                    (8,   'Unmatched / miscellaneous'),
}


# === LAYER RULES === discipline-based (v1) =================================
_DISCIPLINE_RULES = [
    # Ambiguous / fallthrough
    (r'^0$',                                 None),
    (r'^Diverse$',                           None),
    (r'^Layer\s*\d+$',                       None),

    # INT: Internal / project coordination
    (r'^COORDINATION_MARKER',                'KNM_INT_Markers'),
    (r'^Defpoints$',                         'KNM_INT_Markers'),
    (r'^Sporing',                            'KNM_INT_Markers'),
    (r'^00_TRACET',                          'KNM_INT_Markers'),
    (r'^BeskrivendeHjelpe',                  'KNM_INT_Markers'),
    (r'^Hjelpelinje',                        'KNM_INT_Markers'),
    (r'^827-\s*Snittlinjer',                 'KNM_INT_Markers'),
    # 837- Planteplan trær routed to Trees_New (before generic 83X text rule)
    (r'^837-.*Planteplan.*tr',               'KNM_LARK_Trees_New'),
    (r'^83[0-9]?-',                          'KNM_INT_Text'),
    (r'^847-',                               'KNM_INT_Text'),
    (r'\bTekst\b',                           'KNM_INT_Text'),

    # REG (Municipality Borders)
    (r'^TiltakGrense',                       'KNM_Municipality_Borders'),
    (r'^PblTiltak',                          'KNM_Municipality_Borders'),
    (r'^Arealbrukgrense',                    'KNM_Municipality_Borders'),
    (r'^Arealgrense',                        'KNM_Municipality_Borders'),
    (r'^Anleg[gs]*omr.de',                   'KNM_Municipality_Borders'),
    (r'^PLANKART',                           'KNM_Municipality_Borders'),
    (r'regulering',                          'KNM_Municipality_Borders'),
    (r'^Kulturminne',                        'KNM_Municipality_Heritage'),

    # Contours
    (r'^70[12]-',                            'KNM_Municipality_Contours'),
    (r'^71[0-7]-',                           'KNM_Municipality_Contours'),
    (r'^900-\s*RiVei\s*koter',               'KNM_Municipality_Contours'),
    (r'Kotelinjer',                          'KNM_Municipality_Contours'),
    (r'\bkote\b',                            'KNM_Municipality_Contours'),
    (r'^PresH.ydetall',                      'KNM_Municipality_Contours'),
    (r'^Forsenkningskurve',                  'KNM_Municipality_Contours'),

    # Natural water
    (r'^Dam(kant)?$',                        'KNM_Municipality_Water'),
    (r'^ElvBekk',                            'KNM_Municipality_Water'),
    (r'^Floml.pkant',                        'KNM_Municipality_Water'),
    (r'^ElvelinjeFiktiv',                    'KNM_Municipality_Water'),
    (r'^Innsj.',                             'KNM_Municipality_Water'),
    (r'^Kanal(Gr.ft)?$',                     'KNM_Municipality_Water'),
    (r'^Fisketrapp',                         'KNM_Municipality_Water'),

    # Trees
    (r'772-.*\b(Eks|eks)\s*tre',             'KNM_LARK_Trees_Existing'),
    (r'^T.rr?(Lauvtre|Gran|Furu|Bj.rk|Eik|Hegg|Rogn)', 'KNM_LARK_Trees_Existing'),
    (r'^(Lauvtre|Bj.rk|Furu|Eik|Gran|Hegg|Rogn|L.nn|Or)[\b_-]', 'KNM_LARK_Trees_Existing'),
    (r'^(Lauvtre|Bj.rk|Furu|Eik|Gran|Hegg|Rogn|L.nn|Or)$', 'KNM_LARK_Trees_Existing'),
    (r'^Tre(Stamme|Krone|Punkt)',            'KNM_LARK_Trees_Existing'),
    (r'773-.*\b(Nytt|nytt)\s*tre',           'KNM_LARK_Trees_New'),
    (r'transplantert',                       'KNM_LARK_Trees_New'),

    # Paths
    (r'^AnnetGjerde',                        'KNM_LARK_Paths_Existing'),
    (r'^GangSykkelveg',                      'KNM_LARK_Paths_Existing'),
    (r'^Gangvegkant',                        'KNM_LARK_Paths_Existing'),
    (r'^Fortauskant',                        'KNM_LARK_Paths_Existing'),
    (r'^00-?\s*nye?\s*stier',                'KNM_LARK_Paths_New'),
    (r'^00\s*Stier',                         'KNM_LARK_Paths_New'),
    (r'^00-?\s*bes',                         'KNM_LARK_Paths_New'),
    (r'^f-64000',                            'KNM_LARK_Paths_New'),

    # Vegetation
    (r'^771-',                               'KNM_LARK_Vegetation'),
    (r'^00-?\s*ny\s*jord',                   'KNM_LARK_Vegetation'),
    (r'^00-?forflytning\s*jord',             'KNM_LARK_Vegetation'),
    (r'^Arealressurs',                       'KNM_LARK_Vegetation'),
    (r'^Hekk',                               'KNM_LARK_Vegetation'),
    (r'^Lekeplass',                          'KNM_LARK_Vegetation'),
    (r'^777\s*Skj.tsel',                     'KNM_LARK_Vegetation'),

    # Plazas
    (r'^762-',                               'KNM_LARK_Forecourt'),
    # Site furniture
    (r'^790-',                               'KNM_LARK_Site_Furniture'),

    # ARK: Architecture
    (r'^Bygning',                            'KNM_ARK_Buildings'),
    (r'^AnnenBygning',                       'KNM_ARK_Buildings'),
    (r'^Bru\b',                              'KNM_ARK_Buildings'),
    (r'^Bruavgrensning',                     'KNM_ARK_Buildings'),
    (r'^Veranda',                            'KNM_ARK_Buildings'),
    (r'^00-?\s*ARK',                         'KNM_ARK_Buildings'),
    (r'^00-?\s*Hotlink\s*Bygg',              'KNM_ARK_Buildings'),
    (r'^Bygningsdelelinje',                  'KNM_ARK_Buildings'),
    (r'^Bygningslinje',                      'KNM_ARK_Buildings'),
    (r'^Takkant',                            'KNM_ARK_Buildings'),
    (r'^M.nelinje',                          'KNM_ARK_Buildings'),
    (r'^Taksprang',                          'KNM_ARK_Buildings'),
    (r'^TakoverbyggKant',                    'KNM_ARK_Buildings'),
    (r'^Takoverbygg',                        'KNM_ARK_Buildings'),
    (r'^721-',                               'KNM_ARK_Buildings'),
    (r'^72--?\s*Utend',                      'KNM_ARK_Buildings'),
    (r'^Mur\b',                              'KNM_ARK_Buildings'),
    (r'^MurLoddrett',                        'KNM_ARK_Buildings'),
    (r'^00-?\s*Stein\s*og\s*fjell',          'KNM_ARK_Buildings'),
    (r'^Fritt(st.ende)?Trapp',               'KNM_ARK_Buildings'),
    (r'^Fundament',                          'KNM_ARK_Buildings'),
    (r'^Pipe(kant)?$',                       'KNM_ARK_Buildings'),
    (r'^L.vebru',                            'KNM_ARK_Buildings'),
    (r'^TrappBygg',                          'KNM_ARK_Buildings'),
    (r'^Skr.Forst.tning',                    'KNM_ARK_Buildings'),
    (r'^BautaStatue',                        'KNM_ARK_Buildings'),

    # Roads
    (r'^Veg(dekkekant|kant|rekkverk|bom|gr.ft)', 'KNM_RIVeg_Roads_Existing'),
    (r'^Veg$',                               'KNM_RIVeg_Roads_Existing'),
    (r'^Vegskulderkant',                     'KNM_RIVeg_Roads_Existing'),
    (r'^AnnetVegareal',                      'KNM_RIVeg_Roads_Existing'),
    (r'^Kj.rebane',                          'KNM_RIVeg_Roads_Existing'),
    (r'^Spormidt',                           'KNM_RIVeg_Roads_Existing'),
    (r'^761-',                               'KNM_RIVeg_Roads_New'),
    (r'^f-veg_',                             'KNM_RIVeg_Roads_New'),

    # Parking
    (r'^_p-plass',                           'KNM_LARK_Parking_New'),
    (r'^Parkering',                          'KNM_LARK_Parking_New'),
    (r'^00-?\s*(fase\s*\d+\s*)?parkering',   'KNM_LARK_Parking_New'),

    # Veg_TIN earthworks
    (r'Veg_TIN',                             'KNM_RIVeg_Earthworks'),

    # Electric
    (r'^Masteomriss',                        'KNM_Survey_Electric'),

    # Water / VA
    (r'^731-',                               'KNM_RIVA_Water'),
    (r'^733-',                               'KNM_RIVA_Water'),
    (r'^Stikkrenne',                         'KNM_RIVA_Water'),
    (r'^R.rgate',                            'KNM_RIVA_Water'),
    (r'^KaiBrygge',                          'KNM_RIVA_Water'),
    (r'^Sv.mmebasseng',                      'KNM_RIVA_Water'),

    # Art survey (must come before generic Innmålt)
    (r'^Innm.lt\s*kunstinstall',             'KNM_Survey_Art'),
    (r'^Innm.lt\b',                          'KNM_Survey_Electric'),
]
mml.LAYER_RULES_COMPILED = [(re.compile(p, re.IGNORECASE), t) for p, t in _DISCIPLINE_RULES]


# === SOURCE FALLBACK ===
mml.SOURCE_FALLBACK = {unicodedata.normalize('NFC', k): v for k, v in {
    '240806_Nairy Baghramian 1 1':    'KNM_OTHER',
    'landscape plan visitor centre':  'KNM_OTHER',
    'KNM_Stier':                      'KNM_LARK_Paths_New',
    'Kotelinjer_20cm':                'KNM_Municipality_Contours',
    'Kotelinjer_MedPåskrift_2D':      'KNM_Municipality_Contours',
    '230831_Veg_TIN':                 'KNM_RIVeg_Earthworks',
    'Parkering_kistefos_LARK 1':      'KNM_LARK_Parking_New',
}.items()}


# === SOURCE PRIORITY ===
mml.SOURCE_PRIORITY = {k: [unicodedata.normalize('NFC', s) for s in v] for k, v in {
    'KNM_INT_Markers': ['landscape plan visitor centre'],
    'KNM_Municipality_Contours': ['Kotelinjer_MedPåskrift_2D'],
}.items()}


if __name__ == '__main__':
    ACC = 'C:/Users/edkjo/DC/ACCDocs/Skiplum AS/Skiplum Backup/Project Files/10016 - Kistefos'
    mml.run_pipeline(
        dxf_global=f'{ACC}/03_Ut/08_Tegninger/DXF_NTM/Global',
        dwg_global=f'{ACC}/03_Ut/08_Tegninger/DWG_NTM/Global',
        dxf_local=f'{ACC}/03_Ut/08_Tegninger/DXF_NTM/Lokal',
        dwg_local=f'{ACC}/03_Ut/08_Tegninger/DWG_NTM/Lokal',
    )
    print('\nDone.')
