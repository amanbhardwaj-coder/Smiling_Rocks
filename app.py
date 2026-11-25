#!/usr/bin/env python3
"""
MFG → VDB Converter — single-file consolidated app.

Two ways to run:

  Streamlit UI (interactive, with file uploads):
      streamlit run app.py

  Command-line (batch / automation):
      python app.py INPUT.xlsx [INPUT2.xlsx ...] -i IMAGES.csv -o OUTPUT.csv -c Essential

The conversion logic, the Streamlit UI, and the CLI all live here so you
can deploy a single file. The library functions (`convert`, `load_image_map`)
accept either filesystem paths or file-like objects (BytesIO, Streamlit
UploadedFile) — same code path for both contexts.
"""
import argparse
import io
import re
import sys
import zipfile
import shutil
import tempfile
from decimal import Decimal, ROUND_HALF_UP

from collection_mappings import COLLECTION_GROUPS, CLASSIFICATIONS, BASE_CLASSIFICATIONS

import pandas as pd
from openpyxl import load_workbook


# ============================================================================
# CONFIG
# ============================================================================
DEFAULT_COLLECTION = "Auto"                  # Resolve collection/style from source and reference
IMG_BASE     = "https://d2dtfeai6qg5ne.cloudfront.net/Smiling_Rocks/"
MAX_IMAGES   = 8                              # Image Url 1..8
SIDE_COLOR, SIDE_CLARITY = "E-F", "SI1"       # 29-May convention
CEN_COLOR,  CEN_CLARITY  = "E-F", "SI1"       # 29-May convention

# Curated remark → (sub-category rule, Config Field 14, Jewelry Style)
# subcat_rule == "jtype" means: use the MFG jewelry type (Ring/Earring/...) as Sub-Category.
# Anything else is used literally for Sub-Category.
REMARK_MAP = {
    "3 STONE":                    ("jtype", "Three Stone",        "Three Stone"),
    "3 STONE BAZEL":              ("jtype", "Three Stone",        "Three Stone"),
    "BAZEL STUD":                 ("jtype", "Solitaire Stud",     "Solitaire Stud"),
    "BEZEL SOLITAIRE":            ("jtype", "Solitaire",          "Solitaire"),
    "CHANNEL FULL ETERNITY BAND": ("jtype", "Full Eternity Band", "Full Eternity Band"),
    "CHANNEL HOOPS":              ("jtype", "In and Out Hoops",   "In and Out Hoops"),
    "CHENNAL SETTING BAND":       ("Band",  "Band",               "Chennal Setting Band"),
    "ENGAGEMENT":                 ("jtype", "Engagement Ring",    "Engagement Ring"),
    "ESSENTIALS":                 ("jtype", "Fashion",            "Fashion"),
    "FLOWER EARRINGS":            ("jtype", "Flower Earring",     "Flower Earring"),
    "HALO RINGS":                 ("jtype", "Halo",               "Halo"),
    "HOOP EARRINGS":              ("jtype", "In and Out Hoops",   "In and Out Hoops"),
    "IN OUT HOOP EARRINGS":       ("jtype", "In and Out Hoops",   "In and Out Hoops"),
    "SOLITAIRE":                  ("jtype", "Solitaire",          "Solitaire"),
    "SOLITAIRE PENDANTS":         ("jtype", "Solitaire Pendant",  "Solitaire Pendant"),
    "SOLITAIRE STUDS":            ("jtype", "Solitaire Stud",     "Solitaire Stud"),
}

SHAPE_MAP = {
    'BAGUETTE':'Baguette','BAGUTTE':'Baguette','BAGUET':'Baguette','BGT':'Baguette','BAG':'Baguette','BUG':'Baguette',
    'RND':'Round','ROUND':'Round','OVL':'Oval','OVAL':'Oval','EMR':'Emerald','EMERALD':'Emerald',
    'PER':'Pear','PEAR':'Pear','CUS':'Cushion','CUSHION':'Cushion','PRN':'Princess','PRINCESS':'Princess',
    'RAD':'Radiant','RADIANT':'Radiant','MAR':'Marquise','MQS':'Marquise','MARQUISE':'Marquise',
    'TRI':'Triangle','ASR':'Asscher','TAP':'Taper','HMN':'Half Moon','HRT':'Heart','HEART':'Heart',
    'MRQ':'Marquise','ELCUS':'Cushion Long','TRZ':'Trapezoid','MIX':'Mixed',
    'KTE':'Kite','KT':'Kite','KITE':'Kite',
    'OCT':'Octagonal','OCTAGON':'Octagonal','OCTAGONAL':'Octagonal',
    'ROU':'Round','OV':'Oval','EM':'Emerald','EME':'Emerald','EC':'Emerald',
    'CU':'Cushion','CUSH':'Cushion','PRS':'Princess','PRIN':'Princess','PC':'Princess',
    'MQ':'Marquise','MARQ':'Marquise','HT':'Heart','ASS':'Asscher','ASC':'Asscher',
    'TRIL':'Trilliant','TRL':'Trilliant','TRILLIANT':'Trilliant',
    'BQ':'Baguette','BAGT':'Baguette','TB':'Tapered Baguette','TAPB':'Tapered Baguette',
}
COLOR_MAP = {
    'WHSR':'White','WHT':'White','WHITE':'White',
    'BRSR':'Brown',
    'BLSR':'Blue','BSSR':'Blue','BLU':'Blue','BLUE':'Blue',
    'PKSR':'Pink','PNK':'Pink','PINK':'Pink',
    'GNSR':'Green','GRSR':'Green','GSR':'Green','TSSR':'Green','GRN':'Green','GREEN':'Green',
    'RDSR':'Red','RDS':'Red','RD':'Red','RED':'Red',
    'YLSR':'Yellow','YEL':'Yellow','YELLOW':'Yellow',
    'RUSR':'Ruby','EMSR':'Emerald','RNBW':'Rainbow',
}
COLOR_NAME = {'W':'White','Y':'Yellow','R':'Rose'}

# Exact 41-column VDB schema (order matters).
COLUMNS = [
    'Master Stock','Stock Number','Jewelry Sub-Category','Jewelry Style','Config Field 13',
    'Config Field 14','Tags','Short Title','Description','Metal',
    'Image Url 1','Image Url 2','Image Url 3','Image Url 4','Image Url 5','Image Url 6','Image Url 7','Image Url 8',
    'Available Metals','Config Field 17','Total Price','Weight','Config field 11','Available Config field 11',
    'Center Carat Weight','Available Center Carat Weight','Center Color','Center Clarity','Shape','Supported Shape Variations',
    'Side Carat Weight','Side Total Stones','Side Color','Side Clarity','Side Stone Shape',
    'Config Field 6','Available Config Field 6','Comments','barcode No','Available Config Field 7','Available Ring Size',
]


# ============================================================================
# HELPERS
# ============================================================================
def _txt(x):
    if x is None: return ''
    if isinstance(x, float) and x.is_integer(): return str(int(x))
    return str(x).strip()

def _num(x):
    s = _txt(x).replace(',', '')
    try: return float(s)
    except Exception: return 0.0

def _toks(s):
    return re.findall(r'[A-Z0-9]+', _txt(s).upper())

def _r2(v):
    """Round to 2 decimals, half-away-from-zero (spreadsheet convention)."""
    return float(Decimal(str(float(v))).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))

def _fmt(v):
    v = round(float(v), 4)
    if v == 0: return ''
    if float(v).is_integer(): return str(int(v))
    return f"{v:.4f}".rstrip('0').rstrip('.')

def _fmt2(v):
    v = _r2(v)
    if v == 0: return ''
    if float(v).is_integer(): return str(int(v))
    return f"{v:.2f}".rstrip('0').rstrip('.')

def _fmt2_fixed(v):
    """Fixed 2dp, keeps trailing zero (0.5 → '0.50'). Empty for zero."""
    v = _r2(v)
    if v == 0: return ''
    return f"{v:.2f}"


MAIN_COLLECTIONS = ('Essentials', 'Bridals', 'Haute Couture', 'Fashion Collections', "Men's")


def _category_key(value):
    key = re.sub(r'[^a-z0-9]+', '', _txt(value).lower())
    return {'essentials': 'essential', 'fashionbridal': 'fashionbridals',
            'mens': 'men', 'men': 'men', 'stairynight': 'thestarrynight'}.get(key, key)


def _classification_tags(group, collection, style, cf14):
    if group == 'Essentials':
        values = ['Essential', cf14 if cf14 != 'Fashion' else '']
    elif group == 'Bridals':
        values = ['Bridal', collection, cf14]
    elif group == 'Fashion Collections':
        values = ['Fashion', collection if collection != 'Fashion' else style]
    elif group == "Men's":
        values = [collection, 'Men']
    else:
        values = [collection, style]
    return ','.join(dict.fromkeys(v for v in values if v))


def resolve_classification(item, selection=DEFAULT_COLLECTION):
    """Resolve reference tuples without flattening collection/style/filter fields.

    Explicit collection/style context disambiguates existing SKU mappings.
    New items use the same collection rules. Conflicts return a review message
    rather than choosing the most frequent historical mapping.
    """
    selected = _category_key(selection)
    main_aliases = {'essential': 'Essentials', 'bridal': 'Bridals', 'bridals': 'Bridals',
                    'hautecouture': 'Haute Couture', 'fashioncollections': 'Fashion Collections',
                    'men': "Men's"}
    main = main_aliases.get(selected)
    names = {_category_key(c): c for c in COLLECTION_GROUPS}
    explicit_collection = names.get(selected) if not main else None
    source_collection = names.get(_category_key(item.get('source_collection', '')))
    remark = _txt(item.get('remark', ''))
    remark_key = _category_key(remark)
    explicit_style = _txt(item.get('source_style', ''))
    style_key = _category_key(explicit_style or remark)
    # Preserve established legacy remark aliases such as SOLITAIRE STUDS.
    legacy = REMARK_MAP.get(re.sub(r'\s+', ' ', remark.upper()).strip())
    if legacy and not explicit_style:
        style_key = _category_key(legacy[2])
    if not main and selected in ('auto', ''):
        main = main_aliases.get(_category_key(item.get('source_collection', '')))
    remark_collection = names.get(remark_key)
    if remark_collection and main and COLLECTION_GROUPS[remark_collection] != main:
        remark_collection = None
    target = explicit_collection or source_collection or remark_collection

    def matches(row):
        c, st, cf = row
        return (not main or COLLECTION_GROUPS[c] == main) and (not target or c == target)

    candidates = [CLASSIFICATIONS[i] for i in BASE_CLASSIFICATIONS.get(item.get('base', ''), [])]
    candidates = [r for r in candidates if matches(r)]
    has_base_mapping = bool(candidates)
    if not candidates:
        candidates = [r for r in CLASSIFICATIONS if matches(r)]
    specific = [r for r in candidates if _category_key(r[1]) == style_key]
    if not specific:
        specific = [r for r in candidates if _category_key(r[2]) == style_key]
    if specific:
        candidates = specific
    elif explicit_style or (not has_base_mapping and remark and remark_key not in names and not main_aliases.get(remark_key)):
        candidates = []
    candidates = sorted(set(candidates))
    if len(candidates) == 1:
        c, style, cf14 = candidates[0]
        group = COLLECTION_GROUPS[c]
        return c, style, cf14, _classification_tags(group, c, style, cf14), ''

    # New fashion collections follow the named collection / Fashion pattern
    # only when the user explicitly selected that main group.
    if not candidates and main == 'Fashion Collections' and remark and not target:
        c = _txt(item.get('source_collection', '')) or remark
        style = explicit_style or remark
        return c, style, 'Fashion', _classification_tags(main, c, style, 'Fashion'), ''
    if legacy and (not main or main == 'Essentials') and (not target or target == 'Essential'):
        _, cf14, style = legacy
        return 'Essential', style, cf14, _classification_tags('Essentials', 'Essential', style, cf14), ''

    c = target or (selection if selected not in ('auto', '') and not main else '')
    style = explicit_style or remark
    # Preserve shared fields of a conflicting reference set, leaving only
    # the unresolved classification blank for review.
    if candidates:
        fields = [set(r[i] for r in candidates) for i in range(3)]
        c, style, cf14 = [next(iter(v)) if len(v) == 1 else '' for v in fields]
        issue = f"Ambiguous mapping: {item.get('base', '')} / {remark or selection}"
    else:
        cf14 = ''
        issue = f"Unmapped collection/style: {item.get('base', '')} / {remark or selection}"
    return c, style, cf14, '', issue


# ============================================================================
# IMAGE-LIST LOADER
# ============================================================================
# Accept color + index codes, including explicit combinations such as WY1.
# Anything else (WG/YG/RG metal codes, LS/LF loose-stone, digits-only,
# "W1 (1)") is skipped on purpose.
_IMG_CODE_RE = re.compile(r'^([WYR]{1,3})(\d*)$', re.I)

def load_image_map(path_or_file):
    """Read a one-column image-filename list and build {(base, colorLetter): [urls by index]}.

    Filenames look like `SRBA-02008WHT_W2.jpg` or `SR-E04139WHT_Y1.jpg`
    or unnumbered two tone files such as `SRNL-09626BRN_WY.jpg`.
    The CSV may have a header or start directly with a filename.
    (split on the LAST '_'). Returns (image_map, stats) where stats includes
    accepted/skipped counts and a sample of skipped codes for diagnostics.
    """
    if not path_or_file:
        return {}, {'accepted': 0, 'skipped': 0, 'skip_samples': {}}
    df = pd.read_csv(path_or_file, dtype=str, header=None).fillna('')
    col = df.columns[0]
    tmp = {}
    accepted = 0
    skipped = 0
    skip_samples = {}  # one example filename per skipped code
    for fn in df[col]:
        fn = fn.strip()
        if not fn or '_' not in fn:
            skipped += 1
            continue
        base, suffix = fn.rsplit('_', 1)              # 'SRBA-02008WHT', 'W2.jpg'
        code = suffix.split('.')[0].strip()           # 'W2'
        m = _IMG_CODE_RE.match(code)
        if not m:
            skipped += 1
            skip_samples.setdefault(code, fn)
            continue
        C = m.group(1).upper()
        idx = int(m.group(2) or 0)
        tmp.setdefault((base.upper(), C), []).append((idx, fn))
        accepted += 1
    out = {k: list(dict.fromkeys(IMG_BASE + fn for _, fn in sorted(v))) for k, v in tmp.items()}
    stats = {'accepted': accepted, 'skipped': skipped, 'skip_samples': skip_samples}
    return out, stats


# ============================================================================
# MFG PARSER — turns a workbook into a list of item dicts
# ============================================================================
def _norm_header(x):
    """Normalize a header cell for matching: 'D. Pcs' / 'D.Pcs' / ' D Pcs ' -> 'dpcs'."""
    return re.sub(r'[^a-z0-9]+', '', _txt(x).lower())

# Canonical field -> any of these normalized header strings.
# Add aliases here if a future MFG layout uses slightly different wording.
_HEADER_ALIASES = {
    'no':         ('no',),
    'style':      ('styleno',),
    'desc':       ('description',),
    'pcs':        ('dpcs', 'dpc'),
    'size':       ('dsize',),
    'ct':         ('gmcts', 'gmct'),
    'dc':         ('dc',),
    'srp':        ('srp',),
    '14kdc':      ('14kdc',),
    '18kdc':      ('18kdc',),
    '14ksrp':     ('14ksrp',),
    '18ksrp':     ('18ksrp',),
    'default_color': ('defaultmetalcolor', 'bydefaultmetalcolor', 'defaultmetal', 'defaultmetalmasterstock'),
    'variation':  ('variation', 'variations'),
    'remark':     ('remark',),
    'collection': ('collection', 'collectionname'),
    'source_style': ('jewelrystyle', 'jewelrystyles',),
}

def _find_header(rows, max_scan=80):
    """Scan the first `max_scan` rows for the header. Returns (header_row_idx, idx_map) or (None, None).

    Header row must contain 'Style No', 'Description', 'D. Pcs' (or 'D.Pcs'),
    and 'Gm/Cts' (or 'Gm.Cts'). Returns a dict mapping field -> column index.
    """
    for r in range(min(max_scan, len(rows))):
        normalized = [_norm_header(c) for c in rows[r]]
        # Required fields must all be present in this row
        if not all(any(a in normalized for a in _HEADER_ALIASES[f])
                   for f in ('style', 'desc', 'pcs', 'ct')):
            continue
        idx = {}
        for field, aliases in _HEADER_ALIASES.items():
            for a in aliases:
                if a in normalized:
                    idx[field] = normalized.index(a)
                    break
        return r, idx
    return None, None

def _metal_color(value):
    """Canonical stock suffix, including explicit two/three tone combinations."""
    text = _txt(value).upper()
    if not text:
        return 'W'
    compact = re.sub(r'[^A-Z]', '', text)
    aliases = {'WG': 'W', 'YG': 'Y', 'RG': 'R', 'W': 'W', 'Y': 'Y', 'R': 'R'}
    if compact in aliases:
        return aliases[compact]
    if re.fullmatch(r'[WYR]{2,3}', compact):
        return ''.join(dict.fromkeys(compact))
    colors = re.findall(r'WHITE|YELLOW|ROSE|WG|YG|RG', text)
    if colors:
        return ''.join(dict.fromkeys(c[0] for c in colors))
    raise ValueError(f"Unrecognized DEFAULT METAL COLOR {value!r}; specify colors (e.g. WY for two tone).")


def _metal_name(purity, color):
    names = ' '.join(COLOR_NAME[c] for c in color)
    return f"{purity} {names} Gold"


def _item_colors(item):
    default = item.get('default_color', 'W')
    # An explicit combination represents the item's two/three tone metal.
    if len(default) > 1 or item.get('variation', '').upper() in ('NO', 'N', 'FALSE', '0'):
        return (default,)
    return ('W', 'Y', 'R')


def extract(path_or_file, diagnostics=None):
    """Parse one MFG .xlsx (path or file-like) into a list of item dicts.

    Auto-detects each sheet's header row and column positions so it works
    across MFG layout variants (extra leading column, renamed Remark vs
    collection column, etc.).
    """
    diagnostics = diagnostics if diagnostics is not None else []
    # Wrapped XLSX ZIPs can be large because of embedded photographs. Stream
    # the nested workbook to disk rather than holding a second copy in memory.
    with zipfile.ZipFile(path_or_file) as archive:
        if 'xl/workbook.xml' not in archive.namelist():
            names = [n for n in archive.namelist()
                     if n.lower().endswith('.xlsx') and not n.startswith('__MACOSX/')]
            if not names:
                raise ValueError('ZIP contains no .xlsx workbook.')
            nested_items = []
            for name in names:
                with archive.open(name) as source, tempfile.TemporaryFile() as nested:
                    shutil.copyfileobj(source, nested)
                    nested.seek(0)
                    nested_items.extend(extract(nested, diagnostics))
            return nested_items
    wb = load_workbook(path_or_file, read_only=True, data_only=True)
    items = []
    for sheet_name in wb.sheetnames:
        # MFG workbooks may name sheets as `14`, `14 `, `14K`, `18`, `18K`, etc.
        # The earlier logic checked only for `18K`, so a sheet named just `18`
        # was incorrectly treated as 14K. Normalize first, then detect purity.
        sheet_key = re.sub(r'[^0-9A-Z]+', '', _txt(sheet_name).upper())
        if sheet_key.startswith('18'):
            purity = '18K'
        elif sheet_key.startswith('14'):
            purity = '14K'
        else:
            purity = '18K' if '18' in sheet_key else '14K'
        rows = [list(r) for r in wb[sheet_name].iter_rows(values_only=True)]

        header_row, idx = _find_header(rows)
        if header_row is None:
            continue  # sheet doesn't look like an MFG quotation, skip silently

        def cell(row, field):
            j = idx.get(field)
            return row[j] if (j is not None and j < len(row)) else None

        combined = all(f in idx for f in ('14kdc', '18kdc', '14ksrp', '18ksrp'))
        if ('14kdc' in idx or '18kdc' in idx) and not combined:
            raise ValueError(f'{sheet_name}: combined pricing requires 14KDC, 18KDC, 14KSRP and 18KSRP.')

        def finish(item):
            if not item:
                return
            if not item.get('base'):
                diagnostics.append({'severity': 'error', 'issue': 'missing_source_sku',
                                    'sheet': sheet_name, 'source_row': item.get('source_row'),
                                    'details': 'Numbered item block has no recognized SR stock number.'})
                return
            if combined:
                for karat in ('14K', '18K'):
                    items.append(dict(item, purity=karat, dc=item['prices'][karat][0],
                                      srp=item['prices'][karat][1]))
            else:
                items.append(item)

        cur = None
        for r in range(header_row + 1, len(rows)):
            row = rows[r]
            first = _txt(cell(row, 'no'))
            style = _txt(cell(row, 'style'))
            desc  = _txt(cell(row, 'desc'))

            # New item header: numbered row. Some MFG files put the real
            # SR stock number on this row (Pricing 2), while others put a
            # P.O. value here and the SR stock number 1-2 rows below
            # (Pricing 3). Start the item on the numbered row in both cases
            # and fill `base` as soon as an SR-style SKU appears.
            if first.replace('.0', '').isdigit():
                if cur:
                    # Only keep completed items where we found a real SR SKU.
                    finish(cur)
                m = re.search(r'\bSR[A-Z]+-\d+[A-Z0-9]*\b', style.upper())
                # Use Remark if populated, else fall back to collection column
                # (newer MFG files put the category there instead).
                remark_val     = _txt(cell(row, 'remark'))
                collection_val = _txt(cell(row, 'collection'))
                cur = {
                    'sheet': sheet_name, 'source_row': r + 1, 'purity': purity, 'combined': combined,
                    'default_color_raw': _txt(cell(row, 'default_color')),
                    'source_errors': [f'{f}: {_txt(cell(row, f))}' for f in
                                      ('collection', 'dc', 'srp', '14kdc', '18kdc', '14ksrp', '18ksrp')
                                      if _txt(cell(row, f)).startswith('#')],
                    'default_color': _metal_color(cell(row, 'default_color')),
                    'variation': _txt(cell(row, 'variation')),
                    'prices': {k: (_num(cell(row, k.lower() + 'dc')),
                                   _num(cell(row, k.lower() + 'srp')))
                               for k in ('14K', '18K')},
                    'base': (m.group(0) if m else ''),
                    'jtype': '',
                    'remark': remark_val or collection_val,
                    'source_collection': collection_val,
                    'source_style': _txt(cell(row, 'source_style')),
                    'price_raw': {f: _txt(cell(row, f)) for f in
                                  ('dc', 'srp', '14kdc', '18kdc', '14ksrp', '18ksrp')},
                    'dc':  _num(cell(row, 'dc')),
                    'srp': _num(cell(row, 'srp')),
                    'total_dia': 0.0,
                    'centers': [], 'sides': [], 'unrecognized_stones': [],
                }
            if cur is None:
                continue

            # In P.O.-style sheets, the real stock number is on a later detail row.
            if not cur.get('base'):
                m = re.search(r'\bSR[A-Z]+-\d+[A-Z0-9]*\b', style.upper())
                if m:
                    cur['base'] = m.group(0)

            # Jewelry type appears in the style col on a detail line
            for cand in (style, desc):
                if cand.upper() in ['RING', 'NECKLACE', 'EARRING', 'BRACELET', 'PENDANT', 'BAND']:
                    cur['jtype'] = cand.title()

            # Record unknown positive-weight stone rows instead of silently dropping them.
            pcs   = _num(cell(row, 'pcs'))
            carat = _num(cell(row, 'ct'))
            tk = _toks(desc)
            if (tk[:1] == ['CS'] or any(c in SHAPE_MAP for c in tk)) and pcs > 0 and carat > 0:
                rec = {
                    'shape': next((SHAPE_MAP[c] for c in tk if c in SHAPE_MAP), ''),
                    'color': next((COLOR_MAP[c] for c in tk if c in COLOR_MAP), ''),
                    'pcs': pcs, 'carat': carat, 'description': desc,
                }
                if not rec['color']:
                    cur['unrecognized_stones'].append({'description': desc, 'pcs': pcs,
                                                       'carat': carat, 'issue': 'unrecognized_color'})
                (cur['centers'] if tk[:1] == ['CS'] else cur['sides']).append(rec)

            elif pcs > 0 and carat > 0:
                cur['unrecognized_stones'].append({'description': desc, 'pcs': pcs, 'carat': carat})

            # Total Dia. CTS row (anywhere in the row)
            for j, c in enumerate(row):
                if c and 'Total Dia.' in str(c) and j + 1 < len(row):
                    cur['total_dia'] = _num(row[j + 1])

        if cur:
            finish(cur)
            cur = None
    wb.close()
    if not items:
        raise ValueError('No complete MFG items found. Check workbook headers and stock numbers.')
    return items


# ============================================================================
# COMMENT BUILDER
# ============================================================================
def _grouped(stones):
    agg, order = {}, []
    for s in stones:
        k = (s['shape'], s['color'])
        if k not in agg:
            agg[k] = {'pcs': 0.0, 'ct': 0.0}
            order.append(k)
        agg[k]['pcs']  += s['pcs']
        agg[k]['ct']   += s['carat']
    lines = []
    for (shape, color) in order:
        v = agg[(shape, color)]
        sc = f"{shape} {color}".strip()
        # White diamonds → (E-F/VS) ; colored gemstones → no quality grade
        quality = "(E-F/VS)" if color == "White" else ""
        line = f"{_fmt(v['pcs'])} Pcs | {_fmt2_fixed(v['ct'])} ct | {sc}"
        if quality:
            line += f" {quality}"
        lines.append(re.sub(r'\s+', ' ', line).strip())
    return lines

def _build_comment(centers, sides):
    """Headers when center stones exist; side-only items are headerless."""
    cl, sl = _grouped(centers), _grouped(sides)
    if centers:
        parts = []
        if cl: parts.append("Center Stone Details\n" + "\n".join(cl))
        if sl: parts.append("Side Stone Details\n" + "\n".join(sl))
        return "\n\n".join(parts)
    return "\n".join(sl)


# ============================================================================
# MAIN CONVERT
# ============================================================================
def convert(inputs, image_map=None, collection=DEFAULT_COLLECTION):
    """Convert one or more MFG workbooks into a VDB-format DataFrame.

    Args:
        inputs:     list of MFG .xlsx file paths OR file-like objects.
        image_map:  dict from load_image_map(), or None to skip image URLs.
        collection: 'Auto', a main collection group, or a specific Config Field 13 collection.

    Returns:
        (df, unknown_remarks) where df has the 41 VDB columns and
        unknown_remarks contains unresolved/ambiguous classification messages.
        Unresolved fields are left blank for review.
    """
    image_map = image_map or {}
    items = []
    for p in inputs:
        items.extend(extract(p))

    return convert_items(items, image_map, collection)


def convert_items(items, image_map=None, collection=DEFAULT_COLLECTION):
    """Convert already parsed item records (also used by the audit workflow)."""
    image_map = image_map or {}
    # Group by base; collect purities + per-purity snapshot
    by_base = {}
    for it in items:
        purs = by_base.setdefault(it['base'], {})
        previous = purs.get(it['purity'])
        if previous is None or it.get('combined') or not previous.get('combined'):
            purs[it['purity']] = it

    out, unknown = [], set()
    for base in by_base:
        purs = by_base[base]
        # Combined-source snapshots take priority for each matching style/purity.
        master_purity = '14K' if '14K' in purs else min(purs)
        master_color = purs[master_purity].get('default_color', 'W')
        master_stock = f"{base}-{master_purity}{master_color}"
        avail = '#'.join(
            _metal_name(p, C)
            for p in sorted(purs)
            for C in sorted(_item_colors(purs[p]))
        )

        for purity, it in sorted(purs.items()):
            item_collection, style, cf14, tags, issue = resolve_classification(it, collection)
            if issue:
                unknown.add(issue)
            legacy = REMARK_MAP.get(re.sub(r'\s+', ' ', it['remark'].upper()).strip())
            subcat = it['jtype'] or 'Jewelry'
            if legacy and legacy[0] != 'jtype':
                subcat = legacy[0]

            # Stone aggregates
            cen_ct   = sum(c['carat'] for c in it['centers'])
            side_ct  = sum(s['carat'] for s in it['sides'])
            side_pcs = sum(s['pcs']   for s in it['sides'])
            weight   = it['total_dia'] if it['total_dia'] > 0 else (cen_ct + side_ct)
            shape = (it['centers'][0]['shape'] if it['centers']
                     else (it['sides'][0]['shape'] if it['sides'] else 'Round')) or 'Round'
            side_shape = (it['sides'][0]['shape'] if it['sides'] else 'Round') or 'Round'
            comment = _build_comment(it['centers'], it['sides'])

            for C in _item_colors(it):
                metal = _metal_name(purity, C)
                wt = _fmt2_fixed(weight)                # 2dp Weight w/ trailing zero
                title = re.sub(r'\s+', ' ', f"{item_collection} {style} {wt}CT {metal} {subcat}").strip()
                urls = image_map.get((base.upper(), C), [])[:MAX_IMAGES]
                imgcols = {f'Image Url {i}': (urls[i-1] if i <= len(urls) else '')
                           for i in range(1, MAX_IMAGES + 1)}

                out.append({
                    'Master Stock':            master_stock,
                    'Stock Number':            f"{base}-{purity}{C}",
                    'Jewelry Sub-Category':    subcat,
                    'Jewelry Style':           style,
                    'Config Field 13':         item_collection,
                    'Config Field 14':         cf14,
                    'Tags':                    tags,
                    'Short Title':             title,
                    'Description':             '',
                    'Metal':                   metal,
                    **imgcols,
                    'Available Metals':        avail,
                    # 29-May convention: Total Price = DC (wholesale), Config Field 17 = SRP (retail)
                    'Total Price':             _fmt(it['dc']),
                    'Config Field 17':         _fmt(it['srp']),
                    'Weight':                  wt,
                    'Center Carat Weight':     _fmt(cen_ct),
                    'Available Center Carat Weight': _fmt2(cen_ct),
                    'Center Color':            CEN_COLOR if cen_ct > 0 else '',
                    'Center Clarity':          CEN_CLARITY if cen_ct > 0 else '',
                    'Shape':                   shape,            # always populated; default Round
                    'Supported Shape Variations': '',            # blank by default; curated externally
                    'Side Carat Weight':       _fmt(side_ct),    # full precision (e.g. 0.168)
                    'Side Total Stones':       _fmt(side_pcs),
                    'Side Color':              SIDE_COLOR,
                    'Side Clarity':            SIDE_CLARITY,
                    'Side Stone Shape':        side_shape,
                    'Comments':                comment,
                })

    df = pd.DataFrame(out)
    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ''
    df = df[COLUMNS].fillna('')
    return df, unknown


# ============================================================================
# CLI
# ============================================================================
def _run_cli(argv=None):
    ap = argparse.ArgumentParser(description="Convert MFG .xlsx files into VDB-format CSV.")
    ap.add_argument('inputs', nargs='+', help='MFG .xlsx file(s)')
    ap.add_argument('-o', '--output', default='VDB_Output.csv')
    ap.add_argument('--reference', help='Approved VDB inventory CSV; preserves curated groups and mappings, emits audit ZIP.')
    ap.add_argument('-i', '--images', default=None,
                    help='image-filename list CSV (single column)')
    ap.add_argument('-c', '--collection', default=DEFAULT_COLLECTION,
                    help=f"collection name (Config Field 13), default {DEFAULT_COLLECTION!r}")
    a = ap.parse_args(argv)

    imap, stats = load_image_map(a.images)
    if a.images:
        print(f"Image list: accepted {stats['accepted']} W/Y/R+digit files "
              f"-> {len(imap)} (base,color) groups; "
              f"skipped {stats['skipped']} non-conforming entries")
        if stats['skip_samples']:
            top = sorted(stats['skip_samples'].items())[:8]
            print("  examples of skipped codes:",
                  ", ".join(f"{c!r} ({fn})" for c, fn in top))

    if a.reference:
        from pathlib import Path
        from inventory_reconcile import build_inventory, audit_zip
        df, reports = build_inventory(a.inputs, a.reference, imap)
        df.to_csv(a.output, index=False)
        audit_path = Path(a.output).with_suffix('.audit.zip')
        audit_path.write_bytes(audit_zip(df, reports))
        print(reports['summary'])
        print(f'Review inventory: {a.output}; audit reports: {audit_path}')
        if reports['summary']['blocking_issues']:
            raise SystemExit(2)
        return
    df, unknown = convert(a.inputs, imap, collection=a.collection)
    df.to_csv(a.output, index=False)
    print(f"Wrote {len(df)} rows -> {a.output}")
    n_styles = df['Stock Number'].str.replace(r'-..K[WYR]+$', '', regex=True).nunique()
    print(f"Unique styles: {n_styles}")
    print(f"Rows with at least one image: {(df['Image Url 1'] != '').sum()}/{len(df)}")
    if unknown:
        print("\n[!] Collection/style mappings requiring review (unresolved fields are blank):")
        for u in sorted(unknown):
            print("   ", u)


# ============================================================================
# STREAMLIT MODE DETECTION
# ============================================================================
def _running_under_streamlit():
    """True iff this script is being executed by `streamlit run`."""
    # Streamlit sets STREAMLIT_SERVER_PORT in the process env when it spawns the
    # script runner, and bootstraps `streamlit.web.bootstrap` into sys.modules.
    # Check those first so we never import streamlit's scriptrunner in CLI mode
    # (which logs a benign-but-noisy warning).
    import os
    if 'streamlit' not in sys.modules and 'STREAMLIT_SERVER_PORT' not in os.environ:
        return False
    try:
        import logging
        logging.getLogger('streamlit').setLevel(logging.ERROR)  # belt-and-suspenders
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


# ============================================================================
# STREAMLIT UI — only executes when launched via `streamlit run app.py`
# ============================================================================
if _running_under_streamlit():
    import streamlit as st

    st.set_page_config(
        page_title="MFG → VDB Converter",
        page_icon="💎",
        layout="wide",
    )

    st.title("💎 MFG → VDB Converter")
    st.caption(
        "Upload a client MFG quotation and the image-filename list, set the "
        "collection name, and download the VDB-format CSV ready to import."
    )

    # ---- sidebar ----
    with st.sidebar:
        st.header("Settings")
        collection = st.text_input(
            "Main collection / specific collection",
            value=DEFAULT_COLLECTION,
            help="Auto uses the reference mappings. Or enter Essentials, Bridals, Haute Couture, "
                 "Fashion Collections, Men's, or a specific collection such as Flora.",
        ).strip() or DEFAULT_COLLECTION

        st.markdown("---")
        st.caption(f"**Collections:** {len(COLLECTION_GROUPS)}")
        with st.expander("Collection and style mappings"):
            rm_df = pd.DataFrame(
                [(COLLECTION_GROUPS[c], c, style, cf14) for c, style, cf14 in CLASSIFICATIONS],
                columns=["Main collection", "Config Field 13", "Jewelry Style", "Config Field 14"],
            )
            st.dataframe(rm_df, use_container_width=True, hide_index=True)

    # ---- uploads ----
    col_l, col_r = st.columns(2)
    with col_l:
        st.subheader("1. MFG file(s)")
        mfg_files = st.file_uploader(
            "Client quotation workbook (.xlsx or .xlsx.zip). Multiple files OK.",
            type=["xlsx", "zip"],
            accept_multiple_files=True,
            key="mfg",
        )
    with col_r:
        st.subheader("2. Image list (optional)")
        img_file = st.file_uploader(
            "Single-column CSV of image filenames (e.g. SRR-00023WHT_W1.jpg).",
            type=["csv"],
            accept_multiple_files=False,
            key="images",
        )

    reference_file = st.file_uploader(
        "3. Approved VDB inventory CSV (recommended for full inventory updates)", type=["csv"],
        key="reference_inventory", help="Preserves existing master groups, variations, collection/style mappings, "
        "center/side details and images. Client prices and default masters are reconciled with audit reports.",
    )

    # ---- convert ----
    disabled = not mfg_files
    run = st.button("Convert →", type="primary",
                    disabled=disabled, use_container_width=True)
    if disabled:
        st.info("Upload at least one MFG .xlsx to enable conversion.")

    if run:
        # image map
        image_map, img_stats = {}, {'accepted': 0, 'skipped': 0, 'skip_samples': {}}
        if img_file is not None:
            try:
                image_map, img_stats = load_image_map(io.BytesIO(img_file.getvalue()))
            except Exception as e:
                st.error(f"Couldn't read the image list: {e}")
                st.stop()

        # convert
        try:
            inputs = [io.BytesIO(f.getvalue()) for f in mfg_files]
            with st.spinner("Parsing MFG sheets and building VDB rows…"):
                reports = None
                if reference_file is not None:
                    from inventory_reconcile import build_inventory
                    df, reports = build_inventory(inputs, io.BytesIO(reference_file.getvalue()), image_map)
                    unknown = set()
                else:
                    df, unknown = convert(inputs, image_map, collection=collection)
        except Exception as e:
            st.error(f"Conversion failed: {e}")
            st.stop()

        # cache so previews/downloads survive UI reruns
        st.session_state["inventory_audit"] = reports
        st.session_state["result_df"]  = df
        st.session_state["unknown"]    = sorted(unknown)
        st.session_state["img_stats"]  = img_stats
        st.session_state["img_used"]   = img_file is not None
        st.session_state["collection"] = collection

    # ---- results panel ----
    if "result_df" in st.session_state:
        df         = st.session_state["result_df"]
        unknown    = st.session_state["unknown"]
        img_stats  = st.session_state["img_stats"]
        img_used   = st.session_state["img_used"]
        coll_used  = st.session_state["collection"]

        st.markdown("---")
        st.subheader("Results")
        reports = st.session_state.get("inventory_audit")
        needs_review = bool(reports and reports['summary']['blocking_issues'])
        if reports:
            from inventory_reconcile import render_audit
            render_audit(st, df, reports, key='standalone_audit_download')

        # stats cards
        unique_styles = df["Stock Number"].str.replace(r"-..K[WYR]+$", "", regex=True).nunique()
        rows_with_img = (df["Image Url 1"] != "").sum()
        cols = st.columns(4)
        cols[0].metric("Rows", f"{len(df):,}")
        cols[1].metric("Unique styles", f"{unique_styles:,}")
        cols[2].metric(
            "Rows with image",
            f"{rows_with_img:,} / {len(df):,}",
            delta=f"{(rows_with_img / len(df) * 100):.1f}%" if len(df) else None,
            delta_color="off",
        )
        cols[3].metric(
            "Image codes accepted",
            f"{img_stats['accepted']:,}" if img_used else "—",
            delta=f"{img_stats['skipped']:,} skipped" if img_used else None,
            delta_color="off",
        )

        # warnings
        if unknown:
            with st.expander(
                f"⚠️ {len(unknown)} collection/style mapping(s) need review",
                expanded=True,
            ):
                st.write(
                    "These collection/style mappings need review. Unresolved fields are left blank."
                )
                st.code("\n".join(unknown))

        if img_used:
            missing_imgs = df[df["Image Url 1"] == ""]
            if not missing_imgs.empty:
                missing_styles = sorted(
                    set(missing_imgs["Stock Number"].str.replace(r"-..K[WYR]+$", "", regex=True))
                )
                with st.expander(f"📷 {len(missing_styles)} style(s) have no image in the list"):
                    st.code(", ".join(missing_styles))

            if img_stats["skip_samples"]:
                with st.expander(
                    f"🪣 {img_stats['skipped']} image-list entries skipped "
                    "(WG/YG/RG, LS, LF, digits-only, malformed)"
                ):
                    st.caption(
                        "W1/Y1/R1 and two tone codes such as WY or WY1 are accepted. Examples of "
                        "skipped codes (one filename per code):"
                    )
                    top = sorted(img_stats["skip_samples"].items())[:25]
                    st.code("\n".join(f"{code!r:14}  {fn}" for code, fn in top))

        # preview
        st.markdown("**Preview** (first 50 rows)")
        st.dataframe(df.head(50), use_container_width=True, hide_index=True)

        # download
        csv_bytes = df.to_csv(index=False).encode("utf-8")
        fname_slug = re.sub(r"[^A-Za-z0-9_-]+", "_", coll_used) or "VDB"
        st.download_button(
            "Download review CSV" if needs_review else "⬇️ Download VDB CSV",
            data=csv_bytes,
            file_name=f"VDB_{fname_slug}_{'Review' if needs_review else 'Output'}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )


# ============================================================================
# CLI ENTRY POINT — only when invoked as `python app.py …`
# ============================================================================
elif __name__ == "__main__":
    _run_cli()

