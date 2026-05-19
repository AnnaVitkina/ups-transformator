"""MainCosts → lane rows for the transport pricing matrix (Excel MainCosts tab).

Typical source JSON: ``processing/from_rate_card_excel.json`` (see ``conversion-to-json.py`` /
``transformation_to_excel.load_extracted_data``).

Public: ``build_matrix_main_costs``, ``expand_main_costs_lanes_by_zoning``,
``apply_zone_labels_to_main_costs``, ``sort_main_costs_rows_for_layout``,
``global_country``, ``parse_zoning_matrix``, ``MAIN_COSTS_SHIPMENT_COLS``,
``pivot_main_costs`` (legacy flat table).
"""

import re
from collections import defaultdict

from transform_other_tabs import build_zone_label_lookup


# Fixed shipment columns for MainCosts (lane identity + zone context).
MAIN_COSTS_SHIPMENT_COLS = [
    'Lane #',
    'Origin Country Region',
    'Origin Country',
    'Origin Postal Code Zone',
    'Destination Country Region',
    'Destination Country',
    'Destination Postal Code Zone',
    'Original Service',
    'Zone',
]


# ---------------------------------------------------------------------------
# Weight sorting helper
# ---------------------------------------------------------------------------

def _weight_sort_key(w):
    """
    Sort key for weight breakpoint values so they always appear in correct
    numeric order regardless of how they were stored as strings.

    Numeric values (e.g. "0.5", "1", "10.0") are sorted as floats:
        0.5 → 1.0 → 1.5 → 2.0 → 10.0 → 11.0  (correct)
    Non-numeric values (rare edge cases) are sorted alphabetically after
    all numeric values.

    Examples:
        sorted(["10.0", "2.0", "0.5", "1.0"], key=_weight_sort_key)
        → ["0.5", "1.0", "2.0", "10.0"]
    """
    try:
        return (0, float(w))   # numeric: sort by float value
    except (ValueError, TypeError):
        return (1, str(w))     # non-numeric: sort alphabetically after numbers


# ---------------------------------------------------------------------------
# Zone-name helpers
# ---------------------------------------------------------------------------

def _zone_has_letters(zone_name):
    """
    Check whether a zone name uses a letter identifier (e.g. "Zone A") rather than
    a number identifier (e.g. "Zone 1").

    Returns True for "Zone A", "Zone E", etc.
    Returns False for "Zone 1", "Zone 12", etc.
    """
    s = (zone_name or '').strip()
    if not s:
        return False
    if s.upper().startswith('ZONE '):
        suffix = s[5:].strip()
    else:
        suffix = s
    return any(c.isalpha() for c in suffix)


def _zone_is_single_letter(zone_name):
    """
    Return True only when the zone identifier is exactly one letter (e.g. "Zone A", "B").

    This is the fallback criterion used when no matching ZoningMatrix exists for a service.
    A single-letter zone almost certainly refers to a matrix lookup code even when the
    matrix name doesn't match the service name closely enough to be found automatically.

    Examples:
      "Zone A"  -> True   (single letter after "Zone ")
      "Zone AB" -> False  (two letters – probably a real zone name, not a matrix code)
      "Zone 1"  -> False  (number, not a letter)
      "A"       -> True   (bare single letter)
    """
    s = (zone_name or '').strip()
    if not s:
        return False
    if s.upper().startswith('ZONE '):
        suffix = s[5:].strip()
    else:
        suffix = s
    # Exactly one alphabetic character and nothing else
    return len(suffix) == 1 and suffix.isalpha()


def _zone_needs_matrix_lookup(zone_name, service_type, zoning_lookup):
    """
    Decide whether a zone in a given service should be treated as a matrix lookup code
    (i.e. needs to be expanded into real Origin/Destination pairs via the ZoningMatrix).

    NEW TWO-STEP LOGIC:

    Step 1 – Service-matrix match (primary):
      Try to find a ZoningMatrix whose name corresponds to this service type.
      If a match is found, ALL zones for this service are matrix zones – regardless
      of whether their name contains letters or numbers.
      This handles the common case where service "DHL EXPRESS WORLDWIDE THIRD COUNTRY"
      has a matching matrix "DHL EXPRESS THIRD COUNTRY ZONE MATRIX".

    Step 2 – Single-letter fallback:
      If no matrix was found for this service, check whether the zone identifier is
      exactly one letter (e.g. "A", "B", "E").  A bare single letter almost certainly
      means the zone is a matrix lookup code even when the matrix name couldn't be
      matched automatically.

    Returns True if the zone should be flagged as a matrix zone, False otherwise.
    """
    if not zone_name:
        return False

    # Step 1: does a matrix exist for this service?
    if zoning_lookup and _find_matrix_for_service(zoning_lookup, service_type):
        # A matching matrix was found – this zone belongs to it
        return True

    # Step 2: no matrix found for the service; fall back to single-letter check
    return _zone_is_single_letter(zone_name)


def _dhl_express_domestic_single_cost_zone_only(main_costs):
    """
    True if MainCosts contains DHL EXPRESS DOMESTIC pricing and every non-adder section
    for that service uses exactly one distinct zone name (e.g. only "Zone A").

    Then we skip matrix expansion and keep a single lane (see PASS 2;
    expand_main_costs_lanes_by_zoning skips empty Matrix zone).

    Origin/Destination are then filled with either the CountryZoning short label
    (``DOMESTIC_ZONE_*``) when a **DHL EXPRESS DOMESTIC ZONING** block exists in
    CountryZoning, or else the carrier country — see PASS 2 closing steps.
    """
    zones = set()
    seen_domestic = False
    for rate_card in main_costs:
        if _is_adder_section(rate_card):
            continue
        st = (rate_card.get('service_type') or '').strip()
        if st.upper() != 'DHL EXPRESS DOMESTIC':
            continue
        seen_domestic = True
        for zname in (rate_card.get('zone_headers') or {}).values():
            zname = (zname or '').strip()
            if zname:
                zones.add(zname)
    if not seen_domestic:
        return False
    return len(zones) == 1


_DOMESTIC_ZONING_RATE_PHRASE = 'DHL EXPRESS DOMESTIC ZONING'


def _country_zoning_has_dhl_express_domestic_zoning(country_zoning):
    """
    True if CountryZoning contains a **DHL EXPRESS DOMESTIC ZONING** rate block
    (contiguous phrase), i.e. not only ``... DOMESTIC THIRD COUNTRY ZONING``.
    Used with single-zone domestic to fill Origin/Destination from CountryZoning labels.
    """
    if not country_zoning:
        return False
    needle = _DOMESTIC_ZONING_RATE_PHRASE
    for item in country_zoning:
        if not isinstance(item, dict):
            continue
        rn = (item.get('RateName') or '').strip()
        if not rn:
            continue
        norm = ' '.join(rn.split()).upper()
        if needle in norm:
            return True
    return False


def _domestic_zone_short_label(zone_label_lookup, zone_name):
    """
    Return label like ``DOMESTIC_ZONE_A`` for MainCosts zone column (e.g. ``Zone A``)
    using the same (prefix, zone) keys as apply_zone_labels_to_main_costs.
    """
    if not zone_label_lookup or not zone_name:
        return None
    zn = re.sub(r'(?i)^zone\s*', '', str(zone_name).strip()).strip()
    if not zn:
        return None
    return zone_label_lookup.get(('DOMESTIC_ZONE', zn))


def _zone_sort_key(zone_name):
    """
    Generate a sort key for a zone name so that zones appear in a sensible order:
    numeric zones first (Zone 1, Zone 2, Zone 10 …) then letter/other zones after.

    Without this, alphabetical sorting would give: Zone 1, Zone 10, Zone 2 (wrong).
    With this, we get: Zone 1, Zone 2, Zone 10, Zone A (correct).

    UPS toolbox labels (e.g. ``TB\\n4``, ``WW\\n706``) use the last line for a numeric
    sort when possible; otherwise they sort alphabetically in group 1.

    Returns a tuple (group, value) where:
      group=0 means numeric zone (sorted by number as float)
      group=1 means letter/other zone (sorted by str; never mix int/str in value)
    """
    s = (zone_name or '').strip()
    if not s:
        return (1, '')
    if s.upper().startswith('ZONE '):
        suffix = s[5:].strip()
    else:
        # UPS: "TB\n4", "WW\n754" — prefer trailing line for ordering
        parts = [p.strip() for p in re.split(r'[\r\n]+', s) if p.strip()]
        suffix = parts[-1] if parts else s
    try:
        return (0, float(suffix))
    except (ValueError, TypeError):
        return (1, suffix.lower())


def global_country(metadata):
    """
    Extract the country name from the carrier string in the metadata.

    DHL carrier names follow the pattern "DHL Express <Country>" (case-insensitive),
    e.g. "DHL Express France"  -> "France"
         "DHL EXPRESS GERMANY" -> "Germany"
         "DHL express Netherlands" -> "Netherlands"

    The country is everything that comes after the words "DHL" and "EXPRESS"
    (or "EXPRESS" alone), title-cased for consistency.

    If the pattern is not found, the last word of the carrier string is used
    as a fallback so the field is never left empty when a carrier is present.

    This country name is used to fill in the Origin or Destination column for:
    - Domestic lanes (both Origin and Destination = carrier's country)
    - Non-zoned export lanes (Destination = carrier's country)
    - Non-zoned import lanes (Origin = carrier's country)
    """
    import re
    carrier = (metadata.get('carrier') or '').replace('\n', ' ').strip()
    if not carrier:
        return ''

    # Words that signal the end of the country name (non-country suffixes)
    _STOP_WORDS = {
        'customer', 'customers', 'services', 'service', 'surcharges', 'surcharge',
        'export', 'import', 'domestic', 'rates', 'rate', 'ratecard', 'tariff',
        'tariffs', 'zone', 'zones', 'express', 'dhl', 'international', 'standard',
        'priority', 'economy', 'freight', 'air', 'ground', 'parcel', 'and',
    }

    # Match everything after "DHL EXPRESS", then walk word by word until a stop word
    m = re.search(r'\bDHL\s+EXPRESS?\s+(.+)', carrier, re.IGNORECASE)
    if m:
        remainder = m.group(1).strip()
        country_words = []
        for word in remainder.split():
            if word.lower() in _STOP_WORDS:
                break
            country_words.append(word)
        if country_words:
            return ' '.join(country_words).title()   # e.g. "UNITED KINGDOM" -> "United Kingdom"

    # UPS: "UPS Sweden", "UPS United Kingdom"
    m = re.search(r'\bUPS\s+(.+)', carrier, re.IGNORECASE)
    if m:
        remainder = m.group(1).strip()
        country_words = []
        for word in remainder.split():
            if word.lower() in _STOP_WORDS:
                break
            country_words.append(word)
        if country_words:
            return ' '.join(country_words).title()

    # Fallback: return the last word of the carrier string
    parts = carrier.split()
    return parts[-1].title() if parts else ''


def _is_ups_maincosts_context(metadata, main_costs):
    """True when rate cards follow UPS ASSA-style movement + service labels."""
    c = (metadata.get('carrier') or '').lower()
    if 'ups' in c:
        return True
    for rc in main_costs or []:
        st = (rc.get('service_type') or '').lower()
        if 'receiving rates' in st or 'sending rates' in st:
            return True
    return False


def parse_ups_country_zoning_layout(country_zoning_rows):
    """
    Parse CountryZoning rows that define Sending vs Receiving columns and per-service labels.

    Looks for Service1 ≈ 'Zonenumber Sending' and Service8 ≈ 'Zonenumber Receiving', then a row
    where Service* values look like 'express plus\\n2' (service name + sample zone line).
    Returns:
      sending_cut: column index (1-based) where Receiving starts (default 8)
      column_labels: { idx: {'label': str, 'direction': 'Export'|'Import'} }
    """
    rows = [r for r in (country_zoning_rows or []) if isinstance(r, dict)]
    sending_cut = 8
    for r in rows:
        s1 = str(r.get('Service1') or '').lower().replace('nummer', 'number')
        s8 = str(r.get('Service8') or '').lower()
        if 'zonenumber' in s1 and 'sending' in s1 and 'zonenumber' in s8 and 'receiving' in s8:
            break

    column_labels = {}
    for r in rows:
        found_any = False
        for k in range(1, 28):
            key = f'Service{k}'
            if key not in r:
                continue
            raw = str(r.get(key) or '').strip()
            if '\n' not in raw:
                continue
            lines = [ln.strip() for ln in raw.split('\n') if ln.strip()]
            if len(lines) < 2:
                continue
            label = re.sub(r'\s+', ' ', lines[0]).strip()
            if not label or len(label) < 2:
                continue
            if not re.search(r'[a-zA-Z]', label):
                continue
            direc = 'Import' if k >= sending_cut else 'Export'
            column_labels[k] = {'label': label, 'direction': direc}
            found_any = True
        if found_any:
            break

    return {'sending_cut': sending_cut, 'column_labels': column_labels}


def _ups_service_keywords(service_type):
    """Return ordered keyword hints for matching CountryZoning service columns."""
    st = (service_type or '').lower().replace('\n', ' ')
    hints = []
    if 'standard multi' in st or ('standard' in st and 'single' not in st and 'saver' not in st):
        hints.append('standard')
    if 'standard single' in st:
        hints.append('standard')
    if 'expedited' in st:
        hints.append('expedited')
    if 'express saver' in st or 'saver' in st and 'express' in st:
        hints.append('express saver')
    if 'express plus' in st:
        hints.append('express plus')
    if 'freight midday' in st or 'midday' in st:
        hints.append('express freight midday')
    if 'express freight' in st and 'midday' not in st:
        hints.append('express freight')
    if 'express' in st and not any(x in st for x in ('saver', 'plus', 'freight')):
        hints.append('express')
    if not hints and 'standard' in st:
        hints.append('standard')
    return hints


def _match_ups_zoning_column(service_type, column_labels):
    """
    Pick the CountryZoning Service* column index that best matches this MainCosts service block.
    Respects Export (Sending) vs Import (Receiving) from 'Sending rates' / 'Receiving rates'.
    """
    if not column_labels:
        return None
    st = (service_type or '').lower().replace('\n', ' ')
    want_import = 'receiving' in st and 'sending' not in st
    want_export = 'sending' in st and 'receiving' not in st
    if not want_import and not want_export:
        want_import = 'receiving' in st or 'import' in st
        want_export = not want_import

    hints = _ups_service_keywords(service_type)
    best_idx, best_score = None, -1

    for idx, meta in column_labels.items():
        lbl = (meta.get('label') or '').lower()
        direc = meta.get('direction') or ''
        if want_import and direc != 'Import':
            continue
        if want_export and direc != 'Export':
            continue
        score = 0
        for h in hints:
            if h in lbl:
                score += 25
            elif h.split()[0] in lbl:
                score += 12
        # service name contains column label words
        for word in re.findall(r'[a-z]{3,}', lbl):
            if len(word) > 3 and word in st:
                score += 3
        if score > best_score:
            best_score = score
            best_idx = idx
    return best_idx if best_score > 0 else None


def _ups_zone_market_and_suffix(zone_header_val):
    """
    Split MainCosts zone header like 'TB\\n1' or 'TB\\n41\\nCZ41' into (market, suffix_string).

    ``suffix_string`` is only used to build **Origin/Destination Country Region** zone titles.
    Postal Code Zone columns are either empty or a full CZ-style label (see
    ``build_ups_shipment_fields``), never bare numbers or WW/TB alone.
    """
    s = str(zone_header_val or '').strip()
    if not s:
        return '', ''
    parts = [p.strip() for p in s.split('\n') if p.strip()]
    if not parts:
        return '', ''
    if len(parts) == 1:
        return '', parts[0]
    market = parts[0]
    rest = ' '.join(parts[1:])
    return market, rest


def _title_zone_service_words(label):
    """Title-case a zoning service label (e.g. 'express saver' -> 'Express Saver')."""
    if not label:
        return ''
    return ' '.join(s.capitalize() for s in str(label).split())


def _zone_semantic_belongs_in_postal_column(zone_title: str) -> bool:
    """
    True for labels like ``Standard Import Zone 41 CZ41`` — put the full string in
    the postal-code zone column instead of Country Region.
    """
    s = (zone_title or '').strip()
    if not s:
        return False
    return bool(re.search(r'(?i)\bCZ\d+\b', s))


def _apply_plain_country_columns(row: dict, carrier: str) -> None:
    """
    Set **Origin Country** / **Destination Country** only when the corresponding
    side has nothing in both Region and Postal Code Zone columns.

    If either Region or Postal is set for that side, do not fill the plain country column.
    """
    row.setdefault('Origin Country', '')
    row.setdefault('Destination Country', '')

    ocr = (row.get('Origin Country Region') or '').strip()
    ocz = (row.get('Origin Postal Code Zone') or '').strip()
    if not ocr and not ocz:
        if not (row.get('Origin Country') or '').strip():
            row['Origin Country'] = carrier or ''
    else:
        if not (row.get('Origin Country') or '').strip():
            row['Origin Country'] = ''

    dcr = (row.get('Destination Country Region') or '').strip()
    dpz = (row.get('Destination Postal Code Zone') or '').strip()
    if not dcr and not dpz:
        if not (row.get('Destination Country') or '').strip():
            row['Destination Country'] = carrier or ''
    else:
        if not (row.get('Destination Country') or '').strip():
            row['Destination Country'] = ''


def build_ups_shipment_fields(service_type, zone_header_val, ups_layout, metadata):
    """
    Fill MAIN_COSTS_SHIPMENT_COLS for UPS: semantic zone title + carrier country.

    **Postal Code Zone** is only set for CZ-style labels (e.g. ``Standard Import Zone 41 CZ41``);
    otherwise it stays **empty** — never WW/TB and never bare zone numbers.
    Import (Receiving) → Origin Country Region; Export (Sending) → Destination Country Region.
    """
    out = {k: '' for k in MAIN_COSTS_SHIPMENT_COLS if k != 'Lane #'}
    st_full = (service_type or '').strip()
    out['Original Service'] = st_full
    out['Service'] = st_full  # legacy key for downstream

    col_labels = (ups_layout or {}).get('column_labels') or {}
    idx = _match_ups_zoning_column(st_full, col_labels)
    matched_lbl = ''
    if idx is not None and idx in col_labels:
        matched_lbl = _title_zone_service_words(col_labels[idx]['label'])

    hints = _ups_service_keywords(st_full)
    if matched_lbl:
        base = matched_lbl
    elif hints:
        base = _title_zone_service_words(hints[0])
    else:
        base = 'Zone'

    _, z_suffix = _ups_zone_market_and_suffix(zone_header_val)
    st_low = st_full.lower()
    want_import = 'receiving' in st_low and 'sending' not in st_low
    want_export = 'sending' in st_low and 'receiving' not in st_low
    if not want_import and not want_export:
        want_import = 'receiving' in st_low

    direction_word = 'Import' if want_import else 'Export'
    zone_title = f"{base} {direction_word} Zone {z_suffix}".strip() if z_suffix else f"{base} {direction_word} Zone".strip()

    carrier_ctry = global_country(metadata)
    postal_style = _zone_semantic_belongs_in_postal_column(zone_title)

    if want_import:
        if postal_style:
            out['Origin Country Region'] = ''
            out['Origin Postal Code Zone'] = zone_title
        else:
            out['Origin Country Region'] = zone_title
            out['Origin Postal Code Zone'] = ''
        out['Destination Country'] = ''
        dcr = (out.get('Destination Country Region') or '').strip()
        dpz = (out.get('Destination Postal Code Zone') or '').strip()
        if not dcr and not dpz:
            out['Destination Country'] = carrier_ctry
    else:
        out['Origin Country Region'] = carrier_ctry
        if postal_style:
            out['Destination Country Region'] = ''
            out['Destination Postal Code Zone'] = zone_title
        else:
            out['Destination Country Region'] = zone_title
            out['Destination Postal Code Zone'] = ''

    out['Zone'] = (zone_header_val or '').strip()

    return out


# ---------------------------------------------------------------------------
# ZoningMatrix parsing and lane expansion
# ---------------------------------------------------------------------------

def parse_zoning_matrix(zoning_matrix):
    """
    Read the ZoningMatrix data and build a lookup table that answers the question:
    "For zone letter A in matrix X, which (origin zone, destination zone) pairs exist?"

    BACKGROUND – what is a ZoningMatrix?
    The ZoningMatrix is a grid that maps pairs of origin and destination zone numbers
    to a single letter (A, B, C …).  For example:
        Origin 1 -> Destination 3 -> letter "A"
        Origin 2 -> Destination 3 -> letter "A"
        Origin 1 -> Destination 5 -> letter "E"

    The MainCosts pricing table uses those letters as shorthand: instead of listing
    a price for every individual origin/destination pair, it lists one price per letter.
    This function reverses the matrix so we can later expand each letter back into
    all the concrete (origin, destination) pairs it represents.

    THE JSON STRUCTURE:
    The ZoningMatrix arrives as a flat list of rows.  Two types of rows alternate:
      - Header row: has 'MatrixName' filled in + DestinationZone1, DestinationZone2 …
                    whose values are the destination zone numbers (1, 2, 3 …)
      - Data row:   has 'OriginZone' filled in + DestinationZone1, DestinationZone2 …
                    whose values are the zone letters (A, B, E …)

    WHAT THIS FUNCTION RETURNS:
    A dictionary where:
      key   = (matrix_name, zone_letter)   e.g. ("DHL EXPRESS WW ZONE MATRIX", "A")
      value = list of (origin_zone, destination_zone) pairs  e.g. [("1", "3"), ("2", "3")]
    """
    result = {}                    # the lookup table we are building
    dest_cols = None               # ordered list of "DestinationZone1", "DestinationZone2" … keys
    header_dest_nums = None        # the actual destination zone numbers read from the header row
    current_matrix_name = None     # name of the matrix block we are currently inside

    for row in zoning_matrix or []:
        matrix_name = (row.get('MatrixName') or '').strip()
        origin_zone = (row.get('OriginZone') or '').strip()

        # Find DestinationZone* keys in this row (may be in same row as MatrixName or in next row)
        dest_keys = sorted(
            [k for k in row if re.match(r'^DestinationZone\d+$', k)],
            key=lambda k: int(re.search(r'\d+', k).group())
        )

        if matrix_name:
            # ---------------------------------------------------------------
            # This is a HEADER ROW – it starts a new matrix block.
            # Example: MatrixName="DHL EXPRESS WW ZONE MATRIX",
            #          DestinationZone1="1", DestinationZone2="2", DestinationZone3="3"
            # Some PDFs put MatrixName alone in one row; then the next row has the zone columns.
            # ---------------------------------------------------------------
            current_matrix_name = matrix_name

            if dest_keys:
                dest_cols = dest_keys
                header_dest_nums = [str(row.get(k, '')).strip() for k in dest_cols]
            # else: keep previous dest_cols/header_dest_nums so data rows can still be parsed
            continue   # move on to the next row (this header row has no zone letters to add)

        if current_matrix_name and dest_keys and not origin_zone:
            # Row has DestinationZone* but no OriginZone – treat as secondary header (zone column numbers)
            # so the first matrix is not skipped when its header is split across two rows
            dest_cols = dest_keys
            header_dest_nums = [str(row.get(k, '')).strip() for k in dest_cols]
            continue

        if current_matrix_name and origin_zone and dest_cols:
            # ---------------------------------------------------------------
            # This is a DATA ROW – it belongs to the current matrix block.
            # Example: OriginZone="1",
            #          DestinationZone1="A", DestinationZone2="A", DestinationZone3="E"
            # This means: origin 1 -> destination 1 = letter A
            #             origin 1 -> destination 2 = letter A
            #             origin 1 -> destination 3 = letter E
            # ---------------------------------------------------------------
            for col_idx, dest_key in enumerate(dest_cols):
                if col_idx >= len(header_dest_nums):
                    continue   # safety check: don't go past the number of header columns
                dest_zone_num = header_dest_nums[col_idx]   # e.g. "3"
                if not dest_zone_num:
                    continue   # skip if the header had no zone number for this column
                cell_letter = (row.get(dest_key) or '').strip()   # e.g. "A"
                if not cell_letter:
                    continue   # skip empty cells (no zone letter assigned)

                # Build the lookup key: (matrix_name, letter)
                key = (current_matrix_name, cell_letter.upper())
                if key not in result:
                    result[key] = []   # create a new list for this letter if first time seen
                # Record that this (origin, destination) pair maps to this letter
                result[key].append((origin_zone, dest_zone_num))

    return result


def _matrix_zone_to_letter(matrix_zone):
    """
    Extract just the letter part from a zone name like "Zone E" -> "E".
    This is needed because the lookup table is keyed by the letter alone, not the full name.
    If the input is already just a letter (no "Zone " prefix), it is returned as-is in uppercase.
    """
    s = (matrix_zone or '').strip()
    if not s:
        return ''
    if s.upper().startswith('ZONE '):
        return s[5:].strip().upper()   # remove "Zone " and return the rest in uppercase
    return s.upper()


def _main_words(text):
    """
    Split a text string into its meaningful words (all uppercase), ignoring the
    generic words "ZONE" and "MATRIX" which appear in almost every matrix name
    and would cause false matches.

    Example: "DHL EXPRESS THIRD COUNTRY ZONE MATRIX" -> {"DHL", "EXPRESS", "THIRD", "COUNTRY"}
    """
    if not text:
        return set()
    words = set((text or '').upper().split())
    words.discard('ZONE')     # too generic to be useful for matching
    words.discard('MATRIX')   # too generic to be useful for matching
    return words


def _norm_matrix_name(s):
    """Normalize matrix/service strings for comparison (collapse whitespace, upper)."""
    return ' '.join((s or '').strip().split()).upper()


def _find_matrix_name_in_lookup(canonical_name, matrix_names):
    """Return the actual matrix name from lookup if it matches canonical (spacing/case tolerant)."""
    want = _norm_matrix_name(canonical_name)
    for mn in matrix_names:
        if _norm_matrix_name(mn) == want:
            return mn
    return None


def _explicit_third_country_matrix(service_upper, matrix_names):
    """
    Fixed mapping: MainCosts service line -> ZoningMatrix name.

    Without this, Attempt 0 returned the first non-domestic third-country matrix from an
    unordered set — often DHL ECONOMY SELECT THIRD COUNTRY ZONE MATRIX before DHL EXPRESS
    THIRD COUNTRY ZONE MATRIX, so WORLDWIDE THIRD COUNTRY lanes used the wrong grid.

    Mapping (only if that matrix exists in zoning_lookup):
      DHL EXPRESS WORLDWIDE THIRD COUNTRY  -> DHL EXPRESS THIRD COUNTRY ZONE MATRIX
      DHL EXPRESS DOMESTIC THIRD COUNTRY     -> DHL EXPRESS DOMESTIC THIRD COUNTRY ZONE MATRIX
      DHL ECONOMY SELECT THIRD COUNTRY       -> DHL ECONOMY SELECT THIRD COUNTRY ZONE MATRIX
    """
    if 'THIRD' not in service_upper or 'COUNTRY' not in service_upper:
        return None
    if 'ECONOMY' in service_upper and 'SELECT' in service_upper:
        return _find_matrix_name_in_lookup(
            'DHL ECONOMY SELECT THIRD COUNTRY ZONE MATRIX', matrix_names
        )
    if 'DOMESTIC' in service_upper:
        return _find_matrix_name_in_lookup(
            'DHL EXPRESS DOMESTIC THIRD COUNTRY ZONE MATRIX', matrix_names
        )
    if 'WORLDWIDE' in service_upper:
        return _find_matrix_name_in_lookup(
            'DHL EXPRESS THIRD COUNTRY ZONE MATRIX', matrix_names
        )
    return None


def _find_matrix_for_service(zoning_lookup, service):
    """
    Given a service type name (e.g. "DHL EXPRESS THIRD COUNTRY"), find which matrix
    in the zoning_lookup corresponds to it.

    WHY THIS IS NEEDED:
    The service names in MainCosts and the matrix names in ZoningMatrix are written
    slightly differently.  For example:
      - Service:  "DHL EXPRESS THIRD COUNTRY"
      - Matrix:   "DHL EXPRESS THIRD COUNTRY ZONE MATRIX"
    We need to match them up despite these differences.

    MATCHING STRATEGY (tries each approach in order, returns the first match found):
      0. Explicit third-country service -> matrix (_explicit_third_country_matrix)
      1. Attempt 0 legacy: WORLDWIDE THIRD COUNTRY non-domestic matrices
      2. Direct substring: does the service name appear inside the matrix name, or vice versa?
      3. Strip " ZONE MATRIX" from the matrix name, then try substring again.
      4. Word-level match: do all meaningful words from the matrix name appear in the service?
         e.g. {"DHL", "EXPRESS", "THIRD", "COUNTRY"} are all present in "DHL EXPRESS THIRD COUNTRY"

    Returns the matching matrix name, or None if no match is found.
    """
    service = (service or '').strip()
    if not service:
        return None
    service_upper = service.upper()
    service_words = _main_words(service)

    # Get all unique matrix names from the lookup (ignoring the zone letter part of each key)
    matrix_names = {mn for (mn, _) in zoning_lookup}

    explicit = _explicit_third_country_matrix(service_upper, matrix_names)
    if explicit:
        return explicit

    # --- Attempt 0: WORLDWIDE THIRD COUNTRY must use the non-Domestic matrix ---
    # Service "DHL EXPRESS WORLDWIDE THIRD COUNTRY" -> "DHL EXPRESS THIRD COUNTRY ZONE MATRIX"
    # (not "DHL EXPRESS DOMESTIC THIRD COUNTRY ZONE MATRIX"). Prefer matrix that has THIRD COUNTRY but not DOMESTIC.
    if 'WORLDWIDE' in service_upper and 'THIRD' in service_upper and 'COUNTRY' in service_upper:
        for mn in matrix_names:
            mn_upper = mn.upper()
            if 'THIRD' in mn_upper and 'COUNTRY' in mn_upper and 'DOMESTIC' not in mn_upper:
                return mn
        # Fallback: source data often has only DOMESTIC THIRD COUNTRY ZONE MATRIX; use it for WORLDWIDE so expansion runs
        for mn in matrix_names:
            mn_upper = mn.upper()
            if 'THIRD' in mn_upper and 'COUNTRY' in mn_upper:
                return mn

    # --- Attempt 1: direct substring match ---
    for mn in matrix_names:
        if service in mn or mn in service:
            return mn   # found a match, return immediately

    # --- Attempt 2: strip the " ZONE MATRIX" boilerplate and try again ---
    for mn in matrix_names:
        normalized = mn.replace(' ZONE MATRIX', '').strip()
        if service in normalized or normalized in service:
            return mn

    # --- Attempt 3: all meaningful words from the matrix name must be in the service ---
    # This handles cases where word order differs or extra words are present
    for mn in matrix_names:
        matrix_words = _main_words(mn.replace(' ZONE MATRIX', ''))
        # "<=" on sets means "is a subset of": all matrix words appear in service words
        if matrix_words and matrix_words <= service_words:
            return mn

    return None   # no match found in any of the three attempts


# ---------------------------------------------------------------------------
# MainCosts – legacy flat pivot (zones as rows, weights as columns)
# ---------------------------------------------------------------------------

def pivot_main_costs(main_costs, metadata):
    """
    (Legacy / unused view) Convert the MainCosts pricing data into a simple flat table
    where each row = one delivery zone, and each column = one weight bracket.

    Example of what the output looks like:
        Zone    | 0.5 KG | 1 KG | 2 KG
        Zone 1  |  12.50 | 15.00| 18.00
        Zone 2  |  14.00 | 17.50| 21.00

    This is an older, simpler view.  The main view used today is build_matrix_main_costs().
    """
    rows = []   # will hold all the output rows we build

    # Pull the three identity fields that appear on every row
    client = (metadata.get('client') or '')
    carrier = (metadata.get('carrier') or '').replace('\n', ' ')  # remove any line breaks
    validity_date = (metadata.get('validity_date') or '')

    # Loop over each "rate card" block in the MainCosts list.
    # Each rate card covers one service type (e.g. "DHL EXPRESS WORLDWIDE EXPORT")
    # and one cost category (e.g. "Documents").
    for section_idx, rate_card in enumerate(main_costs, 1):
        service_type = rate_card.get('service_type') or ''
        cost_category = rate_card.get('cost_category', '')
        weight_unit = rate_card.get('weight_unit', 'KG')

        # zone_headers maps internal short keys (e.g. "Z1") to display names (e.g. "Zone 1")
        zone_headers = rate_card.get('zone_headers', {})

        # pricing is a list where each entry covers one weight breakpoint.
        # Example entry: { "weight": "0.5", "zone_prices": {"Z1": 12.50, "Z2": 14.00} }
        pricing = rate_card.get('pricing', [])

        # ---------------------------------------------------------------
        # Step 1: Reorganise the data from "weight-first" to "zone-first".
        # ---------------------------------------------------------------
        zone_price_matrix = {}   # zone_name -> { weight -> price }
        weights_set = set()      # collect all unique weight values seen

        for price_entry in pricing:
            weight = price_entry.get('weight', '')
            weights_set.add(weight)
            zone_prices = price_entry.get('zone_prices', {})

            for zone_key, price in zone_prices.items():
                zone_name = zone_headers.get(zone_key, zone_key)
                if zone_name not in zone_price_matrix:
                    zone_price_matrix[zone_name] = {}
                zone_price_matrix[zone_name][weight] = price

        # Sort the weight values numerically
        weights_sorted = sorted(weights_set, key=_weight_sort_key)

        # ---------------------------------------------------------------
        # Step 2: Build one output row per zone.
        # ---------------------------------------------------------------
        for zone_name, weight_prices in zone_price_matrix.items():
            row = {
                'Client': client,
                'Carrier': carrier,
                'Validity Date': validity_date,
                'Section': section_idx,
                'Service Type': service_type,
                'Cost Category': cost_category,
                'Weight Unit': weight_unit,
                'Zone': zone_name
            }

            for weight in weights_sorted:
                col_name = f"{weight} {weight_unit}"   # e.g. "0.5 KG"
                row[col_name] = weight_prices.get(weight, '')

            rows.append(row)

    return rows


# ---------------------------------------------------------------------------
def _format_cost_category(raw_name):
    """
    Wrap a raw cost-category name in the standard "Transport cost (...)" label.

    Examples:
        "Documents up to 2.0 KG"  ->  "Transport cost (Documents up to 2.0 KG)"
        "Envelope up to 300 g"    ->  "Transport cost (Envelope up to 300 g)"
        ""                        ->  ""   (empty stays empty)

    Note: "Adder rate per additional X KG from Y" sections are not formatted
    as a separate cost; they are merged into the previous category (see
    _is_adder_section and adder handling in build_matrix_main_costs).
    """
    raw_name = (raw_name or '').strip()
    if not raw_name:
        return raw_name
    return f"Transport cost ({raw_name})"


def _is_adder_section(rate_card):
    """
    Return True if this rate card is an "adder" table that should be merged
    into the previous cost category instead of creating a new one.

    Adder tables have cost_category like:
      "Adder rate per additional 0.5 KG from 10.1 KG"
      "Adder rate per additional 1 KG from 30.1 KG"
    and weight values like "10.1\n20" (From/To range).
    """
    cost_category = (rate_card.get('cost_category') or '').strip()
    if not cost_category:
        return False
    cost_lower = cost_category.lower()
    return 'adder rate' in cost_lower and 'additional' in cost_lower


def _parse_adder_unit(cost_category_raw):
    """
    Extract the unit value from an adder cost category for the "p/X unit" label.

    Examples:
        "Adder rate per additional 0.5 KG from 10.1 KG"  ->  "0.5"
        "Adder rate per additional 1 KG from 30.1 KG"    ->  "1"
    """
    s = (cost_category_raw or '').strip()
    # Match "additional" followed by optional spaces and a number (int or decimal)
    m = re.search(r'additional\s+([0-9]+(?:\.[0-9]+)?)\s*(?:KG|kg|g)?', s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return '1'


def _normalize_adder_weight(weight_str):
    """
    Convert adder weight range from extracted form to display form.

    The extracted value is often "10.1\\n20" (From\\nTo). We display as "10.1-20".
    """
    if not weight_str:
        return weight_str
    s = str(weight_str).strip().replace('\n', '-').replace('\r', '')
    # Collapse multiple spaces or dashes into one dash
    s = re.sub(r'[\s\-]+', '-', s).strip('-')
    return s if s else weight_str


def _adder_range_sort_key(range_str):
    """
    Sort key for adder weight ranges (e.g. "10.1-20", "70.1-100") so they appear
    in increasing order by the range start. "10.1-20" < "20.1-30" < "70.1-100".
    """
    if not range_str:
        return (0, 0.0)
    s = str(range_str).strip()
    m = re.match(r'^([0-9]+(?:\.[0-9]+)?)', s)
    if m:
        try:
            return (0, float(m.group(1)))
        except ValueError:
            pass
    return (1, s)


def _adder_block_sort_key(block):
    """
    Sort key for category blocks so that Flat (main) is first, then adder blocks
    by unit: p/0.5 unit, p/1 unit, p/5 unit (by numeric value).
    """
    weight_unit, weights, row4_label = block
    if row4_label == 'Flat':
        return (0, 0.0)
    m = re.search(r'p/([0-9]+(?:\.[0-9]+)?)\s*unit', weight_unit, re.IGNORECASE)
    if m:
        try:
            return (1, float(m.group(1)))
        except ValueError:
            pass
    return (1, 999.0)


def _scan_category_merge_meta(main_costs):
    """Collect flat weights per formatted cost category and service (as in the extract)."""
    category_merge_meta = {}

    def _empty_mmeta():
        return {'per_service_flat': defaultdict(set)}

    for rate_card in main_costs:
        if _is_adder_section(rate_card):
            continue

        service_type = (rate_card.get('service_type') or '').strip()
        cost_category_raw = rate_card.get('cost_category') or ''
        pricing = rate_card.get('pricing', [])

        by_base_weights = defaultdict(set)
        for pe in pricing:
            w = pe.get('weight', '')
            if not w:
                continue
            eff_raw = (pe.get('cost_category') or cost_category_raw).strip() or cost_category_raw
            base = _format_cost_category(eff_raw)
            by_base_weights[base].add(w)

        for base, weights_set in by_base_weights.items():
            weights_sorted = sorted(weights_set, key=_weight_sort_key)
            mmeta = category_merge_meta.setdefault(base, _empty_mmeta())
            for w in weights_sorted:
                mmeta['per_service_flat'][service_type].add(w)
            last_base = base

    return category_merge_meta


def _build_service_variant_maps(category_merge_meta):
    """Map each base ``Transport cost (…)`` name to per-service labels (same as base)."""
    base_to_service_variant = {}
    for base, meta in category_merge_meta.items():
        per_flat = meta.get('per_service_flat') or {}
        services = set(per_flat.keys())
        if not services:
            continue
        base_to_service_variant[base] = {svc: base for svc in services}
    return base_to_service_variant


# MainCosts – matrix (lane) view builder
# ---------------------------------------------------------------------------

def build_matrix_main_costs(main_costs, metadata, zoning_matrix=None, country_zoning=None):
    """
    One row per (service, zone lane); price cells keyed by ``(Transport cost (…), weight)``.
    Shipment columns: ``MAIN_COSTS_SHIPMENT_COLS``. UPS uses CountryZoning layout when present.
    Weight brackets are exactly those in the extract (flat columns use ``<= weight`` in Excel).

    Returns ``(rows, category_specs)`` for ``write_matrix_sheet``.
    """
    zoning_lookup = parse_zoning_matrix(zoning_matrix) if zoning_matrix else {}
    use_ups_shipment = _is_ups_maincosts_context(metadata, main_costs)
    ups_layout = parse_ups_country_zoning_layout(country_zoning) if use_ups_shipment else {}

    base_merge_meta = _scan_category_merge_meta(main_costs)
    base_to_service_variant = _build_service_variant_maps(base_merge_meta)

    # PASS 1 — category_specs: [(cost_cat_name, [(weight_unit, weights, row4_label), ...]), ...]
    category_specs = []
    seen_categories = {}  # variant cost_cat_name -> index in category_specs
    seen_adder_per_category = set()
    last_category_idx = -1
    _debug_main_costs = False

    for rate_card in main_costs:
        cost_category_raw = rate_card.get('cost_category') or ''
        service_type = (rate_card.get('service_type') or '').strip()
        pricing = rate_card.get('pricing', [])

        if _is_adder_section(rate_card):
            if not category_specs or last_category_idx < 0:
                if _debug_main_costs:
                    print(f"[DEBUG MainCosts] ADDER skipped (no category yet): service={service_type!r} cost={cost_category_raw!r}")
                continue
            prev_name = category_specs[last_category_idx][0]
            unit = _parse_adder_unit(cost_category_raw)
            rate_by = f"p/{unit} unit"
            weights_adder = []
            for pe in pricing:
                w = pe.get('weight', '')
                if w:
                    weights_adder.append(_normalize_adder_weight(w))
            weights_adder_sorted = sorted(weights_adder, key=_adder_range_sort_key)
            sig = (prev_name, rate_by, tuple(weights_adder_sorted))
            if sig in seen_adder_per_category:
                if _debug_main_costs:
                    print(f"[DEBUG MainCosts] ADDER skipped (duplicate): service={service_type!r} attach_to={prev_name!r} rate_by={rate_by!r} weights={weights_adder_sorted}")
                continue
            seen_adder_per_category.add(sig)
            prev_blocks = category_specs[last_category_idx][1]
            prev_blocks.append((rate_by, weights_adder_sorted, rate_by))
            if _debug_main_costs:
                print(f"[DEBUG MainCosts] ADDER attached: service={service_type!r} cost_raw={cost_category_raw!r} -> ATTACH_TO( last processed category )={prev_name!r} rate_by={rate_by!r} weights={weights_adder_sorted}")
            continue

        # Normal section — use per-pricing-row cost (e.g. UPS Pkg vs header Cntr) when present
        weight_unit = rate_card.get('weight_unit') or 'KG'
        by_eff = defaultdict(set)
        for pe in pricing:
            w = pe.get('weight', '')
            if not w:
                continue
            eff_raw = (pe.get('cost_category') or cost_category_raw).strip() or cost_category_raw
            by_eff[eff_raw].add(w)

        for eff_raw, weights_set in by_eff.items():
            base_category = _format_cost_category(eff_raw)
            svc_map = base_to_service_variant.get(base_category) or {}
            cost_category = svc_map.get(service_type, base_category)
            weights_sorted = sorted(weights_set, key=_weight_sort_key)
            block = (weight_unit, weights_sorted, 'Flat')

            if cost_category not in seen_categories:
                seen_categories[cost_category] = len(category_specs)
                category_specs.append((cost_category, [block]))
                last_category_idx = len(category_specs) - 1
                if _debug_main_costs:
                    print(f"[DEBUG MainCosts] NEW category (now last): service={service_type!r} eff={eff_raw!r} -> category={cost_category!r} (flat weights count={len(weights_sorted)})")
            else:
                idx = seen_categories[cost_category]
                _, blocks = category_specs[idx]
                existing_unit, existing_weights, row4 = blocks[0]
                merged = set(existing_weights) | set(weights_sorted)
                merged_sorted = sorted(merged, key=_weight_sort_key)
                blocks[0] = (existing_unit, merged_sorted, row4)
                last_category_idx = idx
                if _debug_main_costs:
                    print(f"[DEBUG MainCosts] MERGE into existing: service={service_type!r} eff={eff_raw!r} -> category={cost_category!r}")

    if _debug_main_costs:
        print("[DEBUG MainCosts] --- PASS 1 summary: categories and their blocks (order = column order in Excel) ---")
        for i, (cat_name, blocks) in enumerate(category_specs):
            block_labels = []
            for b in blocks:
                unit, weights, row4 = b
                if row4 == 'Flat':
                    block_labels.append(f"Flat({len(weights)} weights)")
                else:
                    block_labels.append(f"{row4}({weights})")
            print(f"  [{i}] {cat_name!r} -> blocks: {block_labels}")

    for _cat_name, blocks in category_specs:
        blocks.sort(key=_adder_block_sort_key)

    domestic_single_zone = _dhl_express_domestic_single_cost_zone_only(main_costs)
    cz = country_zoning or []
    has_domestic_zoning_rate = _country_zoning_has_dhl_express_domestic_zoning(cz)
    use_domestic_rate_labels = bool(
        domestic_single_zone and has_domestic_zoning_rate and cz
    )
    zone_label_lookup_domestic = (
        build_zone_label_lookup(cz) if use_domestic_rate_labels else {}
    )

    # =======================================================================
    # PASS 2 – Build one row per lane (service + zone combination).
    # =======================================================================
    lane_rows = {}   # (service_type, zone_name) -> row dict
    prev_cost_category = None   # variant name for merging adder sections

    for rate_card in main_costs:
        service_type = (rate_card.get('service_type') or '').strip()
        cost_category_raw = rate_card.get('cost_category') or ''
        zone_headers = rate_card.get('zone_headers', {})
        pricing = rate_card.get('pricing', [])

        if _is_adder_section(rate_card):
            if prev_cost_category is None:
                continue
            cost_category = prev_cost_category
            _key_weight = lambda w: _normalize_adder_weight(w)
        else:
            _key_weight = lambda w: w

        service_lower = service_type.lower()
        is_import = 'import' in service_lower
        is_export = 'export' in service_lower

        zone_price_matrix = {}
        for price_entry in pricing:
            weight = price_entry.get('weight', '')
            if not _is_adder_section(rate_card):
                eff_raw = (price_entry.get('cost_category') or cost_category_raw).strip() or cost_category_raw
                base_cat = _format_cost_category(eff_raw)
                svc_map = base_to_service_variant.get(base_cat) or {}
                cost_category = svc_map.get(service_type, base_cat)
                prev_cost_category = cost_category
            zone_prices = price_entry.get('zone_prices', {})
            for zone_key, price in zone_prices.items():
                zone_name = zone_headers.get(zone_key, zone_key)
                if zone_name not in zone_price_matrix:
                    zone_price_matrix[zone_name] = {}
                zone_price_matrix[zone_name][weight] = (price, cost_category, _key_weight)

        for zone_name, weight_map in zone_price_matrix.items():
            key = (service_type, zone_name)

            if key not in lane_rows:
                origin = zone_name if is_import else ''
                destination = zone_name if is_export else ''
                if domestic_single_zone and (service_type or '').strip().upper() == 'DHL EXPRESS DOMESTIC':
                    needs_lookup = False
                    matrix_zone = ''
                else:
                    needs_lookup = _zone_needs_matrix_lookup(zone_name, service_type, zoning_lookup)
                    matrix_zone = zone_name if needs_lookup else ''

                if use_ups_shipment:
                    ship = build_ups_shipment_fields(service_type, zone_name, ups_layout, metadata)
                    ship['Matrix zone'] = matrix_zone
                    lane_rows[key] = ship
                else:
                    lane_rows[key] = {
                        'Origin Country Region': origin,
                        'Origin Country': '',
                        'Origin Postal Code Zone': '',
                        'Destination Country Region': destination,
                        'Destination Country': '',
                        'Destination Postal Code Zone': '',
                        'Original Service': service_type,
                        'Service': service_type,
                        'Zone': matrix_zone,
                        'Matrix zone': matrix_zone,
                    }

            row = lane_rows[key]
            for weight, triple in weight_map.items():
                price, cc_use, kw = triple
                row[(cc_use, kw(weight))] = price

    carrier_last = global_country(metadata)

    sorted_keys = sorted(
        lane_rows.keys(),
        key=lambda k: (k[0] or '', _zone_sort_key(k[1])),
    )

    rows = []
    for lane, key in enumerate(sorted_keys, 1):
        row = lane_rows[key].copy()
        row['Lane #'] = lane

        service = (row.get('Service') or row.get('Original Service') or '').strip()
        matrix_zone = (row.get('Matrix zone') or row.get('Zone') or '').strip()

        if service.upper() == 'DHL EXPRESS DOMESTIC':
            zone_nm = key[1] if isinstance(key, tuple) and len(key) > 1 else ''
            dom_label = _domestic_zone_short_label(zone_label_lookup_domestic, zone_nm)
            if dom_label:
                row['Origin Country Region'] = dom_label
                row['Destination Country Region'] = dom_label
            elif carrier_last:
                row['Origin Country Region'] = carrier_last
                row['Destination Country Region'] = carrier_last
        elif not matrix_zone:
            if carrier_last:
                if not (row.get('Origin Country Region') or '').strip():
                    row['Origin Country Region'] = carrier_last
                if not (row.get('Destination Country Region') or '').strip():
                    row['Destination Country Region'] = carrier_last

        _apply_plain_country_columns(row, carrier_last or '')
        rows.append(row)

    return rows, category_specs


def apply_zone_labels_to_main_costs(matrix_rows, zone_label_lookup):
    """
    Replace raw zone names in Origin/Destination with meaningful short labels.

    PURPOSE:
    After build_matrix_main_costs() runs, zoned lanes have Origin or Destination
    values like "Zone 8".  This function replaces those with a label that includes
    the service context, e.g. "ECONOMY_EXP_ZONE_8", so the analyst can immediately
    see which zoning scheme the zone belongs to.

    HOW IT WORKS:
    For each lane row:
      1. Check if Origin or Destination looks like a zone (starts with "Zone ").
      2. Extract the zone number (e.g. "Zone 8" -> "8").
      3. Convert the Service name to its short prefix using the same
         _transform_rate_name_to_short() logic used to build the lookup.
      4. Look up (short_prefix, zone_number) in the zone_label_lookup dict.
      5. If found, replace the Origin/Destination value with the label.

    Rows where Origin/Destination is a country name (not a zone) are left unchanged.

    Parameters:
      matrix_rows       – list of lane row dicts from build_matrix_main_costs()
      zone_label_lookup – dict built by build_zone_label_lookup() in transform_other_tabs.py
                          keys: (short_prefix, zone_number), values: label string

    Returns the same list of rows with Origin/Destination values updated in place.
    """
    if not zone_label_lookup or not matrix_rows:
        return matrix_rows

    # Import here to avoid circular imports (transform_other_tabs imports nothing from here)
    from transform_other_tabs import _transform_rate_name_to_short

    _zone_re = re.compile(r'(?i)^zone\s+(.+)$')

    for row in matrix_rows:
        service = (row.get('Service') or row.get('Original Service') or '').strip()
        short_prefix = _transform_rate_name_to_short(service)
        if not short_prefix:
            continue

        for field in ('Origin Country Region', 'Destination Country Region', 'Origin', 'Destination'):
            val = (row.get(field) or '').strip()
            m = _zone_re.match(val)
            if not m:
                continue   # not a zone value — leave unchanged

            zone_number = m.group(1).strip()
            label = zone_label_lookup.get((short_prefix, zone_number))
            if label:
                row[field] = label

    return matrix_rows


def _origin_layout_sort_tuple(origin):
    """
    Derive a sortable tuple from Origin for matrix-expanded rows.
    Prefer trailing ``_N`` (e.g. ``WW_EXP_IMP_ZONE_1``), then ``Zone N``, then ``Zone L``.
    """
    s = (origin or '').strip()
    if not s:
        return (2, 0, '')
    m = re.search(r'_(\d+)$', s)
    if m:
        return (0, int(m.group(1)), '')
    m = re.search(r'(?i)zone\s+(\d+)', s)
    if m:
        return (0, int(m.group(1)), '')
    m = re.search(r'(?i)zone\s+([A-Z])\b', s)
    if m:
        return (1, ord(m.group(1).upper()), '')
    return (2, 0, s.upper())


def sort_main_costs_rows_for_layout(matrix_rows):
    """
    Final row order for the MainCosts sheet (does not change pricing or zone logic).

    Rows with a non-empty **Matrix zone** are grouped by **Service**, then sorted by
    **Origin** so numeric zone indices come in order (e.g. ``..._ZONE_1`` before
    ``..._ZONE_2``, and ``Zone 1`` before ``Zone 10``). Letter zones (``Zone A``)
    sort after numeric-style origins. Rows **without** a Matrix zone keep their
    relative order (stable).

    **Lane #** is reassigned 1..n after sorting.
    """
    if not matrix_rows:
        return matrix_rows

    enumerated = list(enumerate(matrix_rows))

    def row_sort_key(entry):
        orig_idx, row = entry
        svc = (row.get('Service') or row.get('Original Service') or '').strip()
        mz = (row.get('Matrix zone') or row.get('Zone') or '').strip()
        if not mz:
            return (svc, 1, orig_idx)
        ocr = (row.get('Origin Country Region') or row.get('Origin') or '').strip()
        ot = _origin_layout_sort_tuple(ocr)
        dest = (row.get('Destination Country Region') or row.get('Destination') or '').strip()
        return (svc, 0, ot, dest, orig_idx)

    enumerated.sort(key=row_sort_key)
    out = [row for _, row in enumerated]
    for lane, row in enumerate(out, 1):
        row['Lane #'] = lane
    return out


def expand_main_costs_lanes_by_zoning(matrix_rows, zoning_matrix):
    """
    Replace abstract letter-zone rows with real Origin/Destination rows.

    PROBLEM THIS SOLVES:
    After build_matrix_main_costs() runs, some lanes have a "Matrix zone" value
    like "Zone A" instead of real origin/destination countries.  "Zone A" is just
    a code that means "all the origin/destination pairs that belong to group A".
    This function looks up those pairs and creates one concrete row per pair.

    EXAMPLE:
    Before expansion:
        Lane | Origin | Destination | Service          | Matrix zone | Price
        1    |        |             | DHL EXPRESS WW   | Zone A      | 12.50

    After expansion (if Zone A covers origin 1->dest 3 and origin 2->dest 3):
        Lane | Origin | Destination | Service          | Matrix zone | Price
        1    | Zone 1 | Zone 3      | DHL EXPRESS WW   | Zone A      | 12.50
        2    | Zone 2 | Zone 3      | DHL EXPRESS WW   | Zone A      | 12.50

    Rows that already have numeric zones (no Matrix zone value) are left unchanged.
    After all expansion is done, Lane numbers are reassigned from 1 upward.
    """
    if not matrix_rows:
        return matrix_rows

    # Build the full (matrix_name, zone_letter) -> [(origin, dest), ...] lookup
    zoning_lookup = parse_zoning_matrix(zoning_matrix)
    if not zoning_lookup:
        print("[DEBUG] expand_matrix_zones: zoning_lookup is empty; no expansion")
        return matrix_rows

    matrix_names_in_lookup = sorted({k[0] for k in zoning_lookup})
    print(f"[DEBUG] expand_matrix_zones: lookup has {len(zoning_lookup)} keys; matrix names: {matrix_names_in_lookup}")

    expanded = []
    debug_logged = set()   # (reason, service_snippet) to avoid repeating same message

    for row in matrix_rows:
        matrix_zone = (row.get('Matrix zone') or '').strip()
        service = (row.get('Service') or '').strip()

        if not matrix_zone:
            expanded.append(row)
            continue

        zone_letter = _matrix_zone_to_letter(matrix_zone)
        if not zone_letter:
            key = ("zone_letter_empty", service[:50])
            if key not in debug_logged:
                debug_logged.add(key)
                print(f"[DEBUG] expand_matrix_zones: SKIP zone_letter empty  service={service!r}  matrix_zone={matrix_zone!r} -> letter={zone_letter!r}")
            expanded.append(row)
            continue

        matrix_name = _find_matrix_for_service(zoning_lookup, service)
        if not matrix_name:
            key = ("no_matrix_name", service[:50])
            if key not in debug_logged:
                debug_logged.add(key)
                print(f"[DEBUG] expand_matrix_zones: SKIP no matrix for service  service={service!r}  matrix_zone={matrix_zone!r}  zone_letter={zone_letter!r}")
            expanded.append(row)
            continue

        key = (matrix_name, zone_letter)
        pairs = zoning_lookup.get(key, [])
        if not pairs:
            key_dbg = ("no_pairs", matrix_name, zone_letter)
            if key_dbg not in debug_logged:
                debug_logged.add(key_dbg)
                available_letters = sorted({k[1] for k in zoning_lookup if k[0] == matrix_name})
                print(f"[DEBUG] expand_matrix_zones: SKIP no pairs  service={service!r}  matrix_name={matrix_name!r}  zone_letter={zone_letter!r}  available_letters_for_this_matrix={available_letters}")
            expanded.append(row)
            continue

        key_ok = ("expanded", matrix_name, zone_letter)
        if key_ok not in debug_logged:
            debug_logged.add(key_ok)
            print(f"[DEBUG] expand_matrix_zones: OK  service={service[:45]!r}  matrix_name={matrix_name!r}  zone_letter={zone_letter!r}  -> {len(pairs)} pair(s)")

        # Create one copy of the row per (origin, destination) pair
        for origin_zone, dest_zone in pairs:
            new_row = row.copy()
            oz = f"Zone {origin_zone}" if origin_zone else ''
            dz = f"Zone {dest_zone}" if dest_zone else ''
            new_row['Origin Country Region'] = oz
            new_row['Destination Country Region'] = dz
            new_row['Origin'] = oz
            new_row['Destination'] = dz
            expanded.append(new_row)

    # Reassign Lane # sequentially after expansion
    for lane, row in enumerate(expanded, 1):
        row['Lane #'] = lane

    return expanded
