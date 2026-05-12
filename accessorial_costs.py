"""
Accessorial Costs tab builder (approved cost-type list + fuzzy match).

Used for PDF-style ``AdditionalCostsPart1`` / ``AdditionalCostsPart2`` and for
**UPS Toolbox** rows from ``from_rate_card_excel.json`` (``AccessorialCosts2``:
``CostName``, ``Movement``, ``Market``, ``Service``, ``RateType``, ``Rate``, ``Section Nbr``).

Reference list (column ``Name`` in .xlsx / .csv):
  - Resolved from ``metadata.client``: any file in the search folders whose **stem**
    matches the client (substring or token-style match — see ``_client_matches_accessorial_ref_filename``).
  - Search order: optional ``accessorial_folder`` argument, then paths from the
    ``UPS_ACCESSORIAL_FOLDER`` environment variable (``os.pathsep``-separated list),
    then :data:`FALLBACK_ACCESSORIAL_FOLDER` if set, then ``addition/Accessorial Costs``,
    then ``addition/`` (next to this file).

Public:
  ``build_accessorial_costs_rows`` — returns ``(rows, ref_file_used)``.

Private helpers:
  ``_load_accessorial_cost_type_names``, ``_token_set``, ``_best_match_cost_type``
"""

# Acceptance thresholds for mapping Original Cost Name → Cost Type (see _best_match_cost_type).
ACCESSORIAL_MATCH_MIN_SCORE = 0.35
ACCESSORIAL_MATCH_MIN_MARGIN = 0.06

import difflib
import os
import re
from pathlib import Path

# Optional: single folder to search (after ``accessorial_folder``). Team Colab default
# below; override with ``UPS_ACCESSORIAL_FOLDER`` or clear to "" for local-only runs.
FALLBACK_ACCESSORIAL_FOLDER: str = (
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT UPS/addition/Accessorial Costs"
)

# Known currency codes (most popular) — used to detect and strip currency from Cost Price
CURRENCY_CODES = frozenset({
    'EUR', 'USD', 'GBP', 'CHF', 'JPY', 'TRY', 'CAD', 'AUD', 'CNY', 'INR',
    'BRL', 'MXN', 'KRW', 'SGD', 'HKD', 'NOK', 'SEK', 'DKK', 'PLN', 'CZK',
    'RUB', 'ZAR', 'AED', 'SAR', 'ILS', 'THB', 'IDR', 'MYR', 'PHP', 'RON',
    'MAD',
})


def _clean_currency_and_price(raw_price, raw_currency):
    """
    Separate the currency code from numeric/text noise. Currency is detected
    from the Cost Price cell using a known list of codes, then stripped from
    the price.

    PROBLEM:
    The extracted values often mix numbers and currency codes together, e.g.:
        CostPrice   = "0.50 EUR met een minimum van 24.00 EUR"
        CostCurrency = "0.50 EUR"

    WHAT WE WANT:
        Currency   = "EUR"          (letters only, from known list)
        Cost Price = "0.50  met een minimum van 24.00"  (currency code removed)

    STEPS:
    1. Look in raw_price for any known currency code (from CURRENCY_CODES)
       as a whole word; use the first match (e.g. "EUR").
    2. If none found in raw_price, try to extract from raw_currency (e.g. "0.50 EUR" -> "EUR").
    3. If a currency code was found, remove every occurrence of it from
       raw_price (case-insensitive).
    4. Collapse double spaces, normalise decimal separator (comma -> dot), strip.

    If no currency code can be extracted, both values are returned unchanged.
    """
    raw_price = str(raw_price or '').strip()
    raw_currency = str(raw_currency or '').strip()
    currency_code = None

    # Step 1: find a known currency code in the Cost Price cell (whole word)
    price_upper = raw_price.upper()
    for code in sorted(CURRENCY_CODES, key=len, reverse=True):  # longer first (e.g. avoid "IN" matching before "INR")
        if re.search(r'\b' + re.escape(code) + r'\b', price_upper):
            currency_code = code
            break

    # Step 2: fallback — extract from raw_currency (e.g. "0.50 EUR" -> "EUR")
    if not currency_code:
        currency_match = re.search(r'\b([A-Z]{2,4})\b', raw_currency)
        if currency_match:
            currency_code = currency_match.group(1)

    if not currency_code:
        return raw_price, raw_currency

    # Step 3: remove the currency code from the price string (all occurrences)
    cleaned_price = re.sub(r'\b' + re.escape(currency_code) + r'\b', '', raw_price, flags=re.IGNORECASE)
    cleaned_price = re.sub(r'  +', ' ', cleaned_price).strip()

    # Step 4: normalise decimal separator
    cleaned_price = cleaned_price.replace(',', '.')

    return cleaned_price, currency_code


def _normalize_client_for_accessorial_match(client: str) -> str:
    """
    Strip suffixes that appear in rate-card metadata but not in reference filenames.

    Examples: ``\"ASICS    - 2020\"`` → ``\"ASICS\"`` so ``\"asics\"`` can match
    ``ASICS_Accessorial.xlsx`` (full client string is no longer required to be a
    substring of the stem).
    """
    c = (client or "").strip()
    if not c:
        return c
    # Trailing " - 2020", " – 2020", " -2020" (common client + validity year labels)
    c = re.sub(r"\s*[-–]\s*\d{4}\s*$", "", c, flags=re.IGNORECASE).strip()
    return c


def _extra_accessorial_search_dirs() -> list[Path]:
    """Paths from env / module fallback, in order, de-duplicated."""
    out: list[Path] = []
    seen: set[str] = set()

    def _add(raw: str) -> None:
        raw = (raw or "").strip()
        if not raw:
            return
        p = Path(raw)
        key = str(p.resolve()) if p.exists() else str(p)
        if key not in seen:
            seen.add(key)
            out.append(p)

    env = os.environ.get("UPS_ACCESSORIAL_FOLDER") or ""
    for part in env.split(os.pathsep):
        _add(part)
    _add(FALLBACK_ACCESSORIAL_FOLDER)
    return out


def _client_matches_accessorial_ref_filename(client: str, file_stem: str) -> bool:
    """
    True if ``file_stem`` plausibly belongs to this client (for picking the approved list file).

    Tries, in order:
      1. Full client string (lower) is a substring of the stem.
      2. Alphanumeric prefix of client (first 12 chars) appears inside alphanumeric stem.
      3. At least two significant tokens (length ≥ 3, not pure digits) from the client
         each appear as substrings in the stem (e.g. ``assa`` + ``abloy`` in ``Assa_Abloy_Accessorial``).
      4. A single significant token (length ≥ 4) appears in the stem (covers short names like ``ASICS``).
    """
    c = _normalize_client_for_accessorial_match(client)
    if not c:
        return False
    cl = c.lower()
    st = (file_stem or "").lower()
    if cl in st:
        return True

    def _alnum(s: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", s.lower())

    ca, sa = _alnum(c), _alnum(file_stem)
    if len(ca) >= 8 and ca[:12] in sa:
        return True

    tokens = [t for t in re.split(r"[\s_,\-]+", cl) if len(t) >= 3 and not t.isdigit()]
    if len(tokens) >= 2:
        if all(t in st for t in tokens[:2]):
            return True
    if len(tokens) == 1 and len(tokens[0]) >= 4 and tokens[0] in st:
        return True
    return False


def _toolbox_accessorial_costs2_item_to_row(item: dict, metadata: dict) -> dict:
    """Map ``AccessorialCosts2`` JSON (UPS Toolbox) to ``ACCESSORIAL_COSTS_COLUMNS`` row shape."""
    carrier = (metadata.get("carrier") or "").replace("\n", " ")
    validity_date = metadata.get("validity_date") or ""
    doc_currency = (metadata.get("document_currency") or "").strip()

    rate = str(item.get("Rate") or "").strip()
    rate_type = str(item.get("RateType") or "").strip()
    movement = str(item.get("Movement") or "").strip()
    market = str(item.get("Market") or "").strip()
    service = str(item.get("Service") or "").strip()
    apply_parts = [p for p in (movement, market, service) if p]
    apply_over = " / ".join(apply_parts)

    return {
        "Original Cost Name": str(item.get("CostName") or "").strip(),
        "Cost Type": "",
        "Cost Price": rate,
        "Minimum": "",
        "Currency": doc_currency,
        "Rate by": rate_type,
        "Apply Over": apply_over,
        "Apply if": "",
        "Additional info(Cost Code)": str(item.get("Section Nbr") or "").strip(),
        "Valid From": validity_date,
        "Valid To": "",
        "Carrier": carrier,
    }


def _split_minimum_from_cost_price(cost_price_str):
    """
    If Cost Price contains a "minimum" phrase (e.g. "0.50 with minimum of 30.00"),
    split into the base price and the minimum value.

    Supported patterns (case-insensitive):
      - "X with minimum of Y"  -> Cost Price = X, Minimum = Y
      - "X minimum of Y"      -> Cost Price = X, Minimum = Y
      - "X met een minimum van Y" (Dutch) -> Cost Price = X, Minimum = Y

    Returns (cost_price_only, minimum_value) where minimum_value is '' if no match.
    """
    s = str(cost_price_str or '').strip()
    if not s:
        return s, ''

    # Match "with minimum of 30.00" or "minimum of 30.00" or "met een minimum van 24.00"
    patterns = [
        (r'\s+with\s+minimum\s+of\s+([0-9]+(?:\.[0-9]+)?)', 1),
        (r'\s+minimum\s+of\s+([0-9]+(?:\.[0-9]+)?)', 1),
        (r'\s+met\s+een\s+minimum\s+van\s+([0-9]+(?:\.[0-9]+)?)', 1),
    ]
    for pattern, group in patterns:
        m = re.search(pattern, s, re.IGNORECASE)
        if m:
            minimum_val = m.group(1)
            cost_only = (s[: m.start()] + s[m.end() :]).strip()
            cost_only = re.sub(r'  +', ' ', cost_only)
            return cost_only, minimum_val
    return s, ''


def _load_accessorial_cost_type_names(ref_path):
    """
    Read the list of approved/canonical cost type names from a reference file.

    PURPOSE:
    The rate card PDF uses its own names for costs (e.g. "Premium 9:00 Delivery").
    The business wants these mapped to standardised names from an approved list
    (e.g. "9:00 Service Fee").  This function loads that approved list.

    The reference file must have a column called 'Name'.  Supported file formats:
      - Excel (.xlsx or .xls)
      - CSV (.csv)

    Returns a deduplicated list of name strings, in the order they appear in the file.
    Returns an empty list [] if the file doesn't exist or has no 'Name' column.
    """
    ref_path = Path(ref_path)
    if not ref_path.exists():
        return []
    names = []
    try:
        if ref_path.suffix.lower() in ('.xlsx', '.xls'):
            import openpyxl
            wb = openpyxl.load_workbook(ref_path, read_only=True, data_only=True)
            ws = wb.active
            header = None
            name_col = None
            for row in ws.iter_rows(values_only=True):
                if header is None:
                    header = [str(c).strip() if c is not None else '' for c in row]
                    for i, h in enumerate(header):
                        if h == 'Name':
                            name_col = i
                            break
                    if name_col is None:
                        break
                    continue
                if name_col is not None and name_col < len(row):
                    val = row[name_col]
                    if val is not None and str(val).strip():
                        names.append(str(val).strip())
            wb.close()
        elif ref_path.suffix.lower() == '.csv':
            import csv
            with open(ref_path, 'r', encoding='utf-8-sig', newline='') as f:
                reader = csv.reader(f)
                header = next(reader, None)
                if header:
                    try:
                        name_col = header.index('Name')
                    except ValueError:
                        name_col = None
                    if name_col is not None:
                        for row in reader:
                            if name_col < len(row) and row[name_col].strip():
                                names.append(row[name_col].strip())
        else:
            return []
    except Exception:
        return []

    # Remove duplicates while keeping the original order
    return list(dict.fromkeys(names))


def _token_set(text):
    """
    Break a text string into a set of individual words (tokens) in lowercase.
    Also handles time-like tokens such as "9:00" as a single token (not split at the colon).

    Example: "Premium 9:00 Delivery Fee" -> {"premium", "9:00", "delivery", "fee"}

    This is used by the fuzzy matching function to compare cost names word-by-word.
    """
    import re
    s = (text or '').lower().strip()
    tokens = set(re.findall(r'[a-z0-9]+(?::[a-z0-9]+)?|[a-z]+', s))
    return tokens


def _best_match_cost_type(
    original_name,
    name_list,
    min_score=None,
    min_margin=None,
):
    """
    Find the best matching canonical cost type name for a given original cost name.

    WHY FUZZY MATCHING?
    The cost names in the rate card PDF (e.g. "Premium 9:00:") don't always match
    exactly the standardised names in the reference file (e.g. "9:00 Service Fee").
    We use a scoring system to find the closest match.

    HOW THE SCORE WORKS:
    For each candidate name in the reference list, we compute a combined score:
      score = character_similarity + token_overlap_bonus

      character_similarity: a 0-to-1 score from Python's difflib library that measures
                            how similar two strings look character by character.
                            e.g. "Premium 9:00" vs "9:00 Service Fee" -> ~0.35

      token_overlap_bonus:  an extra bonus (up to 0.4) for shared meaningful words.
                            "Meaningful" means the word is at least 2 characters long
                            or contains ":" (to catch time codes like "9:00").
                            e.g. both contain "9:00" -> bonus = 0.4

    ACCEPTANCE (stricter than score-only):
      - The highest-scoring candidate must be at least ``min_score`` (default
        ``ACCESSORIAL_MATCH_MIN_SCORE``).
      - If there is a runner-up, its score must be lower by at least ``min_margin``
        (default ``ACCESSORIAL_MATCH_MIN_MARGIN``); otherwise we return '' to avoid
        ambiguous or weak matches (e.g. accidental overlap on common words).
      - If only one candidate exists in the list, only the minimum-score rule applies.
    """
    if min_score is None:
        min_score = ACCESSORIAL_MATCH_MIN_SCORE
    if min_margin is None:
        min_margin = ACCESSORIAL_MATCH_MIN_MARGIN

    if not original_name or not name_list:
        return ''
    original = str(original_name).strip()
    if not original:
        return ''

    orig_tokens = _token_set(original)
    scored = []

    for name in name_list:
        name_str = str(name).strip()
        if not name_str:
            continue

        char_ratio = difflib.SequenceMatcher(None, original.lower(), name_str.lower()).ratio()

        name_tokens = _token_set(name_str)
        shared = orig_tokens & name_tokens

        meaningful_orig = {t for t in orig_tokens if len(t) >= 2 or ':' in t}
        if meaningful_orig:
            token_bonus = (len(shared & meaningful_orig) / len(meaningful_orig)) * 0.4
        else:
            token_bonus = 0.0

        score = char_ratio + token_bonus
        scored.append((score, name_str))

    if not scored:
        return ''

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_name = scored[0]

    if best_score < min_score:
        return ''

    if len(scored) > 1:
        second_score = scored[1][0]
        if (best_score - second_score) < min_margin:
            return ''

    return best_name


def build_accessorial_costs_rows(
    additional_costs_1,
    additional_costs_2,
    metadata,
    cost_type_ref_path=None,
    accessorial_folder=None,
    *,
    accessorial_costs_2_toolbox=None,
):
    """
    Build the rows for the "Accessorial Costs" Excel tab.

    Sources:
      - ``AdditionalCostsPart1`` / ``AdditionalCostsPart2`` (PDF-style JSON fields).
      - ``accessorial_costs_2_toolbox``: ``AccessorialCosts2`` from ``from_rate_card_excel.json``
        (UPS Toolbox: ``CostName``, ``Movement``, ``Market``, ``Service``, ``RateType``, ``Rate``, ``Section Nbr``).

    Cost Type fuzzy match:
      - Loads approved names from a reference file (column ``Name``) chosen by ``metadata.client``:
        first matching file in ``accessorial_folder`` (if passed), then ``UPS_ACCESSORIAL_FOLDER`` /
        :data:`FALLBACK_ACCESSORIAL_FOLDER`, then ``addition/Accessorial Costs/``, then ``addition/``
        (see ``_client_matches_accessorial_ref_filename``).

    Returns: ``(list_of_rows, path_of_reference_file_used_or_None)``
    """
    carrier = (metadata.get('carrier') or '').replace('\n', ' ')
    validity_date = (metadata.get('validity_date') or '')

    def item_to_row(item):
        """Convert one JSON cost item into a row dict matching ACCESSORIAL_COSTS_COLUMNS."""
        raw_price    = item.get('CostPrice') or item.get('CostAmount') or ''
        raw_currency = item.get('CostCurrency', '')
        cost_price, currency = _clean_currency_and_price(raw_price, raw_currency)
        cost_price, minimum = _split_minimum_from_cost_price(cost_price)
        return {
            'Original Cost Name': item.get('CostName', ''),
            'Cost Type': '',                                    # filled later by fuzzy matching
            'Cost Price': cost_price,
            'Minimum': minimum,
            'Currency': currency,
            'Rate by': item.get('PriceMechanism', ''),
            'Apply Over': item.get('ApplyTo', ''),
            'Apply if': '',
            'Additional info(Cost Code)': item.get('CostCode', ''),
            'Valid From': validity_date,
            'Valid To': '',
            'Carrier': carrier,
        }

    rows = []
    for item in additional_costs_1 or []:
        rows.append(item_to_row(item))
    for item in additional_costs_2 or []:
        rows.append(item_to_row(item))
    for item in accessorial_costs_2_toolbox or []:
        if isinstance(item, dict):
            rows.append(_toolbox_accessorial_costs2_item_to_row(item, metadata))

    # -----------------------------------------------------------------------
    # Find the reference file for Cost Type fuzzy matching.
    # Search order:
    #   1. accessorial_folder (if set)
    #   2. UPS_ACCESSORIAL_FOLDER (env) and FALLBACK_ACCESSORIAL_FOLDER (module constant)
    #   3. addition/Accessorial Costs/
    #   4. addition/
    # Filename must match metadata.client (see _client_matches_accessorial_ref_filename).
    # -----------------------------------------------------------------------
    if cost_type_ref_path is None:
        client = (metadata.get('client') or '').strip()
        ext_order = ('.xlsx', '.xls', '.csv')

        search_dirs: list[Path] = []
        if accessorial_folder:
            search_dirs.append(Path(accessorial_folder))
        for d in _extra_accessorial_search_dirs():
            if d not in search_dirs:
                search_dirs.append(d)
        local_addition = Path(__file__).resolve().parent / 'addition'
        local_accessorial_costs = local_addition / 'Accessorial Costs'
        for d in (local_accessorial_costs, local_addition):
            if d not in search_dirs:
                search_dirs.append(d)

        print("[*] Accessorial Cost Type mapping: searching for client reference file...")
        print(f"    Client: {client or '(none)'}")
        print(f"    Search folders: {[str(d) for d in search_dirs]}")

        if not client:
            print("[*] Accessorial cost mapping: no client in metadata, Cost Type left empty")
        else:
            for search_dir in search_dirs:
                if not search_dir.exists() or not search_dir.is_dir():
                    print(f"    [SKIP] Folder not found: {search_dir}")
                    continue
                candidates = [
                    p for p in search_dir.iterdir()
                    if p.is_file()
                    and p.suffix.lower() in ext_order
                    and _client_matches_accessorial_ref_filename(client, p.stem)
                ]
                if candidates:
                    cost_type_ref_path = min(
                        candidates,
                        key=lambda p: ext_order.index(p.suffix.lower()) if p.suffix.lower() in ext_order else 99,
                    )
                    print(f"[*] Accessorial cost mapping: found '{cost_type_ref_path.name}' in {search_dir}")
                    break
                else:
                    print(f"    [MISS] No client-matching reference file in {search_dir}")

            if cost_type_ref_path is None:
                print(f"[*] Accessorial cost mapping: no reference file found for client '{client}', Cost Type left empty")

    if cost_type_ref_path:
        name_list = _load_accessorial_cost_type_names(cost_type_ref_path)
        if name_list:
            for row in rows:
                original = row.get('Original Cost Name', '')
                row['Cost Type'] = _best_match_cost_type(original, name_list)
            print(f"[*] Accessorial Cost Type: filled from {cost_type_ref_path.name} ({len(name_list)} cost types, {len(rows)} rows)")
        else:
            print(f"[*] Accessorial Cost Type: file {cost_type_ref_path.name} has no 'Name' column or is empty, Cost Type left blank")

    return rows, cost_type_ref_path
