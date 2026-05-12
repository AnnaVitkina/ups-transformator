"""
DHL Rate Card Data Extractor

HOW THIS FILE FITS INTO THE BIGGER PICTURE
-------------------------------------------
A DHL rate card PDF is first scanned by Azure Document Intelligence (an AI service
from Microsoft).  That AI service produces a large JSON file containing all the text
and table data it found in the PDF.

This script reads that Azure JSON file and converts it into a clean, structured
Python dictionary.  That dictionary is then saved as extracted_data.json in the
processing/ folder, ready for transformation_to_excel.py (or create_table shim) to build Excel.

In short: Azure JSON  ->  this script  ->  extracted_data.json  ->  transformation_to_excel  ->  Excel
"""

import argparse    # reads command-line arguments (e.g. python main.py myfile.json)
import json        # reads and writes JSON files
import os          # used to check file sizes
import re          # used for pattern matching when searching for the carrier name in raw text
from datetime import datetime   # used to record when the extraction was run
from pathlib import Path        # cross-platform file path handling

# Words that are NOT country names; if they appear after "DHL Express" we skip that match.
# Used both when validating the structured Carrier field and when scanning raw text.
CARRIER_SKIP_FIRST_WORDS = {
    'account', 'manager', 'service', 'services', 'contact',
    'support', 'team', 'representative', 'rep', 'office', 'center', 'centre',
    'hotline', 'helpdesk', 'help', 'desk', 'portal', 'website', 'web',
    'time', 'definite', 'rates', 'rate', 'vereinbarung',
}

# UPS-style tables repeat container/package in CostName on every data row (Pkg, Cntr, …).
# Those must not start a new section; only real DHL-style cost names should (with no RateName).
PACKAGE_ROW_COSTNAME_TOKENS = frozenset({
    'pkg', 'cntr', 'pallet', 'env', 'doc', 'package', 'envelope', 'document',
})


def read_converted_json(filepath):
    """
    Open and parse the Azure Document Intelligence JSON file from disk.

    The Azure service saves its output as a single large JSON file.
    This function reads that file and returns its contents as a Python dictionary.
    It also prints some debug information (file size, top-level keys) to help
    confirm the file was loaded correctly.

    If the file doesn't exist or isn't valid JSON, an error is printed and raised.
    """
    print(f"[*] Reading JSON file: {filepath}")
    try:
        # Get the file size in megabytes so we can see how large the scan result is
        file_size_mb = os.path.getsize(filepath) / (1024 * 1024)
        print(f"    [DEBUG] File size: {file_size_mb:.2f} MB")

        # Open the file and parse it from JSON text into a Python dictionary
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        print(f"[OK] Successfully loaded JSON file")

        # Print the top-level keys so we can confirm the file has the expected structure
        top_keys = list(data.keys()) if isinstance(data, dict) else []
        print(f"    [DEBUG] Top-level keys: {top_keys}")

        # If the file has an 'analyzeResult' section (the main Azure output section),
        # print some extra details about what's inside it
        if 'analyzeResult' in data:
            ar = data['analyzeResult']
            print(f"    [DEBUG] analyzeResult keys: {list(ar.keys())}")
            if 'documents' in ar:
                print(f"    [DEBUG] Number of documents: {len(ar['documents'])}")
            if 'content' in ar:
                content_len = len(ar.get('content', ''))
                print(f"    [DEBUG] Content length (chars): {content_len:,}")

        return data

    except FileNotFoundError:
        print(f"[ERROR] File not found: {filepath}")
        raise
    except json.JSONDecodeError as e:
        print(f"[ERROR] Invalid JSON format: {e}")
        raise
    except Exception as e:
        print(f"[ERROR] Reading file: {e}")
        raise


def extract_fields(data):
    """
    Pull the structured fields out of the Azure Document Intelligence result.

    Azure organises its output like this:
      data
        analyzeResult
          documents[0]
            fields
              MainCosts   <- array of pricing rows
              AddedRates  <- array of surcharge rows
              CountryZoning <- array of country/zone rows
              DemandCosts, DemandSurcharge, DemandSurchargeCountries <- demand surcharge tables (when present)
              ... etc.

    This function navigates to documents[0].fields and returns that dictionary.
    It also prints each field's name, type, and a short preview of its value
    so you can see at a glance what was extracted from the PDF.
    """
    print("[*] Extracting fields from document...")
    try:
        analyze_result = data.get('analyzeResult', {})
        documents = analyze_result.get('documents', [])

        if not documents:
            print("[WARN] No documents found in analyzeResult")
            return {}

        # Azure puts all the structured data in the first document's 'fields' section
        fields = documents[0].get('fields', {})
        print(f"[OK] Found {len(fields)} top-level fields")
        field_names = list(fields.keys())
        print(f"    [DEBUG] Field names: {field_names}")

        # Print a short summary of each field so we can verify the extraction
        for fn in field_names:
            fv = fields[fn]
            ftype = fv.get('type', '?') if isinstance(fv, dict) else type(fv).__name__
            if ftype == 'array':
                # For array fields, show how many items were found
                arr = fv.get('valueArray', [])
                print(f"    [DEBUG]   {fn}: type={ftype}, length={len(arr)}")
            else:
                # For non-array fields, show a preview of the value (truncated to 50 chars)
                val = extract_value(fv) if isinstance(fv, dict) else fv
                preview = str(val)[:50] + "..." if val and len(str(val)) > 50 else val
                print(f"    [DEBUG]   {fn}: type={ftype}, value={repr(preview)}")

        return fields

    except Exception as e:
        print(f"[ERROR] Extracting fields: {e}")
        return {}


def extract_value(field):
    """
    Extract the actual text or number value from an Azure field object.

    Azure wraps each value in an object with a type-specific key, for example:
      { "type": "string", "valueString": "Express Worldwide" }
      { "type": "number", "valueNumber": 12.50 }
      { "type": "date",   "valueDate": "2024-01-01" }
      { "content": "some raw text" }

    This function checks each possible key in order and returns the first value found.
    Returns None if the field is empty or has no recognised value key.
    """
    if not field:
        return None

    # Try each possible value key in order of preference
    if 'valueString' in field:
        return field['valueString']
    elif 'content' in field:
        return field['content']
    elif 'valueNumber' in field:
        return field['valueNumber']
    elif 'valueDate' in field:
        return field['valueDate']

    return None


def process_main_costs_item(value_object, is_header=False):
    """
    Extract the data from a single row in the MainCosts array.

    Each row in the Azure MainCosts array is an object with keys like:
      RateName, CostName, Weight, Zone1, Zone2, Zone3, ...

    This function reads all of those keys and returns a clean dictionary with:
      - 'RateName', 'CostName', 'Weight' as top-level keys
      - all Zone* keys collected together under a 'zones' sub-dictionary

    Example input (Azure object):
      { "RateName": {...}, "Weight": {...}, "Zone1": {...}, "Zone2": {...} }

    Example output:
      { "RateName": "Express", "Weight": "0.5", "zones": {"Zone1": "10.00", "Zone2": "12.00"} }
    """
    result = {}

    # Separate zone columns from the other columns
    zones = {}
    for key, value in value_object.items():
        if key.startswith('Zone'):
            # This is a zone price/name column; collect it in the zones dict
            zone_value = extract_value(value)
            if zone_value:
                zones[key] = zone_value
        elif key in ['RateName', 'CostName', 'Weight']:
            # This is one of the standard identifier columns
            result[key] = extract_value(value)

    if zones:
        result['zones'] = zones

    return result


def process_main_costs(main_costs_field):
    """
    Convert the raw MainCosts array from Azure into a list of structured rate card sections.

    WHAT IS MAINCOSTS?
    MainCosts contains the core pricing table from the DHL rate card PDF.
    It is structured as alternating header rows and data rows:

      Header row: RateName="Express Worldwide", CostName="Fuel Surcharge", Zone1="Zone A", Zone2="Zone B"
      Data row:   Weight="0.5 kg", Zone1="10.00", Zone2="12.00"
      Data row:   Weight="1.0 kg", Zone1="14.00", Zone2="16.00"
      Header row: RateName="Economy Select", ...
      Data row:   ...

    HOW THIS FUNCTION WORKS:
    It loops through every item in the Azure array.  When it sees a row with a
    RateName or CostName, it starts a new "rate card section" (a new pricing block).
    All the data rows that follow are added to that section's 'pricing' list.
    When the next header row is found, the current section is saved and a new one starts.

    WHAT IS RETURNED:
    A list of rate card section dictionaries, each containing:
      - service_type:   the rate name (e.g. "Express Worldwide")
      - cost_category:  the cost name from the header row (e.g. "Fuel Surcharge", or UPS lane row "Cntr")
      - weight_unit:    the weight unit from the header row
      - zone_headers:   a dict mapping Zone1/Zone2/... to zone names like "Zone A"/"Zone B"
      - pricing:        a list of dicts, each with 'weight', 'zone_prices', optional 'rate_type',
                        and optional 'cost_category' when the source row has CostName (e.g. UPS "Pkg")
    """
    print("[*] Processing MainCosts data...")

    if not main_costs_field or main_costs_field.get('type') != 'array':
        print("[WARN] MainCosts is not an array")
        return []

    value_array = main_costs_field.get('valueArray', [])
    if not value_array:
        print("[WARN] MainCosts valueArray is empty")
        return []

    print(f"    [DEBUG] MainCosts valueArray length: {len(value_array)}")

    rate_cards = []          # the list of completed rate card sections we will return
    current_rate_card = None # the section we are currently building
    header_count = 0
    data_row_count = 0

    for item in value_array:
        if item.get('type') != 'object':
            continue   # skip any non-object items (shouldn't happen, but safe to check)

        value_object = item.get('valueObject', {})

        # Decide whether this row is a header row or a data row.
        # DHL: header rows have RateName / CostName (e.g. surcharge name); data rows have only Weight + zones.
        # UPS (ASSA): data rows repeat CostName as container codes (Pkg, Cntr, Pallet) — those must not
        # start a new section; only CostName values outside that set behave like DHL headers.
        cost_name_raw = extract_value(value_object.get('CostName')) if 'CostName' in value_object else None
        cost_token = str(cost_name_raw).strip().lower() if cost_name_raw else ''
        is_package_costname = cost_token in PACKAGE_ROW_COSTNAME_TOKENS

        has_rate_name = bool(extract_value(value_object.get('RateName')))
        has_cost_name_header = bool(cost_name_raw) and not is_package_costname

        if has_rate_name or has_cost_name_header:
            # --- HEADER ROW: start a new rate card section ---
            header_count += 1
            # Save the previous section (if any) before starting a new one
            if current_rate_card and current_rate_card.get('pricing'):
                rate_cards.append(current_rate_card)

            # Create a fresh rate card section for this header
            current_rate_card = {
                'service_type': extract_value(value_object.get('RateName')),
                'cost_category': extract_value(value_object.get('CostName')),
                'weight_unit': extract_value(value_object.get('Weight')),
                'zone_headers': {},   # will map Zone1 -> "Zone A", Zone2 -> "Zone B", etc.
                'pricing': []         # will hold the data rows for this section
            }

            # Extract the zone column names from the header row
            # (e.g. Zone1 -> "Zone A", Zone2 -> "Zone B")
            for key, value in value_object.items():
                if key.startswith('Zone'):
                    zone_name = extract_value(value)
                    if zone_name:
                        current_rate_card['zone_headers'][key] = zone_name

        else:
            # --- DATA ROW: add a price entry to the current section ---
            data_row_count += 1
            if current_rate_card:
                weight = extract_value(value_object.get('Weight'))
                if weight:
                    price_row = {
                        'weight': weight,
                        'zone_prices': {}   # will map Zone1 -> "10.00", Zone2 -> "12.00", etc.
                    }
                    # Per-row CostName (e.g. UPS: Pkg, Cntr, Pallet). Section-level cost_category
                    # still comes from the header row only; DHL data rows usually omit CostName.
                    if cost_name_raw:
                        price_row['cost_category'] = str(cost_name_raw).strip()
                    rt = extract_value(value_object.get('RateType'))
                    if rt:
                        price_row['rate_type'] = rt

                    # Extract the price for each zone column
                    for key, value in value_object.items():
                        if key.startswith('Zone'):
                            price = extract_value(value)
                            if price:
                                price_row['zone_prices'][key] = price

                    # Only add the row if it has at least one zone price
                    if price_row['zone_prices']:
                        current_rate_card['pricing'].append(price_row)

    # The loop ends without saving the last section; save it now
    if current_rate_card and current_rate_card.get('pricing'):
        rate_cards.append(current_rate_card)

    print(f"[OK] Processed {len(rate_cards)} rate card sections")
    print(f"    [DEBUG] Header rows: {header_count}, Data rows: {data_row_count}")

    # Print a preview of the first few sections for debugging
    for i, rc in enumerate(rate_cards[:5], 1):
        svc = (rc.get('service_type') or '(none)')[:35]
        cat = (rc.get('cost_category') or '')[:40]
        cat_suffix = '...' if len(rc.get('cost_category') or '') > 40 else ''
        nprice = len(rc.get('pricing', []))
        nzones = len(rc.get('zone_headers', {}))
        print(f"    [DEBUG]   Section {i}: service={svc!r}, category={cat!r}{cat_suffix}, pricing_rows={nprice}, zones={nzones}")
    if len(rate_cards) > 5:
        print(f"    [DEBUG]   ... and {len(rate_cards) - 5} more sections")

    return rate_cards


def process_array_field(array_field, field_name):
    """
    Convert a generic Azure array field into a plain list of dictionaries.

    This is used for all fields other than MainCosts (which has special header/data
    row logic handled by process_main_costs).  Fields like AddedRates, CountryZoning,
    ZoningMatrix, DemandCosts, DemandSurcharge, DemandSurchargeCountries, etc. are all
    simple arrays of objects.

    Each object in the Azure array looks like:
      { "type": "object", "valueObject": { "Country": {...}, "Zone": {...}, ... } }

    This function extracts the actual values from each object and returns a list of
    plain dictionaries like:
      [ { "Country": "Germany", "Zone": "Zone A" }, { "Country": "France", "Zone": "Zone B" }, ... ]

    Rows where all values are empty are skipped.
    """
    if not array_field or array_field.get('type') != 'array':
        print(f"[WARN] {field_name} is not an array or is empty")
        return []

    value_array = array_field.get('valueArray', [])
    if not value_array:
        print(f"[WARN] {field_name} valueArray is empty")
        return []

    print(f"    [DEBUG] {field_name}: valueArray length={len(value_array)}")
    results = []

    for item in value_array:
        if item.get('type') != 'object':
            continue   # skip non-object items

        value_object = item.get('valueObject', {})
        row = {}

        # Extract the actual value from each key in this object
        for key, value in value_object.items():
            extracted = extract_value(value)
            if extracted:
                row[key] = extracted

        if row:
            results.append(row)   # only add the row if it has at least one non-empty value

    if results:
        # Print the column names from the first row so we can verify the structure
        sample_keys = list(results[0].keys())
        print(f"    [DEBUG] {field_name} sample columns: {sample_keys[:12]}{'...' if len(sample_keys) > 12 else ''}")

    return results


def detect_carrier_from_content(content):
    """
    Search the raw document text for a DHL Express carrier name.

    WHY THIS IS NEEDED:
    Azure Document Intelligence extracts the carrier name into a structured 'Carrier'
    field when it can recognise it.  But sometimes that field is missing or empty
    (e.g. the PDF layout is unusual or the field wasn't labelled in the model).
    In those cases this function scans the full document text directly.

    WHAT IT LOOKS FOR:
    The carrier name always follows the pattern:
      "DHL Express <Country>"  or  "DHL EXPRESS <COUNTRY>"  (case-insensitive)

    Examples that will be matched:
      DHL EXPRESS GERMANY
      DHL Express Netherlands
      DHL Express Belgium
      DHL EXPRESS UNITED KINGDOM

    HOW IT WORKS:
    A regular expression searches the text for "DHL" followed by "EXPRESS" or "Express"
    and then one or more capitalised words (the country name).  The search is
    case-insensitive.  The first match found is returned, with its original casing
    preserved (title-cased for consistency, e.g. "DHL Express Germany").

    Returns the matched carrier name string, or None if no match is found.
    """
    if not content:
        return None

    # Pattern explanation:
    #   DHL[ \t]+       – the word "DHL" followed by one or more spaces/tabs (NOT newlines)
    #   EXPRESS?        – the word "EXPRESS" or "EXPRES" (case-insensitive via re.IGNORECASE)
    #   [ \t]+          – one or more spaces/tabs (NOT newlines)
    #   ([A-Za-z ]{2,40}) – the country name: only letters and plain spaces, no newlines.
    #                       This stops the match at the first newline or punctuation,
    #                       so "DHL EXPRESS DENMARK\nServices" captures only "DENMARK".
    # The trailing strip() in the match handler removes any trailing spaces.
    # Use the shared skip list (see CARRIER_SKIP_FIRST_WORDS at top of file).
    # Words that appear AFTER the country name (e.g. "Belgium Customer") are
    # handled by global_country()'s own stop-word stripping.
    pattern = re.compile(
        r'\bDHL[ \t]+EXPRESS?[ \t]+([A-Za-z][A-Za-z ]{1,39})',
        re.IGNORECASE
    )

    # Iterate over ALL matches and return the first one whose first word
    # is not a stop word (i.e. looks like an actual country name).
    for match in pattern.finditer(content):
        country_part = match.group(1).strip()
        first_word = country_part.split()[0].lower() if country_part else ''
        if first_word in CARRIER_SKIP_FIRST_WORDS:
            # This match is something like "DHL Express account manager" — skip it
            continue
        # Title-case for consistent formatting: "DHL Express Belgium"
        carrier = f"DHL Express {country_part.title()}"
        print(f"[OK] Carrier detected from document text: {carrier!r}")
        start = max(0, match.start() - 20)
        end = min(len(content), match.end() + 20)
        snippet = content[start:end].replace('\n', ' ')
        print(f"    [DEBUG] Context: ...{snippet}...")
        return carrier

    print("[WARN] Could not detect carrier from document text using DHL Express pattern")

    # UPS rate cards: brand appears in the document and "Country:" labels the territory.
    if re.search(r'(?i)\bups\b', content or ''):
        m = re.search(
            r'(?im)^\s*Country:\s*([^\n]+)',
            content or '',
        )
        if m:
            country = m.group(1).strip()
            if country:
                carrier = f"UPS {country}"
                print(f"[OK] Carrier detected from document text (UPS): {carrier!r}")
                return carrier

    return None


def detect_validity_from_content(content):
    """
    Search the raw document text for the validity date after the phrase "Ratecard as of:".

    WHY THIS IS NEEDED:
    Azure Document Intelligence extracts the validity date into a structured 'Validity'
    field when it recognises it.  If that field is missing or empty this function
    scans the full document text directly.

    WHAT IT LOOKS FOR:
    The date always appears immediately after the phrase "Ratecard as of:" (case-insensitive),
    separated by optional whitespace.  The date format is DD-Mon-YYYY, for example:
      Ratecard as of: 01-Mar-2025
      Ratecard as of: 15-Jan-2024

    Returns the matched date string as-is (e.g. "01-Mar-2025"), or None if not found.
    """
    if not content:
        return None

    # Pattern explanation:
    #   rate\s*card      – "Ratecard" or "Rate card" (case-insensitive)
    #   [ \t]*as[ \t]+of – "as of" with optional surrounding spaces/tabs (no newlines)
    #   :?[ \t]*         – optional colon, then optional spaces/tabs
    #   (                – start capturing the date
    #     \d{1,2}        – one or two digit day
    #     [-/]           – separator (dash or slash)
    #     [A-Za-z]{3}    – three-letter month abbreviation (Jan, Feb, Mar, ...)
    #     [-/]           – separator
    #     \d{4}          – four-digit year
    #   )                – end capture
    pattern = re.compile(
        r'rate\s*card[ \t]*as[ \t]+of:?[ \t]*(\d{1,2}[-/][A-Za-z]{3}[-/]\d{4})',
        re.IGNORECASE
    )

    match = pattern.search(content)
    if match:
        date_str = match.group(1).strip()
        print(f"[OK] Validity date detected from document text: {date_str!r}")
        start = max(0, match.start() - 10)
        end = min(len(content), match.end() + 20)
        snippet = content[start:end].replace('\n', ' ')
        print(f"    [DEBUG] Context: ...{snippet}...")
        return date_str

    # UPS zone charts / tariffs often use "Effective 22 Dec, 2024"
    m2 = re.search(
        r'\bEffective\s+(\d{1,2}\s+[A-Za-z]{3},?\s*\d{4})',
        content,
        re.IGNORECASE,
    )
    if m2:
        date_str = m2.group(1).strip()
        print(f"[OK] Validity date detected from document text (Effective …): {date_str!r}")
        return date_str

    print("[WARN] Could not detect validity date from document text using known patterns")
    return None


def detect_currency_from_content(content):
    """
    Search the raw document text for the rate-card currency after the phrase
    "All Rates stated here are in <CODE>" (case-insensitive).

    Example lines from Azure OCR:
      . All Rates stated here are in EUR.
      . All Rates stated here are in EUR. Customs duties, taxes ...

    Returns the currency code (e.g. "EUR", "GBP") or None if not found or not recognised.
    """
    if not content:
        return None

    # Optional strict list; unknown 3-letter codes still accepted below.
    CURRENCY_CODES = frozenset()

    pattern = re.compile(
        r'\bAll\s+Rates\s+stated\s+here\s+are\s+in\s+([A-Za-z]{2,4})\b',
        re.IGNORECASE,
    )
    match = pattern.search(content)
    if not match:
        # UPS: "Rate Chart Currency:\nSEK"
        match = re.search(
            r'Rate Chart Currency:\s*([A-Za-z]{3})\b',
            content,
            re.IGNORECASE,
        )
    if not match:
        print("[WARN] Could not detect document currency from known patterns")
        return None

    code = match.group(1).upper()
    if code in CURRENCY_CODES or (len(code) == 3 and code.isalpha()):
        print(f"[OK] Document currency detected from text: {code!r}")
        start = max(0, match.start() - 10)
        end = min(len(content), match.end() + 30)
        snippet = content[start:end].replace('\n', ' ')
        print(f"    [DEBUG] Context: ...{snippet}...")
        return code

    print(f"[WARN] Document currency candidate {code!r} does not look like a valid code; ignoring")
    return None


def _carrier_is_valid(carrier_value):
    """
    Check whether a carrier value extracted from the Azure 'Carrier' field looks correct.

    A carrier value is considered valid if it:
      - is not None or empty
      - for DHL: contains "DHL" (case-insensitive), and if it looks like "DHL Express <something>",
        <something> is not a skip word
      - for UPS: contains "UPS" and a plausible country/region word after it
    """
    if not carrier_value:
        return False
    s = str(carrier_value).strip()
    low = s.lower()
    if 'ups' in low:
        m = re.match(r'(?i)ups\s+(.+)', s)
        if m:
            after = m.group(1).strip()
            first_word = (after.split()[0] or '').lower()
            if first_word in CARRIER_SKIP_FIRST_WORDS:
                return False
        return True
    if 'dhl' not in low:
        return False
    # If it looks like "DHL Express <something>", check that <something> is not a skip word
    m = re.match(r'(?i)dhl\s+express\s+(.+)', s)
    if m:
        after = m.group(1).strip()
        first_word = (after.split()[0] or '').lower()
        if first_word in CARRIER_SKIP_FIRST_WORDS:
            return False
    return True


def transform_data(fields, client_name, raw_data=None):
    """
    Combine all the extracted fields into a single clean output dictionary.

    This function is the main "assembly" step.  It takes the raw fields dictionary
    from extract_fields() and calls the appropriate processing function for each
    section, then packages everything into one output structure.

    CARRIER DETECTION:
    The carrier name is read from the structured 'Carrier' field first.
    If that field is missing or doesn't look like a valid DHL carrier name,
    the function falls back to searching the full raw document text for the
    pattern "DHL Express <Country>" (handled by detect_carrier_from_content).

    Parameters:
      fields      – the structured fields dict from extract_fields()
      client_name – the detected client name (from detect_client_from_json)
      raw_data    – the full Azure JSON dict (used for carrier text-search fallback);
                    pass None to skip the fallback

    WHAT IS RETURNED:
    A dictionary with these top-level keys:
      metadata          – client name, carrier, validity date, document_currency, extraction timestamp
      MainCosts         – list of rate card sections (from process_main_costs)
      AddedRates        – list of surcharge rows
      AdditionalCostsPart1 – list of additional cost rows (first part)
      CountryZoning     – list of country-to-zone mapping rows
      AdditionalZoning  – list of additional zoning rows
      ZoningMatrix      – list of zone matrix rows
      AdditionalCostsPart2 – list of additional cost rows (second part)
      GoGreenPlusCost   – optional GoGreen Plus surcharge rows (Origin/Destination lists)
      DemandCosts       – demand surcharge flat cost rows per service (Service, Cost, …)
      DemandSurcharge   – demand surcharge matrix rows (Origin × Destination zones)
      DemandSurchargeCountries – demand surcharge zone / country list rows
      statistics        – row counts for each section (used for reporting)
    """
    print("[*] Transforming data...")

    raw_content = (raw_data or {}).get('analyzeResult', {}).get('content', '') if raw_data else ''

    # --- Carrier detection ---
    # First attempt: use the structured 'Carrier' field from Azure (DHL models)
    carrier_value = extract_value(fields.get('Carrier'))
    if _carrier_is_valid(carrier_value):
        print(f"[OK] Carrier from structured field: {carrier_value!r}")
    else:
        # UPS custom models often expose Country but not Carrier; pair with UPS when the doc is UPS.
        country_field = extract_value(fields.get('Country'))
        if country_field and re.search(r'(?i)\bups\b', raw_content or ''):
            carrier_value = f"UPS {country_field.strip()}"
            print(f"[OK] Carrier from Country field + UPS document: {carrier_value!r}")
        else:
            # Fall back to full-text patterns (DHL Express …, then UPS + Country: …).
            if carrier_value:
                print(f"[WARN] Carrier field value {carrier_value!r} does not look valid; trying text search fallback.")
            else:
                print("[WARN] Carrier field is empty; trying text search fallback.")
            carrier_value = detect_carrier_from_content(raw_content)
            if not carrier_value:
                print("[WARN] Carrier could not be determined; leaving as None.")

    # --- Validity date detection ---
    # First attempt: use the structured 'Validity' field from Azure
    validity_value = extract_value(fields.get('Validity'))
    if validity_value:
        print(f"[OK] Validity date from structured field: {validity_value!r}")
    else:
        # The structured field is missing; fall back to searching the raw document text
        # for the pattern "Ratecard as of: DD-Mon-YYYY"
        print("[WARN] Validity date field is empty; trying text search fallback.")
        validity_value = detect_validity_from_content(raw_content)
        if not validity_value:
            print("[WARN] Validity date could not be determined; leaving as None.")

    document_currency = detect_currency_from_content(raw_content) or ''

    output = {
        'metadata': {
            'client': client_name,
            'carrier': carrier_value,       # e.g. "DHL Express Germany"
            'validity_date': validity_value, # e.g. "01-Mar-2025"
            'document_currency': document_currency,  # e.g. "EUR" from "All Rates stated here are in EUR"
            'extraction_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'extraction_source': 'Azure Document Intelligence API'
        },
        'MainCosts': [],
        'AddedRates': [],
        'AdditionalCostsPart1': [],
        'CountryZoning': [],
        'AdditionalZoning': [],
        'ZoningMatrix': [],
        'AdditionalCostsPart2': [],
        'GoGreenPlusCost': [],
        'DemandCosts': [],
        'DemandSurcharge': [],
        'DemandSurchargeCountries': [],
    }

    # MainCosts has special processing (header rows + data rows) so it gets its own function
    main_costs = fields.get('MainCosts')
    if main_costs:
        output['MainCosts'] = process_main_costs(main_costs)
    else:
        print("[WARN] No MainCosts found in fields")

    print(f"    [DEBUG] Metadata: client={output['metadata']['client']!r}, carrier={str(output['metadata']['carrier'])[:40]!r}..., validity={output['metadata']['validity_date']!r}")

    # All other array fields are processed the same way: extract each row as a flat dict
    field_names = ['AddedRates', 'AdditionalCostsPart1', 'CountryZoning',
                   'AdditionalZoning', 'ZoningMatrix', 'AdditionalCostsPart2', 'GoGreenPlusCost',
                   'DemandCosts', 'DemandSurcharge', 'DemandSurchargeCountries']

    for field_name in field_names:
        field = fields.get(field_name)
        if field:
            output[field_name] = process_array_field(field, field_name)
            print(f"[OK] Processed {field_name}: {len(output[field_name])} items")
        else:
            print(f"[WARN] No {field_name} found in fields")

    # Count the total number of pricing rows across all MainCosts sections
    total_main_costs_rows = sum(len(rc.get('pricing', [])) for rc in output['MainCosts'])

    # Build a statistics summary so the pipeline can report how much data was extracted
    output['statistics'] = {
        'MainCosts_sections': len(output['MainCosts']),
        'MainCosts_rows': total_main_costs_rows,
        'AddedRates_rows': len(output['AddedRates']),
        'AdditionalCostsPart1_rows': len(output['AdditionalCostsPart1']),
        'CountryZoning_rows': len(output['CountryZoning']),
        'AdditionalZoning_rows': len(output['AdditionalZoning']),
        'ZoningMatrix_rows': len(output['ZoningMatrix']),
        'AdditionalCostsPart2_rows': len(output['AdditionalCostsPart2']),
        'GoGreenPlusCost_rows': len(output['GoGreenPlusCost']),
        'DemandCosts_rows': len(output['DemandCosts']),
        'DemandSurcharge_rows': len(output['DemandSurcharge']),
        'DemandSurchargeCountries_rows': len(output['DemandSurchargeCountries']),
    }

    print(f"[OK] Transformation complete")
    print(f"  - MainCosts sections: {output['statistics']['MainCosts_sections']}")
    print(f"  - MainCosts rows: {output['statistics']['MainCosts_rows']}")
    print(f"  - AddedRates: {output['statistics']['AddedRates_rows']} rows")
    print(f"  - AdditionalCostsPart1: {output['statistics']['AdditionalCostsPart1_rows']} rows")
    print(f"  - CountryZoning: {output['statistics']['CountryZoning_rows']} rows")
    print(f"  - AdditionalZoning: {output['statistics']['AdditionalZoning_rows']} rows")
    print(f"  - ZoningMatrix: {output['statistics']['ZoningMatrix_rows']} rows")
    print(f"  - AdditionalCostsPart2: {output['statistics']['AdditionalCostsPart2_rows']} rows")
    print(f"  - GoGreenPlusCost: {output['statistics']['GoGreenPlusCost_rows']} rows")
    print(f"  - DemandCosts: {output['statistics']['DemandCosts_rows']} rows")
    print(f"  - DemandSurcharge: {output['statistics']['DemandSurcharge_rows']} rows")
    print(f"  - DemandSurchargeCountries: {output['statistics']['DemandSurchargeCountries_rows']} rows")

    return output


def save_output(data, output_path):
    """
    Write the transformed data dictionary to a JSON file on disk.

    The file is saved with indentation (indent=2) so it's human-readable.
    ensure_ascii=False means special characters (e.g. accented letters in country names)
    are stored as-is rather than being escaped to \\uXXXX codes.

    The output folder is created automatically if it doesn't exist.
    After saving, the file size is printed as a sanity check.
    """
    print(f"[*] Saving output to: {output_path}")

    try:
        # Create the output folder if it doesn't already exist
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)

        # Write the data as formatted JSON (indent=2 makes it readable in a text editor)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        # Report the file size so we can confirm the write was successful
        file_size = os.path.getsize(output_path)
        file_size_kb = file_size / 1024

        print(f"[OK] Successfully saved output file")
        print(f"  - File size: {file_size_kb:.2f} KB")

        # Print the total number of data rows written as a final sanity check
        if 'statistics' in data:
            st = data['statistics']
            total_rows = (st.get('MainCosts_rows', 0) + st.get('AddedRates_rows', 0) +
                         st.get('AdditionalCostsPart1_rows', 0) + st.get('CountryZoning_rows', 0) +
                         st.get('ZoningMatrix_rows', 0))
            print(f"  - [DEBUG] Total data rows written: {total_rows:,}")

    except Exception as e:
        print(f"[ERROR] Saving output: {e}")
        raise


def read_client_list(filepath):
    """
    Read the list of known client names from a plain text file (one name per line).

    The file is expected to look like:
      Airbus
      BMW Group
      Siemens AG

    This list is used by detect_client_from_json() to identify which client's
    rate card is being processed.

    The function tries three different text encodings in order (utf-8, cp1252, latin-1)
    to handle files saved on different operating systems or with special characters.
    Empty lines are ignored.

    Returns a list of client name strings, or an empty list if the file is missing.
    """
    path = Path(filepath)
    if not path.exists():
        print(f"[WARN] Client file not found: {filepath}")
        return []
    try:
        encodings = ["utf-8", "cp1252", "latin-1"]
        for enc in encodings:
            try:
                with open(filepath, "r", encoding=enc) as f:
                    # Read each line, strip whitespace, and skip empty lines
                    names = [line.strip() for line in f if line.strip()]
                if enc != "utf-8":
                    print(f"    [DEBUG] Read clients with encoding: {enc}")
                print(f"    [DEBUG] Client list: {names}")
                return names
            except UnicodeDecodeError:
                continue   # this encoding didn't work; try the next one
        print(f"[WARN] Could not decode client file {filepath} with utf-8, cp1252, or latin-1")
        return []
    except Exception as e:
        print(f"[WARN] Failed to read client file {filepath}: {e}")
        return []


def detect_client_from_json(data, client_list, filename=None):
    """
    Search the full document text to find which client name appears in it.

    WHAT THIS DOES:
    The Azure Document Intelligence result includes a 'content' field that contains
    all the text extracted from the PDF as one long string.  This function searches
    that text for each client name from the client list.

    WHY LONGER NAMES FIRST:
    Names are checked longest-first to avoid false partial matches.
    For example, if the list contains both "DHL" and "DHL LLP AIRBUS", and the
    document contains "DHL LLP AIRBUS", checking "DHL" first would incorrectly
    match "DHL" inside "DHL LLP AIRBUS".  Checking "DHL LLP AIRBUS" first avoids this.

    WHAT IS RETURNED (in priority order):
    1. The first client name found in the document text.
    2. If not found in text, the first client name found in the input filename
       (e.g. "DORM RC.pdf.json" -> tries each client name against "DORM RC").
    3. If not found in filename either, the client name is extracted directly from
       the filename stem by taking the first word(s) before common suffixes like
       "RC", "rate", "ratecard" (e.g. "DORM RC.pdf" -> "DORM").
    4. If the client list is empty, "Unknown" is returned.
    """
    if not client_list:
        print("[WARN] Client list is empty, using 'Unknown'")
        return "Unknown"

    content = data.get('analyzeResult', {}).get('content', '')

    # --- Step 1: search the document text ---
    if content:
        print(f"    [DEBUG] Searching content ({len(content):,} chars) for client name...")
        sorted_names = sorted(client_list, key=len, reverse=True)
        print(f"    [DEBUG] Check order (longest first): {[n[:20] + ('...' if len(n) > 20 else '') for n in sorted_names]}")
        content_lower = content.lower()
        for name in sorted_names:
            if name.lower() in content_lower:
                idx = content_lower.find(name.lower())
                snippet = content[max(0, idx - 15):idx + len(name) + 15].replace('\n', ' ')
                print(f"[OK] Client detected in document text: {name}")
                print(f"    [DEBUG] First occurrence context: ...{snippet}...")
                return name
        print("[WARN] No client name found in document text.")
    else:
        print("[WARN] No content in JSON to search for client name.")

    # --- Step 2: search the filename for a known client name ---
    if filename:
        from pathlib import Path as _Path
        # Strip all extensions: "DORM RC.pdf.json" -> "DORM RC"
        stem = _Path(filename).name
        for ext in ('.json', '.pdf', '.xlsx', '.xls', '.csv'):
            if stem.lower().endswith(ext):
                stem = stem[:-len(ext)]
        stem_lower = stem.lower()
        sorted_names = sorted(client_list, key=len, reverse=True)
        for name in sorted_names:
            if name.lower() in stem_lower:
                print(f"[OK] Client detected from filename '{filename}': {name}")
                return name

        # --- Step 3: derive client name directly from the filename stem ---
        # Strip common non-client suffixes (RC, rate, ratecard, rates, card, dhl, express)
        # and use whatever remains as the client name.
        import re as _re
        _STOP = {'rc', 'rate', 'rates', 'ratecard', 'card', 'dhl', 'express', 'pdf'}
        words = _re.split(r'[\s_\-]+', stem)
        client_words = []
        for w in words:
            if w.lower() in _STOP:
                break
            client_words.append(w)
        if client_words:
            derived = ' '.join(client_words).strip()
            print(f"[OK] Client derived from filename '{filename}': {derived}")
            return derived

    print("[WARN] Could not determine client from document or filename, using first from list")
    return client_list[0]


# The default input folder and file path used when no command-line argument is given
INPUT_DIR = Path('input')
DEFAULT_INPUT = 'input/converted.json'


def list_input_json_files():
    """
    Return a sorted list of all .json files found in the input/ folder.
    Returns an empty list if the input/ folder doesn't exist.
    """
    if not INPUT_DIR.is_dir():
        return []
    return sorted(INPUT_DIR.glob('*.json'), key=lambda p: p.name.lower())


def choose_input_file_interactive():
    """
    Show a numbered menu of JSON files in the input/ folder and ask the user to pick one.

    The menu looks like:
      1. myfile.json  (2.34 MB)
      2. otherfile.json  (1.10 MB)
      0. Default (converted.json)

    The user types a number and presses Enter.  Option 0 uses the default file.
    Keeps asking until a valid number is entered.

    Returns the path of the selected file as a string.
    """
    files = list_input_json_files()
    if not files:
        print("[WARN] No JSON files found in input/. Using default.")
        return DEFAULT_INPUT

    print("Select input file to process:")
    print()
    for i, path in enumerate(files, 1):
        size_mb = path.stat().st_size / (1024 * 1024)
        print(f"  {i}. {path.name}  ({size_mb:.2f} MB)")
    print(f"  0. Default ({Path(DEFAULT_INPUT).name})")
    print()

    while True:
        try:
            choice = input("Enter number (0–{}): ".format(len(files))).strip()
            n = int(choice)
            if n == 0:
                return DEFAULT_INPUT
            if 1 <= n <= len(files):
                return str(files[n - 1])
        except ValueError:
            pass   # user typed something that isn't a number; ask again
        print("Invalid choice. Enter a number from the list.")


def parse_args():
    """
    Read the command-line argument for the input file path.

    If the user runs:   python main.py myfile.json
    the script uses that file.

    If the user runs:   python main.py
    (no argument), the script shows an interactive menu of files in input/.

    If a bare filename is given (no folder path), the script looks for it
    in the input/ folder automatically.

    Returns the resolved input file path as a string.
    """
    parser = argparse.ArgumentParser(
        description='Extract structured data from Azure Document Intelligence JSON.'
    )
    parser.add_argument(
        'input_file',
        nargs='?',   # '?' means the argument is optional
        default=None,
        help='Input JSON file path. If omitted, a menu of files in input/ is shown.'
    )
    args = parser.parse_args()

    if args.input_file is not None:
        p = Path(args.input_file)
        if not p.is_absolute() and len(p.parts) == 1:
            # Just a filename like "myfile.json" -> prepend the input/ folder
            resolved = INPUT_DIR / p
            return str(resolved)
        return str(p)

    # No argument given; show the interactive file picker
    return choose_input_file_interactive()


def main():
    """
    Entry point when this script is run directly from the command line.

    Runs the extraction process in six sequential steps:
      Step 1: Read the client list from addition/clients.txt
      Step 2: Load the Azure Document Intelligence JSON file
      Step 3: Detect which client this rate card belongs to
      Step 4: Extract the structured fields from the JSON
      Step 5: Transform all fields into the clean output structure
      Step 6: Save the result to processing/extracted_data.json

    If any step fails, an error message is printed and the script exits with an error.
    """
    print("=" * 60)
    print("DHL RATE CARD DATA EXTRACTOR")
    print("=" * 60)
    print()

    # Resolve the input file path (from command line or interactive menu)
    input_file = parse_args()
    output_file = 'processing/extracted_data.json'
    client_file = 'addition/clients.txt'

    print(f"[*] Input file: {input_file}")
    print()

    try:
        # Step 1: Load the list of known client names from the clients file
        print("Step 1: Reading client list...")
        client_list = read_client_list(client_file)
        print(f"[OK] Loaded {len(client_list)} client name(s) from list")
        print()

        # Step 2: Load the Azure Document Intelligence JSON file from disk
        print("Step 2: Loading input file...")
        input_data = read_converted_json(input_file)
        print()

        # Step 3: Search the document text for a matching client name
        print("Step 3: Detecting client from document content...")
        client_name = detect_client_from_json(input_data, client_list, filename=input_file)
        print(f"[OK] Client: {client_name}")
        print()

        # Step 4: Navigate to analyzeResult.documents[0].fields and return the fields dict
        print("Step 4: Extracting structured fields...")
        fields = extract_fields(input_data)
        print()

        # Step 5: Convert the raw Azure fields into our clean output structure.
        # Pass input_data as raw_data so the carrier fallback can search the full document text.
        print("Step 5: Processing and transforming data...")
        processed_data = transform_data(fields, client_name, raw_data=input_data)
        print()

        # Step 6: Write the clean output dictionary to a JSON file
        print("Step 6: Saving results...")
        save_output(processed_data, output_file)
        print()

        # Print a success summary
        print("=" * 60)
        print("[SUCCESS] EXTRACTION COMPLETE")
        print("=" * 60)
        print(f"Client: {client_name}")
        print(f"Output: {output_file}")
        print(f"\nExtracted Data:")
        print(f"  - MainCosts: {processed_data['statistics']['MainCosts_sections']} sections, {processed_data['statistics']['MainCosts_rows']} rows")
        print(f"  - AddedRates: {processed_data['statistics']['AddedRates_rows']} rows")
        print(f"  - AdditionalCostsPart1: {processed_data['statistics']['AdditionalCostsPart1_rows']} rows")
        print(f"  - CountryZoning: {processed_data['statistics']['CountryZoning_rows']} rows")
        print(f"  - AdditionalZoning: {processed_data['statistics']['AdditionalZoning_rows']} rows")
        print(f"  - ZoningMatrix: {processed_data['statistics']['ZoningMatrix_rows']} rows")
        print(f"  - AdditionalCostsPart2: {processed_data['statistics']['AdditionalCostsPart2_rows']} rows")
        print(f"  - GoGreenPlusCost: {processed_data['statistics']['GoGreenPlusCost_rows']} rows")
        print(f"  - DemandCosts: {processed_data['statistics'].get('DemandCosts_rows', 0)} rows")
        print(f"  - DemandSurcharge: {processed_data['statistics'].get('DemandSurcharge_rows', 0)} rows")
        print(f"  - DemandSurchargeCountries: {processed_data['statistics'].get('DemandSurchargeCountries_rows', 0)} rows")
        print()
        print("[DEBUG] Extraction summary:")
        print(f"  - Input: {input_file}")
        print(f"  - Output: {output_file}")
        print(f"  - Client source: detected from document content (list: {client_file})")
        print(f"  - Fields used: analyzeResult.documents[0].fields")
        print()

    except Exception as e:
        print()
        print("=" * 60)
        print("[FAILED] EXTRACTION FAILED")
        print("=" * 60)
        print(f"Error: {e}")
        print()
        raise


# Only run main() when this script is executed directly (e.g. python main.py).
# Does NOT run when imported as a module by pipeline_main.py or other scripts.
if __name__ == "__main__":
    main()


