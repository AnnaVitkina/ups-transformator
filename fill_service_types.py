"""
Fill null service_type in MainCosts sections with the previous non-null value.
Reads from processing/from_rate_card_excel.json and writes the result back.
"""

import json
from pathlib import Path

INPUT_OUTPUT_FILE = Path('processing/from_rate_card_excel.json')


def fill_null_service_types(data):
    """
    Walk MainCosts section by section; for any section where service_type is null,
    set it to the previous section's non-null service_type.
    """
    main_costs = data.get('MainCosts', [])
    if not main_costs:
        return 0

    last_service_type = None
    filled_count = 0

    for section in main_costs:
        if section.get('service_type') is None and last_service_type is not None:
            section['service_type'] = last_service_type
            filled_count += 1
        if section.get('service_type') is not None:
            last_service_type = section['service_type']

    return filled_count


def main():
    print("[*] Reading", INPUT_OUTPUT_FILE)
    with open(INPUT_OUTPUT_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)

    filled = fill_null_service_types(data)
    print(f"[OK] Filled {filled} section(s) with previous service_type")

    print("[*] Writing", INPUT_OUTPUT_FILE)
    with open(INPUT_OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print("[OK] Done.")


if __name__ == '__main__':
    main()
