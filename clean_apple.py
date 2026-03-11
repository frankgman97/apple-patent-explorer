import json
from datetime import datetime

PREFIXES = {"Mr.", "Ms.", "Mrs.", "Dr."}
SUFFIXES = {"JR.", "Jr.", "II", "III", "IV", "V", "SR.", "Sr."}


def parse_date(val):
    """Validate and return ISO date string, or None."""
    if not val:
        return None
    try:
        datetime.strptime(val, "%Y-%m-%d")
        return val
    except ValueError:
        return val


def parse_inventor(name_text):
    """Split inventorNameText into structured fields."""
    parts = name_text.strip().split()

    prefix = None
    suffix = None

    # Extract prefix
    if parts and parts[0] in PREFIXES:
        prefix = parts.pop(0)

    # Extract suffix (may have period or not)
    if parts and (parts[-1] in SUFFIXES or parts[-1].rstrip(".") in {"JR", "SR", "II", "III", "IV", "V"}):
        suffix = parts.pop(-1)

    if not parts:
        return {
            "inventorNameText": name_text,
            "firstName": name_text,
            "middleName": None,
            "lastName": name_text.upper(),
            "prefix": prefix,
            "suffix": suffix,
        }

    # Determine last name: consecutive ALL CAPS words at the end
    # Check from the end backwards for uppercase words
    # Skip single-letter initials like "R.", "K.", "A." — those are middle initials
    last_name_parts = []
    i = len(parts) - 1
    while i >= 1:  # keep at least one word for firstName
        stripped = parts[i].replace("-", "").replace("'", "").rstrip(".")
        is_initial = len(stripped) <= 1  # single letter like "R." or "K."
        if not is_initial and stripped.isupper() and len(stripped) > 1:
            last_name_parts.insert(0, parts[i])
            i -= 1
        else:
            break

    if last_name_parts:
        # ALL CAPS surname detected
        remaining = parts[: i + 1]
        first_name = remaining[0] if remaining else None
        middle_name = " ".join(remaining[1:]) if len(remaining) > 1 else None
        last_name = " ".join(last_name_parts).upper()
    else:
        # Mixed case — last word is surname
        last_name = parts[-1].upper()
        first_name = parts[0]
        middle_name = " ".join(parts[1:-1]) if len(parts) > 2 else None

    return {
        "inventorNameText": name_text,
        "firstName": first_name,
        "middleName": middle_name,
        "lastName": last_name,
        "prefix": prefix,
        "suffix": suffix,
    }


def clean_record(rec):
    meta = rec["applicationMetaData"]

    # Parse inventors
    meta["inventorBag"] = [parse_inventor(inv["inventorNameText"]) for inv in meta.get("inventorBag", [])]

    # Uppercase inventionTitle
    if "inventionTitle" in meta:
        meta["inventionTitle"] = meta["inventionTitle"].upper()

    # Convert date fields
    for date_field in ("filingDate", "applicationStatusDate", "grantDate", "internationalRegistrationPublicationDate"):
        if date_field in meta:
            meta[date_field] = parse_date(meta[date_field])

    return rec


def main():
    with open("/Users/fgomez/Developer/Apple/Apple.json") as f:
        data = json.load(f)

    data["patentdata"] = [clean_record(rec) for rec in data["patentdata"]]

    with open("/Users/fgomez/Developer/Apple/Apple_clean.json", "w") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"Done. Cleaned {len(data['patentdata'])} records → Apple_clean.json")


if __name__ == "__main__":
    main()
