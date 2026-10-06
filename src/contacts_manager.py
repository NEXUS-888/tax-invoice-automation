import os
import json
import re


def normalize_phone(raw):
    """
    Returns a WhatsApp number as country code + digits (e.g. 919876543210), or "" if unusable.
    Accepts 98765 43210, 098765 43210, +91 98765 43210, 0091 98765 43210 and, when written with
    + or 00, numbers from other countries.
    """
    text = str(raw or "").strip()
    digits = ''.join(c for c in text if c.isdigit())
    international = text.startswith('+') or digits.startswith('00')
    if digits.startswith('00'):
        digits = digits[2:]
    if len(digits) == 11 and digits[0] == '0' and digits[1] in '6789':
        digits = digits[1:]  # domestic trunk prefix
    if len(digits) == 10 and digits[0] in '6789':
        return f"91{digits}"
    if len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
        return digits
    if international and 8 <= len(digits) <= 15 and digits[0] != '0' and not digits.startswith('91'):
        return digits
    return ""


class ContactsManager:
    def __init__(self, contacts_json_path):
        self.contacts_json_path = contacts_json_path
        directory = os.path.dirname(os.path.abspath(self.contacts_json_path))
        os.makedirs(directory, exist_ok=True)
        self.contacts = self.load_contacts()

    def load_contacts(self):
        """Loads contacts dictionary from JSON file."""
        if os.path.exists(self.contacts_json_path):
            try:
                with open(self.contacts_json_path, 'r', encoding='utf-8') as f:
                    contacts = json.load(f)
                if not isinstance(contacts, dict):
                    raise ValueError("contacts JSON must contain an object")
                return contacts
            except (json.JSONDecodeError, OSError, ValueError) as exc:
                raise ValueError(
                    f"Could not read contacts file '{self.contacts_json_path}': {exc}"
                ) from exc
        return {}

    def save_contacts(self):
        """Saves current contacts dictionary to JSON file."""
        with open(self.contacts_json_path, 'w', encoding='utf-8') as f:
            json.dump(self.contacts, f, indent=4, ensure_ascii=False)

    def parse_phone_list(self, raw_input):
        """
        Parses input string or list into clean WhatsApp phone numbers list.
        - Splits ONLY on commas, semicolons, slashes, or newlines.
        - Preserves single 10-digit / 12-digit phone numbers cleanly.
        """
        if not raw_input:
            return []
        if isinstance(raw_input, list):
            raw_str = ",".join(str(x) for x in raw_input)
        else:
            raw_str = str(raw_input)

        # Split ONLY by comma, semicolon, slash, or newline (NOT spaces inside a number)
        parts = re.split(r'[,;/|\n]+', raw_str)
        cleaned = []

        for p in parts:
            p = p.strip()
            if not p:
                continue
            digits = normalize_phone(p)
            if digits:
                cleaned.append(digits)

        # Deduplicate preserving order
        seen = set()
        result = []
        for d in cleaned:
            if d not in seen:
                seen.add(d)
                result.append(d)
        return result

    def _find_key(self, sheet_name):
        if not sheet_name:
            return None
        if sheet_name in self.contacts:
            return sheet_name
        stripped = str(sheet_name).strip()
        if stripped in self.contacts:
            return stripped
        clean = stripped.lower()
        for k in self.contacts.keys():
            if str(k).strip().lower() == clean:
                return k
        return None

    def get_phones(self, sheet_name):
        """Returns a list of phone number strings for an agency."""
        key = self._find_key(sheet_name)
        info = self.contacts.get(key, {}) if key else {}
        raw = info.get('phones', info.get('phone', ''))
        return self.parse_phone_list(raw)

    def get_phone_str(self, sheet_name):
        """Returns a comma-separated string of phone numbers for display."""
        phones = self.get_phones(sheet_name)
        return ", ".join(phones)

    def update_phone(self, sheet_name, phone_number_input, agency_name=""):
        """Updates or sets phone numbers for an agency."""
        key = self._find_key(sheet_name) or str(sheet_name).strip()
        if key not in self.contacts:
            self.contacts[key] = {}
        
        phone_list = self.parse_phone_list(phone_number_input)
        self.contacts[key]['phones'] = phone_list
        self.contacts[key]['phone'] = ", ".join(phone_list)
        
        if agency_name:
            self.contacts[key]['agency_name'] = agency_name
        self.save_contacts()

    def sync_agencies(self, agencies_list):
        """Ensures all detected agencies are populated in contacts file."""
        changed = False
        for a in agencies_list:
            s_name = a['sheet_name']
            key = self._find_key(s_name)
            if not key:
                clean_key = str(s_name).strip()
                self.contacts[clean_key] = {
                    'agency_name': a['agency_name'],
                    'phones': [],
                    'phone': ''
                }
                changed = True
        if changed:
            self.save_contacts()
