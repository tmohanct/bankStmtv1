"""Positioned, first-page-only evidence for the PDF_Status worksheet.

This module never reads transaction rows to fill missing identity fields.
Bank-specific names and layout hints belong to src/parsers/*_parser.py.
"""
from __future__ import annotations

import importlib
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from functools import lru_cache
from pathlib import Path

import fitz

LABELS = ('Customer Name', 'Bank Name', 'Account Number', 'Address', 'Statement Date Between')
REQUIRED = ('Customer Name', 'Bank Name', 'Account Number', 'Statement Date Between')
NAME_LABEL = r'(?:(?:Main|Primary)\s+)?(?:Account\s+Holder(?:\(s\)|s)?\s+Names?|Account\s+Name|Customer\s+Name|A\s*/\s*C\s+Name|(?<!Branch )(?<!Nominee )(?<!Product )(?<!Scheme )Name)'
ACCOUNT_LABEL = r'(?:Account|Acc(?:ount|t)?\.?|A\s*/\s*C)\s*(?:Number|No\.?|#)(?:\s*\(\s*\d+\s*DIGIT\s*\))?|A\s*/\s*C(?=\s*[:\-]?\s*[0-9Xx*]{6})'
ADDRESS_LABEL = r"(?:Customer(?:'s)?\s+Address|Address\s+of\s+Customer|Communication(?:\s+Address)?|Mailing\s+Address|Address)"
STOP_LABEL = re.compile(
    rf'\b(?:{NAME_LABEL}|{ACCOUNT_LABEL}|{ADDRESS_LABEL}|Branch(?:\s+(?:Name|Address|Code|Email|Phone))?|'
    r'(?:Account|A\s*/\s*C|Acc\.?)\s+(?:Type|Currency|Status|Open(?:ing)?\s+Date|Description|Branch)|Customer\s*(?:/CIF)?\s*ID|'
    r'CIF\s*(?:No\.?|Number)|(?:RTGS/NEFT\s+)?IFS(?:C|\s+Code)?(?:\s+Code)?|MICR(?:\s+Code)?|'
    r'(?:CKYC|CKYCR|Entity\s+CKYC|Primary\s+GSTIN|GSTIN|Email|E-mail|Mobile|Phone|Contact|Tel|Fax|Nomination|Nominee|'
    r'Joint\s+Holders?|Mode\s+of\s+Operation|Scheme|Product|Currency|Drawing\s+Power|Interest\s+Rate|'
    r'Opening\s+Balance|Closing\s+Balance|Current\s+Balance|Cleared\s+Balance|Effective\s+Available\s+Balance|'
    r'Address\s+Last\s+Updated|Regd\.\s+Mobile|Statement\s+Date|Statement\s+Period|Transaction\s+Period|'
    r'From\s+Date|To\s+Date|Searched\s+by|Home\s+Branch\s+Details|CRN)\b)', re.I)
MONTH = r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
DATE_TOKEN = rf'(?:\d{{4}}[-/]\d{{1,2}}[-/]\d{{1,2}}|\d{{1,2}}[-/.]\d{{1,2}}[-/.]\d{{2,4}}|\d{{1,2}}[-\s]+{MONTH}[-\s]+[\x27\u2019]?\d{{2,4}}|{MONTH}\s+\d{{1,2}},?\s+\d{{4}})'
DATE_RE = re.compile(DATE_TOKEN, re.I)


def clean(value: str) -> str:
    return re.sub(r'\s+', ' ', value or '').strip(' :\t\n')


@lru_cache(maxsize=1)
def bank_profiles() -> tuple[dict, ...]:
    # The legacy CLI puts src/code before src. Keep that order for its imports.
    source = Path(__file__).resolve().parents[1]
    if str(source) not in sys.path:
        sys.path.append(str(source))
    profiles = []
    for path in sorted((source / 'parsers').glob('*_parser.py')):
        module = importlib.import_module('parsers.' + path.stem)
        profile = getattr(module, 'PDF_STATUS_PROFILE', None)
        if profile:
            profiles.append(profile)
    return tuple(profiles)


@dataclass
class Row:
    words: list

    @property
    def text(self):
        return ' '.join(w[4] for w in self.words)

    @property
    def y(self):
        return min(w[1] for w in self.words)

    @property
    def height(self):
        return max(w[3] - w[1] for w in self.words)

    def x_at(self, character):
        offset = 0
        for word in self.words:
            if offset + len(word[4]) > character:
                return word[0]
            offset += len(word[4]) + 1
        return self.words[-1][2]

    def crop(self, left=0, right=float('inf')):
        return clean(' '.join(w[4] for w in self.words if left - 2 <= w[0] < right - 2))


def positioned_rows(words) -> list[Row]:
    groups = []
    for word in sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        if not clean(word[4]):
            continue
        center = (word[1] + word[3]) / 2
        if groups and abs(center - groups[-1][0]) <= max(3, (word[3] - word[1]) * .75):
            groups[-1][1].append(word)
        else:
            groups.append((center, [word]))
    return [Row(sorted(group, key=lambda w: w[0])) for _, group in groups]


def header_rows(rows: list[Row]) -> list[Row]:
    for index, row in enumerate(rows):
        # Transaction table headers can wrap over adjacent lines.
        nearby = ' '.join(r.text for r in rows[index:index + 3] if r.y - row.y < row.height * 3)
        tokens = sum(bool(re.search(pattern, nearby, re.I)) for pattern in (
            r'\b(?:date|dated)\b', r'\b(?:particulars|narration|description|remarks|transaction\s+details)\b',
            r'\b(?:debit|withdrawals?|withdrawl|dr)\b', r'\b(?:credit|deposits?|cr)\b', r'\b(?:balance|amount)\b'))
        if re.search(r'(?i)statement|period|\bfrom\b|balance\s+as\s+on', row.text):
            continue
        if not re.search(r'(?i)\b(date|particulars|narration|description|remarks|debit|credit|withdrawals?)\b', row.text):
            continue
        if tokens >= 4 or (tokens >= 3 and re.search(r'\b(?:chq|cheque|txn|tran)\b', nearby, re.I)):
            return rows[:index]
    # Do not treat a page without a recognizable table header as all identity data.
    for index, row in enumerate(rows):
        if DATE_RE.match(row.text) and re.search(r'\b(?:UPI|NEFT|IMPS|RTGS|TRANSFER)\b', row.text, re.I):
            return rows[:index]
    return rows


@dataclass
class Header:
    rows: list[Row]
    width: float
    is_ocr: bool = False

    @property
    def text(self):
        return '\n'.join(r.text for r in self.rows)

    def field(self, pattern, multiline=False) -> tuple[str, int | None]:
        for index, row in enumerate(self.rows):
            match = re.search(r'(?<!\w)(?:' + pattern + r')\s*[:\-]?', row.text, re.I)
            if not match:
                continue
            left = row.x_at(match.start())
            next_labels = [m for m in STOP_LABEL.finditer(row.text, match.end())]
            end = next_labels[0].start() if next_labels else len(row.text)
            right = row.x_at(end) if next_labels else self.width
            # A label/value pair in the left column must not consume the right column.
            if left < self.width * .4:
                peer_labels = [r.x_at(m.start()) for r in self.rows
                               for m in STOP_LABEL.finditer(r.text)
                               if self.width * .45 < r.x_at(m.start()) < self.width * .98]
                if right == self.width and len(peer_labels) >= 2:
                    right = min(peer_labels)
                # Some banks right-align both the branch labels and their values.
                if right > self.width * .80 and peer_labels:
                    right = self.width * .55
            value = clean(row.text[match.end():end]).lstrip('-: ')
            # Re-crop at the actual boundary if a neighbouring value has no label on this row.
            if right < self.width:
                first_value_x = row.x_at(match.end())
                value_words = [w for w in row.words if first_value_x - 1 <= w[0] < row.x_at(end)]
                # Preserve a full-width inline value if no physical column gap exists.
                crossing = [i for i, w in enumerate(value_words) if w[0] >= right]
                if not next_labels and crossing and (crossing[0] == 0 or value_words[crossing[0]][0] - value_words[crossing[0] - 1][2] < 14):
                    right = self.width
                value = row.crop(first_value_x, right).lstrip('-: ')
                if first_value_x <= row.x_at(match.start()):
                    value = clean(row.text[match.end():end]).lstrip('-: ')
            parts = [value] if value else []
            previous_y = row.y
            for following in self.rows[index + 1:]:
                if parts and not multiline:
                    break
                if following.y - previous_y > max(row.height * 2.8, 24):
                    break
                text = following.crop(left, right)
                if not text:
                    continue
                # Multiline labels such as "Communication / address" may span two rows.
                if re.search('Communication', pattern, re.I):
                    text = re.sub(r'(?i)^address\b(?!\s+Last)\s*:?\s*', '', text)
                nearby_label = any(
                    following.y <= later.y <= following.y + row.height * .7
                    and any(abs(later.x_at(m.start()) - left) < 12 for m in STOP_LABEL.finditer(later.text))
                    for later in self.rows[index + 1:]
                )
                if re.search('Communication', pattern, re.I) and re.match(r'(?i)^address\b(?!\s+Last)', following.crop(left, right)):
                    nearby_label = False
                if nearby_label or STOP_LABEL.match(text) or re.match(r'(?i)(?:statement|account\s+(?:statement|summary|activity|details)|date\b|type\s*:|status\s*:)', text):
                    break
                parts.append(text.lstrip('-: '))
                previous_y = following.y
                if len(parts) >= (8 if multiline else 1):
                    break
            return clean(' '.join(parts)), index
        return '', None


def normalized_date(value: str) -> str:
    value = re.sub(r"[\x27\u2019]", '', clean(value)).replace('Sept ', 'Sep ')
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y',
                '%d/%m/%y', '%d-%m-%y', '%d %b %Y', '%d %B %Y', '%d-%b-%Y', '%d-%B-%Y',
                '%d %b %y', '%d %B %y', '%d-%b-%y', '%b %d, %Y', '%B %d, %Y', '%b %d %Y', '%B %d %Y'):
        try:
            return datetime.strptime(value, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return ''


def statement_period(header: Header) -> str:
    candidates = []
    for label in (r'Transaction\s+Period', r'Statement\s+Period', r'St\.\s*Period'):
        value, _ = header.field(label, multiline=True)
        if value:
            candidates.append(value)
    candidates.append(header.text)
    for text in candidates:
        for match in re.finditer(rf'(?P<start>{DATE_TOKEN})\s*(?:to\s*:?|and|[-\u2013\u2014])\s*(?P<end>{DATE_TOKEN})', text, re.I):
            start, end = normalized_date(match['start']), normalized_date(match['end'])
            if start and end and start <= end:
                return f'{start} to {end}'
        match = re.search(rf'from\s*(?:date)?\s*[:\-]?\s*({DATE_TOKEN})\s*to\s*(?:date)?\s*[:\-]?\s*({DATE_TOKEN})', text, re.I)
        if match:
            start, end = map(normalized_date, match.groups())
            if start and end and start <= end:
                return f'{start} to {end}'
    start, _ = header.field(r'From\s+Date|Period\s+From')
    end, _ = header.field(r'To\s+Date|Period\s+To')
    start, end = normalized_date(start), normalized_date(end)
    return f'{start} to {end}' if start and end and start <= end else ''


def identify_bank(header: Header) -> tuple[dict, str]:
    profiles = bank_profiles()
    # Only branch IFSC labels in the first-page header, never transaction counterparties.
    ifsc, _ = header.field(r'(?:RTGS/NEFT\s+)?IFS(?:C|\s+Code)(?:\s+Code)?')
    for profile in profiles:
        if re.search(r'\b' + profile['ifsc'] + r'0[A-Z0-9]{6}\b', ifsc, re.I):
            return profile, 'Page 1 branch IFSC'
    # Some forms print the IFSC value before its label.
    if re.search(r'\bIFS(?:C|\s+Code)\b', header.text, re.I):
        found = [p for p in profiles if re.search(r'\b' + p['ifsc'] + r'0[A-Z0-9]{6}\b', header.text, re.I)]
        if len(found) == 1:
            return found[0], 'Page 1 header IFSC'
    matches = []
    for profile in profiles:
        for name in profile['aliases']:
            match = re.search(r'\b' + (r'[A-Z]?' if header.is_ocr and len(name) > 6 else '') + re.escape(name) + r'\b', header.text, re.I)
            if match:
                matches.append((match.start(), -len(name), profile))
        if profile.get('header_pattern') and re.search(profile['header_pattern'], header.text, re.I):
            return profile, 'Page 1 bank product label'
    if matches:
        return min(matches, key=lambda m: m[:2])[2], 'Page 1 bank name'
    return {}, ''


def _unlabelled_block(header: Header, profile: dict, name: str) -> tuple[str, str]:
    rows = header.rows
    anchor = profile.get('name_after')
    start = next((i + 1 for i, r in enumerate(rows) if anchor and anchor.lower() in r.text.lower()), None)
    if start is None:
        if not profile.get('unlabelled_left') and not name:
            return '', ''
        start = 0
    left_rows = [(i, r.crop(0, header.width * .50)) for i, r in enumerate(rows)]
    if name:
        compact_name = re.sub(r'\s+', '', name).lower()
        hit = next((i for i, text in left_rows if compact_name in re.sub(r'\s+', '', text).lower() and not STOP_LABEL.search(text)), None)
        if hit is not None:
            start = hit
        else:
            return name, ''
    elif not (anchor or profile.get('unlabelled_left')):
        return '', ''
    selected = []
    for i, text in left_rows[start:]:
        if not text:
            continue
        if re.match(r'(?i)^(?:page\b|as of\b|account statement|statement of account)', text) or any(normalized_date(m[0]) for m in DATE_RE.finditer(text)):
            if selected:
                break
            continue
        if re.match(r'(?i)^CRN\b', text):
            continue
        if '@' in text:
            break
        if re.match(r'(?i)^(?:Uncleared|Cleared|Date of|Time of|\+?MOD|Lien|Limit|Monthly|Drawing|Interest|Account open|Records|INDIAN RUPEES)', text):
            if selected:
                break
            continue
        if STOP_LABEL.match(text) or re.search(r'(?i)\b(?:statement|balance)\b', text) or (not selected and re.search(r'(?i)\bbank\b', text)):
            if selected:
                break
            continue
        if not selected and (not re.search('[A-Za-z]', text) or '@' in text):
            continue
        selected.append(text)
        if len(selected) >= 8:
            break
    return (selected[0], clean(' '.join(selected[1:]))) if selected else (name, '')


def extract_values(header: Header, profile_override=None) -> tuple[dict[str, str], dict[str, str]]:
    profile, bank_source = identify_bank(header)
    if profile_override and not profile:
        profile, bank_source = profile_override, 'Page 1 OCR bank name/logo'
    name, name_index = header.field(NAME_LABEL, multiline=True)
    if profile.get('customer_label') and not name:
        block, _ = header.field(re.escape(profile['customer_label']), multiline=False)
        name = block
    account_raw, account_index = header.field(ACCOUNT_LABEL)
    match = re.match(r'(?:[A-Z]{1,3}-)?([0-9Xx*][0-9Xx*\s-]{4,32})', account_raw.lstrip("'"))
    account = re.sub(r'[\s-]', '', match[1]) if match else ''
    if not (6 <= len(account) <= 34 and re.search(r'\d', account)):
        account = ''
    address = ''
    for address_pattern in (r"Customer(?:'s)?\s+Address|Address\s+of\s+Customer", r'Communication(?:\s+Address)?', r'Mailing\s+Address', r'(?<!Branch\s)Address(?!\s+Last)'):
        address_header = Header(header.rows[name_index:], header.width) if name_index is not None and address_pattern.startswith('(?<!') else header
        address, _ = address_header.field(address_pattern, multiline=True)
        if address:
            break
    if profile.get('unlabelled_left') and (not name or profile.get('address_is_branch')):
        # Some layouts label only the branch address.
        name, address = _unlabelled_block(header, profile, name)
    else:
        inferred_name, inferred_address = _unlabelled_block(header, profile, name)
        name, address = name or inferred_name, address or inferred_address
    if profile.get('account_name_tail') and not name and match:
        name = clean(account_raw[match.end():])
        if name.upper() in ('INR', 'USD'):
            name = ''
        if name and account_index is not None and account_index + 1 < len(header.rows):
            following = header.rows[account_index + 1].crop(0, header.width * .5)
            if following and not STOP_LABEL.search(following):
                name = clean(name + ' ' + following)
    if profile.get('address_label'):
        extra_address, _ = header.field(re.escape(profile['address_label']), multiline=True)
        if extra_address:
            address = re.split(r'(?i)\bScheme\s*:', extra_address)[0].strip('- ')
    if profile.get('customer_label') and name and not address:
        _, index = header.field(re.escape(profile['customer_label']))
        if index is not None:
            block, _ = header.field(re.escape(profile['customer_label']), multiline=True)
            address = clean(block[len(name):])
    # A bank may repeat the account/currency after the customer's name.
    if account:
        name = re.split(re.escape(account), name)[0].strip(' -')
    name = re.split(r'(?i)\b(?:Primary\s+GSTIN|Account\s+Type|Contact\s+No)\b', name)[0].strip()
    period = statement_period(header)
    sources = {label: 'Page 1 text' for label in LABELS}
    sources['Bank Name'] = bank_source
    if not period and profile.get('balance_period'):
        opening, _ = header.field(r'Opening\s+Balance')
        closing, _ = header.field(r'Ledger\s+Balance')
        first, last = DATE_RE.search(opening), DATE_RE.search(closing)
        if first and last:
            a, b = normalized_date(first[0]), normalized_date(last[0])
            if a and b and a <= b:
                period = f'{a} to {b}'
                sources['Statement Date Between'] = 'Page 1 opening/ledger balance dates (inferred period)'
    return dict(zip(LABELS, (name, profile.get('name', ''), account, address, period))), sources


def ocr_words(page, clip=None, zoom=2.5, min_confidence=20):
    """OCR the first page in memory. No altered/decrypted PDF is written."""
    import pytesseract
    from PIL import Image, ImageOps
    executable = shutil.which('tesseract')
    if not executable:
        executable = next((str(p) for p in (Path(r'C:\Program Files\Tesseract-OCR\tesseract.exe'),
                          Path(r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe')) if p.is_file()), None)
    if not executable:
        raise RuntimeError('Tesseract OCR is unavailable')
    pytesseract.pytesseract.tesseract_cmd = executable
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
    image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
    data = pytesseract.image_to_data(ImageOps.grayscale(image), config='--psm 11', output_type=pytesseract.Output.DICT, timeout=45)
    xoff, yoff = (clip.x0, clip.y0) if clip is not None else (0, 0)
    return [(data['left'][i] / zoom + xoff, data['top'][i] / zoom + yoff,
             (data['left'][i] + data['width'][i]) / zoom + xoff,
             (data['top'][i] + data['height'][i]) / zoom + yoff, text)
            for i, text in enumerate(data['text']) if clean(text) and float(data['conf'][i]) >= min_confidence]


@dataclass
class FirstPageResult:
    values: dict[str, str] = field(default_factory=lambda: dict.fromkeys(LABELS, ''))
    sources: dict[str, str] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    error: str = ''


def read_first_page(path: Path, password: str | None = None, *, allow_ocr=True) -> FirstPageResult:
    result = FirstPageResult()
    try:
        with fitz.open(path) as doc:
            if doc.needs_pass and (not password or not doc.authenticate(password)):
                result.error = 'PDF password is missing or incorrect'
                return result
            if not len(doc):
                result.error = 'PDF has no pages'
                return result
            page = doc[0]
            rows = header_rows(positioned_rows(page.get_text('words')))
            result.values, result.sources = extract_values(Header(rows, page.rect.width))
            if allow_ocr and any(not result.values[k] for k in REQUIRED):
                try:
                    bottom = max((r.y + r.height for r in rows), default=page.rect.height)
                    clip = fitz.Rect(0, 0, page.rect.width, min(page.rect.height, bottom + 6))
                    ocr_header = Header(header_rows(positioned_rows(ocr_words(page, clip))), page.rect.width, True)
                    profile, _ = identify_bank(ocr_header)
                    if not profile:
                        logo_clip = fitz.Rect(0, 0, page.rect.width, page.rect.height * .18)
                        logo_header = Header(positioned_rows(ocr_words(page, logo_clip, zoom=4, min_confidence=0)), page.rect.width, True)
                        profile, _ = identify_bank(logo_header)
                    if profile and rows:
                        native_values, native_sources = extract_values(Header(rows, page.rect.width), profile)
                        for key in LABELS:
                            if native_values[key] and not result.values[key]:
                                result.values[key] = native_values[key]
                                result.sources[key] = native_sources[key]
                    values, sources = extract_values(ocr_header, profile)
                    for key in LABELS:
                        if values[key] and not result.values[key]:
                            result.values[key] = values[key]
                            result.sources[key] = sources.get(key, 'Page 1').replace('Page 1', 'Page 1 OCR').replace('OCR OCR', 'OCR')
                    result.notes.append('First-page OCR used. OCR-derived fields require visual verification.')
                except Exception as exc:
                    result.notes.append(f'First-page OCR could not complete: {type(exc).__name__}: {exc}')
            for key in LABELS:
                if not re.search(r'[A-Za-z0-9]', result.values[key]):
                    result.values[key] = ''
    except Exception as exc:
        result.error = f'First page could not be read: {type(exc).__name__}: {exc}'
    return result
