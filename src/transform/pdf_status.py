"""Evidence-based PDF inspection, used exclusively by the PDF_Status sheet.

An unsigned PDF cannot establish who edited it, or exclude a flattened edit.
Recorded content changes, suspicious features and unavailable checks are reported
separately. Metadata, encryption and ordinary save history are not proof of fraud.
"""
from __future__ import annotations

import difflib
import hashlib
import re
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz

from src.utils.pdf_status_reader import LABELS, REQUIRED, FirstPageResult, read_first_page
from src.utils.statement_utils import safe_pdf_display_name
from src.transform.validate import BalanceCheckResult
from src.transform.pdf_signature import verify_reference_hash, verify_signatures
from src.transform.pdf_ai_provenance import AIProvenanceResult, inspect_ai_provenance

COLUMNS = ['PDF', 'Check', 'Status', 'Result', 'Details']
MAX_REVISION_CANDIDATES = 64
MAX_REVISION_PAGE_COMPARISONS = 2000
MAX_REVISION_TOTAL_BYTES = 256 * 1024 * 1024
MAX_VISUAL_OCR_PAGES = 2
EDITOR_PATTERN = re.compile(
    r'Microsoft.*Word|LibreOffice|OpenOffice|WPS\s+(?:Office|Writer)|Foxit.*Editor|'
    r'PDF-XChange|Acrobat.*(?:Pro|PDFMaker)|Sejda|PDFescape|Master PDF Editor|'
    r'Wondershare|PDFelement', re.I)
LIMITATION = ('These checks cannot establish who made a change or prove that the PDF is an untouched bank original. '
              'Full rewrites, flattened edits and image edits can leave no detectable history. '
              'A signature establishes bank origin only when its signer and trust chain are independently configured and verified.')


def row(name, check, status, result, details=''):
    return dict(zip(COLUMNS, (name, check, status, result, str(details))))


def metadata_datetime(value):
    """Normalize explicit PDF/ISO offsets to UTC; leave unknown zones naive."""
    text = str(value or '').strip()
    if not text:
        return None
    if re.match(r'^\d{4}-', text):
        try:
            parsed = datetime.fromisoformat(text.replace('Z', '+00:00'))
            return parsed.astimezone(timezone.utc) if parsed.tzinfo else parsed
        except ValueError:
            return None
    match = re.fullmatch(r'(?:D:)?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(Z|[+-]\d{2}(?:\x27?\d{2}\x27?)?)?(?:00\x2700\x27)?', text)
    if not match:
        return None
    parts = match.groups()
    try:
        parsed = datetime(*(int(v or default) for v, default in zip(parts[:6], (1, 1, 1, 0, 0, 0))))
        zone = parts[6]
        if zone:
            if zone == 'Z':
                offset = 0
            else:
                digits = zone[1:].replace("'", '')
                hours, minutes = int(digits[:2]), int(digits[2:] or 0)
                if hours > 23 or minutes > 59:
                    return None
                offset = (hours * 60 + minutes) * (1 if zone[0] == '+' else -1)
            parsed = parsed.replace(tzinfo=timezone(timedelta(minutes=offset))).astimezone(timezone.utc)
        return parsed
    except ValueError:
        return None


def date_evidence(created_raw, modified_raw):
    created, modified = metadata_datetime(created_raw), metadata_datetime(modified_raw)
    detail = f'CreationDate: {created_raw or "not recorded"}\nModDate: {modified_raw or "not recorded"}'
    if (created_raw and not created) or (modified_raw and not modified):
        return 'WARNING', 'Invalid metadata date', detail + '\nA recorded date could not be parsed.'
    if not created or not modified:
        return 'PASS', 'Date comparison unavailable', detail + '\nMissing dates are not evidence of editing.'
    if bool(created.tzinfo) != bool(modified.tzinfo):
        return 'WARNING', 'Date time zones cannot be compared reliably', detail
    delta = (modified - created).total_seconds()
    detail += f'\nComparable difference: {delta:g} seconds (explicit offsets normalized to UTC).'
    if abs(delta) > 60:
        return 'WARNING', 'Later modification recorded' if delta > 0 else 'Modification predates creation', detail + '\nTimestamps alone do not establish a content edit.'
    return 'PASS', 'Creation and modification dates are consistent', detail


def _page_digest(page):
    # Bound memory for unusually large page dimensions. Text is compared at
    # full precision; rendering checks appearance at up to 144 dpi / 1600 px.
    longest = max(page.rect.width, page.rect.height, 1)
    scale = min(2, 1600 / longest)
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False, annots=False)
    return tuple(page.rect), page.rotation, pixmap.width, pixmap.height, pixmap.digest


def revision_evidence(raw, current, password):
    """Compare all retained versions, including edits subsequently reverted.

    Candidate byte markers are accepted only when a parser reopens them without
    repair. A failed comparison cannot discard changes already established.
    """
    count = current.version_count
    if count <= 1:
        return [], [], [], 'No earlier saved revision is available for comparison.'
    boundaries, errors = [], []
    for match in re.finditer(rb'(?:\r\n|\r|\n)startxref\s+(\d+)\s+%%EOF', raw):
        offset = int(match[1])
        if offset >= match.start():
            continue
        target = raw[offset:offset + 160]
        if not (target.startswith(b'xref') or re.match(rb'\d+\s+\d+\s+obj\b', target)):
            continue
        boundaries.append(match.end())
        if len(boundaries) >= MAX_REVISION_CANDIDATES:
            errors.append('Revision candidate limit reached; history comparison is incomplete.')
            break
    versions = {}
    changes, samples = [], []
    bytes_opened = 0
    try:
        for end in boundaries:
            if bytes_opened + end > MAX_REVISION_TOTAL_BYTES:
                errors.append('Revision byte budget reached; history comparison is incomplete.')
                break
            bytes_opened += end
            previous = None
            try:
                previous = fitz.open(stream=raw[:end], filetype='pdf')
                if previous.needs_pass and not previous.authenticate(password or ''):
                    errors.append('An earlier revision could not be authenticated.')
                    continue
                version = previous.version_count
                if previous.is_repaired or version >= count or version in versions:
                    continue
                versions[version] = previous
                previous = None  # ownership passes to versions
            except Exception as exc:
                errors.append(f'Revision candidate could not be opened: {type(exc).__name__}')
            finally:
                if previous is not None:
                    previous.close()
        if set(versions) != set(range(1, count)):
            errors.append(f'Could compare {len(versions)} of {count - 1} earlier revision(s).')
        chain = sorted(versions.items()) + [(count, current)]
        comparisons = 0
        for (old_version, old), (new_version, new) in zip(chain, chain[1:]):
            if len(old) != len(new):
                changes.append(f'Revisions {old_version}->{new_version}: page count {len(old)}->{len(new)}')
            for index in range(min(len(old), len(new))):
                comparisons += 1
                if comparisons > MAX_REVISION_PAGE_COMPARISONS:
                    errors.append('Revision page comparison limit reached; comparison is incomplete.')
                    return changes, samples, errors, 'Only part of the preserved history was compared.'
                label = f'Page {index + 1}, revisions {old_version}->{new_version}'
                try:
                    a, b = old[index], new[index]
                    before, after = a.get_text(sort=True), b.get_text(sort=True)
                    if before != after:
                        changes.append(f'{label}: text changed')
                        if len(samples) < 6:
                            # Limit diff work on pathological text streams.
                            removed = [line[2:].strip() for line in difflib.ndiff(before[:20000].splitlines(), after[:20000].splitlines()) if line.startswith('- ')]
                            added = [line[2:].strip() for line in difflib.ndiff(before[:20000].splitlines(), after[:20000].splitlines()) if line.startswith('+ ')]
                            samples.append(f'Page {index + 1}: before: {" / ".join(removed)[:250]}\nafter: {" / ".join(added)[:250]}')
                    # Exclude annotations/signature widgets from page comparison.
                    if _page_digest(a) != _page_digest(b):
                        changes.append(f'{label}: appearance changed')
                except Exception as exc:
                    errors.append(f'{label}: comparison failed ({type(exc).__name__})')
    finally:
        for document in versions.values():
            document.close()
    return changes, samples, errors, f'Compared preserved revisions against subsequent versions ({count} total).'


def _span_text(span):
    return ''.join(chr(c[0]) for c in span['chars'] if 0 <= c[0] <= 0x10ffff)


def _visible_span(span):
    return span['type'] in (0, 1) and span['opacity'] >= .05


def _covered_characters(span, rect):
    return [c for c in span['chars'] if 0 <= c[0] <= 0x10ffff and chr(c[0]).strip()
            and fitz.Rect(c[3]).get_area() > 0
            and (rect & fitz.Rect(c[3])).get_area() / fitz.Rect(c[3]).get_area() > .8]


def _numeric_overprints(spans, number):
    """Conflicting monetary text at the same position, ignoring duplicate paint.

    Same-text fill/stroke and synthetic bold are normal. Font variation alone
    is never enough to flag a statement.
    """
    buckets = {}
    hits = []
    for span in spans:
        if not _visible_span(span):
            continue
        text = _span_text(span)
        for match in re.finditer(r'(?<![\w.])[+-]?\d[\d,]*\.\d{2}(?![\w.])', text):
            chars = span['chars'][match.start():match.end()]
            rect = fitz.Rect(chars[0][3])
            for char in chars[1:]:
                rect |= fitz.Rect(char[3])
            if rect.is_empty:
                continue
            # Nearby baseline buckets avoid all-pairs comparison on long pages.
            bucket = int(rect.y1 // 4)
            value = match[0].replace(',', '')
            for key in range(bucket - 1, bucket + 2):
                for old_rect, old_value, seqno in buckets.get(key, ()):
                    if old_value == value or seqno == span['seqno']:
                        continue
                    overlap = (rect & old_rect).get_area()
                    if overlap / max(rect.get_area(), old_rect.get_area()) >= .7:
                        hits.append(f'Page {number}, box ({rect.x0:.0f},{rect.y0:.0f},{rect.x1:.0f},{rect.y1:.0f}): '
                                    f'conflicting amounts {old_value!r} and {value!r}')
                        if len(hits) >= 4:
                            return hits
            buckets.setdefault(bucket, []).append((rect, value, span['seqno']))
    return hits


def page_evidence(document):
    """Inspect every page, retaining page references and bounded examples."""
    covered, hidden, annotations, errors, scanned = [], [], [], [], []
    fonts_by_page = []
    image_overlays, overprints, raster_pages = [], [], []
    for number, page in enumerate(document, 1):
        try:
            spans = page.get_texttrace()
            drawings = page.get_drawings()
            overprints.extend(_numeric_overprints(spans, number))
            image_hits = 0
            image_area = 0.0
            image_count = 0
            for seqno, (kind, bbox, *_) in enumerate(page.get_bboxlog()):
                if kind != 'fill-image':
                    continue
                rect = fitz.Rect(bbox)
                image_count += 1
                image_area += (rect & page.rect).get_area()
                if (rect & page.rect).get_area() >= .5 * page.rect.get_area():
                    if number not in raster_pages:
                        raster_pages.append(number)
                if image_hits >= 4:
                    continue
                for span in spans:
                    if not _visible_span(span) or span['seqno'] >= seqno:
                        continue
                    chars = _covered_characters(span, rect)
                    if len(chars) >= 2:
                        image_overlays.append(
                            f'Page {number}, box ({rect.x0:.0f},{rect.y0:.0f},{rect.x1:.0f},{rect.y1:.0f}): '
                            f'later image overlaps text {"".join(chr(c[0]) for c in chars)[:90]!r}')
                        image_hits += 1
                        break
            if image_count > 1 and image_area >= .5 * page.rect.get_area() and number not in raster_pages:
                raster_pages.append(number)
            if not page.get_text().strip():
                scanned.append(number)
            fonts = sorted({s['font'] for s in spans if s['type'] != 3})
            fonts_by_page.append(fonts)
            invisible = [s for s in spans if s['type'] == 3 or s['opacity'] < .05]
            if invisible:
                hidden.append(f'Page {number}: {len(invisible)} invisible text run(s) (may be an OCR layer)')
            # Inspect opaque rectangles in drawing order, not just overlapping
            # bounding boxes: an original table background precedes its text.
            rectangles = [d for d in drawings if d.get('fill') is not None
                          and d.get('fill_opacity', 1) >= .98
                          and any(item[0] == 're' for item in d.get('items', []))]
            page_hits = 0
            for drawing in rectangles:
                for item in drawing['items']:
                    if item[0] != 're':
                        continue
                    rect = fitz.Rect(item[1])
                    if rect.is_empty:
                        continue
                    earlier = [s for s in spans if s['type'] in (0, 1) and s['opacity'] >= .05 and s['seqno'] < drawing['seqno'] and rect.intersects(s['bbox'])]
                    for span in earlier:
                        occluded = _covered_characters(span, rect)
                        if len(occluded) < 2:
                            continue
                        replacement = [s for s in spans if s['seqno'] > drawing['seqno'] and s['type'] in (0, 1) and rect.intersects(s['bbox'])]
                        text = ''.join(chr(c[0]) for c in occluded)[:90]
                        covered.append(f'Page {number}, box ({rect.x0:.0f},{rect.y0:.0f},{rect.x1:.0f},{rect.y1:.0f}): '
                                       f'covered text {text!r}' + (f'; later text: {_span_text(replacement[0])[:90]!r}' if replacement else ''))
                        page_hits += 1
                        if page_hits >= 4:
                            break
                    if page_hits >= 4:
                        break
                if page_hits >= 4:
                    break
            for annotation in page.annots() or ():
                annotations.append(f'Page {number}: {annotation.type[1]} annotation')
        except Exception as exc:
            errors.append(f'Page {number}: {type(exc).__name__}: {exc}')
    return covered, hidden, annotations, errors, scanned, fonts_by_page, image_overlays, overprints, raster_pages


def visual_text_evidence(document, candidate_pages):
    """Compare high-confidence OCR amounts with the PDF text layer on risky pages."""
    from src.utils.pdf_status_reader import ocr_words

    candidates = sorted(set(candidate_pages))
    reviewed, errors, findings = [], [], []
    if len(candidates) > MAX_VISUAL_OCR_PAGES:
        errors.append(f'Visual OCR limited to {MAX_VISUAL_OCR_PAGES} of {len(candidates)} candidate pages.')
    amount_pattern = re.compile(r'(?<!\w)\d[\d,]*\.\d{2}(?!\w)')
    for number in candidates[:MAX_VISUAL_OCR_PAGES]:
        try:
            page = document[number - 1]
            visible = ' '.join(word[4] for word in ocr_words(page, zoom=2, min_confidence=85))
            extracted = page.get_text(sort=True)
            ocr_amounts = Counter(value.replace(',', '') for value in amount_pattern.findall(visible))
            text_amounts = Counter(value.replace(',', '') for value in amount_pattern.findall(extracted))
            unmatched = list((ocr_amounts - text_amounts).elements())
            reviewed.append(f'Page {number}: {sum(ocr_amounts.values())} high-confidence visual amount(s), '
                            f'{sum(text_amounts.values())} extractable amount(s).')
            if unmatched and text_amounts:
                findings.append(f'Page {number}: visually read amounts absent from text layer: {unmatched[:8]}')
            elif not text_amounts:
                errors.append(f'Page {number}: no extractable amounts to compare with visual OCR.')
        except Exception as exc:
            errors.append(f'Page {number}: visual OCR unavailable ({type(exc).__name__}).')
    return reviewed, findings, errors


def object_evidence(document, raw):
    raw_size = len(raw)
    signatures, actions, forms, errors = [], [], [], []
    for xref in range(1, document.xref_length()):
        try:
            # Arrays, strings and freed objects are ordinary PDF objects too.
            if not document.xref_object(xref).lstrip().startswith('<<'):
                continue
            keys = document.xref_get_keys(xref)
            if not keys:
                continue
            get = lambda key: document.xref_get_key(xref, key)[1]
            action = get('S')
            subtype = get('Subtype')
            if action in ('/JavaScript', '/Launch') or get('Type') == '/EmbeddedFile' or subtype == '/RichMedia':
                actions.append(f'Object {xref}: {action if action != "null" else get("Type") + " " + subtype}')
            if 'XFA' in keys and get('XFA') != 'null':
                forms.append(f'Object {xref}: dynamic XFA form')
            if get('FT') in ('/Tx', '/Ch', '/Btn'):
                forms.append(f'Object {xref}: editable {get("FT")} field')
            if 'ByteRange' in keys or get('Type') in ('/Sig', '/DocTimeStamp'):
                # Do not extract digits from malformed arrays (e.g. decimals,
                # extra values or indirect references) and call them valid.
                array = get('ByteRange')
                exact = re.fullmatch(r'\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]', array)
                nums = [int(n) for n in exact.groups()] if exact else []
                valid = (len(nums) == 4 and nums[0] == 0 and nums[1] > 0
                         and nums[2] > nums[1] and nums[3] > 0 and nums[2] + nums[3] <= raw_size)
                if valid:
                    gap = raw[nums[1]:nums[2]]
                    # Only a populated hex Contents value may be excluded from
                    # the signed byte ranges. This is still NOT CMS validation.
                    hex_gap = re.fullmatch(rb'<([0-9a-fA-F\s]+)>', gap)
                    populated = hex_gap and re.search(rb'[1-9a-fA-F]', hex_gap[1])
                    has_content = document.xref_get_key(xref, 'Contents')[0] == 'string' and bool(get('Contents'))
                    valid = bool(populated and has_content)
                signatures.append((xref, valid, raw_size - nums[2] - nums[3] if exact else None))
        except Exception as exc:
            errors.append(f'Object {xref}: {type(exc).__name__}')
    return signatures, actions, forms, errors


def _inspect_pdf(path: Path, password: str | None = None, *, first_page: FirstPageResult | None = None,
                 audit: dict[str, object] | None = None):
    path = Path(path)
    name = safe_pdf_display_name(path)
    rows = []
    suspicion, confirmed, incomplete, integrity, signature_failure, reference_mismatch = [], [], [], [], [], []
    authenticated = False
    ai_provenance = None
    try:
        raw = path.read_bytes()
        document = fitz.open(stream=raw, filetype='pdf')
    except Exception as exc:
        return [row(name, 'Overall PDF modification status', 'UNASSESSABLE', 'Cannot assess - PDF is inaccessible', str(exc)),
                row(name, 'File access', 'UNASSESSABLE', 'PDF could not be read', str(exc))]
    with document:
        encrypted = bool(document.needs_pass)
        if encrypted and (not password or not document.authenticate(password)):
            return [row(name, 'Overall PDF modification status', 'UNASSESSABLE', 'Cannot assess - password required or incorrect', LIMITATION),
                    row(name, 'File access', 'UNASSESSABLE', 'Password authentication failed'),
                    row(name, 'File fingerprint (SHA-256)', 'PASS', hashlib.sha256(raw).hexdigest(), 'Fingerprint of the original source bytes.')]
        if not document.is_pdf or not len(document):
            return [row(name, 'Overall PDF modification status', 'UNASSESSABLE', 'Cannot assess - no PDF pages')]
        rows.append(row(name, 'File access', 'PASS', f'Opened {len(document)} page(s)', 'Original PDF inspected in memory.'))
        rows.append(row(name, 'File fingerprint (SHA-256)', 'PASS', hashlib.sha256(raw).hexdigest(), 'Identifies this exact file for later comparison; not proof of bank origin.'))
        reference = verify_reference_hash(raw, path)
        if reference:
            ref_status, ref_result, ref_detail = reference
            rows.append(row(name, 'Trusted source fingerprint', ref_status, ref_result, ref_detail))
            if ref_result == 'Matches configured reference PDF':
                authenticated = True
            elif ref_result == 'Differs from configured reference PDF':
                reference_mismatch.append(ref_result)
            else:
                incomplete.append(ref_result)
        rows.append(row(name, 'Encryption', 'PASS', 'Password authenticated' if encrypted else 'No opening password required', 'Encryption itself is not evidence of editing.'))
        try:
            ai_provenance = inspect_ai_provenance(document)
        except Exception as exc:
            ai_provenance = AIProvenanceResult(
                'INCONCLUSIVE', 'AI provenance inspection incomplete',
                f'AI metadata check failed: {type(exc).__name__}', incomplete=True)
        rows.append(row(name, 'AI provenance', ai_provenance.status,
                        ai_provenance.result, ai_provenance.details))
        if ai_provenance.incomplete:
            incomplete.append('AI provenance inspection incomplete; see AI provenance details')
        if audit and isinstance(audit.get('balance_check'), BalanceCheckResult):
            balance = audit['balance_check']
            if balance.status == 'mismatch':
                integrity.append(f'{len(balance.mismatches)} running balance mismatch(es)')
            elif balance.status in ('partial', 'unavailable'):
                incomplete.append('Running balance comparison incomplete or unavailable')
            rows.append(row(name, 'Running balance consistency',
                            'WARNING' if balance.status != 'passed' else 'PASS',
                            f'{len(balance.mismatches)} mismatch(es); {balance.compared} adjacent pair(s) checked',
                            f'Order: {balance.direction}; pairs without usable balances: {balance.missing}.\n'
                            + '\n'.join(balance.mismatches[:40])))
        if document.is_repaired:
            incomplete.append('PDF structure required repair')
        rows.append(row(name, 'PDF structure repair', 'WARNING' if document.is_repaired else 'PASS',
                        'PDF structure required repair' if document.is_repaired else 'No repair required'))
        try:
            header = first_page if first_page is not None else read_first_page(path, password)
        except Exception as exc:
            header = FirstPageResult()
            header.error = f'First-page inspection failed: {type(exc).__name__}'
        missing = [key for key in REQUIRED if not header.values[key]]
        if header.error:
            incomplete.append(header.error)
        if missing:
            incomplete.append('Required first-page fields missing: ' + ', '.join(missing))
        rows.append(row(name, 'Required first-page details', 'WARNING' if missing or header.error else 'PASS',
                        'Missing: ' + ', '.join(missing) if missing else 'Bank, customer, account and statement period found',
                        '\n'.join(filter(None, [header.error, *header.notes]))))
        for key in LABELS:
            value = header.values.get(key, '')
            source = header.sources.get(key, '')
            needs_review = not value or 'OCR' in source or 'inferred' in source
            rows.append(row(name, 'First page: ' + key, 'WARNING' if needs_review else 'PASS',
                            value or 'Not found on first page', source))
            if value and ('OCR' in source or 'inferred' in source) and key in REQUIRED:
                incomplete.append(key + ': ' + source + '; verification required')
        count = document.version_count
        rows.append(row(name, 'Incremental updates', 'PASS', f'{count} saved revision(s)',
                        f'Parsed revision count; linearized: {bool(document.is_fast_webaccess)}. Signing and normal saves can create revisions.'))
        try:
            changes, examples, revision_errors, detail = revision_evidence(raw, document, password)
        except Exception as exc:
            changes, examples = [], []
            revision_errors = [f'Revision inspection failed: {type(exc).__name__}']
            detail = 'Preserved revision inspection could not complete.'
        if changes:
            confirmed.extend(changes)
        incomplete.extend(revision_errors)
        rows.append(row(name, 'Saved revision content comparison', 'FAIL' if changes else 'WARNING' if revision_errors else 'PASS',
                        'Page content changed between saved revisions' if changes else
                        'Comparison incomplete' if revision_errors else 'No page content change found in saved revisions' if count > 1 else 'No earlier revision retained',
                        '\n'.join([detail, *changes[:30], *examples, *revision_errors])[:12000]))
        status, result, detail = date_evidence(document.metadata.get('creationDate'), document.metadata.get('modDate'))
        rows.append(row(name, 'Metadata dates', status, result, detail))
        if status == 'WARNING':
            suspicion.append(result)
        creator = document.metadata.get('creator') or ''
        producer = document.metadata.get('producer') or ''
        editor = EDITOR_PATTERN.search(creator + ' ' + producer)
        conversion = re.search(r'(?i)iLovePDF|Smallpdf|Print\s*To\s*PDF|Quartz\s+PDFContext|Scanner|intsig', creator + ' ' + producer)
        result = 'Editing software recorded' if editor else 'Conversion/processing software recorded' if conversion else 'Creator/producer metadata'
        rows.append(row(name, 'PDF creator', 'WARNING' if editor or conversion else 'PASS', result,
                        f'Creator: {creator or "not recorded"}\nProducer: {producer or "not recorded"}\nSoftware metadata is editable and does not establish a manual content change.'))
        if editor:
            suspicion.append(result)
        if conversion:
            incomplete.append('PDF was converted/processed; original edit history may be lost')
        try:
            xmp_text = document.get_xml_metadata() or ''
        except Exception as exc:
            incomplete.append(f'XMP metadata could not be read: {type(exc).__name__}')
            rows.append(row(name, 'XMP metadata access', 'WARNING', 'XMP metadata could not be read'))
            xmp_text = None
        if xmp_text:
            try:
                root = ET.fromstring(xmp_text)
                fields = {}
                for element in root.iter():
                    for key, value in [(element.tag, element.text or ''), *element.attrib.items()]:
                        local = key.rsplit('}', 1)[-1]
                        if local in ('CreateDate', 'ModifyDate', 'MetadataDate', 'CreatorTool', 'action', 'softwareAgent', 'when') and value.strip():
                            fields.setdefault(local, []).append(value.strip())
                events = fields.get('action', [])
                edited = any(action.lower() in ('edited', 'modified') for action in events)
                if edited:
                    suspicion.append('XMP history records an edit')
                xmp_editor = any(EDITOR_PATTERN.search(value) for key in ('CreatorTool', 'softwareAgent')
                                 for value in fields.get(key, []))
                if xmp_editor:
                    suspicion.append('Editing software recorded in XMP metadata')
                if fields.get('CreateDate') or fields.get('ModifyDate'):
                    xmp_status, xmp_result, xmp_detail = date_evidence(
                        next(iter(fields.get('CreateDate', [])), None),
                        next(iter(fields.get('ModifyDate', [])), None))
                    rows.append(row(name, 'XMP dates', xmp_status, xmp_result, xmp_detail))
                    if xmp_status == 'WARNING':
                        suspicion.append('XMP: ' + xmp_result)
                rows.append(row(name, 'XMP metadata', 'WARNING' if edited or xmp_editor else 'PASS', 'XMP edit history recorded' if edited else 'XMP metadata read',
                                '\n'.join(f'{key}: {", ".join(values[:6])}' for key, values in fields.items())))
            except ET.ParseError:
                incomplete.append('XMP metadata could not be parsed')
                rows.append(row(name, 'XMP metadata', 'WARNING', 'XMP metadata could not be parsed'))
        elif xmp_text is not None:
            rows.append(row(name, 'XMP metadata', 'PASS', 'No XMP packet recorded', 'Absence of XMP is not evidence of editing.'))
        try:
            covered, hidden, annotations, page_errors, scanned, fonts, image_overlays, overprints, raster_pages = page_evidence(document)
        except Exception as exc:
            covered, hidden, annotations, scanned, fonts, image_overlays, overprints, raster_pages = ([] for _ in range(8))
            page_errors = [f'Page inspection failed: {type(exc).__name__}']
        if image_overlays:
            suspicion.append('Later images overlap existing visible text')
        if overprints:
            suspicion.append('Different monetary amounts occupy the same position')
        if raster_pages:
            incomplete.append(f'Large raster images on pages {raster_pages[:30]}; image contents cannot be authenticated')
        rows.append(row(name, 'Image overlays', 'WARNING' if image_overlays or page_errors else 'PASS',
                        f'{len(image_overlays)} possible image cover-up(s)' if image_overlays else 'No later images overlapping visible text found',
                        '\n'.join([*image_overlays[:40], 'Checks drawing order. Image transparency and clipping may produce benign overlaps.'])))
        rows.append(row(name, 'Conflicting amount overprints', 'WARNING' if overprints or page_errors else 'PASS',
                        f'{len(overprints)} conflicting amount area(s)' if overprints else 'No conflicting monetary text overprints found',
                        '\n'.join(overprints[:40]) or 'Different amounts at substantially overlapping positions are review indicators, not proof of fraud.'))
        rows.append(row(name, 'Raster image coverage', 'WARNING' if raster_pages or page_errors else 'PASS',
                        f'{len(raster_pages)} page(s) with large raster images',
                        f'Pages: {raster_pages[:40]}. Flattened image edits cannot be established from text checks.' if raster_pages
                        else 'No single image covers half or more of a page. Smaller image edits remain possible.'))
        candidate_pages = set(raster_pages) | set(scanned)
        candidate_pages.update(int(match[1]) for item in image_overlays
                               if (match := re.match(r'Page (\d+)', item)))
        if candidate_pages:
            visual_review, visual_findings, visual_errors = visual_text_evidence(document, candidate_pages)
            if visual_findings:
                suspicion.append('Visible monetary amounts disagree with the extractable text layer')
            incomplete.extend(visual_errors)
            rows.append(row(name, 'Visual OCR versus PDF text',
                            'WARNING' if visual_findings or visual_errors else 'PASS',
                            f'{len(visual_findings)} possible mismatch(es) on {len(visual_review)} checked page(s)',
                            '\n'.join([*visual_review, *visual_findings, *visual_errors])))
        else:
            rows.append(row(name, 'Visual OCR versus PDF text', 'PASS',
                            'No image-heavy page selected for OCR comparison',
                            'OCR is targeted to raster pages, textless pages and image overlays.'))
        if covered:
            suspicion.append('Text is covered by later opaque rectangles')
        incomplete.extend(page_errors)
        if scanned:
            incomplete.append(f'{len(scanned)} page(s) have no extractable text; image edits cannot be excluded')
        rows.append(row(name, 'All-page inspection coverage', 'WARNING' if page_errors or scanned else 'PASS',
                        f'Inspected {len(document)} page(s); {len(scanned)} without extractable text; {len(page_errors)} page error(s)',
                        '\n'.join([f'Pages without text: {scanned}' if scanned else '', *page_errors])))
        rows.append(row(name, 'Text overlays and cover-ups', 'WARNING' if covered or page_errors else 'PASS',
                        f'{len(covered)} possible covered-text area(s)' if covered else 'No later opaque rectangles covering text found',
                        '\n'.join(covered[:40]) or 'Checks drawing order and character coverage. This heuristic cannot detect every flattened or image edit.'))
        rows.append(row(name, 'Hidden text', 'WARNING' if hidden or page_errors else 'PASS', f'{len(hidden)} page(s) with invisible text' if hidden else 'No invisible text runs found',
                        '\n'.join(hidden[:30]) or 'Normal OCR layers may contain invisible text.'))
        if hidden:
            incomplete.append('Invisible text may be OCR or concealed content; visual review required')
        rows.append(row(name, 'Annotations', 'WARNING' if annotations or page_errors else 'PASS', f'{len(annotations)} annotation(s) found', '\n'.join(annotations[:30])))
        if annotations:
            suspicion.append('Annotations are present')
        # Record font changes as inspection evidence, not a tampering verdict.
        font_sets = {tuple(f) for f in fonts}
        rows.append(row(name, 'Page font inventory', 'WARNING' if page_errors else 'PASS', f'{len(font_sets)} distinct page font set(s)',
                        'Mixed fonts are normal in bank statements.\n' + '\n'.join(f'Page {i + 1}: {", ".join(f) or "image only"}' for i, f in enumerate(fonts[:20]))))
        try:
            signatures, actions, forms, object_errors = object_evidence(document, raw)
        except Exception as exc:
            signatures, actions, forms = [], [], []
            object_errors = [f'Object inspection failed: {type(exc).__name__}']
        incomplete.extend(object_errors)
        rows.append(row(name, 'Active PDF objects', 'WARNING' if actions or forms or object_errors else 'PASS',
                        f'{len(actions)} active object(s); {len(forms)} editable form field(s)',
                        '\n'.join([*actions[:20], *forms[:20], *object_errors[:10]]) or 'Parsed object dictionaries checked. Plain text markers and ordinary open-page destinations are not active code.'))
        if actions or forms:
            suspicion.append('Active content or editable form fields are present')
        rows.append(row(name, 'Digital signature', 'WARNING' if signatures or object_errors else 'PASS',
                        f'{len(signatures)} signature object(s) found; see cryptographic check' if signatures else 'No digital signature byte range found',
                        '\n'.join(f'Object {xref}: structural range {"valid" if valid else "invalid"}; unsigned trailing bytes: {trailing}' for xref, valid, trailing in signatures)
                        or 'No verified bank signature or trusted original was supplied.'))
        if any(not valid or trailing for _, valid, trailing in signatures):
            suspicion.append('Invalid signature byte range or bytes after a signed revision')
        if signatures:
            sig_status, sig_result, sig_detail = verify_signatures(
                raw, path, password, (audit or {}).get('bank'))
            if sig_status == 'FAIL':
                signature_failure.append(sig_result)
            elif sig_status == 'WARNING':
                incomplete.append(sig_result)
            elif sig_status == 'PASS' and all(valid and not trailing for _, valid, trailing in signatures):
                rows[-1]['Status'] = 'PASS'
                authenticated = True
            rows.append(row(name, 'Cryptographic signature verification', sig_status, sig_result, sig_detail))
        else:
            rows.append(row(name, 'Cryptographic signature verification', 'PASS',
                            'No embedded signature detected', 'No bank signature is available to establish authenticity.'))
        if page_errors or object_errors:
            page_checks = {'Image overlays', 'Conflicting amount overprints', 'Raster image coverage',
                           'Text overlays and cover-ups', 'Hidden text', 'Annotations', 'Page font inventory'}
            for check_row in rows:
                if page_errors and check_row['Check'] in page_checks:
                    check_row['Result'] = 'Partial inspection: ' + check_row['Result']
                    check_row['Details'] += '\nInspection incomplete; see All-page inspection coverage.'
                if object_errors and check_row['Check'] == 'Digital signature':
                    check_row['Result'] = 'Partial inspection: ' + check_row['Result']
                    check_row['Details'] += '\nObject inspection incomplete; no signature assurance is available.'
    if confirmed:
        status, result = 'FAIL', 'Content changes detected after an earlier save'
        explanation = 'PDF page content changed. This establishes a saved content change, not who made it or whether it was authorized.'
    elif signature_failure:
        status, result = 'FAIL', 'Cryptographic signature integrity failed'
        explanation = '\n'.join(dict.fromkeys(signature_failure))
    elif ai_provenance and ai_provenance.detected:
        status, result = 'WARNING', 'AI processing evidence detected - review required'
        explanation = ai_provenance.details
        if reference_mismatch or integrity or suspicion:
            explanation += '\nOther indicators: ' + '; '.join(dict.fromkeys(reference_mismatch + integrity + suspicion))
    elif reference_mismatch:
        status, result = 'WARNING', 'Source differs from configured reference'
        explanation = '\n'.join(dict.fromkeys(reference_mismatch + integrity + suspicion))
    elif integrity:
        status, result = 'WARNING', 'Transaction integrity requires review'
        explanation = '\n'.join(dict.fromkeys(integrity + suspicion))
    elif suspicion:
        status, result = 'WARNING', 'Possible manual editing - review evidence'
        explanation = '\n'.join(dict.fromkeys(suspicion))
    elif incomplete:
        status, result = 'WARNING', 'Inconclusive - inspection has limitations'
        explanation = '\n'.join(dict.fromkeys(incomplete))
    else:
        status, result = 'PASS', 'Reference match or bank signature verified' if authenticated else 'No evidence of manual editing found'
        explanation = 'All available page, revision, metadata and object checks completed.'
    if confirmed and (signature_failure or reference_mismatch or integrity or suspicion):
        explanation += '\nOther indicators: ' + '; '.join(dict.fromkeys(
            signature_failure + reference_mismatch + integrity + suspicion))
    if (confirmed or signature_failure) and ai_provenance and ai_provenance.detected:
        explanation += '\nAI provenance: ' + ai_provenance.result + '; see AI provenance details.'
    if (confirmed or signature_failure or reference_mismatch or suspicion or integrity or
            (ai_provenance and ai_provenance.detected)) and incomplete:
        explanation += '\nLimitations: ' + '; '.join(dict.fromkeys(incomplete))
    return [row(name, 'Overall PDF modification status', status, result, explanation + '\n' + LIMITATION), *rows]


def inspect_pdf(path: Path, password: str | None = None, *, first_page: FirstPageResult | None = None,
                audit: dict[str, object] | None = None):
    """A failed check must never abort export or imply a clean PDF."""
    try:
        rows = _inspect_pdf(path, password, first_page=first_page, audit=audit)
    except Exception as exc:
        rows = [row(safe_pdf_display_name(path), 'Overall PDF modification status', 'WARNING',
                    'Inconclusive - inspection could not complete',
                    f'{type(exc).__name__}: {exc}\n{LIMITATION}')]
    for item in rows:
        for key in COLUMNS:
            value = str(item[key])
            value = value.replace(str(path), safe_pdf_display_name(path))
            value = value.replace(Path(path).name, safe_pdf_display_name(path))
            item[key] = value
    return rows
