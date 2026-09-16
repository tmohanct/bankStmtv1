"""Evidence-based PDF inspection, used exclusively by the PDF_Status sheet.

An unsigned PDF cannot establish who edited it, or exclude a flattened edit.
Recorded content changes, suspicious features and unavailable checks are reported
separately. Metadata, encryption and ordinary save history are not proof of fraud.
"""
from __future__ import annotations

import difflib
import hashlib
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

import fitz

from src.utils.pdf_status_reader import LABELS, REQUIRED, FirstPageResult, read_first_page

COLUMNS = ['PDF', 'Check', 'Status', 'Result', 'Details']
LIMITATION = ('These checks cannot establish who made a change or prove that the PDF is an untouched bank original. '
              'Full rewrites, flattened edits and image edits can leave no detectable history. '
              'Digital signatures are inspected structurally, not cryptographically validated.')


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
    if not created or not modified:
        return 'PASS', 'Date comparison unavailable', detail + '\nMissing dates are not evidence of editing.'
    if bool(created.tzinfo) != bool(modified.tzinfo):
        return 'WARNING', 'Date time zones cannot be compared reliably', detail
    delta = (modified - created).total_seconds()
    detail += f'\nComparable difference: {delta:g} seconds (explicit offsets normalized to UTC).'
    if abs(delta) > 60:
        return 'WARNING', 'Later modification recorded' if delta > 0 else 'Modification predates creation', detail + '\nTimestamps alone do not establish a content edit.'
    return 'PASS', 'Creation and modification dates are consistent', detail


def revision_evidence(raw, current, password):
    """Compare preserved revisions, rejecting literal markers inside streams.

The parsed version count decides whether revisions exist. Raw byte boundaries
are only candidates, checked by reopening and checking their xref/version data.
"""
    count = current.version_count
    if count <= 1:
        return [], [], [], 'No earlier saved revision is available for comparison.'
    boundaries = []
    for match in re.finditer(rb'(?:\r?\n)startxref\s+(\d+)\s+%%EOF', raw):
        offset = int(match[1])
        if offset >= match.start():
            continue
        target = raw[offset:offset + 160]
        if not (target.startswith(b'xref') or re.match(rb'\d+\s+\d+\s+obj\b', target)):
            continue
        boundaries.append(match.end())
    versions = {}
    errors = []
    for end in boundaries:
        try:
            previous = fitz.open(stream=raw[:end], filetype='pdf')
            if previous.needs_pass and not previous.authenticate(password or ''):
                previous.close()
                continue
            version = previous.version_count
            if previous.is_repaired or version >= count or version in versions:
                previous.close()
                continue
            versions[version] = previous
        except Exception as exc:
            errors.append(f'Revision candidate could not be opened: {type(exc).__name__}')
    changes, samples = [], []
    try:
        if set(versions) != set(range(1, count)):
            errors.append(f'Could compare {len(versions)} of {count - 1} earlier revision(s).')
        chain = sorted(versions.items()) + [(count, current)]
        for (old_version, old), (new_version, new) in zip(chain, chain[1:]):
            if len(old) != len(new):
                changes.append(f'Revisions {old_version}->{new_version}: page count {len(old)}->{len(new)}')
            for index in range(min(len(old), len(new))):
                a, b = old[index], new[index]
                before, after = a.get_text(sort=True), b.get_text(sort=True)
                text_changed = before != after
                # Render without annotations so adding a signature widget is not
                # misreported as an edit of the underlying statement page.
                ap = a.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False, annots=False)
                bp = b.get_pixmap(matrix=fitz.Matrix(1, 1), alpha=False, annots=False)
                visual_changed = (ap.width, ap.height, ap.digest) != (bp.width, bp.height, bp.digest)
                if text_changed or visual_changed:
                    kinds = '/'.join(k for k, present in (('text', text_changed), ('appearance', visual_changed)) if present)
                    changes.append(f'Page {index + 1}, revisions {old_version}->{new_version}: {kinds} changed')
                    if text_changed and len(samples) < 6:
                        removed = [line[2:].strip() for line in difflib.ndiff(before.splitlines(), after.splitlines()) if line.startswith('- ')]
                        added = [line[2:].strip() for line in difflib.ndiff(before.splitlines(), after.splitlines()) if line.startswith('+ ')]
                        samples.append(f'Page {index + 1}: before: {" / ".join(removed)[:250]}\nafter: {" / ".join(added)[:250]}')
    finally:
        for document in versions.values():
            document.close()
    return changes, samples, errors, f'Compared preserved revisions against subsequent versions ({count} total).'


def _span_text(span):
    return ''.join(chr(c[0]) for c in span['chars'] if 0 <= c[0] <= 0x10ffff)


def page_evidence(document):
    """Inspect every page, retaining page references and bounded examples."""
    covered, hidden, annotations, errors, scanned = [], [], [], [], []
    fonts_by_page = []
    for number, page in enumerate(document, 1):
        try:
            spans = page.get_texttrace()
            drawings = page.get_drawings()
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
                        occluded = [c for c in span['chars'] if chr(c[0]).strip() and fitz.Rect(c[3]).get_area() > 0
                                    and (rect & fitz.Rect(c[3])).get_area() / fitz.Rect(c[3]).get_area() > .8]
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
    return covered, hidden, annotations, errors, scanned, fonts_by_page


def object_evidence(document, raw_size):
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
            if get('FT') in ('/Tx', '/Ch', '/Btn'):
                forms.append(f'Object {xref}: editable {get("FT")} field')
            if 'ByteRange' in keys:
                nums = [int(n) for n in re.findall(r'-?\d+', get('ByteRange'))]
                has_content = get('Contents') not in ('null', '', '<>')
                valid = (len(nums) == 4 and nums[0] == 0 and nums[1] > 0
                         and nums[2] > nums[1] and nums[3] >= 0 and nums[2] + nums[3] <= raw_size)
                signatures.append((xref, valid and has_content, raw_size - nums[2] - nums[3] if valid else None))
        except Exception as exc:
            errors.append(f'Object {xref}: {type(exc).__name__}')
    return signatures, actions, forms, errors


def _inspect_pdf(path: Path, password: str | None = None, *, first_page: FirstPageResult | None = None):
    path = Path(path)
    name = path.name
    rows = []
    suspicion, confirmed, incomplete = [], [], []
    try:
        raw = path.read_bytes()
        document = fitz.open(stream=raw, filetype='pdf')
    except Exception as exc:
        return [row(name, 'Overall PDF modification status', 'FAIL', 'Cannot assess - PDF is inaccessible', str(exc)),
                row(name, 'File access', 'FAIL', 'PDF could not be read', str(exc))]
    with document:
        encrypted = bool(document.needs_pass)
        if encrypted and (not password or not document.authenticate(password)):
            return [row(name, 'Overall PDF modification status', 'FAIL', 'Cannot assess - password required or incorrect', LIMITATION),
                    row(name, 'File access', 'FAIL', 'Password authentication failed'),
                    row(name, 'File fingerprint (SHA-256)', 'PASS', hashlib.sha256(raw).hexdigest(), 'Fingerprint of the original source bytes.')]
        if not document.is_pdf or not len(document):
            return [row(name, 'Overall PDF modification status', 'FAIL', 'Cannot assess - no PDF pages')]
        rows.append(row(name, 'File access', 'PASS', f'Opened {len(document)} page(s)', 'Original PDF inspected in memory.'))
        rows.append(row(name, 'File fingerprint (SHA-256)', 'PASS', hashlib.sha256(raw).hexdigest(), 'Identifies this exact file for later comparison; not proof of bank origin.'))
        rows.append(row(name, 'Encryption', 'PASS', 'Password authenticated' if encrypted else 'No opening password required', 'Encryption itself is not evidence of editing.'))
        if document.is_repaired:
            incomplete.append('PDF structure required repair')
        rows.append(row(name, 'PDF structure repair', 'WARNING' if document.is_repaired else 'PASS',
                        'PDF structure required repair' if document.is_repaired else 'No repair required'))
        header = first_page if first_page is not None else read_first_page(path, password)
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
        changes, examples, revision_errors, detail = revision_evidence(raw, document, password)
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
        editor = re.search(r'(?i)Microsoft.*Word|LibreOffice|OpenOffice|WPS\s+(?:Office|Writer)|Foxit.*Editor|PDF-XChange|Acrobat.*(?:Pro|PDFMaker)|Sejda|PDFescape|Master PDF Editor|Wondershare|PDFelement', creator + ' ' + producer)
        conversion = re.search(r'(?i)iLovePDF|Smallpdf|Print\s*To\s*PDF|Quartz\s+PDFContext|Scanner|intsig', creator + ' ' + producer)
        result = 'Editing software recorded' if editor else 'Conversion/processing software recorded' if conversion else 'Creator/producer metadata'
        rows.append(row(name, 'PDF creator', 'WARNING' if editor or conversion else 'PASS', result,
                        f'Creator: {creator or "not recorded"}\nProducer: {producer or "not recorded"}\nSoftware metadata is editable and does not establish a manual content change.'))
        if editor:
            suspicion.append(result)
        if conversion:
            incomplete.append('PDF was converted/processed; original edit history may be lost')
        xmp_text = document.get_xml_metadata() or ''
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
                rows.append(row(name, 'XMP metadata', 'WARNING' if edited else 'PASS', 'XMP edit history recorded' if edited else 'XMP metadata read',
                                '\n'.join(f'{key}: {", ".join(values[:6])}' for key, values in fields.items())))
            except ET.ParseError:
                incomplete.append('XMP metadata could not be parsed')
                rows.append(row(name, 'XMP metadata', 'WARNING', 'XMP metadata could not be parsed'))
        else:
            rows.append(row(name, 'XMP metadata', 'PASS', 'No XMP packet recorded', 'Absence of XMP is not evidence of editing.'))
        covered, hidden, annotations, page_errors, scanned, fonts = page_evidence(document)
        if covered:
            suspicion.append('Text is covered by later opaque rectangles')
        incomplete.extend(page_errors)
        if scanned:
            incomplete.append(f'{len(scanned)} page(s) have no extractable text; image edits cannot be excluded')
        rows.append(row(name, 'All-page inspection coverage', 'WARNING' if page_errors or scanned else 'PASS',
                        f'Inspected {len(document)} page(s); {len(scanned)} without extractable text; {len(page_errors)} page error(s)',
                        '\n'.join([f'Pages without text: {scanned}' if scanned else '', *page_errors])))
        rows.append(row(name, 'Text overlays and cover-ups', 'WARNING' if covered else 'PASS',
                        f'{len(covered)} possible covered-text area(s)' if covered else 'No later opaque rectangles covering text found',
                        '\n'.join(covered[:40]) or 'Checks drawing order and character coverage. This heuristic cannot detect every flattened or image edit.'))
        rows.append(row(name, 'Hidden text', 'WARNING' if hidden else 'PASS', f'{len(hidden)} page(s) with invisible text' if hidden else 'No invisible text runs found',
                        '\n'.join(hidden[:30]) or 'Normal OCR layers may contain invisible text.'))
        if hidden:
            incomplete.append('Invisible text may be OCR or concealed content; visual review required')
        rows.append(row(name, 'Annotations', 'WARNING' if annotations else 'PASS', f'{len(annotations)} annotation(s) found', '\n'.join(annotations[:30])))
        if annotations:
            suspicion.append('Annotations are present')
        # Record font changes as inspection evidence, not a tampering verdict.
        font_sets = {tuple(f) for f in fonts}
        rows.append(row(name, 'Page font inventory', 'PASS', f'{len(font_sets)} distinct page font set(s)',
                        'Mixed fonts are normal in bank statements.\n' + '\n'.join(f'Page {i + 1}: {", ".join(f) or "image only"}' for i, f in enumerate(fonts[:20]))))
        signatures, actions, forms, object_errors = object_evidence(document, len(raw))
        incomplete.extend(object_errors)
        rows.append(row(name, 'Active PDF objects', 'WARNING' if actions or forms or object_errors else 'PASS',
                        f'{len(actions)} active object(s); {len(forms)} editable form field(s)',
                        '\n'.join([*actions[:20], *forms[:20], *object_errors[:10]]) or 'Parsed object dictionaries checked. Plain text markers and ordinary open-page destinations are not active code.'))
        if actions or forms:
            suspicion.append('Active content or editable form fields are present')
        rows.append(row(name, 'Digital signature', 'WARNING' if signatures else 'PASS',
                        f'{len(signatures)} signature byte range(s) found; not cryptographically validated' if signatures else 'No digital signature byte range found',
                        '\n'.join(f'Object {xref}: structural range {"valid" if valid else "invalid"}; unsigned trailing bytes: {trailing}' for xref, valid, trailing in signatures)
                        or 'No verified bank signature or trusted original was supplied.'))
        if any(not valid or trailing for _, valid, trailing in signatures):
            suspicion.append('Invalid signature byte range or bytes after a signed revision')
    if confirmed:
        status, result = 'FAIL', 'Content changes detected after an earlier save'
        explanation = 'PDF page content changed. This establishes a saved content change, not who made it or whether it was authorized.'
    elif suspicion:
        status, result = 'WARNING', 'Possible manual editing - review evidence'
        explanation = '\n'.join(dict.fromkeys(suspicion))
    elif incomplete:
        status, result = 'WARNING', 'Inconclusive - inspection has limitations'
        explanation = '\n'.join(dict.fromkeys(incomplete))
    else:
        status, result = 'PASS', 'No evidence of manual editing found'
        explanation = 'All available page, revision, metadata and object checks completed.'
    if suspicion and incomplete:
        explanation += '\nLimitations: ' + '; '.join(dict.fromkeys(incomplete))
    return [row(name, 'Overall PDF modification status', status, result, explanation + '\n' + LIMITATION), *rows]


def inspect_pdf(path: Path, password: str | None = None, *, first_page: FirstPageResult | None = None):
    """A failed check must never abort export or imply a clean PDF."""
    try:
        return _inspect_pdf(path, password, first_page=first_page)
    except Exception as exc:
        return [row(Path(path).name, 'Overall PDF modification status', 'WARNING',
                    'Inconclusive - inspection could not complete',
                    f'{type(exc).__name__}: {exc}\n{LIMITATION}')]
