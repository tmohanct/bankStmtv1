"""Read embedded C2PA declarations without extracting PDF page content.

This is a bounded metadata reader, not a C2PA signature/trust validator.
"""
from __future__ import annotations

from io import BytesIO
import json
import re
import struct

MAX_ATTACHMENT_BYTES = 8 * 1024 * 1024
MAX_TOTAL_BYTES = 16 * 1024 * 1024
MAX_ATTACHMENTS = 32
MAX_BOXES = 2048
MAX_DEPTH = 12


def parse_c2pa(blob: bytes) -> list[tuple[str, dict]]:
    """Return action/claim declarations only; reject malformed containers."""
    if len(blob) > MAX_ATTACHMENT_BYTES:
        raise ValueError('C2PA attachment exceeds inspection size limit')
    budget = [MAX_BOXES]
    records = []
    store_seen = False

    def walk(data, path=(), depth=0):
        nonlocal store_seen
        if depth > MAX_DEPTH:
            raise ValueError('C2PA nesting limit reached')
        offset, label = 0, None
        while offset < len(data):
            budget[0] -= 1
            if budget[0] < 0 or len(data) - offset < 8:
                raise ValueError('Invalid or excessive C2PA boxes')
            size, kind = struct.unpack_from('>I4s', data, offset)
            header = 8
            if size == 1:
                if len(data) - offset < 16:
                    raise ValueError('Truncated C2PA extended box')
                size = struct.unpack_from('>Q', data, offset + 8)[0]
                header = 16
            elif size == 0:
                size = len(data) - offset
            if size < header or size > len(data) - offset:
                raise ValueError('Invalid C2PA box length')
            body = data[offset + header:offset + size]
            offset += size
            if kind == b'jumd':
                if label is not None or len(body) < 18 or not body[16] & 2:
                    raise ValueError('Invalid C2PA description')
                end = body.find(b'\0', 17)
                if end < 0:
                    raise ValueError('Unterminated C2PA label')
                label = body[17:end].decode('utf-8')
                if not path and label == 'c2pa':
                    store_seen = True
            elif kind == b'jumb':
                walk(body, path + ((label,) if label else ()), depth + 1)
            elif kind in (b'cbor', b'json') and path and path[0] == 'c2pa':
                if label not in ('c2pa.claim', 'c2pa.claim.v2', 'c2pa.actions', 'c2pa.actions.v2'):
                    continue
                if kind == b'cbor':
                    import cbor2
                    stream = BytesIO(body)
                    value = cbor2.CBORDecoder(stream).decode()
                    if stream.read(1):
                        raise ValueError('Trailing CBOR data')
                else:
                    value = json.loads(body)
                if not isinstance(value, dict):
                    raise ValueError('C2PA declaration is not a map')
                records.append(('/'.join((*path, label)), value))

    if len(blob) < 8 or blob[4:8] != b'jumb':
        raise ValueError('Expected JUMBF manifest')
    walk(blob)
    if not store_seen:
        raise ValueError('No C2PA manifest store found')
    if not records:
        raise ValueError('No supported C2PA claims/actions found')
    return records


def read_c2pa(document) -> tuple[list[tuple[str, dict]], list[str], bool]:
    """Inspect attachments and catalog-associated manifests; never page streams.

    Unknown/oversized attachments make coverage incomplete, not clean.
    Remote manifests are not fetched and certificates are not trusted here.
    """
    records, errors, seen = [], [], set()
    present, total = False, 0

    def consume(blob, expected=False):
        nonlocal total, present
        total += len(blob)
        if total > MAX_TOTAL_BYTES:
            raise ValueError('Attachment byte budget reached')
        if not expected and (len(blob) < 8 or blob[4:8] != b'jumb'):
            return
        present = True
        if blob in seen:
            return
        seen.add(blob)
        records.extend(parse_c2pa(blob))

    try:
        names = document.embfile_names()
    except Exception as exc:
        names = []
        errors.append(f'Attachment list unavailable: {type(exc).__name__}')
    if len(names) > MAX_ATTACHMENTS:
        errors.append('Attachment count limit reached')
    for index, name in enumerate(names[:MAX_ATTACHMENTS]):
        expected = name.casefold() == 'content credentials' or name.casefold().endswith('.c2pa')
        present |= expected
        try:
            info = document.embfile_info(index)
            if max(info.get('size', 0), info.get('length', 0)) > MAX_ATTACHMENT_BYTES:
                raise ValueError('Attachment exceeds inspection size limit')
            consume(document.embfile_get(index), expected)
        except Exception as exc:
            errors.append(f'Embedded attachment {index + 1}: {type(exc).__name__}: {exc}')

    # Associated files need not also appear in the attachment name tree.
    try:
        kind, value = document.xref_get_key(document.pdf_catalog(), 'AF')
        if kind == 'xref':
            value = document.xref_object(int(value.split()[0]))
        refs = re.findall(r'(\d+)\s+\d+\s+R\b', value)
    except Exception as exc:
        refs = []
        errors.append(f'Associated-file list unavailable: {type(exc).__name__}')
    if len(refs) > MAX_ATTACHMENTS:
        errors.append('Associated-file count limit reached')
    for ref in refs[:MAX_ATTACHMENTS]:
        try:
            xref = int(ref)
            if document.xref_get_key(xref, 'AFRelationship')[1] != '/C2PA_Manifest':
                continue
            present = True
            kind, value = document.xref_get_key(xref, 'EF/F')
            if kind != 'xref':
                kind, value = document.xref_get_key(xref, 'EF/UF')
            if kind != 'xref':
                raise ValueError('Associated C2PA stream missing')
            stream_ref = int(value.split()[0])
            _, length = document.xref_get_key(stream_ref, 'Length')
            if int(length) > MAX_ATTACHMENT_BYTES:
                raise ValueError('Associated manifest exceeds inspection size limit')
            consume(document.xref_stream(stream_ref), True)
        except Exception as exc:
            errors.append(f'Associated manifest: {type(exc).__name__}: {exc}')
    return records, list(dict.fromkeys(errors)), present
