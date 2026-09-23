"""Synthetic metadata regressions; no customer statements in fixtures."""
import hashlib
import struct
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import cbor2
import fitz

from src.export.final_excel_builder import _build_pdf_status_sheet
from src.transform.pdf_ai_provenance import inspect_ai_provenance
from src.transform.pdf_status import inspect_pdf
from src.utils.pdf_provenance_reader import parse_c2pa
from src.utils.pdf_status_reader import FirstPageResult, LABELS


def box(kind, data):
    return struct.pack('>I4s', len(data) + 8, kind) + data


def jumb(label, *children):
    desc = b'\0' * 16 + b'\3' + label.encode() + b'\0'
    return box(b'jumb', box(b'jumd', desc) + b''.join(children))


def manifest(agent='gpt-5-6-thinking', source='trainedAlgorithmicMedia', generator='ChatGPT'):
    action = {'action': 'c2pa.created', 'softwareAgent': {'name': agent},
              'digitalSourceType': 'http://cv.iptc.org/newscodes/digitalsourcetype/' + source,
              'when': '2026-09-23T00:06:18Z'}
    return jumb('c2pa', jumb('urn:c2pa:test',
        jumb('c2pa.assertions', jumb('c2pa.actions.v2',
             box(b'cbor', cbor2.dumps({'actions': [action], 'allActionsIncluded': False})))),
        jumb('c2pa.claim.v2', box(b'cbor', cbor2.dumps({'claim_generator_info': {'name': generator}})))))


class AIProvenanceTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.path = Path(self.directory.name) / 'statement.pdf'

    def tearDown(self):
        self.directory.cleanup()

    def pdf(self, attachment=None, name='Content Credentials', metadata=None, text='TEST 100.00', xmp=None):
        with fitz.open() as doc:
            doc.new_page().insert_text((40, 50), text)
            doc.set_metadata(metadata or {'creator': 'JasperReports', 'producer': 'OpenPDF'})
            if attachment is not None:
                doc.embfile_add(name, attachment)
            if xmp:
                doc.set_xml_metadata(xmp)
            doc.save(self.path)

    def inspect(self):
        with fitz.open(self.path) as doc:
            return inspect_ai_provenance(doc)

    def test_ai_manifest_with_unchanged_standard_metadata(self):
        self.pdf(manifest())
        before = hashlib.sha256(self.path.read_bytes()).digest()
        with fitz.open(self.path) as doc, patch.object(fitz.Page, 'get_text', side_effect=AssertionError('page read')):
            result = inspect_ai_provenance(doc)
        self.assertEqual(result.code, 'AI_PROVENANCE_RECORDED')
        self.assertEqual(result.status, 'WARNING')
        self.assertTrue(result.detected)
        self.assertFalse(result.incomplete)
        self.assertIn('gpt-5-6-thinking', result.details)
        self.assertIn('2026-09-23T00:06:18Z', result.details)
        self.assertIn('not validated', result.details)
        self.assertEqual(hashlib.sha256(self.path.read_bytes()).digest(), before)

    def test_no_metadata_is_not_proof_of_original(self):
        self.pdf()
        result = self.inspect()
        self.assertEqual(result.code, 'NO_AI_EVIDENCE')
        self.assertFalse(result.detected)
        self.assertIn('no evidence does not mean', result.details)

    def test_generic_content_credentials_are_not_ai(self):
        self.pdf(manifest('Camera', 'digitalCapture', 'Camera'))
        result = self.inspect()
        self.assertEqual(result.code, 'NO_AI_EVIDENCE')
        self.assertFalse(result.incomplete)

    def test_algorithmic_correction_alone_is_not_ai(self):
        self.pdf(manifest('Ordinary Editor', 'algorithmicallyEnhanced', 'Ordinary Editor'))
        self.assertEqual(self.inspect().code, 'NO_AI_EVIDENCE')

    def test_associated_file_error_preserves_attachment_evidence(self):
        self.pdf(manifest())
        original = fitz.Document.xref_get_key
        def fail_af(document, xref, key):
            if key == 'AF':
                raise RuntimeError('broken')
            return original(document, xref, key)
        with fitz.open(self.path) as doc:
            with patch.object(fitz.Document, 'xref_get_key', fail_af):
                result = inspect_ai_provenance(doc)
        self.assertTrue(result.detected)
        self.assertTrue(result.incomplete)

    def test_keywords_in_page_title_and_unrelated_attachment_do_not_trigger(self):
        self.pdf(b'ChatGPT trainedAlgorithmicMedia', name='notes.txt',
                 text='Payment to OpenAI ChatGPT', metadata={'title': 'ChatGPT statement'})
        self.assertEqual(self.inspect().code, 'NO_AI_EVIDENCE')

    def test_ai_source_without_known_tool_is_detected(self):
        self.pdf(manifest('Unknown Tool', 'compositeWithTrainedAlgorithmicMedia', 'Unknown'))
        self.assertTrue(self.inspect().detected)

    def test_creator_is_unsigned_evidence(self):
        self.pdf(metadata={'creator': 'ChatGPT'})
        self.assertEqual(self.inspect().code, 'AI_METADATA_RECORDED')

    def test_renamed_attachment_is_detected(self):
        self.pdf(manifest(), name='provenance.bin')
        self.assertEqual(self.inspect().code, 'AI_PROVENANCE_RECORDED')

    def test_associated_manifest_without_attachment_name_tree(self):
        with fitz.open() as doc:
            doc.new_page()
            stream = doc.get_new_xref()
            doc.update_object(stream, '<< /Type /EmbeddedFile /Subtype /application#2Fc2pa >>')
            doc.update_stream(stream, manifest())
            spec = doc.get_new_xref()
            doc.update_object(spec, f'<< /Type /Filespec /AFRelationship /C2PA_Manifest /EF << /F {stream} 0 R >> >>')
            doc.xref_set_key(doc.pdf_catalog(), 'AF', f'[{spec} 0 R]')
            doc.save(self.path)
        self.assertEqual(self.inspect().code, 'AI_PROVENANCE_RECORDED')

    def test_damaged_manifest_is_inconclusive(self):
        self.pdf(manifest()[:-10])
        result = self.inspect()
        self.assertEqual(result.code, 'INCONCLUSIVE')
        self.assertEqual(result.status, 'WARNING')

    def test_missing_decoder_does_not_report_clean(self):
        self.pdf(manifest())
        with patch.dict(sys.modules, {'cbor2': None}):
            self.assertEqual(self.inspect().code, 'INCONCLUSIVE')

    def test_attachment_budget_is_explicit(self):
        self.pdf(manifest())
        with patch('src.utils.pdf_provenance_reader.MAX_ATTACHMENT_BYTES', 20):
            result = self.inspect()
        self.assertEqual(result.code, 'INCONCLUSIVE')
        self.assertIn('limit', result.details)

    def test_jumbf_length_and_depth_are_bounded(self):
        for raw in (b'\0\0\0\x04jumb', struct.pack('>I4s', 1000000, b'jumb')):
            with self.assertRaises(ValueError):
                parse_c2pa(raw)
        deep = manifest()
        for _ in range(15):
            deep = jumb('extra', deep)
        with self.assertRaisesRegex(ValueError, 'nesting'):
            parse_c2pa(deep)

    def test_ai_evidence_survives_second_broken_attachment(self):
        with fitz.open() as doc:
            doc.new_page()
            doc.embfile_add('Content Credentials', manifest())
            doc.embfile_add('broken.c2pa', b'bad')
            doc.save(self.path)
        result = self.inspect()
        self.assertTrue(result.detected)
        self.assertTrue(result.incomplete)

    def test_training_permission_is_not_generation_evidence(self):
        blob = jumb('c2pa', jumb('test',
            jumb('c2pa.assertions', jumb('c2pa.ai_training',
                box(b'cbor', cbor2.dumps({'use': 'allowed', 'name': 'ChatGPT'})))),
            jumb('c2pa.claim.v2', box(b'cbor', cbor2.dumps({'claim_generator_info': {'name': 'Camera'}})))))
        self.pdf(blob)
        self.assertEqual(self.inspect().code, 'NO_AI_EVIDENCE')

    def test_xmp_ai_source_resource_attribute(self):
        self.pdf(xmp='<x:xmpmeta xmlns:x="adobe:ns:meta/" '
                     'xmlns:iptc="http://iptc.org/std/Iptc4xmpExt/2008-02-29/" '
                     'xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
                     '<iptc:DigitalSourceType rdf:resource="http://cv.iptc.org/newscodes/'
                     'digitalsourcetype/trainedAlgorithmicMedia"/></x:xmpmeta>')
        self.assertEqual(self.inspect().code, 'AI_METADATA_RECORDED')

    def test_audit_and_export_expose_ai_status(self):
        self.pdf(manifest())
        header = FirstPageResult()
        header.values.update({label: 'Test' for label in LABELS})
        rows = inspect_pdf(self.path, first_page=header)
        self.assertEqual(rows[0]['Result'], 'AI processing evidence detected - review required')
        self.assertEqual(rows[0]['Status'], 'WARNING')
        frame = _build_pdf_status_sheet([self.path])
        ai_row = frame.loc[frame['Check'] == 'AI provenance'].iloc[0]
        self.assertEqual(ai_row['Status'], 'WARNING')
        self.assertIn('AI_PROVENANCE_RECORDED', ai_row['Details'])

    def test_confirmed_content_change_retains_fail_priority(self):
        self.pdf(manifest())
        with fitz.open(self.path) as doc:
            doc[0].insert_text((40, 80), 'Changed 900.00')
            doc.saveIncr()
        rows = inspect_pdf(self.path, first_page=FirstPageResult())
        self.assertEqual(rows[0]['Status'], 'FAIL')
        self.assertIn('AI provenance:', rows[0]['Details'])

    def test_failed_ai_check_cannot_silently_pass(self):
        self.pdf()
        with patch('src.transform.pdf_status.inspect_ai_provenance', side_effect=RuntimeError('broken')):
            rows = inspect_pdf(self.path, first_page=FirstPageResult())
        self.assertEqual(next(r['Status'] for r in rows if r['Check'] == 'AI provenance'), 'WARNING')


if __name__ == '__main__':
    unittest.main()
