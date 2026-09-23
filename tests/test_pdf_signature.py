"""Real CMS signature verification with a disposable, self-signed test key."""
from __future__ import annotations

import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path

import fitz

from src.transform.pdf_signature import verify_signatures


class SignatureVerificationTests(unittest.TestCase):
    def test_signed_pdf_is_cryptographically_checked_but_not_called_bank_trusted(self):
        try:
            from cryptography import x509
            from cryptography.hazmat.primitives import hashes, serialization
            from cryptography.hazmat.primitives.asymmetric import rsa
            from cryptography.hazmat.primitives.serialization import pkcs12
            from cryptography.x509.oid import NameOID
            from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
            from pyhanko.sign import signers
        except ImportError:
            self.skipTest('pyHanko test dependency is unavailable')

        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
            name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, 'Synthetic Test Signer')])
            now = datetime.now(timezone.utc)
            cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name)
                    .public_key(key.public_key()).serial_number(x509.random_serial_number())
                    .not_valid_before(now - timedelta(days=1))
                    .not_valid_after(now + timedelta(days=1))
                    .add_extension(x509.BasicConstraints(ca=True, path_length=None), critical=True)
                    .sign(key, hashes.SHA256()))
            pfx = folder / 'test.pfx'
            pfx.write_bytes(pkcs12.serialize_key_and_certificates(
                b'test', key, cert, None, serialization.NoEncryption()))
            with fitz.open() as doc:
                doc.new_page().insert_text((40, 50), 'Synthetic statement')
                doc.set_metadata({'title': 'TESTMARKER'})
                unsigned = doc.tobytes()
            signer = signers.SimpleSigner.load_pkcs12(str(pfx))
            output = BytesIO()
            signers.sign_pdf(IncrementalPdfFileWriter(BytesIO(unsigned)),
                             signers.PdfSignatureMetadata(field_name='Signature1'),
                             signer=signer, output=output)
            status, result, details = verify_signatures(output.getvalue(), folder / 'statement.pdf', None, 'axis')
            self.assertEqual(status, 'WARNING')
            self.assertIn('unverified', result.lower())
            self.assertIn('cryptographic integrity=valid', details)
            self.assertIn('coverage=ENTIRE_FILE', details)
            altered = output.getvalue().replace(b'TESTMARKER', b'FAKEMARKER', 1)
            self.assertNotEqual(altered, output.getvalue())
            status, result, details = verify_signatures(altered, folder / 'statement.pdf', None, 'axis')
            self.assertEqual(status, 'FAIL')
            self.assertIn('cryptographic integrity=invalid', details)
