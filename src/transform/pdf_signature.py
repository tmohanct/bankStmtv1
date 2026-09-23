"""Verify PDF signatures against explicitly supplied bank trust material."""
from __future__ import annotations

import hashlib
import json
import os
from io import BytesIO
from pathlib import Path

from src.utils.statement_utils import safe_pdf_display_name


def _config_path(source: Path) -> Path:
    return Path(os.environ.get("BANKSTMT_TRUST_CONFIG", source.parent / "BankTrust.json"))


def verify_reference_hash(raw: bytes, source: Path) -> tuple[str, str, str] | None:
    """Compare with an operator-supplied, independently obtained source hash."""
    config_path = _config_path(source)
    if not config_path.is_file():
        return None
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
        expected = config.get("trusted_sources", {}).get(safe_pdf_display_name(source))
        if expected is None:
            return None
        expected = str(expected).lower().strip()
        actual = hashlib.sha256(raw).hexdigest()
        if len(expected) != 64 or any(char not in "0123456789abcdef" for char in expected):
            return "WARNING", "Configured source hash is invalid", "Expected a 64-character SHA-256 hex digest."
        if actual == expected:
            return "PASS", "Matches configured reference PDF", "SHA-256 matches the independently supplied reference digest."
        return "WARNING", "Differs from configured reference PDF", f"Expected SHA-256: {expected}\nActual SHA-256: {actual}"
    except (OSError, ValueError, TypeError, AttributeError) as exc:
        return "WARNING", "Reference hash comparison unavailable", f"Trust configuration could not be read: {type(exc).__name__}."


def _trust_profile(bank_code: str | None, source: Path) -> tuple[list[Path], set[str], str]:
    config_path = _config_path(source)
    if not config_path.is_file():
        return [], set(), "No bank trust configuration was supplied."
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
        profile = config.get("banks", {}).get(bank_code or "", {})
        if not profile:
            return [], set(), f"No bank signature trust material is configured for {bank_code or 'unknown bank'}."
        roots = [(config_path.parent / item).resolve() for item in profile.get("trust_roots", [])]
        signer_hashes = {str(item).lower().replace(":", "") for item in profile.get("signer_sha256", [])}
        if any(not path.is_file() for path in roots):
            return [], set(), "A configured bank trust certificate could not be read."
        return roots, signer_hashes, ""
    except (OSError, ValueError, TypeError, AttributeError) as exc:
        return [], set(), f"Bank trust configuration is invalid: {type(exc).__name__}."


def verify_signatures(raw: bytes, source: Path, password: str | None, bank_code: str | None) -> tuple[str, str, str]:
    """Return status, result and bounded evidence for signatures in source bytes.

    Trust requires both a bank-approved root and a pinned signer certificate.
    Network fetching is disabled so this audit is reproducible offline.
    """
    try:
        from pyhanko.keys import load_certs_from_pemder
        from pyhanko.pdf_utils.reader import PdfFileReader
        from pyhanko.sign.validation import validate_pdf_signature
        from pyhanko_certvalidator import ValidationContext
    except ImportError:
        return "WARNING", "Signature verification unavailable", "Install pyHanko to validate CMS signatures."

    roots, signer_hashes, config_error = _trust_profile(bank_code, source)
    try:
        reader = PdfFileReader(BytesIO(raw), strict=False)
        if reader.encrypted:
            if not password:
                return "WARNING", "Signature verification unavailable", "The PDF password is required."
            reader.decrypt(password)
        signatures = reader.embedded_regular_signatures
        if not signatures:
            return "WARNING", "Signature object could not be verified", "No complete embedded signature field was found."
        root_certs = list(load_certs_from_pemder([str(path) for path in roots])) if roots else []
        findings = []
        verified = 0
        altered = 0
        for index, signature in enumerate(signatures, 1):
            try:
                context = ValidationContext(trust_roots=root_certs, allow_fetching=False,
                                            revocation_mode="soft-fail")
                status = validate_pdf_signature(signature, context, skip_diff=True)
                signer_hash = hashlib.sha256(status.signing_cert.dump()).hexdigest()
                bank_match = bool(signer_hashes and signer_hash in signer_hashes)
                coverage = str(getattr(status.coverage, "name", status.coverage))
                full_file = coverage == "ENTIRE_FILE"
                intact = bool(status.intact and status.valid)
                trusted = False
                if intact and bank_match and roots and full_file:
                    try:
                        strict_context = ValidationContext(trust_roots=root_certs, allow_fetching=False,
                                                           revocation_mode="hard-fail")
                        trusted = bool(validate_pdf_signature(signature, strict_context, skip_diff=True).bottom_line)
                    except Exception:
                        trusted = False
                verified += trusted
                altered += not intact
                findings.append(f"Signature {index}: cryptographic integrity={'valid' if intact else 'invalid'}; "
                                f"bank signer={'matched' if bank_match else 'unverified'}; "
                                f"bank trust and revocation={'verified' if trusted else 'unverified'}; "
                                f"coverage={coverage}; signer SHA-256={signer_hash}.")
            except Exception as exc:
                findings.append(f"Signature {index}: validation failed ({type(exc).__name__}: {exc}).")
        details = "\n".join(findings[:20] + ([config_error] if config_error else []))
        if altered:
            return "FAIL", "Cryptographic signature integrity failed", details
        if verified == len(signatures):
            return "PASS", "Bank signatures verified for entire file", details
        return "WARNING", "Signature or bank identity remains unverified", details
    except Exception as exc:
        return "WARNING", "Signature verification could not complete", f"{type(exc).__name__}: {exc}"
