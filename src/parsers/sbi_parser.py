"""First-page header profile used only by PDF_Status."""

# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "State Bank of India",
    "ifsc": "SBIN",
    "aliases": [
        "State Bank of India"
    ],
    "name_after": "STATEMENT OF ACCOUNT"
}



PDF_STATUS_PROFILE.update({'aliases': ['State Bank of India', 'SBI']})
