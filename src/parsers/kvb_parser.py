"""First-page header profile used only by PDF_Status."""

# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "Karur Vysya Bank",
    "ifsc": "KVBL",
    "aliases": [
        "Karur Vysya Bank",
        "Karur Vysya",
        "KVB"
    ],
    "name_after": "Account Statement"
}



PDF_STATUS_PROFILE.update({'aliases': ['Karur Vysya Bank', 'Karur Vysya', 'KVB', 'KarurVysyaBank'], 'name_after': 'STATEMENT'})
