"""First-page header profile used only by PDF_Status."""

# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "IndusInd Bank",
    "ifsc": "INDB",
    "aliases": [
        "IndusInd Bank"
    ],
    "unlabelled_left": True
}



PDF_STATUS_PROFILE.update({'address_is_branch': True})
