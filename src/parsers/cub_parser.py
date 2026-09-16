"""First-page header profile used only by PDF_Status."""

# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "City Union Bank",
    "ifsc": "CIUB",
    "aliases": [
        "City Union Bank"
    ],
    "customer_label": "Customer Details"
}



PDF_STATUS_PROFILE.update({'unlabelled_left': True})
