# Bulk Invoice Generator

One Excel row = one PDF invoice.

- **Bill To** (constant): entered in the left sidebar (name, address, optional GSTIN).
- **From** (varies): Creator Name, PAN, Mobile Number, Email Address from each Excel row.
- **Bank details** (varies): Bank Name, Account Number, IFSC, Branch from each Excel row.
- **Item**: Campaign Name. **Amount**: shown as-is (no tax).
- **Dates**: Invoice Date / Due Date from the row; current date if empty.
- **Download**: radio button, ZIP (default) or individual PDFs.

## Excel columns (headers are case-insensitive)

Required: Creator Name, PAN, Mobile Number, Invoice Number, Campaign Name, Amount,
Bank Account Number, IFSC, Bank Name, Email Address, Branch

Optional: Sl No, Invoice Date, Due Date

## Deploy (Streamlit Cloud)

Upload to the GitHub repo: `Invoice generator.py`, `requirements.txt`, `packages.txt`.
`packages.txt` installs a font so the rupee symbol prints correctly.
