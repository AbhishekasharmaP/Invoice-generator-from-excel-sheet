import os
import re
import zipfile
from datetime import datetime, date
from io import BytesIO
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image as RLImage, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

# ---------------------------------------------------------------------------
# Fonts (so the rupee symbol prints correctly instead of a black square)
# ---------------------------------------------------------------------------
FONT = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
RUPEE = "Rs. "  # fallback if no Unicode font is found


def setup_fonts():
    global FONT, FONT_BOLD, RUPEE
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        (os.path.join(here, "fonts", "DejaVuSans.ttf"), os.path.join(here, "fonts", "DejaVuSans-Bold.ttf")),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ]
    for regular, bold in candidates:
        if os.path.exists(regular) and os.path.exists(bold):
            try:
                pdfmetrics.registerFont(TTFont("InvFont", regular))
                pdfmetrics.registerFont(TTFont("InvFont-Bold", bold))
                pdfmetrics.registerFontFamily(
                    "InvFont", normal="InvFont", bold="InvFont-Bold",
                    italic="InvFont", boldItalic="InvFont-Bold")
                FONT, FONT_BOLD, RUPEE = "InvFont", "InvFont-Bold", "₹"
                return
            except Exception:
                continue


setup_fonts()

# ---------------------------------------------------------------------------
# Excel columns
# ---------------------------------------------------------------------------
# internal key -> label shown to the user
REQUIRED_COLUMNS = {
    "creator_name": "Creator Name",
    "pan": "PAN",
    "mobile_number": "Mobile Number",
    "invoice_number": "Invoice Number",
    "campaign_name": "Campaign Name",
    "amount": "Amount",
    "bank_account_number": "Bank Account Number",
    "ifsc": "IFSC",
    "bank_name": "Bank Name",
    "email": "Email Address",
    "branch": "Branch",
}

# alternate header spellings (after normalising) -> internal key
COLUMN_ALIASES = {
    "si_no": "sl_no", "s_no": "sl_no", "sno": "sl_no", "serial_no": "sl_no",
    "creator": "creator_name",
    "mobile": "mobile_number", "phone": "mobile_number", "phone_number": "mobile_number",
    "invoice_no": "invoice_number", "invoice": "invoice_number",
    "campaign": "campaign_name",
    "account_number": "bank_account_number", "account_no": "bank_account_number",
    "bank_account_no": "bank_account_number",
    "ifsc_code": "ifsc",
    "bank": "bank_name",
    "email_address": "email", "email_id": "email", "e_mail": "email",
    "branch_name": "branch",
}


def normalize_columns(df):
    """Case-insensitive, space/punctuation-insensitive header matching."""
    cols = []
    for c in df.columns:
        key = re.sub(r"[^a-z0-9]+", "_", str(c).strip().lower()).strip("_")
        cols.append(COLUMN_ALIASES.get(key, key))
    df.columns = cols
    return df


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def clean(val, default=""):
    """Turn any Excel cell value into tidy text (no 'nan', no '123.0')."""
    if val is None:
        return default
    if isinstance(val, (datetime, pd.Timestamp)):
        return default if pd.isna(val) else val.strftime("%d %b %Y")
    if isinstance(val, date):
        return val.strftime("%d %b %Y")
    try:
        if pd.isna(val):
            return default
    except (TypeError, ValueError):
        pass
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    text = str(val).strip()
    return text if text else default


def indian_format(n):
    sign = "-" if n < 0 else ""
    int_part, dec = f"{abs(n):.2f}".split(".")
    if len(int_part) > 3:
        head, tail = int_part[:-3], int_part[-3:]
        head = re.sub(r"(\d+?)(?=(\d{2})+$)", r"\1,", head)
        int_part = head + "," + tail
    return f"{sign}{int_part}.{dec}"


def money(n):
    return f"{RUPEE}{indian_format(n)}"


def number_to_words(amount):
    """Indian numbering system, with paise."""
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    teens = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
             "Seventeen", "Eighteen", "Nineteen"]

    def below_thousand(n):
        if n == 0:
            return ""
        if n < 10:
            return ones[n]
        if n < 20:
            return teens[n - 10]
        if n < 100:
            return tens[n // 10] + (" " + ones[n % 10] if n % 10 else "")
        return ones[n // 100] + " Hundred" + (" " + below_thousand(n % 100) if n % 100 else "")

    def whole(num):
        if num == 0:
            return "Zero"
        parts = []
        for size, name in ((10000000, "Crore"), (100000, "Lakh"), (1000, "Thousand")):
            chunk, num = divmod(num, size)
            if chunk:
                parts.append(below_thousand(chunk) + " " + name)
        if num:
            parts.append(below_thousand(num))
        return " ".join(parts)

    total_paise = int(round(amount * 100))
    rupees, paise = divmod(total_paise, 100)
    words = whole(rupees) + " Rupees"
    if paise:
        words += " and " + whole(paise) + " Paise"
    return words + " Only"


def safe_filename(text):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("_") or "invoice"


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def generate_invoice_pdf(bill_to, frm, inv, logo_path=None):
    """bill_to: constant (sidebar). frm + inv: from one Excel row."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5 * inch,
                            bottomMargin=0.5 * inch, leftMargin=0.5 * inch,
                            rightMargin=0.5 * inch)
    styles = getSampleStyleSheet()
    base = ParagraphStyle("Base", parent=styles["Normal"], fontName=FONT)
    right = ParagraphStyle("Right", parent=base, alignment=TA_RIGHT)
    from_style = ParagraphStyle("From", parent=base, fontSize=8, leading=11, alignment=TA_RIGHT)
    title_style = ParagraphStyle("Title", parent=base, fontName=FONT_BOLD, fontSize=18,
                                 spaceAfter=20, spaceBefore=10)
    cell = ParagraphStyle("Cell", parent=base, fontSize=9, leading=12)
    e = lambda v: escape(str(v))

    elements = []

    # --- FROM (varies per row, top right) ---
    from_lines = [f"<b>{e(frm['creator_name'])}</b>"]
    if frm["pan"]:
        from_lines.append(f"PAN: {e(frm['pan'])}")
    if frm["mobile"]:
        from_lines.append(f"Mobile: {e(frm['mobile'])}")
    if frm["email"]:
        from_lines.append(f"Email: {e(frm['email'])}")
    from_para = Paragraph("<br/>".join(from_lines), from_style)

    if logo_path and os.path.exists(logo_path):
        left_cell = RLImage(logo_path, width=1.2 * inch, height=0.8 * inch)
    else:
        left_cell = ""
    header = Table([[left_cell, from_para]], colWidths=[1.5 * inch, 5.5 * inch])
    header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements += [header, Spacer(1, 0.1 * inch), Paragraph("INVOICE", title_style)]

    # --- BILL TO (constant) + invoice details ---
    bill_lines = ["<b>Bill To:</b>", f"<b>{e(bill_to['name'])}</b>"]
    if bill_to["gstin"]:
        bill_lines.append(f"<font size=8>GSTIN: {e(bill_to['gstin'])}</font>")
    bill_lines.append(f"<font size=8>{e(bill_to['address']).replace(chr(10), '<br/>')}</font>")

    details = (f"<b>Invoice #:</b> {e(inv['invoice_number'])}<br/>"
               f"<b>Invoice Date:</b> {e(inv['invoice_date'])}<br/>"
               f"<b>Due Date:</b> {e(inv['due_date'])}")
    info = Table([[Paragraph("<br/>".join(bill_lines), base), Paragraph(details, right)]],
                 colWidths=[3.8 * inch, 3.2 * inch])
    info.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                              ("TOPPADDING", (0, 0), (-1, -1), 10),
                              ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
    elements += [info, Spacer(1, 0.2 * inch)]

    # --- Items ---
    amount = inv["amount"]
    rows = [
        ["#", "Item", "Amount"],
        ["1", Paragraph(e(inv["campaign_name"]), cell), money(amount)],
        ["", "Total", money(amount)],
        ["", "Total Items / Qty: 1 / 1", ""],
        ["", "", ""],
        ["", "Amount Payable:", money(amount)],
        ["", f"Total amount (in words): INR {number_to_words(amount)}", ""],
    ]
    items = Table(rows, colWidths=[0.5 * inch, 4.8 * inch, 1.7 * inch])
    items.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, 2), 0.5, colors.grey),
        ("ALIGN", (2, 1), (2, -1), "RIGHT"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("FONTNAME", (1, 2), (-1, 2), FONT_BOLD),
        ("FONTNAME", (1, 5), (-1, 5), FONT_BOLD),
        ("FONTSIZE", (1, 5), (-1, 5), 11),
        ("LINEABOVE", (1, 5), (-1, 5), 1, colors.black),
        ("SPAN", (1, 6), (2, 6)),
        ("ALIGN", (1, 6), (1, 6), "LEFT"),
        ("FONTSIZE", (1, 6), (2, 6), 8),
        ("TOPPADDING", (1, 6), (2, 6), 10),
    ]))
    elements += [items, Spacer(1, 0.3 * inch)]

    # --- Bank details (from the same Excel row) + signature ---
    bank_lines = ["<b>Bank Details:</b>"]
    bank_lines.append(f"<font size=8>Bank: {e(frm['bank_name'])}</font>")
    bank_lines.append(f"<font size=8>Account Holder: {e(frm['creator_name'])}</font>")
    bank_lines.append(f"<font size=8>Account #: {e(frm['account'])}</font>")
    bank_lines.append(f"<font size=8>IFSC Code: {e(frm['ifsc'])}</font>")
    bank_lines.append(f"<font size=8>Branch: {e(frm['branch'])}</font>")

    signature = (f"<font size=8>For {e(frm['creator_name'])}<br/><br/><br/><br/>"
                 f"Authorized Signatory</font>")
    footer = Table([[Paragraph("<br/>".join(bank_lines), base), Paragraph(signature, right)]],
                   colWidths=[3.8 * inch, 3.2 * inch])
    footer.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements.append(footer)

    doc.build(elements)
    buffer.seek(0)
    return buffer


def build_template_bytes():
    sample = pd.DataFrame({
        "Sl No": [1, 2, 3],
        "Creator Name": ["John Doe", "Jane Smith", "Bob Johnson"],
        "PAN": ["ABCDE1234F", "FGHIJ5678K", "KLMNO9012P"],
        "Mobile Number": ["9876543210", "9876543211", "9876543212"],
        "Invoice Number": ["INV-001", "INV-002", "INV-003"],
        "Campaign Name": ["Twitter Campaign - Product Launch",
                          "Instagram Ads - Festival Sale",
                          "LinkedIn Campaign - B2B Lead Gen"],
        "Amount": [110000, 75000, 50000],
        "Bank Account Number": ["1234567890", "0987654321", "1122334455"],
        "IFSC": ["KARB0000123", "HDFC0000124", "SBIN0000125"],
        "Bank Name": ["Karnataka Bank", "HDFC Bank", "State Bank of India"],
        "Email Address": ["john.doe@example.com", "jane.smith@example.com", "bob.johnson@example.com"],
        "Branch": ["Bangalore", "Mumbai", "Chennai"],
        "Invoice Date": ["15 Feb 2026", "15 Feb 2026", "15 Feb 2026"],
        "Due Date": ["28 Feb 2026", "28 Feb 2026", "28 Feb 2026"],
    })
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sample.to_excel(writer, index=False, sheet_name="Sheet1")
    buf.seek(0)
    return buf.getvalue()


# ===========================================================================
# STREAMLIT UI
# ===========================================================================
st.set_page_config(page_title="Bulk Invoice Generator", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f77b4; margin-bottom: 0.5rem; }
    .sub-header { font-size: 1.1rem; color: #666; margin-bottom: 2rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">📄 Bulk Invoice Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Generate multiple invoices from one Excel file</p>',
            unsafe_allow_html=True)

# ---- Sidebar: constant BILL TO only ----
with st.sidebar:
    st.header("📋 BILL TO (Constant)")
    st.markdown("*This appears on all invoices*")
    bill_to_name = st.text_input("Client Name", "")
    bill_to_address = st.text_area(
        "Client Address",
        "")
    bill_to_gstin = st.text_input("Client GSTIN (optional)", "")

    st.markdown("---")
    st.subheader("🎨 Logo (Optional)")
    logo_file = st.file_uploader("Upload Logo", type=["png", "jpg", "jpeg"])
    logo_path = None
    if logo_file:
        logo_path = os.path.join("/tmp", f"logo_{safe_filename(logo_file.name)}")
        with open(logo_path, "wb") as f:
            f.write(logo_file.getbuffer())

# ---- Main ----
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📊 Upload Excel File")
    st.markdown("""
**Required columns (case-insensitive):**
Creator Name, PAN, Mobile Number, Invoice Number, Campaign Name, Amount,
Bank Account Number, IFSC, Bank Name, Email Address, Branch

**Optional columns:** Sl No, Invoice Date, Due Date (current date is used if empty)

Creator Name, PAN, Mobile Number, Email Address and all bank details come from each row
and appear in the **From** block of that row's invoice.
""")
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

with col2:
    st.header("📥 Sample Template")
    st.download_button(
        label="📄 Download Excel Template",
        data=build_template_bytes(),
        file_name="invoice_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, dtype=object)
        df = normalize_columns(df).dropna(how="all")

        missing = [label for key, label in REQUIRED_COLUMNS.items() if key not in df.columns]
        if missing:
            st.error(f"❌ Missing required columns: {', '.join(missing)}")
            st.info("💡 Header names are not case-sensitive. Download the template to compare.")
        else:
            file_sig = f"{uploaded_file.name}-{uploaded_file.size}"

            st.header("👀 Preview Data")
            st.dataframe(df.head(10))
            st.info(f"📊 Rows found: **{len(df)}** → will generate **{len(df)} invoices**")

            st.header("📦 Download Options")
            download_option = st.radio(
                "Choose download format:",
                ["ZIP File (All PDFs)", "Individual PDFs"],
                index=0,
            )

            if st.button("🎨 Generate Invoices", type="primary"):
                bill_to = {
                    "name": bill_to_name,
                    "address": bill_to_address,
                    "gstin": bill_to_gstin.strip(),
                }
                today = datetime.now().strftime("%d %b %Y")
                results, errors, used_names = [], [], set()
                progress = st.progress(0)

                for pos, (idx, row) in enumerate(df.iterrows(), start=1):
                    excel_row = idx + 2  # header is row 1
                    try:
                        creator = clean(row["creator_name"])
                        inv_no = clean(row["invoice_number"])
                        campaign = clean(row["campaign_name"])
                        if not creator or not inv_no or not campaign:
                            raise ValueError("Creator Name, Invoice Number and Campaign Name are required")

                        amount = float(str(row["amount"]).replace(",", "").strip())
                        if amount <= 0:
                            raise ValueError(f"Amount must be greater than 0 (found {amount:g})")

                        frm = {
                            "creator_name": creator,
                            "pan": clean(row["pan"]),
                            "mobile": clean(row["mobile_number"]),
                            "email": clean(row["email"]),
                            "bank_name": clean(row["bank_name"]),
                            "branch": clean(row["branch"]),
                            "account": clean(row["bank_account_number"]),
                            "ifsc": clean(row["ifsc"]),
                        }
                        inv = {
                            "invoice_number": inv_no,
                            "invoice_date": clean(row.get("invoice_date"), today),
                            "due_date": clean(row.get("due_date"), today),
                            "campaign_name": campaign,
                            "amount": amount,
                        }

                        pdf = generate_invoice_pdf(bill_to, frm, inv, logo_path)

                        name = safe_filename(f"Invoice_{inv_no}")
                        unique, n = name, 2
                        while unique in used_names:
                            unique, n = f"{name}_{n}", n + 1
                        used_names.add(unique)
                        results.append({"filename": unique + ".pdf", "data": pdf.getvalue()})
                    except Exception as err:
                        errors.append(f"Row {excel_row}: {err}")
                    progress.progress(pos / len(df))

                st.session_state["generated"] = {"sig": file_sig, "results": results, "errors": errors}

            gen = st.session_state.get("generated")
            if gen and gen["sig"] == file_sig:
                results, errors = gen["results"], gen["errors"]
                if results:
                    st.success(f"✅ Generated {len(results)} invoice(s) successfully!")
                for msg in errors:
                    st.warning(f"⚠️ Skipped - {msg}")

                if results:
                    if download_option.startswith("ZIP"):
                        zip_buf = BytesIO()
                        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                            for item in results:
                                zf.writestr(item["filename"], item["data"])
                        st.download_button(
                            label=f"📦 Download All {len(results)} Invoices (ZIP)",
                            data=zip_buf.getvalue(),
                            file_name=f"Invoices_{datetime.now().strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                        )
                    else:
                        st.subheader("📄 Download Individual PDFs")
                        cols = st.columns(3)
                        for i, item in enumerate(results):
                            with cols[i % 3]:
                                st.download_button(
                                    label=f"📄 {item['filename']}",
                                    data=item["data"],
                                    file_name=item["filename"],
                                    mime="application/pdf",
                                    key=f"pdf_{i}",
                                )
    except Exception as err:
        st.error(f"❌ Error processing file: {err}")
        st.info("💡 Please make sure your Excel file matches the template format.")
else:
    st.info("👆 Upload an Excel file to get started, or download the sample template")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>One Excel row → one invoice PDF</p>
    <p style='font-size: 0.8rem;'>Bill To (constant, sidebar) • From + bank details (per row, Excel)</p>
</div>
""", unsafe_allow_html=True)
