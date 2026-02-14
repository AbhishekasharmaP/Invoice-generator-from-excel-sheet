import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_LEFT
from io import BytesIO
from datetime import datetime
import os
import zipfile

# Page configuration
st.set_page_config(
    page_title="Bulk Invoice Generator - APJ Digital Solutions",
    page_icon="📄",
    layout="wide"
)

def number_to_words(num):
    """Convert number to words (Indian system)"""
    ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine']
    tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
    teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
    
    def convert_below_thousand(n):
        if n == 0:
            return ''
        elif n < 10:
            return ones[n]
        elif n < 20:
            return teens[n - 10]
        elif n < 100:
            return tens[n // 10] + (' ' + ones[n % 10] if n % 10 != 0 else '')
        else:
            return ones[n // 100] + ' Hundred' + (' ' + convert_below_thousand(n % 100) if n % 100 != 0 else '')
    
    if num == 0:
        return 'Zero Rupees Only'
    
    # Indian system: crore, lakh, thousand
    crore = num // 10000000
    num %= 10000000
    lakh = num // 100000
    num %= 100000
    thousand = num // 1000
    num %= 1000
    
    result = []
    if crore:
        result.append(convert_below_thousand(crore) + ' Crore')
    if lakh:
        result.append(convert_below_thousand(lakh) + ' Lakh')
    if thousand:
        result.append(convert_below_thousand(thousand) + ' Thousand')
    if num:
        result.append(convert_below_thousand(num))
    
    return ' '.join(result) + ' Rupees Only'

def generate_invoice_pdf(bill_to_info, from_info, invoice_data, company_info, logo_path=None):
    """Generate single invoice PDF"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, 
                           topMargin=0.5*inch, bottomMargin=0.5*inch,
                           leftMargin=0.5*inch, rightMargin=0.5*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    company_detail_style = ParagraphStyle(
        'CompanyDetail',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.black,
        alignment=TA_RIGHT,
        leading=10,
    )
    
    invoice_title_style = ParagraphStyle(
        'InvoiceTitle',
        parent=styles['Normal'],
        fontSize=18,
        fontName='Helvetica-Bold',
        spaceAfter=20,
        spaceBefore=10,
    )
    
    # Header with logo and company info (FROM - varies per row)
    header_data = []
    if logo_path and os.path.exists(logo_path):
        logo = RLImage(logo_path, width=1.2*inch, height=0.8*inch)
        from_info_text = f"""<b>{from_info['creator_name']}</b><br/>
        <font size=8>PAN: {from_info['pan']}<br/>
        Mobile: {from_info['mobile']}<br/>
        Email: {company_info.get('email', '')}</font>"""
        header_data = [[logo, Paragraph(from_info_text, company_detail_style)]]
    else:
        from_info_text = f"""<b>{from_info['creator_name']}</b><br/>
        <font size=8>PAN: {from_info['pan']}<br/>
        Mobile: {from_info['mobile']}<br/>
        Email: {company_info.get('email', '')}</font>"""
        header_data = [['', Paragraph(from_info_text, company_detail_style)]]
    
    header_table = Table(header_data, colWidths=[1.5*inch, 5*inch])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 0.1*inch))
    
    # Invoice title
    elements.append(Paragraph("INVOICE", invoice_title_style))
    
    # Bill To (CONSTANT) and Invoice details
    bill_to_text = f"""<b>Bill To:</b><br/>
    <b>{bill_to_info['name']}</b><br/>
    <font size=8>{bill_to_info['address']}</font>"""
    
    if bill_to_info.get('gstin'):
        bill_to_text += f"<br/><font size=8>GSTIN: {bill_to_info['gstin']}</font>"
    
    invoice_details_text = f"""<b>Invoice #:</b> {invoice_data['invoice_number']}<br/>
    <b>Invoice Date:</b> {invoice_data['invoice_date']}<br/>
    <b>Due Date:</b> {invoice_data['due_date']}"""
    
    info_data = [[
        Paragraph(bill_to_text, styles['Normal']),
        Paragraph(invoice_details_text, ParagraphStyle('RightAlign', parent=styles['Normal'], alignment=TA_RIGHT))
    ]]
    
    info_table = Table(info_data, colWidths=[3.5*inch, 3*inch])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 0.2*inch))
    
    # Items table
    items_data = [['#', 'Item', 'Amount']]
    items_data.append([
        '1',
        str(invoice_data['campaign_name']),
        f"₹{invoice_data['amount']:,.2f}"
    ])
    
    # Total
    total = invoice_data['amount']
    items_data.append(['', 'Total', f"₹{total:,.2f}"])
    items_data.append(['', f'Total Items / Qty: 1 / 1', ''])
    items_data.append(['', '', ''])
    items_data.append(['', 'Amount Payable:', f"₹{total:,.2f}"])
    
    # Amount in words
    total_int = int(total)
    amount_words = number_to_words(total_int)
    items_data.append(['', f'Total amount (in words): INR {amount_words}', ''])
    
    items_table = Table(items_data, colWidths=[0.5*inch, 4.5*inch, 1.5*inch])
    items_table.setStyle(TableStyle([
        # Header row
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Data rows
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -5), 0.5, colors.grey),
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        
        # Total rows styling
        ('FONTNAME', (1, -5), (-1, -5), 'Helvetica-Bold'),
        ('FONTNAME', (1, -2), (-1, -2), 'Helvetica-Bold'),
        ('FONTSIZE', (1, -2), (-1, -2), 11),
        ('LINEABOVE', (1, -2), (-1, -2), 1, colors.black),
        
        # Amount in words
        ('SPAN', (1, -1), (2, -1)),
        ('FONTSIZE', (1, -1), (2, -1), 8),
        ('TOPPADDING', (1, -1), (2, -1), 10),
    ]))
    
    elements.append(items_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # Bank details (varies per row) and signature
    bank_text = f"""<b>Bank Details:</b><br/>
    <font size=8>Bank: {company_info.get('bank_name', '')}<br/>
    Account Holder: {from_info['creator_name']}<br/>
    Account #: {invoice_data.get('bank_account_number', '')}<br/>
    IFSC Code: {invoice_data.get('ifsc', '')}<br/>
    Branch: {company_info.get('branch', '')}</font>"""
    
    signature_text = f"""<br/><br/><br/>
    <font size=8>For {from_info['creator_name']}<br/><br/><br/><br/>
    Authorized Signatory</font>"""
    
    footer_data = [[
        Paragraph(bank_text, styles['Normal']),
        Paragraph(signature_text, ParagraphStyle('RightAlign', parent=styles['Normal'], alignment=TA_RIGHT))
    ]]
    
    footer_table = Table(footer_data, colWidths=[3.5*inch, 3*inch])
    footer_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    elements.append(footer_table)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

def normalize_column_names(df):
    """Normalize column names to lowercase and remove spaces"""
    df.columns = df.columns.str.lower().str.strip().str.replace(' ', '_').str.replace('.', '')
    return df

# App styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #666;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# App title
st.markdown('<p class="main-header">📄 Bulk Invoice Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Generate multiple invoices from one Excel file</p>', unsafe_allow_html=True)

# Sidebar for BILL TO (constant) and Company info
with st.sidebar:
    st.header("📋 BILL TO (Constant)")
    st.markdown("*This appears on all invoices*")
    
    bill_to_name = st.text_input("Client Name", "NEXGROW DIGITAL PRIVATE LIMITED")
    bill_to_address = st.text_area("Client Address", 
        "Block C, 4th Floor, 56/6\nSector 62, Noida\nGautambuddha Nagar, UTTAR PRADESH, 201309")
    bill_to_gstin = st.text_input("Client GSTIN (optional)", "09AAKCN1659F1Z8")
    
    st.markdown("---")
    st.header("🏢 Company Details")
    st.markdown("*Additional info for invoices*")
    
    company_email = st.text_input("Company Email", "apjdigitalsol@gmail.com")
    bank_name = st.text_input("Bank Name", "Karnataka Bank")
    branch = st.text_input("Branch", "Bangalore")
    
    # Logo upload
    st.markdown("---")
    st.subheader("🎨 Company Logo (Optional)")
    logo_file = st.file_uploader("Upload Logo", type=['png', 'jpg', 'jpeg'])
    logo_path = None
    if logo_file:
        logo_path = f"/tmp/logo_{logo_file.name}"
        with open(logo_path, "wb") as f:
            f.write(logo_file.getbuffer())

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📊 Upload Excel File")
    st.markdown("""
    **Required columns (case-insensitive):**
    - Sl No / SI No
    - Creator Name
    - PAN
    - Mobile Number
    - Invoice Number
    - Campaign Name
    - Amount
    - Bank Account Number
    - IFSC
    
    **Optional columns:**
    - Invoice Date (auto-generated if empty)
    - Due Date (auto-generated if empty)
    """)
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

with col2:
    st.header("📥 Sample Template")
    if st.button("Download Template", type="secondary", use_container_width=True):
        # Create sample template
        sample_data = pd.DataFrame({
            'Sl No': [1, 2, 3],
            'Creator Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
            'PAN': ['ABCDE1234F', 'FGHIJ5678K', 'KLMNO9012P'],
            'Mobile Number': ['9876543210', '9876543211', '9876543212'],
            'Invoice Number': ['INV-001', 'INV-002', 'INV-003'],
            'Campaign Name': ['Twitter Campaign - Product Launch', 'Instagram Ads - Festival Sale', 'LinkedIn Campaign - B2B Lead Gen'],
            'Amount': [110000, 75000, 50000],
            'Bank Account Number': ['1234567890', '0987654321', '1122334455'],
            'IFSC': ['KARB0000123', 'KARB0000124', 'KARB0000125'],
            'Invoice Date': ['15 Feb 2026', '15 Feb 2026', '15 Feb 2026'],
            'Due Date': ['28 Feb 2026', '28 Feb 2026', '28 Feb 2026']
        })
        
        sample_buffer = BytesIO()
        with pd.ExcelWriter(sample_buffer, engine='openpyxl') as writer:
            sample_data.to_excel(writer, index=False, sheet_name='Sheet1')
        
        sample_buffer.seek(0)
        st.download_button(
            label="📄 Download Excel Template",
            data=sample_buffer,
            file_name="Bulk_Invoice_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Process uploaded file
if uploaded_file is not None:
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file)
        
        # Normalize column names
        df = normalize_column_names(df)
        
        # Check required columns
        required_cols = ['creator_name', 'pan', 'mobile_number', 'invoice_number', 
                        'campaign_name', 'amount', 'bank_account_number', 'ifsc']
        
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
            st.info("💡 Column names are case-insensitive. Check your Excel file format.")
        else:
            # Preview data
            st.header("👀 Preview Data")
            st.dataframe(df.head(10), use_container_width=True)
            st.info(f"📊 Total rows found: **{len(df)}** → Will generate **{len(df)} invoices**")
            
            # Download option
            st.header("📦 Download Options")
            download_option = st.radio(
                "Choose download format:",
                ["ZIP File (All PDFs)", "Individual PDFs"],
                index=0
            )
            
            # Generate invoices button
            if st.button("🎨 Generate Invoices", type="primary", use_container_width=True):
                with st.spinner(f"Generating {len(df)} invoices..."):
                    # Prepare constant data
                    bill_to_info = {
                        'name': bill_to_name,
                        'address': bill_to_address,
                        'gstin': bill_to_gstin if bill_to_gstin else None
                    }
                    
                    company_info = {
                        'email': company_email,
                        'bank_name': bank_name,
                        'branch': branch
                    }
                    
                    # Generate PDFs
                    pdf_buffers = []
                    current_date = datetime.now().strftime('%d %b %Y')
                    
                    progress_bar = st.progress(0)
                    
                    for idx, row in df.iterrows():
                        # Prepare FROM info (varies per row)
                        from_info = {
                            'creator_name': str(row['creator_name']),
                            'pan': str(row['pan']),
                            'mobile': str(row['mobile_number'])
                        }
                        
                        # Prepare invoice data
                        invoice_data = {
                            'invoice_number': str(row['invoice_number']),
                            'invoice_date': str(row.get('invoice_date', current_date)),
                            'due_date': str(row.get('due_date', current_date)),
                            'campaign_name': str(row['campaign_name']),
                            'amount': float(row['amount']),
                            'bank_account_number': str(row['bank_account_number']),
                            'ifsc': str(row['ifsc'])
                        }
                        
                        # Generate PDF
                        pdf_buffer = generate_invoice_pdf(
                            bill_to_info, from_info, invoice_data, company_info, logo_path
                        )
                        
                        pdf_buffers.append({
                            'buffer': pdf_buffer,
                            'filename': f"Invoice_{invoice_data['invoice_number']}.pdf"
                        })
                        
                        # Update progress
                        progress_bar.progress((idx + 1) / len(df))
                    
                    st.success(f"✅ Generated {len(pdf_buffers)} invoices successfully!")
                    
                    # Download based on option
                    if download_option == "ZIP File (All PDFs)":
                        # Create ZIP file
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for pdf_data in pdf_buffers:
                                zip_file.writestr(pdf_data['filename'], pdf_data['buffer'].getvalue())
                        
                        zip_buffer.seek(0)
                        
                        st.download_button(
                            label=f"📦 Download All {len(pdf_buffers)} Invoices (ZIP)",
                            data=zip_buffer,
                            file_name=f"Invoices_{datetime.now().strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                    else:
                        # Individual downloads
                        st.subheader("📄 Download Individual PDFs")
                        cols = st.columns(3)
                        for idx, pdf_data in enumerate(pdf_buffers):
                            with cols[idx % 3]:
                                st.download_button(
                                    label=f"📄 {pdf_data['filename']}",
                                    data=pdf_data['buffer'],
                                    file_name=pdf_data['filename'],
                                    mime="application/pdf",
                                    use_container_width=True,
                                    key=f"pdf_{idx}"
                                )
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("💡 Please ensure your Excel file matches the template format.")

else:
    st.info("👆 Upload an Excel file to get started, or download the sample template")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>Built with Streamlit | One Excel Row → One Invoice PDF</p>
    <p style='font-size: 0.8rem;'>Bill To (constant from sidebar) • From (varies per row from Excel)</p>
</div>
""", unsafe_allow_html=True)
