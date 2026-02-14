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

# Page configuration
st.set_page_config(
    page_title="Invoice Generator - APJ Digital Solutions",
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

def generate_invoice_pdf(vendor_info, client_info, invoice_data, items_df, logo_path=None):
    """Generate invoice PDF matching APJ Digital Solutions format"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, 
                           topMargin=0.5*inch, bottomMargin=0.5*inch,
                           leftMargin=0.5*inch, rightMargin=0.5*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    company_name_style = ParagraphStyle(
        'CompanyName',
        parent=styles['Normal'],
        fontSize=16,
        textColor=colors.black,
        alignment=TA_RIGHT,
        fontName='Helvetica-Bold',
        spaceAfter=2,
    )
    
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
    
    # Header with logo and company info
    header_data = []
    if logo_path and os.path.exists(logo_path):
        logo = RLImage(logo_path, width=1.2*inch, height=0.8*inch)
        company_info_text = f"""<b>{vendor_info['name']}</b><br/>
        <font size=8>PAN {vendor_info['pan']}<br/>
        {vendor_info['address']}<br/>
        Email: {vendor_info['email']}</font>"""
        header_data = [[logo, Paragraph(company_info_text, company_detail_style)]]
    else:
        # Without logo
        company_info_text = f"""<b>{vendor_info['name']}</b><br/>
        <font size=8>PAN {vendor_info['pan']}<br/>
        {vendor_info['address']}<br/>
        Email: {vendor_info['email']}</font>"""
        header_data = [['', Paragraph(company_info_text, company_detail_style)]]
    
    header_table = Table(header_data, colWidths=[1.5*inch, 5*inch])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 0.1*inch))
    
    # Invoice title
    elements.append(Paragraph("INVOICE", invoice_title_style))
    
    # Bill To and Invoice details table
    bill_to_text = f"""<b>Bill To:</b><br/>
    <b>{client_info['name']}</b><br/>
    <font size=8>GSTIN: {client_info['gstin']}<br/>
    {client_info['address']}</font>"""
    
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
    items_data = [['#', 'Item', 'HSN/SAC', 'Amount']]
    
    for idx, row in items_df.iterrows():
        items_data.append([
            str(idx + 1),
            str(row['Description']),
            str(row.get('HSN_SAC', '-')),
            f"₹{row['Amount']:,.2f}"
        ])
    
    # Calculate totals
    subtotal = items_df['Amount'].sum()
    cgst_rate = invoice_data.get('cgst_rate', 0)
    sgst_rate = invoice_data.get('sgst_rate', 0)
    cgst_amount = subtotal * (cgst_rate / 100)
    sgst_amount = subtotal * (sgst_rate / 100)
    total = subtotal + cgst_amount + sgst_amount
    
    # Add totals
    items_data.append(['', '', 'Total', f"₹{subtotal:,.2f}"])
    if cgst_rate > 0:
        items_data.append(['', '', f'CGST ({cgst_rate}%)', f"₹{cgst_amount:,.2f}"])
    if sgst_rate > 0:
        items_data.append(['', '', f'SGST ({sgst_rate}%)', f"₹{sgst_amount:,.2f}"])
    
    # Total items count
    total_items = len(items_df)
    items_data.append(['', f'Total Items / Qty: {total_items} / {total_items}', '', ''])
    items_data.append(['', '', '', ''])
    items_data.append(['', '', 'Amount Payable:', f"₹{total:,.2f}"])
    
    # Amount in words
    total_int = int(total)
    amount_words = number_to_words(total_int)
    items_data.append(['', f'Total amount (in words): INR {amount_words}', '', ''])
    
    items_table = Table(items_data, colWidths=[0.5*inch, 3.5*inch, 1.5*inch, 1*inch])
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
        ('GRID', (0, 0), (-1, -(6)), 0.5, colors.grey),
        ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        
        # Total rows styling
        ('FONTNAME', (2, -6), (-1, -6), 'Helvetica-Bold'),
        ('FONTNAME', (2, -4), (-1, -4), 'Helvetica-Bold'),
        ('FONTNAME', (2, -2), (-1, -2), 'Helvetica-Bold'),
        ('FONTSIZE', (2, -2), (-1, -2), 11),
        ('LINEABOVE', (2, -2), (-1, -2), 1, colors.black),
        
        # Amount in words
        ('SPAN', (1, -1), (3, -1)),
        ('FONTSIZE', (1, -1), (3, -1), 8),
        ('TOPPADDING', (1, -1), (3, -1), 10),
    ]))
    
    elements.append(items_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # Bank details and signature
    bank_text = f"""<b>Bank Details:</b><br/>
    <font size=8>Bank: {vendor_info.get('bank_name', '')}<br/>
    Account Holder: {vendor_info.get('account_holder', '')}<br/>
    Account #: {vendor_info.get('account_number', '')}<br/>
    IFSC Code: {vendor_info.get('ifsc_code', '')}<br/>
    Branch: {vendor_info.get('branch', '')}</font>"""
    
    signature_text = f"""<br/><br/><br/>
    <font size=8>For {vendor_info['name']}<br/><br/><br/><br/>
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
    return buffer, total

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
st.markdown('<p class="main-header">📄 Invoice Generator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Upload your Excel file to generate professional invoices</p>', unsafe_allow_html=True)

# Sidebar for vendor information
with st.sidebar:
    st.header("🏢 Vendor Information")
    
    vendor_name = st.text_input("Vendor Name", "M/S APJ DIGITAL SOLUTIONS LLP")
    vendor_pan = st.text_input("PAN Number", "ACNFA0470H")
    vendor_gstin = st.text_input("GSTIN (if applicable)", "")
    vendor_address = st.text_area("Address", 
        "No 245, 3RD MAIN, 10TH CROSS, 7 MELEKHANENGERS\nBangalore West, KARNATAKA, 560085")
    vendor_email = st.text_input("Email", "apjdigitalsol@gmail.com")
    
    st.subheader("🏦 Bank Details")
    bank_name = st.text_input("Bank Name", "Karnataka Bank")
    account_holder = st.text_input("Account Holder", "APJ DIGITAL SOLUTIONS")
    account_number = st.text_input("Account Number", "")
    ifsc_code = st.text_input("IFSC Code", "")
    branch = st.text_input("Branch", "Bangalore")
    
    # Logo upload
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
    st.header("📊 Upload Invoice Data")
    st.markdown("""
    Your Excel file should contain the following sheets:
    - **Vendor_Info**: Your company details (optional, uses sidebar if not provided)
    - **Client_Info**: Client name, GSTIN, address
    - **Invoice_Info**: Invoice number, date, due date, CGST%, SGST%
    - **Items**: Description, HSN_SAC, Amount
    """)
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

with col2:
    st.header("📥 Sample Template")
    if st.button("Download Template", type="secondary"):
        # Create sample template
        sample_vendor = pd.DataFrame({
            'name': ['M/S APJ DIGITAL SOLUTIONS LLP'],
            'pan': ['ACNFA0470H'],
            'address': ['No 245, 3RD MAIN, 10TH CROSS, Bangalore'],
            'email': ['apjdigitalsol@gmail.com'],
            'bank_name': ['Karnataka Bank'],
            'account_holder': ['APJ DIGITAL SOLUTIONS'],
            'account_number': ['1234567890'],
            'ifsc_code': ['KARB0000123'],
            'branch': ['Bangalore']
        })
        
        sample_client = pd.DataFrame({
            'name': ['NEXGROW DIGITAL PRIVATE LIMITED'],
            'gstin': ['09AAKCN1659F1Z8'],
            'address': ['Block C, 4th Floor, 56/6\nSector 62, Noida\nGautambuddha Nagar, UTTAR PRADESH, 201309']
        })
        
        sample_invoice = pd.DataFrame({
            'invoice_number': ['INV-102'],
            'invoice_date': ['09 Feb 2026'],
            'due_date': ['09 Feb 2026'],
            'cgst_rate': [0],
            'sgst_rate': [0]
        })
        
        sample_items = pd.DataFrame({
            'Description': ['Twitter Campaign Tamil Nadu - Modi visit'],
            'HSN_SAC': ['-'],
            'Amount': [110000.00]
        })
        
        sample_buffer = BytesIO()
        with pd.ExcelWriter(sample_buffer, engine='openpyxl') as writer:
            sample_vendor.to_excel(writer, sheet_name='Vendor_Info', index=False)
            sample_client.to_excel(writer, sheet_name='Client_Info', index=False)
            sample_invoice.to_excel(writer, sheet_name='Invoice_Info', index=False)
            sample_items.to_excel(writer, sheet_name='Items', index=False)
        
        sample_buffer.seek(0)
        st.download_button(
            label="📄 Download Excel Template",
            data=sample_buffer,
            file_name="invoice_template_apj.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Process uploaded file
if uploaded_file is not None:
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Check required sheets
        required_sheets = ['Client_Info', 'Invoice_Info', 'Items']
        missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
        
        if missing_sheets:
            st.error(f"❌ Missing required sheets: {', '.join(missing_sheets)}")
        else:
            # Read data
            if 'Vendor_Info' in excel_file.sheet_names:
                vendor_info_df = pd.read_excel(uploaded_file, sheet_name='Vendor_Info')
                vendor_info = vendor_info_df.iloc[0].to_dict()
            else:
                # Use sidebar values
                vendor_info = {
                    'name': vendor_name,
                    'pan': vendor_pan,
                    'address': vendor_address,
                    'email': vendor_email,
                    'bank_name': bank_name,
                    'account_holder': account_holder,
                    'account_number': account_number,
                    'ifsc_code': ifsc_code,
                    'branch': branch
                }
            
            client_info_df = pd.read_excel(uploaded_file, sheet_name='Client_Info')
            invoice_info_df = pd.read_excel(uploaded_file, sheet_name='Invoice_Info')
            items_df = pd.read_excel(uploaded_file, sheet_name='Items')
            
            # Preview section
            st.header("👀 Preview Data")
            
            tab1, tab2, tab3 = st.tabs(["📋 Invoice Info", "👤 Client Info", "📦 Items"])
            
            with tab1:
                st.dataframe(invoice_info_df, use_container_width=True)
            
            with tab2:
                st.dataframe(client_info_df, use_container_width=True)
            
            with tab3:
                st.dataframe(items_df, use_container_width=True)
            
            # Prepare data
            client_info = client_info_df.iloc[0].to_dict()
            invoice_data = invoice_info_df.iloc[0].to_dict()
            
            # Generate invoice button
            st.header("🎨 Generate Invoice")
            
            col1, col2 = st.columns([1, 3])
            
            with col1:
                if st.button("📄 Generate PDF", type="primary", use_container_width=True):
                    with st.spinner("Generating invoice..."):
                        pdf_buffer, total = generate_invoice_pdf(
                            vendor_info, client_info, invoice_data, items_df, logo_path
                        )
                        
                        st.success(f"✅ Invoice generated! Total: ₹{total:,.2f}")
                        
                        st.download_button(
                            label="💾 Download Invoice PDF",
                            data=pdf_buffer,
                            file_name=f"Invoice_{invoice_data['invoice_number']}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.info("💡 Please ensure your Excel file matches the template format. Download the sample template to see the expected structure.")

else:
    st.info("👆 Upload an Excel file to get started, or download the sample template")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p>Built with Streamlit | Upload Excel → Generate Professional Invoice PDF</p>
    <p style='font-size: 0.8rem;'>Supports GST invoicing with CGST/SGST calculations</p>
</div>
""", unsafe_allow_html=True)
