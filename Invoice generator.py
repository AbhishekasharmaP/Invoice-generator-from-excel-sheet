import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Invoice Generator",
    page_icon="📄",
    layout="wide"
)

def generate_invoice_pdf(company_info, client_info, invoice_data, items_df):
    """Generate a professional invoice PDF"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5*inch)
    elements = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#2E4057'),
        spaceAfter=30,
    )
    
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#444444'),
    )
    
    right_align_style = ParagraphStyle(
        'RightAlign',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_RIGHT,
    )
    
    # Title
    elements.append(Paragraph("INVOICE", title_style))
    elements.append(Spacer(1, 0.2*inch))
    
    # Company and Client Info Table
    info_data = [
        [
            Paragraph(f"<b>From:</b><br/>{company_info['name']}<br/>{company_info['address']}<br/>{company_info['email']}<br/>{company_info['phone']}", header_style),
            Paragraph(f"<b>Invoice #:</b> {invoice_data['invoice_number']}<br/><b>Date:</b> {invoice_data['date']}<br/><b>Due Date:</b> {invoice_data['due_date']}", right_align_style)
        ],
        [
            Paragraph(f"<b>Bill To:</b><br/>{client_info['name']}<br/>{client_info['address']}<br/>{client_info['email']}", header_style),
            ""
        ]
    ]
    
    info_table = Table(info_data, colWidths=[4*inch, 2.5*inch])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 0.3*inch))
    
    # Items table
    items_data = [['Description', 'Quantity', 'Unit Price', 'Amount']]
    
    for _, row in items_df.iterrows():
        items_data.append([
            str(row['Description']),
            str(row['Quantity']),
            f"${row['Unit_Price']:.2f}",
            f"${row['Amount']:.2f}"
        ])
    
    # Calculate totals
    subtotal = items_df['Amount'].sum()
    tax_rate = invoice_data.get('tax_rate', 0) / 100
    tax_amount = subtotal * tax_rate
    total = subtotal + tax_amount
    
    # Add summary rows
    items_data.append(['', '', 'Subtotal:', f"${subtotal:.2f}"])
    if tax_rate > 0:
        items_data.append(['', '', f'Tax ({invoice_data.get("tax_rate", 0)}%):', f"${tax_amount:.2f}"])
    items_data.append(['', '', 'Total:', f"${total:.2f}"])
    
    items_table = Table(items_data, colWidths=[3.5*inch, 1*inch, 1.2*inch, 1*inch])
    items_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2E4057')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -4), colors.beige),
        ('GRID', (0, 0), (-1, -4), 1, colors.black),
        ('LINEBELOW', (2, -3), (-1, -3), 1, colors.black),
        ('LINEBELOW', (2, -2), (-1, -2), 1, colors.black),
        ('LINEBELOW', (2, -1), (-1, -1), 2, colors.black),
        ('FONTNAME', (2, -3), (-1, -1), 'Helvetica-Bold'),
    ]))
    
    elements.append(items_table)
    elements.append(Spacer(1, 0.5*inch))
    
    # Payment terms
    if invoice_data.get('notes'):
        elements.append(Paragraph(f"<b>Notes:</b><br/>{invoice_data['notes']}", header_style))
        elements.append(Spacer(1, 0.2*inch))
    
    if invoice_data.get('payment_terms'):
        elements.append(Paragraph(f"<b>Payment Terms:</b><br/>{invoice_data['payment_terms']}", header_style))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer, total

# App title
st.title("📄 Invoice Generator")
st.markdown("Upload your Excel file with invoice details to generate a professional PDF invoice")

# Sidebar for company information
with st.sidebar:
    st.header("Company Information")
    company_name = st.text_input("Company Name", "Your Company Name")
    company_address = st.text_area("Company Address", "123 Business St\nCity, State 12345")
    company_email = st.text_input("Company Email", "info@company.com")
    company_phone = st.text_input("Company Phone", "+1 (555) 123-4567")

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.header("Upload Excel File")
    st.markdown("Your Excel file should contain the following sheets:")
    st.markdown("""
    - **Invoice_Info**: Invoice number, date, due date, tax rate, notes, payment terms
    - **Client_Info**: Client name, address, email
    - **Items**: Description, Quantity, Unit_Price (Amount will be calculated)
    """)
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

with col2:
    st.header("Quick Actions")
    if st.button("📥 Download Sample Template"):
        # Create sample template
        sample_invoice_info = pd.DataFrame({
            'invoice_number': ['INV-001'],
            'date': [datetime.now().strftime('%Y-%m-%d')],
            'due_date': [(datetime.now()).strftime('%Y-%m-%d')],
            'tax_rate': [10],
            'notes': ['Thank you for your business!'],
            'payment_terms': ['Payment due within 30 days']
        })
        
        sample_client_info = pd.DataFrame({
            'name': ['Client Company Ltd'],
            'address': ['456 Client Ave\nCity, State 54321'],
            'email': ['client@example.com']
        })
        
        sample_items = pd.DataFrame({
            'Description': ['Web Development Services', 'Logo Design', 'Consulting Hours'],
            'Quantity': [1, 1, 10],
            'Unit_Price': [1500.00, 500.00, 100.00]
        })
        
        sample_buffer = BytesIO()
        with pd.ExcelWriter(sample_buffer, engine='openpyxl') as writer:
            sample_invoice_info.to_excel(writer, sheet_name='Invoice_Info', index=False)
            sample_client_info.to_excel(writer, sheet_name='Client_Info', index=False)
            sample_items.to_excel(writer, sheet_name='Items', index=False)
        
        sample_buffer.seek(0)
        st.download_button(
            label="Download Template",
            data=sample_buffer,
            file_name="invoice_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Process uploaded file
if uploaded_file is not None:
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Check required sheets
        required_sheets = ['Invoice_Info', 'Client_Info', 'Items']
        missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
        
        if missing_sheets:
            st.error(f"Missing required sheets: {', '.join(missing_sheets)}")
        else:
            # Read data
            invoice_info_df = pd.read_excel(uploaded_file, sheet_name='Invoice_Info')
            client_info_df = pd.read_excel(uploaded_file, sheet_name='Client_Info')
            items_df = pd.read_excel(uploaded_file, sheet_name='Items')
            
            # Calculate amounts
            items_df['Amount'] = items_df['Quantity'] * items_df['Unit_Price']
            
            # Preview section
            st.header("Preview Data")
            
            tab1, tab2, tab3 = st.tabs(["Invoice Info", "Client Info", "Items"])
            
            with tab1:
                st.dataframe(invoice_info_df, use_container_width=True)
            
            with tab2:
                st.dataframe(client_info_df, use_container_width=True)
            
            with tab3:
                st.dataframe(items_df, use_container_width=True)
            
            # Prepare data for PDF
            company_info = {
                'name': company_name,
                'address': company_address,
                'email': company_email,
                'phone': company_phone
            }
            
            client_info = {
                'name': client_info_df['name'].iloc[0],
                'address': client_info_df['address'].iloc[0],
                'email': client_info_df['email'].iloc[0]
            }
            
            invoice_data = {
                'invoice_number': invoice_info_df['invoice_number'].iloc[0],
                'date': str(invoice_info_df['date'].iloc[0]),
                'due_date': str(invoice_info_df['due_date'].iloc[0]),
                'tax_rate': invoice_info_df['tax_rate'].iloc[0] if 'tax_rate' in invoice_info_df.columns else 0,
                'notes': invoice_info_df['notes'].iloc[0] if 'notes' in invoice_info_df.columns else '',
                'payment_terms': invoice_info_df['payment_terms'].iloc[0] if 'payment_terms' in invoice_info_df.columns else ''
            }
            
            # Generate invoice button
            st.header("Generate Invoice")
            
            col1, col2, col3 = st.columns([1, 1, 2])
            
            with col1:
                if st.button("🎨 Generate PDF Invoice", type="primary"):
                    with st.spinner("Generating invoice..."):
                        pdf_buffer, total = generate_invoice_pdf(company_info, client_info, invoice_data, items_df)
                        
                        st.success(f"✅ Invoice generated successfully! Total: ${total:.2f}")
                        
                        st.download_button(
                            label="📄 Download Invoice PDF",
                            data=pdf_buffer,
                            file_name=f"Invoice_{invoice_data['invoice_number']}.pdf",
                            mime="application/pdf"
                        )
            
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.info("Please make sure your Excel file has the correct format. Download the sample template to see the expected structure.")

else:
    st.info("👆 Upload an Excel file to get started, or download the sample template from the sidebar")

# Footer
st.markdown("---")
st.markdown("Built with Streamlit • Upload Excel → Generate PDF Invoice")