import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.units import inch
import io
from datetime import datetime

def process_excel_data(df):
    """Process Excel data to group by Style and sum quantities and values"""
    # Filter out rows where Style is NaN or empty
    df_clean = df.dropna(subset=['Style'])
    
    # Group by Style and sum the required columns
    grouped = df_clean.groupby('Style').agg({
        'Description': 'first',  # Take first description for each style
        'Composition': 'first',  # Take first composition for each style
        'USD Fob$': 'first',    # Take first unit price for each style
        'Total Qty': 'sum',     # Sum all quantities for same style
        'Total Value': 'sum'    # Sum all values for same style
    }).reset_index()
    
    return grouped

def generate_proforma_invoice(df, pi_number="SAR/LG/0148", date_str="14-10-2024", cpo_number="CPO/47062/25"):
    """Generate PDF invoice matching the exact reference format"""
    
    # Process the data
    processed_df = process_excel_data(df)
    
    # Create PDF in memory
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    
    # Create custom styles
    styles = getSampleStyleSheet()
    
    # Header style
    header_style = ParagraphStyle(
        'CustomHeader',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_LEFT,
        spaceAfter=3
    )
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Normal'],
        fontSize=14,
        alignment=TA_CENTER,
        spaceAfter=12,
        fontName='Helvetica-Bold'
    )
    
    # Create header table with supplier and invoice info
    header_data = [
        ["Supplier Name No. & date of PI", f"{pi_number} Dt. {date_str}"],
        ["SAR APPARELS INDIA PVT.LTD.", ""],
        ["ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.", f"Landmark order Reference: {cpo_number}"],
        ["PHONE : 9874173373", "Buyer Name: LANDMARK GROUP"],
        ["FAX : N.A.", "Brand Name: Juniors"],
        ["Consignee:-", "Payment Term: T/T"],
        ["RNA Resources Group Ltd- Landmark (Babyshop),", ""],
        ["P O Box 25030, Dubai, UAE,", "Bank Details (Including Swift/IBAN)"],
        ["Tel: 00971 4 8095500, Fax: 00971 4 8095555/66", ":- SAR APPARELS INDIA PVT.LTD"],
        ["", ":- 2112819952"],
        ["", "BANK'S NAME :- KOTAK MAHINDRA BANK LTD"],
        ["", "BANK ADDRESS :- 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR,"],
        ["", "KOLKATA-700001"],
        ["", ":- KKBKINBBCPC"],
        ["", "BANK CODE :- 0323"],
        ["Loading Country: India", "L/C Advicing Bank (If Payment term LC Applicable )"],
        ["Port of loading: Mumbai", ""],
        ["Agreed Shipment Date: 07-02-2025", ""],
        ["", ""],
        ["REMARKS if ANY:-", ""],
        ["Description of goods: Value Packs", ""],
        ["CURRENCY: USD", ""]
    ]
    
    header_table = Table(header_data, colWidths=[4*inch, 4*inch])
    header_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    
    elements.append(header_table)
    elements.append(Spacer(1, 12))
    
    # Add title
    elements.append(Paragraph("Proforma Invoice", title_style))
    elements.append(Spacer(1, 6))
    
    # Create main data table
    table_data = [
        ["STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE\nKNITTED /\nWOVEN", "H.S NO\n(8digit)", 
         "COMPOSITION OF\nMATERIAL", "COUNTRY OF\nORIGIN", "QTY", "UNIT PRICE\nFOB", "AMOUNT"]
    ]
    
    total_qty = 0
    total_amount = 0
    
    for _, row in processed_df.iterrows():
        qty = int(row['Total Qty'])
        unit_price = float(row['USD Fob$'])
        amount = float(row['Total Value'])
        
        total_qty += qty
        total_amount += amount
        
        table_data.append([
            str(row['Style']),
            str(row['Description']),
            "KNITTED",  # Default to KNITTED as per reference
            "61112000",  # Default HS code as per reference
            str(row['Composition']),
            "India",
            f"{qty:,}",
            f"{unit_price:.2f}",
            f"{amount:.2f}"
        ])
    
    # Add totals row
    table_data.append([
        "", "", "", "", "", "Total", f"{total_qty:,}", "", f"{total_amount:.2f}USD"
    ])
    
    # Create and style the main table
    main_table = Table(table_data, repeatRows=1)
    main_table.setStyle(TableStyle([
        # Header row styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
        
        # Data rows styling
        ('FONTSIZE', (0, 1), (-1, -2), 9),
        ('ALIGN', (0, 1), (-1, -2), 'CENTER'),
        ('VALIGN', (0, 1), (-1, -2), 'MIDDLE'),
        
        # Total row styling
        ('FONTSIZE', (0, -1), (-1, -1), 10),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('ALIGN', (0, -1), (-1, -1), 'CENTER'),
        
        # Grid lines
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
    ]))
    
    elements.append(main_table)
    elements.append(Spacer(1, 12))
    
    # Add total in words
    def number_to_words(amount):
        # Simple implementation for demo - you might want to use a library like num2words
        return f"US DOLLAR {amount:,.2f}"
    
    total_words = f"TOTAL {number_to_words(total_amount).upper()} DOLLARS"
    elements.append(Paragraph(total_words, header_style))
    elements.append(Spacer(1, 24))
    
    # Footer signature section
    signature_data = [
        ["Signed by …………………….(Affix Stamp here)", "for RNA Resources Group Ltd-Landmark (Babyshop)"],
        ["", ""],
        ["Terms & Conditions (If Any)", ""]
    ]
    
    signature_table = Table(signature_data, colWidths=[4*inch, 4*inch])
    signature_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    
    elements.append(signature_table)
    
    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

# Streamlit App
def main():
    st.set_page_config(page_title="Proforma Invoice Generator", page_icon="📋", layout="wide")
    
    st.title("📋 Proforma Invoice Generator")
    st.markdown("Upload your Excel file to generate a professional Proforma Invoice PDF")
    
    # Sidebar for additional options
    with st.sidebar:
        st.header("Invoice Details")
        pi_number = st.text_input("PI Number", value="SAR/LG/0148")
        invoice_date = st.date_input("Invoice Date", value=datetime.now())
        cpo_number = st.text_input("CPO Number", value="CPO/47062/25")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Upload Excel File")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload your Excel file containing Style, Description, Composition, USD Fob$, Total Qty, and Total Value columns"
        )
        
        if uploaded_file is not None:
            try:
                # Read the Excel file
                df = pd.read_excel(uploaded_file)
                
                # Display file info
                st.success(f"✅ File uploaded successfully!")
                st.info(f"📊 Found {len(df)} rows in the Excel file")
                
                # Show preview of data
                st.subheader("📋 Data Preview")
                
                # Check required columns
                required_columns = ['Style', 'Description', 'Composition', 'USD Fob$', 'Total Qty', 'Total Value']
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    st.error(f"❌ Missing required columns: {', '.join(missing_columns)}")
                    st.info("📝 Required columns: Style, Description, Composition, USD Fob$, Total Qty, Total Value")
                else:
                    # Show processed data preview
                    processed_df = process_excel_data(df)
                    st.dataframe(processed_df, use_container_width=True)
                    
                    st.subheader("📈 Summary")
                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.metric("Unique Styles", len(processed_df))
                    with col_b:
                        st.metric("Total Quantity", f"{processed_df['Total Qty'].sum():,}")
                    with col_c:
                        st.metric("Total Value", f"${processed_df['Total Value'].sum():,.2f}")
            
            except Exception as e:
                st.error(f"❌ Error reading Excel file: {str(e)}")
    
    with col2:
        st.header("Generate Invoice")
        
        if uploaded_file is not None and not missing_columns:
            if st.button("🚀 Generate PDF Invoice", type="primary", use_container_width=True):
                try:
                    with st.spinner("Generating PDF..."):
                        # Generate PDF
                        date_str = invoice_date.strftime("%d-%m-%Y")
                        pdf_buffer = generate_proforma_invoice(df, pi_number, date_str, cpo_number)
                        
                        # Create download button
                        filename = f"PI_{pi_number.replace('/', '_')}_{date_str}.pdf"
                        
                        st.success("✅ PDF Generated Successfully!")
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_buffer.getvalue(),
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                        
                except Exception as e:
                    st.error(f"❌ Error generating PDF: {str(e)}")
        else:
            st.info("📤 Upload a valid Excel file to generate PDF")

if __name__ == "__main__":
    main()