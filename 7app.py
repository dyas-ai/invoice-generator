import streamlit as st
import pandas as pd
from datetime import datetime
import base64

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

def generate_html_invoice(df, pi_number="SAR/LG/0148", date_str="14-10-2024", cpo_number="CPO/47062/25"):
    """Generate HTML invoice matching the exact reference format"""
    
    # Process the data
    processed_df = process_excel_data(df)
    
    total_qty = processed_df['Total Qty'].sum()
    total_amount = processed_df['Total Value'].sum()
    
    # Generate table rows
    table_rows = ""
    for _, row in processed_df.iterrows():
        qty = int(row['Total Qty'])
        unit_price = float(row['USD Fob$'])
        amount = float(row['Total Value'])
        
        table_rows += f"""
        <tr>
            <td>{row['Style']}</td>
            <td>{row['Description']}</td>
            <td>KNITTED</td>
            <td>61112000</td>
            <td>{row['Composition']}</td>
            <td>India</td>
            <td>{qty:,}</td>
            <td>{unit_price:.2f}</td>
            <td>{amount:.2f}</td>
        </tr>
        """
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Proforma Invoice</title>
        <style>
            @page {{
                size: A4;
                margin: 0.5in;
            }}
            body {{
                font-family: Arial, sans-serif;
                font-size: 10px;
                margin: 0;
                padding: 20px;
                line-height: 1.2;
            }}
            .header-table {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 15px;
            }}
            .header-table td {{
                padding: 2px 5px;
                vertical-align: top;
                font-size: 9px;
            }}
            .title {{
                text-align: center;
                font-size: 14px;
                font-weight: bold;
                margin: 10px 0;
            }}
            .main-table {{
                width: 100%;
                border-collapse: collapse;
                border: 2px solid black;
                margin: 10px 0;
            }}
            .main-table th {{
                background-color: #d3d3d3;
                border: 1px solid black;
                padding: 8px 4px;
                text-align: center;
                font-weight: bold;
                font-size: 8px;
                vertical-align: middle;
            }}
            .main-table td {{
                border: 1px solid black;
                padding: 6px 4px;
                text-align: center;
                font-size: 9px;
                vertical-align: middle;
            }}
            .total-row {{
                font-weight: bold;
                font-size: 10px;
            }}
            .total-words {{
                font-weight: bold;
                margin: 10px 0;
                font-size: 10px;
            }}
            .signature-section {{
                margin-top: 30px;
                font-size: 10px;
            }}
            .signature-table {{
                width: 100%;
                border-collapse: collapse;
            }}
            .signature-table td {{
                padding: 5px;
                vertical-align: top;
            }}
            @media print {{
                body {{ margin: 0; }}
                .no-print {{ display: none; }}
            }}
        </style>
    </head>
    <body>
        <!-- Header Section -->
        <table class="header-table">
            <tr>
                <td style="width: 50%;"><strong>Supplier Name No. & date of PI</strong></td>
                <td style="width: 50%;"><strong>{pi_number} Dt. {date_str}</strong></td>
            </tr>
            <tr>
                <td><strong>SAR APPARELS INDIA PVT.LTD.</strong></td>
                <td></td>
            </tr>
            <tr>
                <td><strong>ADDRESS :</strong> 6, Picaso Bithi, KOLKATA - 700017.</td>
                <td>Landmark order Reference: <strong>{cpo_number}</strong></td>
            </tr>
            <tr>
                <td><strong>PHONE :</strong> 9874173373</td>
                <td>Buyer Name: <strong>LANDMARK GROUP</strong></td>
            </tr>
            <tr>
                <td><strong>FAX :</strong> N.A.</td>
                <td>Brand Name: <strong>Juniors</strong></td>
            </tr>
            <tr>
                <td><strong>Consignee:-</strong></td>
                <td>Payment Term: <strong>T/T</strong></td>
            </tr>
            <tr>
                <td>RNA Resources Group Ltd- Landmark (Babyshop),</td>
                <td></td>
            </tr>
            <tr>
                <td>P O Box 25030, Dubai, UAE,</td>
                <td>Bank Details (Including Swift/IBAN)</td>
            </tr>
            <tr>
                <td>Tel: 00971 4 8095500, Fax: 00971 4 8095555/66</td>
                <td><strong>:- SAR APPARELS INDIA PVT.LTD</strong></td>
            </tr>
            <tr>
                <td></td>
                <td><strong>:- 2112819952</strong></td>
            </tr>
            <tr>
                <td></td>
                <td><strong>BANK'S NAME :- KOTAK MAHINDRA BANK LTD</strong></td>
            </tr>
            <tr>
                <td></td>
                <td><strong>BANK ADDRESS :- 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR,</strong></td>
            </tr>
            <tr>
                <td></td>
                <td><strong>KOLKATA-700001</strong></td>
            </tr>
            <tr>
                <td></td>
                <td><strong>:- KKBKINBBCPC</strong></td>
            </tr>
            <tr>
                <td></td>
                <td><strong>BANK CODE :- 0323</strong></td>
            </tr>
            <tr>
                <td>Loading Country: India</td>
                <td>L/C Advicing Bank (If Payment term LC Applicable )</td>
            </tr>
            <tr>
                <td>Port of loading: Mumbai</td>
                <td></td>
            </tr>
            <tr>
                <td>Agreed Shipment Date: 07-02-2025</td>
                <td></td>
            </tr>
            <tr>
                <td></td>
                <td></td>
            </tr>
            <tr>
                <td><strong>REMARKS if ANY:-</strong></td>
                <td></td>
            </tr>
            <tr>
                <td>Description of goods: Value Packs</td>
                <td></td>
            </tr>
            <tr>
                <td><strong>CURRENCY: USD</strong></td>
                <td></td>
            </tr>
        </table>

        <!-- Title -->
        <div class="title">Proforma Invoice</div>

        <!-- Main Data Table -->
        <table class="main-table">
            <thead>
                <tr>
                    <th>STYLE NO.</th>
                    <th>ITEM DESCRIPTION</th>
                    <th>FABRIC TYPE<br>KNITTED /<br>WOVEN</th>
                    <th>H.S NO<br>(8digit)</th>
                    <th>COMPOSITION OF<br>MATERIAL</th>
                    <th>COUNTRY OF<br>ORIGIN</th>
                    <th>QTY</th>
                    <th>UNIT PRICE<br>FOB</th>
                    <th>AMOUNT</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
                <tr class="total-row">
                    <td colspan="6" style="text-align: right; font-weight: bold;">Total</td>
                    <td><strong>{total_qty:,}</strong></td>
                    <td></td>
                    <td><strong>{total_amount:.2f}USD</strong></td>
                </tr>
            </tbody>
        </table>

        <!-- Total in Words -->
        <div class="total-words">
            TOTAL US DOLLAR {total_amount:,.2f} DOLLARS
        </div>

        <!-- Signature Section -->
        <div class="signature-section">
            <table class="signature-table">
                <tr>
                    <td style="width: 50%;">Signed by …………………….(Affix Stamp here)</td>
                    <td style="width: 50%; text-align: right;">for RNA Resources Group Ltd-Landmark (Babyshop)</td>
                </tr>
                <tr>
                    <td colspan="2" style="height: 30px;"></td>
                </tr>
                <tr>
                    <td colspan="2">Terms & Conditions (If Any)</td>
                </tr>
            </table>
        </div>
    </body>
    </html>
    """
    
    return html_content

# Streamlit App
def main():
    st.set_page_config(page_title="Proforma Invoice Generator", page_icon="📋", layout="wide")
    
    st.title("📋 Proforma Invoice Generator")
    st.markdown("Upload your Excel file to generate a professional Proforma Invoice")
    
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
                    st.write("Available columns:", df.columns.tolist())
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
        
        if uploaded_file is not None and 'missing_columns' in locals() and not missing_columns:
            if st.button("🚀 Generate Invoice", type="primary", use_container_width=True):
                try:
                    with st.spinner("Generating Invoice..."):
                        # Generate HTML invoice
                        date_str = invoice_date.strftime("%d-%m-%Y")
                        html_content = generate_html_invoice(df, pi_number, date_str, cpo_number)
                        
                        st.success("✅ Invoice Generated Successfully!")
                        
                        # Display the invoice
                        st.subheader("📋 Proforma Invoice")
                        st.components.v1.html(html_content, height=800, scrolling=True)
                        
                        # Download button
                        filename = f"PI_{pi_number.replace('/', '_')}_{date_str}.html"
                        st.download_button(
                            label="📥 Download Invoice (HTML)",
                            data=html_content,
                            file_name=filename,
                            mime="text/html",
                            use_container_width=True,
                            help="Download as HTML file. You can open it in any browser and print/save as PDF using Ctrl+P"
                        )
                        
                        st.info("💡 **To save as PDF:** Download the HTML file, open it in any browser, then press Ctrl+P (Cmd+P on Mac) and select 'Save as PDF'")
                        
                except Exception as e:
                    st.error(f"❌ Error generating invoice: {str(e)}")
        else:
            st.info("📤 Upload a valid Excel file to generate invoice")

if __name__ == "__main__":
    main()
