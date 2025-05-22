import streamlit as st
import pandas as pd
from jinja2 import Template
import pdfkit
import os
from num2words import num2words
import io # To handle binary data for download button

st.set_page_config(layout="wide", page_title="Hand Receipt Generator")

# --- Configure wkhtmltopdf path ---
# This path is common for Linux environments where wkhtmltopdf is installed
# (e.g., via apt-get, which Streamlit Community Cloud uses with packages.txt)
WKHTMLTOPDF_PATH = '/usr/bin/wkhtmltopdf'

# Fallback for local Windows development if wkhtmltopdf is in a custom path
if os.name == 'nt' and not os.path.exists(WKHTMLTOPDF_PATH):
    # Adjust this path if your wkhtmltopdf.exe is elsewhere on Windows
    WKHTMLTOPDF_PATH = 'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'

# Try to set the config, handle if wkhtmltopdf is not found during initial config
# The actual error for pdf generation will be caught later.
try:
    config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
except OSError as e:
    st.error(f"Error configuring wkhtmltopdf. Make sure it's installed and accessible: {e}")
    st.info("For Streamlit Community Cloud, ensure 'wkhtmltopdf' is in your `packages.txt` file.")
    config = None # Set config to None if initialization fails


# Improved HTML template for the receipt (Same as your Flask version)
receipt_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=210mm, height: 297mm">
    <title>Hand Receipt (RPWA 28)</title>
    <style>
        body { font-family: sans-serif; margin: 0; }
        @page {
            margin: 10mm;  /* Page margins */
        }
        .container {
            width: 210mm !important; /* Added !important */
            min-height: 297mm;
            margin: 10mm 20mm !important; /* Added !important */
            border: 2px solid #ccc !important;
            padding: 0mm; /* Changed to 0mm */
            box-sizing: border-box;
            position: relative;
            page-break-before: always;
        }
        .container:first-child {
            page-break-before: auto;
        }
        .header { text-align: center; margin-bottom: 2px; }
        .details { margin-bottom: 1px; }
        .amount-words { font-style: italic; }
        .signature-area { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .signature-area td, .signature-area th {
            border: 1px solid #ccc !important;
            padding: 5px;
            text-align: left;
        }
        .offices { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .offices td, .offices th {
            border: 1px solid black !important;
            padding: 5px;
            text-align: left;
            word-wrap: break-word;
        }
        .input-field { border-bottom: 1px dotted #ccc; padding: 3px; width: calc(100% - 10px); display: inline-block; }
        .seal-container { position: absolute; left: 10mm; bottom: 10mm; width: 40mm; height: 25mm; z-index: 10; }
        .seal { max-width: 100%; max-height: 100%; text-align: center; line-height: 40mm; color: blue; display: flex; justify-content: space-around; align-items: center; }
        .bottom-left-box {
            position: absolute; bottom: 40mm; left: 40mm;
            border: 2px solid blue; padding: 10px;
            width: 450px; text-align: left; height: 55mm;
            color: blue;
        }
        .bottom-left-box p { margin: 3px 0; }
          @media print {
            .container {
                border: none;
                width: 210mm;
                min-height: 297mm;
                margin: 0;
                padding: 0;
            }
        }
    </style>
</head>
<body>
    {% for receipt in receipts %}
    <div class="container">
        <div class="header">
            <h2>Payable to: - {{ receipt.payee }} ( Electric Contractor)</h2>
            <h2>HAND RECEIPT (RPWA 28)</h2>
            <p>(Referred to in PWF&A Rules 418,424,436 & 438)</p>
            <p>Division - PWD Electric Division, Udaipur</p>
        </div>
        <div class="details">
            <p>(1)Cash Book Voucher No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Date &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
            <p>(2)Cheque No. and Date &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
            <p>(3) Pay for ECS Rs.{{ receipt.amount }}/- (Rupees <span class="amount-words">{{ receipt.amount_words }} Only</span>)</p>
            <p>(4) Paid by me</p>
            <p>(5) Received from The Executive Engineer PWD Electric Division, Udaipur the sum of Rs. {{ receipt.amount }}/- (Rupees <span class="amount-words">{{ receipt.amount_words }} Only</span>)</p>
            <p> Name of work for which payment is made: <span id="work-name" class="input-field">{{ receipt.work }}</span></p>
            <p> Chargeable to Head:- 8443 [EMD- Refund] </p>
            <table class="signature-area">
                <tr>
                    <td>Witness</td>
                    <td>Stamp</td>
                    <td>Signature of payee</td>
                </tr>
                <tr>
                    <td>Cash Book No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Page No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    <td></td>
                    <td></td>
                </tr>
            </table>
            <table class="offices">
                <tr>
                    <td>For use in the Divisional Office</td>
                    <td>For use in the Accountant General's office</td>
                </tr>
                <tr>
                    <td>Checked</td>
                    <td>Audited/Reviewed</td>
                </tr>
                <tr>
                    <td>Accounts Clerk</td>
                    <td>DA &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Auditor &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Supdt. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; G.O.</td>
                </tr>
            </table>
        </div>
        <div class="seal-container">
            <div class="seal">
                <p></p>
                <p></p>
                <p></p>
            </div>
        </div>
        <div class="bottom-left-box">
                <p></p>
                <p></p>
                <p></p>
            <p> Passed for Rs. {{ receipt.amount }}</p>
            <p> In Words Rupees: {{ receipt.amount_words }} Only</p>
            <p> Chargeable to Head:- 8443 [EMD- Refund]</p>
            <div class="seal">
                <p>Ar.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;D.A.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E.E.</p>
            </div>
        </div>
    </div>
    {% endfor %}
</body>
</html>
"""

st.title("Hand Receipt (RPWA 28) Generator")
st.write("Upload an Excel file with 'Payee Name', 'Amount', and 'Work' columns to generate PDF receipts.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    if config is None:
        st.warning("`wkhtmltopdf` configuration failed. PDF generation might not work.")
        st.stop() # Stop execution if wkhtmltopdf is not properly configured

    try:
        df = pd.read_excel(uploaded_file)
        # Limiting to first 10 rows as per your original Flask code
        df = df.head(10)

        if df.empty:
            st.warning("The uploaded Excel file is empty or has no data in the first 10 rows.")
            st.stop()

        required_columns = ["Payee Name", "Amount", "Work"]
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns in Excel file. Please ensure it has: {', '.join(required_columns)}")
            st.stop()

        receipts_data = []
        for _, row in df.iterrows():
            try:
                amount_val = float(row["Amount"]) # Ensure amount is numeric for num2words
                receipts_data.append({
                    "payee": str(row["Payee Name"]),
                    "amount": f"{amount_val:,.2f}", # Format amount for display
                    "amount_words": num2words(amount_val, lang='en').title(),
                    "work": str(row["Work"])
                })
            except ValueError:
                st.warning(f"Skipping row due to invalid 'Amount': {row['Amount']}")
                continue


        if not receipts_data:
            st.error("No valid receipt data could be processed from the Excel file.")
        else:
            rendered_html = Template(receipt_template).render(receipts=receipts_data)

            # Generate PDF
            try:
                # pdfkit.from_string returns bytes
                pdf_bytes = pdfkit.from_string(
                    rendered_html,
                    False, # Output to string (bytes) instead of file
                    options={"page-size": "A4", "enable-local-file-access": ""}, # Added local file access for potential CSS/image loading if needed
                    configuration=config
                )

                st.success("PDF generated successfully!")

                # Offer the PDF for download
                st.download_button(
                    label="Download Receipts PDF",
                    data=pdf_bytes,
                    file_name="generated_receipts.pdf",
                    mime="application/pdf"
                )

            except Exception as e:
                st.error(f"An error occurred during PDF generation. Please check your `wkhtmltopdf` setup and input data: {e}")
                st.info("Common issues: `wkhtmltopdf` not installed, incorrect path, or invalid HTML/data.")

    except pd.errors.EmptyDataError:
        st.error("The uploaded file is empty.")
    except Exception as e:
        st.error(f"An unexpected error occurred while processing the Excel file: {e}")
