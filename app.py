from flask import Flask, render_template, request, send_file
import pandas as pd
from jinja2 import Template
import pdfkit
import os
from num2words import num2words

app = Flask(__name__)

# Configure wkhtmltopdf path (OS-aware)
import os

if os.name == 'nt':  # Windows
    wkhtmltopdf_path = 'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'  # Adjust if needed
else:  # Linux or macOS
    wkhtmltopdf_path = '/usr/bin/wkhtmltopdf'  # Or the path from 'which wkhtmltopdf'

config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

# Improved HTML template for the receipt
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

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            try:
                df = pd.read_excel(file)
                df = df.head(10)

                receipts = []
                for _, row in df.iterrows():
                    receipts.append({
                        "payee": row["Payee Name"],
                        "amount": row["Amount"],
                        "amount_words": num2words(row["Amount"], lang='en').title(),
                        "work": row["Work"]
                    })

                rendered_html = Template(receipt_template).render(receipts=receipts)

                script_dir = os.path.dirname(__file__)
                pdf_file = os.path.join(script_dir, "receipts.pdf")

                pdfkit.from_string(rendered_html, pdf_file, options={"page-size": "A4"}, configuration=config)

                return send_file(pdf_file, as_attachment=True)

            except Exception as e:
                return f"An error occurred: {e}"

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)