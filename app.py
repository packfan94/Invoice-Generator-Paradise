import streamlit as st
import base64
import os
import re
from datetime import datetime
from jinja2 import Template
import weasyprint
from docx import Document

st.set_page_config(page_title="Paradise Bar Events - Invoice Generator", layout="centered")

INVOICE_HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
  @page {
    size: letter;
    margin: 8mm 12mm 8mm 12mm;
    background-color: #fcfbf9;
  }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: #2c3036;
    margin: 0;
    padding: 0;
    font-size: 8.4pt;
    line-height: 1.35;
    background-color: #fcfbf9;
  }
  .header-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 8px;
    border-bottom: 2px solid #5a6572;
    padding-bottom: 6px;
  }
  .logo-img { height: 52px; width: auto; border-radius: 3px; }
  .invoice-title {
    font-size: 24pt;
    font-weight: 900;
    color: #3b444f;
    text-align: right;
    letter-spacing: 2px;
    margin: 0 0 4px 0;
    font-family: 'Copperplate', 'Georgia', serif, sans-serif;
  }
  .meta-table { float: right; border-collapse: collapse; font-size: 8.2pt; }
  .meta-table td { padding: 1.5px 4px; }
  .meta-label { font-weight: 600; color: #6d7784; text-align: right; }
  .meta-val { font-weight: 700; color: #2c3036; text-align: right; }
  .two-col-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 8px 0;
    margin-left: -8px;
    margin-right: -8px;
    margin-bottom: 8px;
  }
  .two-col-table > tbody > tr > td { width: 50%; vertical-align: top; }
  .info-card {
    background-color: #f4f1ea;
    border: 1px solid #ddd6c8;
    border-radius: 4px;
    padding: 7px 10px;
    min-height: 122px;
  }
  .info-list { width: 100%; border-collapse: collapse; font-size: 8pt; }
  .info-list td { padding: 1.8px 0; vertical-align: top; }
  .info-lbl { width: 28%; font-weight: 700; color: #3b444f; }
  .info-data { width: 72%; color: #2c3036; }
  .spec-card {
    background-color: #f7f4ed;
    border: 1px solid #ddd6c8;
    border-left: 3.5px solid #6c7886;
    border-radius: 4px;
    padding: 7px 10px;
    margin-bottom: 8px;
    font-size: 7.8pt;
    line-height: 1.32;
  }
  .spec-title {
    font-weight: 800;
    color: #3b444f;
    margin-bottom: 4px;
    text-transform: uppercase;
    font-size: 8.2pt;
    letter-spacing: 0.5px;
  }
  .inc-exc-card {
    background-color: #f4f1ea;
    border: 1px solid #ddd6c8;
    border-radius: 4px;
    padding: 6px 10px;
    margin-bottom: 8px;
    font-size: 7.8pt;
  }
  .items-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 8px;
  }
  .items-table th {
    background-color: #4b5563;
    color: #ffffff;
    font-size: 7.8pt;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    padding: 4px 6px;
    text-align: left;
  }
  .items-table td {
    padding: 4px 6px;
    border-bottom: 1px solid #e4decb;
    font-size: 7.8pt;
    vertical-align: middle;
  }
  .items-table tbody tr:nth-child(even) { background-color: #f7f4ed; }
  .text-right { text-align: right !important; }
  .text-center { text-align: center !important; }
  .fin-table { width: 100%; border-collapse: collapse; margin-top: 1px; }
  .fin-table > tbody > tr > td { vertical-align: top; }
  .deposit-callout {
    background: #ede8dc;
    border: 1.5px solid #8e9aab;
    border-radius: 4px;
    padding: 8px 10px;
  }
  .deposit-tag {
    display: inline-block;
    background-color: #5a6572;
    color: #ffffff;
    font-size: 7pt;
    font-weight: 700;
    padding: 1px 5px;
    border-radius: 2px;
    text-transform: uppercase;
    letter-spacing: 0.4px;
    margin-bottom: 3px;
  }
  .deposit-title { font-size: 8.8pt; font-weight: 800; color: #3b444f; margin-bottom: 2px; }
  .deposit-amount { font-size: 17pt; font-weight: 900; color: #2c3036; line-height: 1.1; margin-bottom: 4px; }
  .deposit-desc {
    font-size: 7.5pt;
    color: #4b5563;
    line-height: 1.3;
    border-top: 1px dashed #cfc7b6;
    padding-top: 4px;
  }
  .summary-table { width: 100%; border-collapse: collapse; font-size: 7.9pt; }
  .summary-table td { padding: 1.5px 3px; }
  .sum-label { text-align: right; color: #6d7784; font-weight: 600; }
  .sum-val { text-align: right; font-weight: 700; color: #2c3036; width: 80px; }
  .sum-total-row {
    border-top: 1px solid #cfc7b6;
    border-bottom: 1px solid #cfc7b6;
    background-color: #f4f1ea;
  }
  .sum-total-row td { font-weight: 800; font-size: 8.2pt; color: #2c3036; padding: 2.5px 3px; }
  .sum-due-row { background-color: #4b5563; }
  .sum-due-row td { font-weight: 800; font-size: 8.8pt; color: #ffffff; padding: 3.5px 5px; }
  .footer-notice {
    background-color: #f5eedc;
    border: 1px solid #c9b48f;
    border-radius: 4px;
    padding: 5px 8px;
    margin-top: 8px;
    text-align: center;
    font-size: 8.5pt;
    font-weight: 800;
    color: #634f2d;
  }
</style>
</head>
<body>
<table class="header-table">
  <tr>
    <td style="width: 50%;">
      {% if logo_b64 %}
      <img src="data:image/png;base64,{{ logo_b64 }}" class="logo-img" alt="Paradise Bar Events" />
      {% endif %}
    </td>
    <td style="width: 50%;">
      <div class="invoice-title">INVOICE</div>
      <table class="meta-table">
        <tr><td class="meta-label">Invoice Number:</td><td class="meta-val">{{ invoice_num }}</td></tr>
        <tr><td class="meta-label">Invoice Date:</td><td class="meta-val">{{ invoice_date }}</td></tr>
      </table>
    </td>
  </tr>
</table>

<table class="two-col-table">
  <tr>
    <td>
      <div class="info-card">
        <table class="info-list">
          <tr><td class="info-lbl">Client:</td><td class="info-data">{{ client_name }}</td></tr>
          <tr><td class="info-lbl">Contact:</td><td class="info-data">{{ contact_name }}</td></tr>
          <tr><td class="info-lbl">Phone:</td><td class="info-data">{{ phone }}</td></tr>
          <tr><td class="info-lbl">Email:</td><td class="info-data">{{ email }}</td></tr>
          <tr><td class="info-lbl">Location:</td><td class="info-data">{{ location }}</td></tr>
        </table>
      </div>
    </td>
    <td>
      <div class="info-card">
        <table class="info-list">
          <tr><td class="info-lbl">Event Name:</td><td class="info-data">{{ event_name }}</td></tr>
          <tr><td class="info-lbl">Date:</td><td class="info-data">{{ event_date }}</td></tr>
          <tr><td class="info-lbl">Start Time:</td><td class="info-data">{{ start_time }}</td></tr>
          <tr><td class="info-lbl">End Time:</td><td class="info-data">{{ end_time }}</td></tr>
          <tr><td class="info-lbl">Event Theme:</td><td class="info-data">{{ theme }}</td></tr>
          <tr><td class="info-lbl">Guest Count:</td><td class="info-data">{{ guest_count }}</td></tr>
        </table>
      </div>
    </td>
  </tr>
</table>

<div class="spec-card">
  <div class="spec-title">Specialty Bar</div>
  <div>{{ specialty_bar_html | safe }}</div>
</div>

<div class="inc-exc-card">
  <div style="font-weight: 700; color: #3b444f; text-transform: uppercase; font-size: 7.8pt; margin-bottom: 3px;">Event Items &amp; Costing</div>
  <table style="width: 100%; border-collapse: collapse; font-size: 7.8pt;">
    <tr>
      <td style="width: 50%; vertical-align: top; padding-right: 8px;">
        <strong>Included:</strong><br>{{ included_html | safe }}
      </td>
      <td style="width: 50%; vertical-align: top; padding-left: 8px; border-left: 1px solid #ddd6c8;">
        <strong>Excluded (Client Provides):</strong><br>{{ excluded_html | safe }}
      </td>
    </tr>
  </table>
</div>

<table class="items-table">
  <thead>
    <tr>
      <th style="width: 52%;">Cost Breakdown</th>
      <th style="width: 12%;" class="text-center">Qty.</th>
      <th style="width: 18%;" class="text-right">Price Ea.</th>
      <th style="width: 18%;" class="text-right">Extended</th>
    </tr>
  </thead>
  <tbody>
    {% for item in line_items %}
    <tr>
      <td>{{ item.desc }}</td>
      <td class="text-center">{{ item.qty }}</td>
      <td class="text-right">{{ item.price }}</td>
      <td class="text-right">{{ item.extended }}</td>
    </tr>
    {% endfor %}
  </tbody>
</table>

<table class="fin-table">
  <tr>
    <td style="width: 48%;">
      <div class="deposit-callout">
        <div class="deposit-tag">Amount Due (50%)</div>
        <div class="deposit-amount">${{ amount_due_50 }}</div>
        <div class="deposit-desc">
          <strong>Discounted Total:</strong> ${{ final_total }}<br>
          <strong>50% Due Now:</strong> ${{ amount_due_50 }}<br>
          <strong>Remaining Balance:</strong> ${{ amount_due_50 }}
        </div>
      </div>
    </td>
    <td style="width: 4%;"></td>
    <td style="width: 48%;">
      <table class="summary-table">
        <tr><td class="sum-label">Sub-total:</td><td class="sum-val">${{ subtotal }}</td></tr>
        <tr><td class="sum-label">Coordination Fee ({{ coord_rate }}):</td><td class="sum-val">${{ coord_fee }}</td></tr>
        <tr><td class="sum-label">Sales Tax ({{ tax_rate }}):</td><td class="sum-val">${{ tax_amount }}</td></tr>
        <tr><td class="sum-label">TOTAL:</td><td class="sum-val">${{ gross_total }}</td></tr>
        {% if discount and discount != '0.00' %}
        <tr><td class="sum-label" style="color: #4b6b4b;">Special approved discount:</td><td class="sum-val" style="color: #4b6b4b;">-${{ discount }}</td></tr>
        {% endif %}
        <tr class="sum-total-row"><td class="sum-label" style="color: #2c3036;">Discounted Total:</td><td class="sum-val">${{ final_total }}</td></tr>
        <tr class="sum-due-row"><td style="color: #ffffff; text-align: right;">AMOUNT DUE (50%):</td><td style="color: #ffffff; text-align: right;" class="sum-val">${{ amount_due_50 }}</td></tr>
      </table>
    </td>
  </tr>
</table>

<div class="footer-notice">
  Final balance will be due 7 days before the event.
</div>
</body>
</html>
"""

def parse_quote_docx(file_bytes):
    doc = Document(file_bytes)
    
    # Collect all text from paragraphs and table cells
    doc_lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            doc_lines.append(p.text.strip())
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                if c.text.strip() and c.text.strip() not in doc_lines:
                    doc_lines.append(c.text.strip())
                    
    full_text = "\n".join(doc_lines)
    
    def extract_field(pattern, default=""):
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            if match.groups():
                return match.group(1).strip()
            return match.group(0).strip()
        return default

    client_name = extract_field(r"Client:\s*\n*(.*?)(?=\n|Contact:)", "Valued Client")
    contact_name = extract_field(r"Contact:\s*\n*(.*?)(?=\n|Phone:)", client_name)
    phone = extract_field(r"Phone\s*\n*(.*?)(?=\n|Email)", "")
    email = extract_field(r"Email\s*\n*(.*?)(?=\n|Event Theme|Guest Count)", "")
    location = extract_field(r"Location\s*\n*(.*?)(?=\n|Specialty Bar|Open Bar)", "")
    location = location.replace("\n", "<br>")

    event_name = extract_field(r"Event Name\s*\n*(.*?)(?=\n|Date)", "Special Event")
    event_date = extract_field(r"Date\s*\n*(.*?)(?=\n|Start Time)", "")
    start_time = extract_field(r"Start Time\s*\n*(.*?)(?=\n|End Time)", "")
    end_time = extract_field(r"End Time\s*\n*(.*?)(?=\n|Phone|Event Theme)", "")
    start_time = start_time.replace("\n", "<br>")
    end_time = end_time.replace("\n", "<br>")
    
    theme = extract_field(r"Event Theme\s*\n*(.*?)(?=\n|Guest Count)", "")
    guest_count = extract_field(r"Guest Count\s*\n*(.*?)(?=\n|Location)", "")

    # Extract Specialty Bar section text safely
    spec_bar_raw = extract_field(r"((?:Specialty Bar|Open Bar for First Hour)[\s\S]*?)(?=Event Items & Costing)", "")
    spec_lines = [l.strip() for l in spec_bar_raw.split("\n") if l.strip()]
    specialty_bar_html = "<br>".join(spec_lines)

    # Extract Included and Excluded
    inc_raw = extract_field(r"Included\s*\n*([\s\S]*?)(?=Excluded)", "")
    exc_raw = extract_field(r"Excluded \(Client Provides\)\s*\n*([\s\S]*?)(?=Cost Breakdown)", "")
    
    included_html = "<br>".join([l.strip() for l in inc_raw.split("\n") if l.strip()])
    excluded_html = "<br>".join([l.strip() for l in exc_raw.split("\n") if l.strip()])

    line_items = []
    subtotal = "0.00"
    coord_rate = "19%"
    coord_fee = "0.00"
    tax_rate = "9.50%"
    tax_amount = "0.00"
    gross_total = "0.00"
    discount = "0.00"
    final_total = "0.00"

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            row_str = " ".join(cells)
            
            if "Sub-total" in row_str or "Subtotal" in row_str:
                for c in reversed(cells):
                    if "$" in c or re.search(r"\d+\.\d{2}", c):
                        subtotal = c.replace("$", "").strip()
                        break
            elif "Coordination Fee" in row_str:
                for c in cells:
                    if "%" in c:
                        coord_rate = c
                    elif "$" in c or re.search(r"\d+\.\d{2}", c):
                        coord_fee = c.replace("$", "").strip()
            elif "Tax" in row_str:
                for c in cells:
                    if "%" in c or "[" in c:
                        tax_rate = c
                    elif "$" in c or re.search(r"\d+\.\d{2}", c):
                        tax_amount = c.replace("$", "").strip()
            elif "TOTAL" in row_str and "DISCOUNT" not in row_str and "REVISED" not in row_str:
                for c in reversed(cells):
                    if "$" in c or re.search(r"\d+\.\d{2}", c):
                        gross_total = c.replace("$", "").strip()
                        break
            elif "discount" in row_str.lower():
                for c in reversed(cells):
                    if "$" in c or re.search(r"\d+\.\d{2}", c):
                        discount = c.replace("$", "").replace("-", "").strip()
                        break
            elif "Discounted Total" in row_str or "REVISED TOTAL" in row_str:
                for c in reversed(cells):
                    if "$" in c or re.search(r"\d+\.\d{2}", c):
                        final_total = c.replace("$", "").strip()
                        break
            elif len(cells) >= 3 and any(char.isdigit() for char in row_str):
                if not any(k in row_str for k in ["Sub-total", "Fee", "Tax", "TOTAL", "discount", "Included", "Excluded"]):
                    line_items.append({
                        "desc": cells[0].replace("\n", "<br>"),
                        "qty": cells[1] if len(cells) > 1 else "",
                        "price": cells[2] if len(cells) > 2 else "",
                        "extended": cells[3] if len(cells) > 3 else ""
                    })

    if final_total == "0.00" and gross_total != "0.00":
        final_total = gross_total

    return {
        "client_name": client_name,
        "contact_name": contact_name,
        "phone": phone,
        "email": email,
        "location": location,
        "event_name": event_name,
        "event_date": event_date,
        "start_time": start_time,
        "end_time": end_time,
        "theme": theme,
        "guest_count": guest_count,
        "specialty_bar_html": specialty_bar_html,
        "included_html": included_html,
        "excluded_html": excluded_html,
        "line_items": line_items,
        "subtotal": subtotal,
        "coord_rate": coord_rate,
        "coord_fee": coord_fee,
        "tax_rate": tax_rate,
        "tax_amount": tax_amount,
        "gross_total": gross_total,
        "discount": discount,
        "final_total": final_total
    }

st.title("Paradise Bar Events — Invoice Generator")

uploaded_file = st.file_uploader("Upload Word Quote (.docx)", type=["docx"])

if uploaded_file:
    with st.spinner("Processing document..."):
        data = parse_quote_docx(uploaded_file)
        
        final_clean = str(data.get("final_total", "0")).replace(",", "").replace("$", "").strip()
        final_float = float(final_clean) if final_clean else 0.0
        amount_due_50 = f"{final_float * 0.5:,.2f}"

        logo_b64 = ""
        if os.path.exists("brand_logo.png"):
            with open("brand_logo.png", "rb") as f:
                logo_b64 = base64.b64encode(f.read()).decode("utf-8")

        template = Template(INVOICE_HTML_TEMPLATE)
        rendered_html = template.render(
            logo_b64=logo_b64,
            invoice_num=f"INV-{datetime.now().strftime('%Y%m%d')}",
            invoice_date=datetime.now().strftime('%B %d, %Y'),
            amount_due_50=amount_due_50,
            **data
        )

        pdf_bytes = weasyprint.HTML(string=rendered_html).write_pdf()

    st.success("Invoice generated successfully!")
    st.download_button(
        label="Download 1-Page PDF Invoice",
        data=pdf_bytes,
        file_name=f"Invoice_{data.get('client_name', 'Event').replace(' ', '_')}.pdf",
        mime="application/pdf"
    )
