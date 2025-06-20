
import streamlit as st
from docx import Document
import datetime
from io import BytesIO

st.title("Solar Proposal Generator")

client_name = st.text_input("Client Name")
site_location = st.text_input("Site Location")
proposal_date = st.date_input("Proposal Date", value=datetime.date.today())
capacity_kwp = st.number_input("System Capacity (kWp)", min_value=0.0)
system_type = st.selectbox("System Type", ["On-grid", "Hybrid", "Off-grid"])
base_price = st.number_input("Base Price (₹)", min_value=0.0)
gst = st.number_input("GST Amount (₹)", min_value=0.0)
total_after_gst = base_price + gst
subsidy = st.number_input("Subsidy (₹)", min_value=0.0)
net_payable = total_after_gst - subsidy
monthly_generation = st.number_input("Monthly Generation (kWh)", min_value=0.0)
year_1_savings = st.number_input("First Year Savings (₹)", min_value=0.0)
lifetime_savings = st.number_input("25 Year Lifetime Savings (₹)", min_value=0.0)

template_upload = st.file_uploader("Upload Template DOCX", type=["docx"])

if st.button("Generate Proposal") and template_upload:
    doc = Document(template_upload)
    replacements = {
        "{{client_name}}": client_name,
        "{{site_location}}": site_location,
        "{{proposal_date}}": proposal_date.strftime("%d-%m-%Y"),
        "{{capacity_kwp}}": str(capacity_kwp),
        "{{system_type}}": system_type,
        "{{base_price}}": f"{base_price:,.2f}",
        "{{gst}}": f"{gst:,.2f}",
        "{{total_after_gst}}": f"{total_after_gst:,.2f}",
        "{{subsidy}}": f"{subsidy:,.2f}",
        "{{net_payable}}": f"{net_payable:,.2f}",
        "{{monthly_generation}}": str(monthly_generation),
        "{{year_1_savings}}": f"{year_1_savings:,.2f}",
        "{{lifetime_savings}}": f"{lifetime_savings:,.2f}",
    }
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    output = BytesIO()
    doc.save(output)
    st.download_button(
        label="Download Proposal",
        data=output.getvalue(),
        file_name=f"Proposal_{client_name.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
