
import streamlit as st
from docx import Document
import datetime
from io import BytesIO

st.title("Solar Proposal Generator")

client_name = st.text_input("Client Name")
site_location = st.text_input("Address")
proposal_date = st.date_input("Proposal Date", value=datetime.date.today())
aio_solar_kit_price = st.number_input("All-in-One Solar Installation Kit (3.3 kWp) ₹", min_value=0.0)
total_price = st.number_input("Total Price (₹)", min_value=0.0)
discounted_price= st.number_input("Discounted Price (₹)", min_value=0.0)
net_effective_price = st.number_input("Net Effective Price (₹)", min_value=0.0)
template_upload = st.file_uploader("Upload Template DOCX", type=["docx"])

if st.button("Generate Proposal") and template_upload:
    doc = Document(template_upload)
    replacements = {
        "{{client_name}}": client_name,
        "{{site_location}}": site_location,
        "{{proposal_date}}": proposal_date.strftime("%d-%m-%Y"),
        "{{aio_solar_kit_price}}": f"{aio_solar_kit_price:,.2f}",
        "{{total_price}}": f"{total_price:,.2f}",
        "{{discounted_price}}": f"{discounted_price:,.2f}",
        "{{net_effective_price}}": f"{net_effective_price:,.2f}",
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

