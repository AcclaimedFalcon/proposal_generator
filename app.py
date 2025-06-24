import streamlit as st
from docx import Document
import datetime
from io import BytesIO
import zipfile
from lxml import etree

st.title("Solar Proposal Generator")

client_name = st.text_input("Client Name")
site_location = st.text_input("Address")
proposal_date = st.date_input("Proposal Date", value=datetime.date.today())
aio_solar_kit_price = st.number_input("All-in-One Solar Installation Kit (3.3 kWp) ₹", min_value=0.0)
total_price = st.number_input("Total Price (₹)", min_value=0.0)
discounted_price = st.number_input("Discounted Price (₹)", min_value=0.0)
net_effective_price = st.number_input("Net Effective Price (₹)", min_value=0.0)
template_upload = st.file_uploader("Upload Template DOCX", type=["docx"])

def replace_text_in_textboxes(docx_file, replacements):
    with zipfile.ZipFile(docx_file) as docx_zip:
        modified_docx = BytesIO()
        with zipfile.ZipFile(modified_docx, 'w') as modified_zip:
            for item in docx_zip.infolist():
                xml = docx_zip.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        tree = etree.fromstring(xml)
                        for node in tree.xpath("//w:t", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
                            if node.text:
                                for key, val in replacements.items():
                                    if key in node.text:
                                        node.text = node.text.replace(key, val)
                        xml = etree.tostring(tree, encoding='utf-8')
                    except Exception:
                        pass  # Skip non-XML or malformed files
                modified_zip.writestr(item.filename, xml)
        modified_docx.seek(0)
        return modified_docx

if st.button("Generate Proposal") and template_upload:
    replacements = {
        "{{client_name}}": client_name,
        "{{site_location}}": site_location,
        "{{proposal_date}}": proposal_date.strftime("%d-%m-%Y"),
        "{{aio_solar_kit_price}}": f"{aio_solar_kit_price:,.2f}",
        "{{total_price}}": f"{total_price:,.2f}",
        "{{discounted_price}}": f"{discounted_price:,.2f}",
        "{{net_effective_price}}": f"{net_effective_price:,.2f}",
    }

    # Replace text in all parts of the DOCX including textboxes
    patched_docx_stream = replace_text_in_textboxes(template_upload, replacements)

    st.download_button(
        label="Download Proposal",
        data=patched_docx_stream.getvalue(),
        file_name=f"Proposal_{client_name.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
