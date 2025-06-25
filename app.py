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

# Comprehensive namespace map for DOCX XML
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
}

def replace_in_text_nodes(tree, replacements):
    """Replace placeholders in regular text nodes"""
    text_nodes = tree.xpath(".//w:t", namespaces=NAMESPACES)
    for node in text_nodes:
        if node.text:
            original_text = node.text
            for placeholder, replacement in replacements.items():
                if placeholder in node.text:
                    node.text = node.text.replace(placeholder, replacement)
                    #print(f"Replaced '{placeholder}' in regular text: {original_text} -> {node.text}")

def replace_in_textboxes(tree, replacements):
    """Replace placeholders in textboxes and shapes"""
    # Search for textboxes in various XML structures
    textbox_paths = [
        ".//wps:txbx//w:t",  # Word 2010+ textboxes
        ".//v:textbox//w:t",  # Legacy VML textboxes
        ".//w:txbxContent//w:t",  # Direct textbox content
        ".//a:t",  # DrawingML text
    ]
    
    for path in textbox_paths:
        text_nodes = tree.xpath(path, namespaces=NAMESPACES)
        for node in text_nodes:
            if node.text:
                original_text = node.text
                for placeholder, replacement in replacements.items():
                    if placeholder in node.text:
                        node.text = node.text.replace(placeholder, replacement)
                        #print(f"Replaced '{placeholder}' in textbox: {original_text} -> {node.text}")

def replace_in_drawing_objects(tree, replacements):
    """Handle text in drawing objects and shapes"""
    # Search for text in various drawing structures
    drawing_paths = [
        ".//w:drawing//a:t",  # DrawingML text
        ".//w:pict//v:shape//w:t",  # VML shapes with text
        ".//mc:AlternateContent//w:t",  # Alternate content blocks
    ]
    
    for path in drawing_paths:
        text_nodes = tree.xpath(path, namespaces=NAMESPACES)
        for node in text_nodes:
            if node.text:
                original_text = node.text
                for placeholder, replacement in replacements.items():
                    if placeholder in node.text:
                        node.text = node.text.replace(placeholder, replacement)
                        #print(f"Replaced '{placeholder}' in drawing object: {original_text} -> {node.text}")

def handle_split_placeholders(tree, replacements):
    """Handle placeholders that might be split across multiple text nodes"""
    # Get all text content and rebuild split placeholders
    all_text_nodes = tree.xpath(".//w:t | .//a:t", namespaces=NAMESPACES)
    
    # Create a map of nodes and their text content
    node_texts = []
    for node in all_text_nodes:
        if node.text:
            node_texts.append((node, node.text))
    
    # Look for split placeholders across consecutive nodes
    for placeholder, replacement in replacements.items():
        # Simple approach: look for opening braces and try to reconstruct
        for i, (node, text) in enumerate(node_texts):
            if "{{" in text and "}}" not in text:
                # Potential start of split placeholder
                combined_text = text
                nodes_to_update = [node]
                
                # Look ahead to find the complete placeholder
                for j in range(i + 1, min(i + 10, len(node_texts))):  # Look ahead max 10 nodes
                    next_node, next_text = node_texts[j]
                    combined_text += next_text
                    nodes_to_update.append(next_node)
                    
                    if placeholder in combined_text:
                        # Found complete placeholder, replace it
                        replaced_text = combined_text.replace(placeholder, replacement)
                        
                        # Clear all nodes except the first
                        for k, update_node in enumerate(nodes_to_update):
                            if k == 0:
                                update_node.text = replaced_text
                            else:
                                update_node.text = ""
                        
                        #print(f"Replaced split placeholder '{placeholder}': {combined_text} -> {replaced_text}")
                        break
                    
                    if "}}" in next_text:
                        break

def replace_placeholders_in_docx(input_path,  replacements):
    """Main function to replace placeholders in DOCX file"""
    #print(f"Processing DOCX file: {input_path}")
    
    with zipfile.ZipFile(input_path) as docx_zip:
        output_stream = BytesIO()
        
        with zipfile.ZipFile(output_stream, 'w', zipfile.ZIP_DEFLATED) as modified_zip:
            for item in docx_zip.infolist():
                data = docx_zip.read(item.filename)
                
                # Process XML files that might contain text
                if (item.filename.startswith("word/") and 
                    item.filename.endswith(".xml") and
                    item.filename != "word/fontTable.xml"):  # Skip font table
                    
                    try:
                        print(f"Processing: {item.filename}")
                        tree = etree.fromstring(data)
                        
                        # Apply all replacement strategies
                        replace_in_text_nodes(tree, replacements)
                        replace_in_textboxes(tree, replacements)
                        replace_in_drawing_objects(tree, replacements)
                        handle_split_placeholders(tree, replacements)
                        
                        # Convert back to bytes
                        data = etree.tostring(
                            tree, 
                            encoding="utf-8", 
                            xml_declaration=True, 
                            standalone=False
                        )
                        
                    except etree.XMLSyntaxError as e:
                        print(f"Skipping non-XML file: {item.filename}")
                    except Exception as e:
                        print(f"Error processing {item.filename}: {e}")
                
                modified_zip.writestr(item.filename, data)
        return output_stream


if st.button("Generate Proposal") and template_upload:
    REPLACEMENTS = {
        "{{client_name}}": client_name,
        "{{site_location}}": site_location,
        "{{proposal_date}}": proposal_date.strftime("%d-%m-%Y"),
        "{{aio_price}}": f"{aio_solar_kit_price:,.2f}",
        "{{total_price}}": f"{total_price:,.2f}",
        "{{disc_price}}": f"{discounted_price:,.2f}",
        "{{nt_eff_price}}": f"{net_effective_price:,.2f}",
    }

    # Replace text in all parts of the DOCX including textboxes
    #patched_docx_stream = replace_text_in_textboxes(template_upload, replacements)
    patched_docx_stream = replace_placeholders_in_docx(template_upload, REPLACEMENTS)
    st.download_button(
        label="Download Proposal",
        data=patched_docx_stream.getvalue(),
        file_name=f"Proposal_{client_name.replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
