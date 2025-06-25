import streamlit as st
from docx import Document
import datetime
from io import BytesIO
import zipfile
from lxml import etree

st.title("Solar Proposal Generator")

# Basic input fields
client_name = st.text_input("Client Name")
site_location = st.text_input("Address")
proposal_date = st.date_input("Proposal Date", value=datetime.date.today())

# State and Wattage selection
state = st.selectbox("Select State", ["UP", "UK"])
wattage = st.selectbox("Select Wattage", ["3.3kWp", "4.4kWp", "5.5kWp"])

# Extract numeric wattage for calculations
wattage_numeric = float(wattage.replace("kWp", ""))

# Price inputs
aio_solar_kit_price = st.number_input("All-in-One Solar Installation Kit Price (₹)", min_value=0.0)
total_price = st.number_input("Total Price (₹)", min_value=0.0)
discounted_price = st.number_input("Discounted Price (₹)", min_value=0.0)

# Template upload
template_upload = st.file_uploader("Upload Template DOCX", type=["docx"])

# Subsidy and pricing lookup table
SUBSIDY_TABLE = {
    "UP": {
        3.3: {"mnre_subsidy": 78000, "state_subsidy": 30000, "net_eff_price": 123000},
        4.4: {"mnre_subsidy": 78000, "state_subsidy": 30000, "net_eff_price": 200000},
        5.5: {"mnre_subsidy": 78000, "state_subsidy": 30000, "net_eff_price": 277000}
    },
    "UK": {
        3.3: {"mnre_subsidy": 85800, "state_subsidy": 0, "net_eff_price": 145200},
        4.4: {"mnre_subsidy": 85800, "state_subsidy": 0, "net_eff_price": 222200},
        5.5: {"mnre_subsidy": 85800, "state_subsidy": 0, "net_eff_price": 299200}
    }
}

# Get subsidy values based on selection
selected_subsidies = SUBSIDY_TABLE[state][wattage_numeric]

# Display calculated values
st.subheader("Calculated Subsidies and Pricing")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("MNRE Subsidy", f"₹{selected_subsidies['mnre_subsidy']:,}")
with col2:
    st.metric("State Subsidy", f"₹{selected_subsidies['state_subsidy']:,}")
with col3:
    st.metric("Net Effective Price", f"₹{selected_subsidies['net_eff_price']:,}")

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
                        break
                    
                    if "}}" in next_text:
                        break

def replace_placeholders_in_docx(input_path, replacements):
    """Main function to replace placeholders in DOCX file"""
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
                        
                    except etree.XMLSyntaxError:
                        # Skip non-XML files
                        pass
                    except Exception as e:
                        st.error(f"Error processing {item.filename}: {e}")
                
                modified_zip.writestr(item.filename, data)
        return output_stream

# Generate proposal button
if st.button("Generate Proposal") and template_upload:
    # Create replacements dictionary with all placeholders
    REPLACEMENTS = {
        "{{client_name}}": client_name,
        "{{site_location}}": site_location,
        "{{proposal_date}}": proposal_date.strftime("%d-%m-%Y"),
        "{{aio_price}}": f"{aio_solar_kit_price:,.2f}",
        "{{total_price}}": f"{total_price:,.2f}",
        "{{disc_price}}": f"{discounted_price:,.2f}",
        "{{mnre_subsidy}}": f"{selected_subsidies['mnre_subsidy']:,}",
        "{{state_subsidy}}": f"{selected_subsidies['state_subsidy']:,}",
        "{{nt_eff_price}}": f"{selected_subsidies['net_eff_price']:,}",
    }

    try:
        # Process the document
        with st.spinner("Generating proposal..."):
            patched_docx_stream = replace_placeholders_in_docx(template_upload, REPLACEMENTS)
        
        # Provide download button
        st.success("Proposal generated successfully!")
        st.download_button(
            label="Download Proposal",
            data=patched_docx_stream.getvalue(),
            file_name=f"Proposal_{client_name.replace(' ', '_')}_{state}_{wattage}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        
        # Show summary of replacements
        st.subheader("Proposal Details Summary")
        summary_data = {
            "Field": ["Client Name", "Site Location", "Proposal Date", "State", "Wattage", 
                     "MNRE Subsidy", "State Subsidy", "Net Effective Price", 
                     "AIO Price", "Total Price", "Discounted Price"],
            "Value": [client_name, site_location, proposal_date.strftime("%d-%m-%Y"), 
                     state, wattage, f"₹{selected_subsidies['mnre_subsidy']:,}", 
                     f"₹{selected_subsidies['state_subsidy']:,}", 
                     f"₹{selected_subsidies['net_eff_price']:,}",
                     f"₹{aio_solar_kit_price:,.2f}", f"₹{total_price:,.2f}", 
                     f"₹{discounted_price:,.2f}"]
        }
        st.table(summary_data)
        
    except Exception as e:
        st.error(f"Error generating proposal: {str(e)}")
        st.error("Please check your template file and try again.")

# Information section
with st.expander("ℹ️ How to use"):
    st.markdown("""
    **Steps to generate a proposal:**
    1. Fill in the client details (name, address, date)
    2. Select the appropriate state (UP or UK)
    3. Choose the solar system wattage (3.3kWp, 4.4kWp, or 5.5kWp)
    4. Enter pricing information
    5. Upload your DOCX template with placeholders
    6. Click "Generate Proposal"
    
    **Supported placeholders in your template:**
    - `{{client_name}}` - Client's name
    - `{{site_location}}` - Installation address
    - `{{proposal_date}}` - Proposal date
    - `{{aio_price}}` - All-in-One kit price
    - `{{total_price}}` - Total project price
    - `{{disc_price}}` - Discounted price
    - `{{mnre_subsidy}}` - MNRE subsidy amount (auto-calculated)
    - `{{state_subsidy}}` - State subsidy amount (auto-calculated)
    - `{{nt_eff_price}}` - Net effective price (auto-calculated)
    
    **Note:** The MNRE subsidy, state subsidy, and net effective price are automatically calculated based on your state and wattage selection.
    """)

# Debug section (optional - can be removed in production)
if st.checkbox("Show Debug Info"):
    st.subheader("Debug Information")
    st.write("Selected State:", state)
    st.write("Selected Wattage:", wattage_numeric)
    st.write("Calculated Subsidies:", selected_subsidies)
    if template_upload:
        st.write("Template file uploaded:", template_upload.name)
