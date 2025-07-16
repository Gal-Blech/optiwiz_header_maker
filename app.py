# Optiwiz Header Maker - Web Application
#
# This script creates a simple web UI to translate Excel files into YAML.
#
# Required Libraries:
# - streamlit: To create the web application UI.
# - openpyxl: To read .xlsx files and access cell formatting.
# - ruamel.yaml: To generate clean, well-formatted YAML output.
#
# How to Install:
# pip install streamlit openpyxl ruamel.yaml
#
# How to Run:
# 1. Save this code as a Python file (e.g., app.py).
# 2. Open your terminal or command prompt.
# 3. Navigate to the directory where you saved the file.
# 4. Run the command: streamlit run app.py

import streamlit as st
from openpyxl import load_workbook
from ruamel.yaml import YAML
import io

# --- Helper Functions (from original script) ---

def get_merged_range(sheet, cell):
    """Checks if a cell is part of a merged range and returns the range string."""
    for merged_cell_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_cell_range:
            return str(merged_cell_range)
    return None

def format_color_hex(argb_hex):
    """Converts openpyxl's ARGB hex to a standard #RRGGBB hex."""
    if isinstance(argb_hex, str) and len(argb_hex) == 8:
        return f"#{argb_hex[2:]}"
    return None

# --- Core Translation Logic (adapted for Streamlit) ---

def generate_yaml_from_file(file_object):
    """
    Reads an in-memory Excel file object, translates it, and returns the YAML as a string.
    Also returns any warnings generated during the process.
    """
    warnings = []
    
    # Load the workbook from the in-memory file object
    workbook = load_workbook(file_object)
    sheet = workbook.active

    # Prepare for YAML Output
    yaml = YAML()
    yaml.indent(mapping=4, sequence=4, offset=2)
    yaml.preserve_quotes = True
    
    data = {'template': {'format': {'page_header': []}}}
    page_header = data['template']['format']['page_header']
    processed_merges = set()

    # Iterate Through Excel Rows and Cells
    for row in sheet.iter_rows():
        row_data = []
        is_empty_row = all(cell.value is None for cell in row)

        if is_empty_row:
            page_header.append([])
            continue

        for cell in row:
            cell_obj = {}
            
            # Merged Cells
            merged_range = get_merged_range(sheet, cell)
            if merged_range:
                if merged_range in processed_merges:
                    row_data.append(None)
                    continue
                processed_merges.add(merged_range)
                cell_obj['merge'] = {'from_to': merged_range}

            # Cell Value and Special Keywords
            value = cell.value
            if value is None and not merged_range:
                row_data.append(None)
                continue
            
            if value is not None:
                if str(value).strip() == '<Logo>':
                    cell_obj['type'] = 'logo'
                    cell_obj['value'] = True
                    # Logo Warning Logic
                    cell_below = sheet.cell(row=cell.row + 1, column=cell.column)
                    if cell_below.value:
                        warnings.append(f"**Logo Warning:** Data '{cell_below.value}' in cell {cell_below.coordinate} "
                                        f"may be obscured by the Logo in {cell.coordinate}.")
                elif str(value).strip() == '<placeholder>':
                    cell_obj['type'] = 'expert'
                    # Use a block scalar style for the placeholder value
                    cell_obj['value'] = 'return "<placeholder>"'
                else:
                    cell_obj['value'] = value
            
            # Formatting (only add if not default)
            if cell.font.bold: cell_obj['bold'] = True
            if cell.font.name and cell.font.name.lower() != 'calibri': cell_obj['font_name'] = cell.font.name.lower()
            if cell.font.size and cell.font.size != 11: cell_obj['font_size'] = int(cell.font.size)
            
            # **FIXED** Check for color type before processing
            if cell.font.color and cell.font.color.type == 'rgb':
                 font_color = format_color_hex(cell.font.color.rgb)
                 if font_color and font_color.upper() != '#000000':
                    cell_obj['font_color'] = font_color

            # **FIXED** Check for color type before processing
            if cell.fill.fill_type == 'solid' and cell.fill.start_color.type == 'rgb':
                bg_color = format_color_hex(cell.fill.start_color.rgb)
                if bg_color and bg_color.upper() != '#FFFFFF':
                    cell_obj['bg_color'] = bg_color

            if cell.alignment.horizontal and cell.alignment.horizontal != 'left':
                cell_obj['align'] = cell.alignment.horizontal
            if cell.alignment.vertical and cell.alignment.vertical != 'bottom':
                cell_obj['valign'] = 'vcenter' if cell.alignment.vertical == 'center' else cell.alignment.vertical
            
            if cell.border.left.style or cell.border.right.style or cell.border.top.style or cell.border.bottom.style:
                cell_obj['border'] = 1
                # **FIXED** Check for color type before processing
                if cell.border.left.color and cell.border.left.color.type == 'rgb':
                    border_color = format_color_hex(cell.border.left.color.rgb)
                    if border_color and border_color.upper() != '#000000':
                        cell_obj['border_color'] = border_color

            # Only append if the object is not empty, otherwise it's a simple null cell
            row_data.append(cell_obj if cell_obj else None)

        page_header.append(row_data)

    page_header.append([]) # Final required empty row

    # Convert the YAML data to a string
    string_stream = io.StringIO()
    yaml.dump(data, string_stream)
    yaml_string = string_stream.getvalue()
    
    return yaml_string, warnings

# --- Streamlit User Interface ---

st.set_page_config(page_title="Optiwiz Header Maker", layout="wide")

st.title("📄 Optiwiz Header Maker")
st.write("Upload your Excel design file to instantly translate it into Optiwiz YAML code.")

# Initialize session state
if 'yaml_output' not in st.session_state:
    st.session_state['yaml_output'] = ""
if 'file_name' not in st.session_state:
    st.session_state['file_name'] = ""

uploaded_file = st.file_uploader(
    "Choose an Excel file (.xlsx)",
    type="xlsx",
    accept_multiple_files=False
)

if uploaded_file is not None:
    st.success(f"File '{uploaded_file.name}' uploaded successfully!")

    if st.button("Translate to YAML", type="primary"):
        with st.spinner("Translating..."):
            # Process the file
            yaml_output, warnings = generate_yaml_from_file(uploaded_file)
            
            # Display any warnings
            for warning in warnings:
                st.warning(warning)
            
            # Store results in session state
            st.session_state['yaml_output'] = yaml_output
            st.session_state['file_name'] = uploaded_file.name.rsplit('.', 1)[0] + ".yaml"

# Display the output and download button only if there is content
if st.session_state['yaml_output']:
    st.subheader("Generated YAML Code")
    st.code(st.session_state['yaml_output'], language='yaml')
    
    st.download_button(
        label="⬇️ Download YAML File",
        data=st.session_state['yaml_output'],
        file_name=st.session_state['file_name'],
        mime='text/yaml'
    )
