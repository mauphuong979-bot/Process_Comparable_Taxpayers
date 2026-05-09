import streamlit as st
import io
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, Mm
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook

DEFAULT_INCLUDED_HEADERS = [
    "Company Name", "Country", "Registration Number", "Primary US SIC",
    "Total Net Sales are", "Operating Income Loss For 'X' Or More Consecutive Years",
    "Rejection Reason", "Accepted Companies", "Rejected Companies"
]

DEFAULT_EXCLUDED_HEADERS = [
    "Data Source", "Publication Date", "City", "Registration Name",
    "Latest Consolidated Tax Year Available", "Latest Unconsolidated Tax Year Available",
    "Financial Type"
]

def calculate_column_widths(data_matrix, start_row_index, usable_width, header_row_idx):
    """
    Calculate optimal column widths based on content length of first 10 rows.
    usable_width is in inches.
    header_row_idx is used to identify specific columns that need more space.
    """
    if not data_matrix:
        return []
    
    num_cols = len(data_matrix[0])
    if num_cols == 0:
        return []
        
    # Content sample: 10 rows starting from start_row_index
    sample_start = min(start_row_index, len(data_matrix))
    sample_end = min(sample_start + 10, len(data_matrix))
    content_sample = data_matrix[sample_start:sample_end]
    
    widths = []
    for col_idx in range(num_cols):
        # Measure max length in the sample
        max_content_len = 0
        if content_sample:
            for row in content_sample:
                val_str = row[col_idx]
                max_content_len = max(max_content_len, len(val_str))
        else:
            max_content_len = 5 # Default fallback
            
        # Heuristic: base width on content length
        # Using a slightly higher multiplier since we ignore headers
        multiplier = 0.8
        min_val = 5
        
        # Specific columns that should not wrap
        if header_row_idx is not None and header_row_idx < len(data_matrix):
            header_text = data_matrix[header_row_idx][col_idx].lower()
            if "country" in header_text:
                multiplier = 1.0
                min_val = 9
            elif "registration" in header_text:
                multiplier = 1.1
                min_val = 11
            elif any(k in header_text for k in ["operating income loss", "accepted companies", "rejected companies"]):
                multiplier = 1.3
                min_val = 14
                
        widths.append(max(min_val, min(max_content_len * multiplier, 50)))
    
    total_raw_width = sum(widths)
    if total_raw_width == 0:
        return [usable_width / num_cols] * num_cols
    
    # Convert to inches and scale
    scaled_widths = [(w / total_raw_width) * usable_width for w in widths]
    
    # Ensure min width (inches)
    for i in range(len(scaled_widths)):
        target_min = 0.5
        if header_row_idx is not None and header_row_idx < len(data_matrix):
            h_text = data_matrix[header_row_idx][i].lower()
            if "country" in h_text:
                target_min = 1
            elif "registration" in h_text:
                target_min = 1
            elif any(k in h_text for k in ["operating income loss", "accepted companies", "rejected companies"]):
                target_min = 0.5
        
        if scaled_widths[i] < target_min:
            scaled_widths[i] = target_min
            
    # Re-scale to fit usable width if we went over due to target_min
    current_total = sum(scaled_widths)
    if current_total > usable_width:
        scaled_widths = [(w / current_total) * usable_width for w in scaled_widths]
        
    return scaled_widths


def find_company_name_row(ws):
    """
    Automatically find the row index (1-based) containing "Company Name".
    Searches first 50 rows.
    """
    for row in ws.iter_rows(min_row=1, max_row=50):
        for cell in row:
            if cell.value:
                val_str = str(cell.value).strip().lower()
                if "company name" in val_str:
                    return cell.row
    return 3 # Default fallback if not found

def excel_to_word_landscape(uploaded_file, excluded_keywords=None):
    """
    Convert Excel sheet to a landscape Word document with optimized table and merged cells.
    """
    # 1. Read Excel with openpyxl to get merge info
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.worksheets[0]
    
    # 2. Automatically detect start row
    header_row = find_company_name_row(ws)
    start_row = header_row + 1

    # 3. Determine which columns to keep
    if excluded_keywords is None:
        excluded_keywords = []
    
    # Get all headers from the detected header row
    headers = []
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        headers.append(str(val).strip() if val is not None else "")
    
    keep_col_indices = []
    for i, h_text in enumerate(headers):
        should_exclude = False
        h_lower = h_text.lower()
        for kw in excluded_keywords:
            if kw.lower() in h_lower:
                should_exclude = True
                break
        if not should_exclude:
            keep_col_indices.append(i)
            
    if not keep_col_indices:
        raise ValueError("No columns selected. Please select at least one column.")

    # 4. Get Merged Ranges
    merged_ranges = ws.merged_cells.ranges
    
    # 5. Determine actual data bounds (for data matrix extraction)
    max_r = 0
    for row in ws.iter_rows(min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None:
                max_r = max(max_r, cell.row)
    
    for m_range in merged_ranges:
        max_r = max(max_r, m_range.max_row)
    
    if max_r == 0:
        raise ValueError("The Excel file seems to be empty.")

    # 6. Filter and Remap Merged Ranges (Headers only)
    # Map old col index (0-based) to new col index (0-based)
    col_map = {old_idx: new_idx for new_idx, old_idx in enumerate(keep_col_indices)}
    
    header_merged_ranges = [m for m in merged_ranges if m.min_row < start_row]
    remapped_header_merges = []
    for m_range in header_merged_ranges:
        # Check if ALL columns in the range are kept
        old_indices = list(range(m_range.min_col - 1, m_range.max_col))
        if all(idx in col_map for idx in old_indices):
            # All columns kept. Check if they are contiguous in new mapping.
            new_indices = [col_map[idx] for idx in old_indices]
            if max(new_indices) - min(new_indices) == len(new_indices) - 1:
                remapped_header_merges.append({
                    'min_row': m_range.min_row,
                    'max_row': m_range.max_row,
                    'min_col': min(new_indices) + 1,
                    'max_col': max(new_indices) + 1
                })

    # 6. Extract data matrix efficiently and filter columns
    data_matrix = []
    # ws.iter_rows with values_only is much faster
    for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=ws.max_column, values_only=True):
        full_row = [str(val) if val is not None else "" for val in row]
        # Only keep selected columns
        filtered_row = [full_row[i] for i in keep_col_indices]
        data_matrix.append(filtered_row)

    # 7. Create Word Document
    doc = Document()
    
    # 5. Set Margins and Orientation (A4 Landscape)
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Mm(297)
    section.page_height = Mm(210)
    
    section.top_margin = Mm(25.4)
    section.bottom_margin = Mm(22.9)
    section.left_margin = Mm(17.8)
    section.right_margin = Mm(17.8)
    
    usable_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    
    max_c = len(keep_col_indices)
    max_r = len(data_matrix)
    
    table = doc.add_table(rows=max_r, cols=max_c)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Set Table to Fixed Layout for stability
    tbl = table._element
    tblPr = tbl.xpath('w:tblPr')
    if not tblPr:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    else:
        tblPr = tblPr[0]
    layout = tblPr.xpath('w:tblLayout')
    if not layout:
        layout = OxmlElement('w:tblLayout')
        layout.set(qn('w:val'), 'fixed')
        tblPr.append(layout)
    else:
        layout[0].set(qn('w:val'), 'fixed')

    # 8. Extract column widths
    col_widths = calculate_column_widths(data_matrix, start_row - 1, usable_width, header_row - 1)
    
    # Pre-identify columns to center
    centered_col_indices = set()
    if header_row - 1 < len(data_matrix):
        center_keywords = [
            "country", "registration", "total net sales are", 
            "consecutive years", "accepted companies", "rejected companies"
        ]
        for i, h_text in enumerate(data_matrix[header_row - 1]):
            h_text_lower = h_text.lower()
            if any(k in h_text_lower for k in center_keywords):
                centered_col_indices.add(i)
    
    # 9. Fill Table and Handle Styling
    merged_cells_to_skip = set()
    for m in remapped_header_merges:
        for lr in range(m['min_row'], m['max_row'] + 1):
            for lc in range(m['min_col'], m['max_col'] + 1):
                if lr == m['min_row'] and lc == m['min_col']:
                    continue
                merged_cells_to_skip.add((lr, lc))

    # last_2_start_idx for bold total logic
    last_2_start_idx = max(0, max_r - 2)

    for r_idx, word_row in enumerate(table.rows):
        r_num = r_idx + 1
        
        # Determine if this is a "Total" row (only for last 2 rows)
        is_total_row = False
        if r_idx >= last_2_start_idx:
            row_text = " ".join(str(v).lower() for v in data_matrix[r_idx])
            is_total_row = "total" in row_text
        
        for c_idx, word_cell in enumerate(word_row.cells):
            c_num = c_idx + 1
            
            # Set width
            word_cell.width = Inches(col_widths[c_idx])
            
            if (r_num, c_num) in merged_cells_to_skip:
                continue
            
            val_str = data_matrix[r_idx][c_idx]
            word_cell.text = val_str
            
            is_header = r_num < start_row
            word_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            # Cell margins
            tc = word_cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for m_name in ['top', 'left', 'bottom', 'right']:
                node = OxmlElement(f'w:{m_name}')
                node.set(qn('w:w'), '60')
                node.set(qn('w:type'), 'dxa')
                tcMar.append(node)
            tcPr.append(tcMar)
            
            # Styling
            for para in word_cell.paragraphs:
                if is_header or c_idx in centered_col_indices:
                    para.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                run = para.runs[0] if para.runs else para.add_run()
                run.font.name = 'Times New Roman'
                run.font.size = Pt(10)
                is_bold = bool(is_header or is_total_row)
                run.font.bold = is_bold

    # 10. Execute Merges
    for m in remapped_header_merges:
        top_left = table.cell(m['min_row'] - 1, m['min_col'] - 1)
        bottom_right = table.cell(m['max_row'] - 1, m['max_col'] - 1)
        if top_left != bottom_right:
            top_left.merge(bottom_right)

    # 11. Save to BytesIO
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

    # 10. Save to BytesIO (NO REPEATING HEADER)
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def render_matrix_processor_tab():
    st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>Matrix Processor</h2>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Convert Excel matrices to professional landscape Word documents</p>", unsafe_allow_html=True)

    # Settings
    auto_process = st.checkbox("Auto-process after upload", value=True, key="matrix_auto_process_check")

    # File Uploader
    uploaded_file = st.file_uploader(
        "Upload Excel file (.xlsx)",
        type=["xlsx"],
        help="Reads the first sheet of the Excel file",
        key="matrix_uploader"
    )

    if uploaded_file is not None:
        # Quick read to get headers for UI
        try:
            uploaded_file.seek(0)
            wb_temp = load_workbook(uploaded_file, data_only=True, read_only=True)
            ws_temp = wb_temp.worksheets[0]
            header_row_idx = find_company_name_row(ws_temp)
            
            all_headers = []
            # Read first row that has "Company Name" to get all column titles
            for col in range(1, ws_temp.max_column + 1):
                val = ws_temp.cell(row=header_row_idx, column=col).value
                all_headers.append(str(val).strip() if val is not None else f"Column {col}")
            wb_temp.close()
        except Exception as e:
            st.error(f"Error reading headers: {e}")
            all_headers = []

        if all_headers:
            st.write("### Column Selection")
            # Determine default selection
            default_selection = []
            for h in all_headers:
                h_lower = h.lower()
                is_excluded = any(exc.lower() in h_lower for exc in DEFAULT_EXCLUDED_HEADERS)
                is_included = any(inc.lower() in h_lower for inc in DEFAULT_INCLUDED_HEADERS)
                
                # If it's in excluded list, don't select. Otherwise, if it's in included or not in either, select.
                if not is_excluded:
                    default_selection.append(h)
            
            selected_headers = st.multiselect(
                "Choose columns to keep in Word document:",
                options=all_headers,
                default=default_selection,
                help="Columns matching excluded list are unchecked by default."
            )
            
            # Identify which keywords to exclude based on what's NOT selected
            excluded_from_selection = [h for h in all_headers if h not in selected_headers]
        else:
            excluded_from_selection = []

        file_id = f"matrix_{uploaded_file.name}_{uploaded_file.size}_{hash(tuple(excluded_from_selection))}"
        
        # Process button
        start_manual = st.button(
            "🚀 Convert to Word", 
            key="btn_matrix_process", 
            use_container_width=True,
            disabled=auto_process
        )
        
        # Trigger processing if manual button clicked OR (auto-process is ON and file changed)
        should_run = start_manual or (auto_process and st.session_state.get('last_matrix_id') != file_id)

        if should_run:
            st.session_state['last_matrix_id'] = file_id
            
            with st.status("Processing Excel to Word...", expanded=True) as status:
                try:
                    status.write("Reading Excel data...")
                    # Reset pointer just in case
                    uploaded_file.seek(0)
                    
                    status.write("Generating Word document (Landscape)...")
                    processed_stream = excel_to_word_landscape(uploaded_file, excluded_keywords=excluded_from_selection)
                    
                    status.write("Finalizing...")
                    st.session_state['matrix_processed_data'] = processed_stream.getvalue()
                    
                    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
                    base_name = uploaded_file.name.rsplit('.', 1)[0]
                    st.session_state['matrix_output_filename'] = f"{base_name}_{timestamp}.docx"
                    
                    status.update(label="✅ Conversion complete!", state="complete", expanded=False)
                    st.success("Successfully converted to Word!")
                except Exception as e:
                    st.error(f"❌ Error during conversion: {str(e)}")
                    st.session_state['last_matrix_id'] = None

        # Download button
        if st.session_state.get('last_matrix_id') == file_id and 'matrix_processed_data' in st.session_state:
            st.download_button(
                label="📥 Download Word Document",
                data=st.session_state['matrix_processed_data'],
                file_name=st.session_state['matrix_output_filename'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="btn_matrix_download"
            )
