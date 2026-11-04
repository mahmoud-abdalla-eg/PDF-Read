import sys
import os
import re
import warnings
warnings.filterwarnings("ignore")

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def contains_arabic(text: str) -> bool:
    if not text:
        return False
    return bool(re.search(r'[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]', text))

def clean_cell_value(val):
    if val is None:
        return ""
    val_str = str(val).strip()
    if not val_str:
        return ""

    # Check for integer or float
    clean_num = val_str.replace(",", "").replace("$", "").replace("€", "").replace("£", "").replace("ج.م", "").strip()
    if re.match(r'^-?\d+(\.\d+)?$', clean_num):
        try:
            if "." in clean_num:
                return float(clean_num)
            return int(clean_num)
        except ValueError:
            pass
    return val_str

def convert_pdf_to_excel(pdf_path: str, output_xlsx: str):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default blank sheet

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    header_fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
    header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    regular_font = Font(name='Segoe UI', size=10, color='0F172A')
    alt_fill = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            sheet_title = f"Page {page_idx + 1}"
            ws = wb.create_sheet(title=sheet_title)

            full_text = page.extract_text() or ""
            is_rtl = contains_arabic(full_text)
            if is_rtl:
                ws.sheet_view.rightToLeft = True

            # Strategy 1: Explicit tables by lines
            tables = page.extract_tables()

            # Strategy 2: If no explicit tables, try text-alignment table detection
            if not tables:
                table_settings = {
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_y_tolerance": 4,
                    "intersection_x_tolerance": 15
                }
                tables = page.extract_tables(table_settings)

            current_row = 1

            if tables:
                for table in tables:
                    if not table or len(table) == 0:
                        continue

                    valid_rows = [r for r in table if any(c and str(c).strip() for c in r)]
                    if not valid_rows:
                        continue

                    for r_idx, row in enumerate(valid_rows):
                        is_header = (r_idx == 0)
                        for c_idx, cell_value in enumerate(row):
                            cleaned = clean_cell_value(cell_value)
                            col_num = c_idx + 1
                            cell = ws.cell(row=current_row, column=col_num)
                            cell.value = cleaned
                            cell.border = thin_border

                            if is_header:
                                cell.fill = header_fill
                                cell.font = header_font
                                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            else:
                                cell.font = regular_font
                                if r_idx % 2 == 1:
                                    cell.fill = alt_fill
                                
                                is_num = isinstance(cleaned, (int, float))
                                cell.alignment = Alignment(
                                    horizontal='right' if (is_rtl or is_num) else 'left',
                                    vertical='center'
                                )

                        ws.row_dimensions[current_row].height = 24 if is_header else 20
                        current_row += 1

                    current_row += 2 # Spacing between multiple tables
            else:
                # Fallback: extract words and construct visually aligned rows and columns
                words = page.extract_words()
                if words:
                    lines = []
                    curr_line = []
                    last_top = None
                    for w in sorted(words, key=lambda x: (x['top'], x['x0'])):
                        top = w['top']
                        if last_top is None or abs(top - last_top) <= 4:
                            curr_line.append(w)
                            if last_top is None:
                                last_top = top
                        else:
                            if curr_line:
                                lines.append(sorted(curr_line, key=lambda x: x['x0']))
                            curr_line = [w]
                            last_top = top
                    if curr_line:
                        lines.append(sorted(curr_line, key=lambda x: x['x0']))

                    for line in lines:
                        cols = []
                        curr_col = []
                        last_x1 = None
                        for w in line:
                            if last_x1 is not None and (w['x0'] - last_x1) > 14:
                                cols.append(" ".join(curr_col))
                                curr_col = []
                            curr_col.append(w['text'])
                            last_x1 = w['x1']
                        if curr_col:
                            cols.append(" ".join(curr_col))

                        for c_idx, text in enumerate(cols):
                            val = clean_cell_value(text)
                            cell = ws.cell(row=current_row, column=c_idx + 1)
                            cell.value = val
                            cell.font = regular_font
                            cell.alignment = Alignment(horizontal='right' if is_rtl else 'left', vertical='center')

                        current_row += 1

            # Auto-fit column widths
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.value is not None:
                        s = str(cell.value)
                        max_len = max(max_len, len(s))
                ws.column_dimensions[col_letter].width = max(min(max_len + 4, 50), 12)

    if len(wb.worksheets) == 0:
        ws = wb.create_sheet(title="Sheet 1")
        ws.cell(row=1, column=1, value="No table data detected in PDF.")

    wb.save(output_xlsx)
    print(f"EXCEL_EXPORT_SUCCESS: {output_xlsx}")

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: convert_pdf_to_excel.py <input_pdf> <output_xlsx>")
        sys.exit(1)
    convert_pdf_to_excel(sys.argv[1], sys.argv[2])
