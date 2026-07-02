"""
PO Splitter API — Vercel Serverless Function (Flask + openpyxl)
Pure Python, no COM dependency. Works in-memory (no temp files).
Support dynamic column selection: source, PO1, HNA, Master Box.
"""

import io
import json
import traceback
from copy import copy

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string

app = Flask(__name__)
CORS(app)

# Default column constants
DEFAULT_SOURCE_COL = 'AS'   # po sisa
DEFAULT_PO1_COL = 'AR'      # po 1
DEFAULT_HNA_COL = 'F'       # HNA UM
DEFAULT_MB_COL = 'P'        # MASTER BOX
DEFAULT_ENABLE_PO1 = False

COL_AT = 46  # start of new columns


def copy_cell_format(src_cell, dst_cell):
    try: dst_cell.font = copy(src_cell.font)
    except: pass
    try: dst_cell.fill = copy(src_cell.fill)
    except: pass
    try: dst_cell.border = copy(src_cell.border)
    except: pass
    try: dst_cell.number_format = src_cell.number_format
    except: pass
    try: dst_cell.alignment = copy(src_cell.alignment)
    except: pass
    try: dst_cell.protection = copy(src_cell.protection)
    except: pass


def find_sheet_and_header(wb):
    pp_name = None
    for name in wb.sheetnames:
        if name.upper() == 'PP' or 'PP' in name.upper():
            pp_name = name
            break
    if not pp_name:
        raise ValueError('Sheet "PP" tidak ditemukan.')
    ws = wb[pp_name]
    max_row = ws.max_row or 1
    max_col = ws.max_column or 50
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and 'po sisa' in cell.value.lower():
                header_row = cell.row
                break
        if header_row:
            break
    if not header_row:
        raise ValueError('Header "po sisa" tidak ditemukan.')
    return ws, pp_name, header_row


def find_data_range(ws, header_row):
    max_row = ws.max_row or 1
    data_start = None
    data_end = None
    for row_num in range(header_row + 1, max_row + 1):
        cell_a = ws.cell(row=row_num, column=1).value
        if cell_a is not None:
            is_data = False
            if isinstance(cell_a, (int, float)):
                is_data = cell_a > 0
            elif isinstance(cell_a, str):
                try:
                    val = float(cell_a)
                    is_data = val > 0
                except (ValueError, TypeError):
                    if cell_a.startswith('='):
                        is_data = True
            if is_data:
                if data_start is None:
                    data_start = row_num
                data_end = row_num
    if data_start is None:
        raise ValueError('Data produk tidak ditemukan.')
    return data_start, data_end


def process_excel(file_bytes, num_po, percentages, source_col='AS', po1_col='AR',
                  enable_po1=False, hna_col='F', mb_col='P'):
    """Process Excel in-memory. Returns BytesIO."""
    wb = load_workbook(io.BytesIO(file_bytes))
    ws, pp_name, header_row = find_sheet_and_header(wb)

    if ws.protection.sheet:
        ws.protection.sheet = False
    if ws.protection.password:
        ws.protection.password = None

    data_start, data_end = find_data_range(ws, header_row)

    src_idx = column_index_from_string(source_col)
    po1_idx = column_index_from_string(po1_col)
    hna_idx = column_index_from_string(hna_col)
    mb_idx = column_index_from_string(mb_col)

    start_col = COL_AT
    summary_row_val = header_row - 2
    total_col = start_col + num_po
    data_number_format = '#,##0'

    ref_hdr = ws.cell(row=header_row, column=src_idx)
    for i in range(num_po + 1):
        copy_cell_format(ref_hdr, ws.cell(row=header_row, column=start_col + i))

    for i in range(num_po):
        ws.cell(row=header_row, column=start_col + i).value = f"PO Sisa {i + 1} ({percentages[i]}%)"
    ws.cell(row=header_row, column=total_col).value = "TOTAL PO Sisa"

    src_letter = get_column_letter(src_idx)
    po1_letter = get_column_letter(po1_idx)
    mb_letter = get_column_letter(mb_idx)
    hna_letter = get_column_letter(hna_idx)

    def source_expr(row_num):
        if enable_po1:
            return f"({src_letter}{row_num}-{po1_letter}{row_num})"
        return f"{src_letter}{row_num}"

    for row_num in range(data_start, data_end + 1):
        s_expr = source_expr(row_num)
        for i in range(num_po):
            col_idx = start_col + i
            cell = ws.cell(row=row_num, column=col_idx)
            if i < num_po - 1:
                pct = percentages[i] / 100
                cell.value = f"=MROUND({s_expr}*{pct},{mb_letter}{row_num})"
            else:
                prev_refs = [f"{get_column_letter(start_col + j)}{row_num}" for j in range(num_po - 1)]
                cell.value = f"={s_expr}-({'+'.join(prev_refs)})"
            cell.number_format = data_number_format

        total_cell = ws.cell(row=row_num, column=total_col)
        po_refs = [f"{get_column_letter(start_col + j)}{row_num}" for j in range(num_po)]
        total_cell.value = f"={'+'.join(po_refs)}"
        total_cell.number_format = data_number_format

    ws.cell(row=summary_row_val, column=1).value = "TOTAL NILAI RUPIAH (Qty x HNA)"
    for i in range(num_po):
        cl = get_column_letter(start_col + i)
        ws.cell(row=summary_row_val, column=start_col + i).value = \
            f"=SUMPRODUCT({cl}{data_start}:{cl}{data_end},{hna_letter}{data_start}:{hna_letter}{data_end})"
    tcl = get_column_letter(total_col)
    ws.cell(row=summary_row_val, column=total_col).value = \
        f"=SUMPRODUCT({tcl}{data_start}:{tcl}{data_end},{hna_letter}{data_start}:{hna_letter}{data_end})"

    output = io.BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    return output


def preview_excel(file_bytes):
    """Read Excel, return headers + ALL column values (vals) for dynamic access."""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws, pp_name, header_row = find_sheet_and_header(wb)
    max_col = ws.max_column or 50

    headers = []
    for col_idx in range(1, max_col + 1):
        cell_val = ws.cell(row=header_row, column=col_idx).value
        if cell_val and str(cell_val).strip():
            letter = get_column_letter(col_idx)
            headers.append({
                'col': col_idx,
                'letter': letter,
                'label': f"{letter} — {str(cell_val).strip()}"
            })

    products = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        a_val = row[0].value
        if a_val is None:
            continue
        try:
            no = int(float(a_val))
        except (ValueError, TypeError):
            continue
        if no <= 0:
            continue

        # Build vals array (0-indexed, ALL columns)
        vals = []
        for cell in row:
            v = cell.value
            if v is None:
                vals.append(0)
            elif isinstance(v, (int, float)):
                vals.append(v)
            else:
                try:
                    vals.append(float(v))
                except:
                    vals.append(0)

        def str_val(v):
            return str(v) if v is not None else ''

        products.append({
            'no': no,
            'supplier': str_val(row[1].value) if len(row) > 1 else '',
            'partNo': str_val(row[2].value) if len(row) > 2 else '',
            'desc': str_val(row[3].value) if len(row) > 3 else '',
            'vals': vals,
        })

    wb.close()
    return products, pp_name, headers


# ============================================================
# API Routes
# Both /api/* and /* paths (Vercel may strip /api prefix)
# ============================================================

@app.route('/api/health', methods=['GET'])
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/preview', methods=['POST'])
@app.route('/preview', methods=['POST'])
def api_preview():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['file']
        products, sheet_name, headers = preview_excel(file.read())
        return jsonify({
            'products': products,
            'sheetName': sheet_name,
            'headers': headers,
            'defaults': {
                'sourceCol': DEFAULT_SOURCE_COL,
                'po1Col': DEFAULT_PO1_COL,
                'enablePO1': DEFAULT_ENABLE_PO1,
                'hnaCol': DEFAULT_HNA_COL,
                'mbCol': DEFAULT_MB_COL,
            }
        })
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/process', methods=['POST'])
@app.route('/process', methods=['POST'])
def api_process():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        num_po = int(request.form.get('numPO', 4))
        percentages = json.loads(str(request.form.get('percentages', '[]')))
        source_col = request.form.get('sourceCol', DEFAULT_SOURCE_COL).upper()
        po1_col = request.form.get('po1Col', DEFAULT_PO1_COL).upper()
        enable_po1 = request.form.get('enablePO1', 'false').lower() == 'true'
        hna_col = request.form.get('hnaCol', DEFAULT_HNA_COL).upper()
        mb_col = request.form.get('mbCol', DEFAULT_MB_COL).upper()

        if len(percentages) < num_po:
            return jsonify({'error': 'Persentase tidak lengkap'}), 400
        total_pct = sum(percentages[:num_po])
        if total_pct != 100:
            return jsonify({'error': f'Total persentase harus 100%, sekarang {total_pct}%'}), 400

        output = process_excel(file.read(), num_po, percentages,
                               source_col, po1_col, enable_po1, hna_col, mb_col)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"PP_Output_PO_Split_{num_po}PO.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# Vercel auto-detects Flask `app` variable — no handler needed
