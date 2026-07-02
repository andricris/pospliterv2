"""
PO Splitter API — Vercel Serverless Function (Flask + openpyxl)
Pure Python, no COM dependency. Works in-memory (no temp files).
"""

import io
import json
import traceback
from copy import copy

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Column constants (1-based)
COL_A = 1
COL_F = 6    # HNA UM
COL_P = 16   # MASTER BOX
COL_W = 23   # Purch Plan by kelipatan
COL_AR = 44  # po 1 0,75
COL_AS = 45  # po sisa
COL_AT = 46  # start of new columns


def copy_cell_format(src_cell, dst_cell):
    """Copy format from src to dst using openpyxl copy()."""
    try:
        dst_cell.font = copy(src_cell.font)
    except:
        pass
    try:
        dst_cell.fill = copy(src_cell.fill)
    except:
        pass
    try:
        dst_cell.border = copy(src_cell.border)
    except:
        pass
    try:
        dst_cell.number_format = src_cell.number_format
    except:
        pass
    try:
        dst_cell.alignment = copy(src_cell.alignment)
    except:
        pass
    try:
        dst_cell.protection = copy(src_cell.protection)
    except:
        pass


def process_excel(file_bytes, num_po, percentages):
    """
    Process Excel file in-memory with openpyxl.
    Returns processed file as BytesIO.
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    wb = load_workbook(io.BytesIO(file_bytes))

    # Find PP sheet
    pp_name = None
    for name in wb.sheetnames:
        if name.upper() == 'PP' or 'PP' in name.upper():
            pp_name = name
            break
    if not pp_name:
        wb.close()
        raise ValueError('Sheet "PP" tidak ditemukan.')

    ws = wb[pp_name]

    # Unprotect
    if ws.protection.sheet:
        ws.protection.sheet = False
    if ws.protection.password:
        ws.protection.password = None

    max_row = ws.max_row or 1
    max_col = ws.max_column or 50

    # Find header row
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and 'po sisa' in cell.value.lower():
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        wb.close()
        raise ValueError('Header "po sisa" tidak ditemukan.')

    # Find data rows (handle formulas in column A)
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
        wb.close()
        raise ValueError('Data produk tidak ditemukan.')

    start_col = COL_AT
    summary_row_val = header_row - 2
    total_col = start_col + num_po

    # Number format: plain comma integers
    data_number_format = '#,##0'

    # Copy header format from AS
    ref_hdr = ws.cell(row=header_row, column=COL_AS)
    for i in range(num_po + 1):
        col_idx = start_col + i
        dst_hdr = ws.cell(row=header_row, column=col_idx)
        copy_cell_format(ref_hdr, dst_hdr)

    # Header labels
    for i in range(num_po):
        ws.cell(row=header_row, column=start_col + i).value = f"PO Sisa {i + 1} ({percentages[i]}%)"
    ws.cell(row=header_row, column=total_col).value = "TOTAL PO Sisa"

    # Column letters
    as_col = get_column_letter(COL_AS)
    p_col = get_column_letter(COL_P)
    f_col = get_column_letter(COL_F)

    # Data formulas
    for row_num in range(data_start, data_end + 1):
        for i in range(num_po):
            col_idx = start_col + i
            cell = ws.cell(row=row_num, column=col_idx)

            if i < num_po - 1:
                pct = percentages[i] / 100
                cell.value = f"=MROUND({as_col}{row_num}*{pct},{p_col}{row_num})"
            else:
                prev_refs = [f"{get_column_letter(start_col + j)}{row_num}" for j in range(num_po - 1)]
                cell.value = f"={as_col}{row_num}-({'+'.join(prev_refs)})"

            cell.number_format = data_number_format

        # Total
        total_cell = ws.cell(row=row_num, column=total_col)
        po_refs = [f"{get_column_letter(start_col + j)}{row_num}" for j in range(num_po)]
        total_cell.value = f"={'+'.join(po_refs)}"
        total_cell.number_format = data_number_format

    # Summary: TOTAL NILAI RUPIAH
    ws.cell(row=summary_row_val, column=1).value = "TOTAL NILAI RUPIAH (Qty x HNA)"
    for i in range(num_po):
        cl = get_column_letter(start_col + i)
        ws.cell(row=summary_row_val, column=start_col + i).value = \
            f"=SUMPRODUCT({cl}{data_start}:{cl}{data_end},{f_col}{data_start}:{f_col}{data_end})"

    tcl = get_column_letter(total_col)
    ws.cell(row=summary_row_val, column=total_col).value = \
        f"=SUMPRODUCT({tcl}{data_start}:{tcl}{data_end},{f_col}{data_start}:{f_col}{data_end})"

    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    return output


def preview_excel(file_bytes):
    """Read Excel file via openpyxl, return product data for preview."""
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    pp_name = None
    for name in wb.sheetnames:
        if name.upper() == 'PP' or 'PP' in name.upper():
            pp_name = name
            break
    if not pp_name:
        wb.close()
        raise ValueError('Sheet "PP" tidak ditemukan.')

    ws = wb[pp_name]

    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and 'po sisa' in cell.value.lower():
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        wb.close()
        raise ValueError('Header "po sisa" tidak ditemukan.')

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

        def gv(r, idx):
            v = r[idx - 1].value if idx - 1 < len(r) else None
            return v

        def num(v):
            if v is None:
                return 0
            try:
                return float(v)
            except:
                return 0

        def str_val(v):
            if v is None:
                return ''
            return str(v)

        products.append({
            'no': no,
            'supplier': str_val(gv(row, 2)),
            'partNo': str_val(gv(row, 3)),
            'desc': str_val(gv(row, 4)),
            'hna': num(gv(row, COL_F)),
            'masterBox': num(gv(row, COL_P)),
            'purchPlan': num(gv(row, COL_W)),
            'po1': num(gv(row, COL_AR)),
            'poSisa': num(gv(row, COL_AS)),
        })

    wb.close()
    return products, pp_name


# ============================================================
# API Routes
# ============================================================

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/api/preview', methods=['POST'])
def api_preview():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        file_bytes = file.read()
        products, sheet_name = preview_excel(file_bytes)
        return jsonify({'products': products, 'sheetName': sheet_name})

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/process', methods=['POST'])
def api_process():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']
        num_po = int(request.form.get('numPO', 4))
        percentages = json.loads(str(request.form.get('percentages', '[]')))

        if len(percentages) < num_po:
            return jsonify({'error': 'Persentase tidak lengkap'}), 400

        total_pct = sum(percentages[:num_po])
        if total_pct != 100:
            return jsonify({'error': f'Total persentase harus 100%, sekarang {total_pct}%'}), 400

        file_bytes = file.read()
        output = process_excel(file_bytes, num_po, percentages)

        output_name = f"PP_Output_PO_Split_{num_po}PO.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=output_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# Vercel handler
def handler(request):
    return app
