"""Show G1-G5, E3, A-col near 主布颜色 for a specific PO's Brazil sheet."""
import openpyxl, re
from pathlib import Path

pl_dir = Path('../测试数据/debug/PL')
target_pos = {'PO149178', 'PO151208'}

for xlsx in sorted(pl_dir.glob('*.xlsx')):
    if xlsx.name.startswith('~'):
        continue
    wb = openpyxl.load_workbook(str(xlsx), data_only=True)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        e3 = ws['E3'].value
        a3 = ws['A3'].value
        if not e3 or not a3 or str(a3).strip() != '合同号:':
            continue
        e3_str = str(e3).strip()
        po_base = re.split(r'[-_]', e3_str)[0]
        if po_base not in target_pos:
            continue
        # Find 主布颜色 row
        zhbu_row = None
        for r in range(1, min(ws.max_row+1, 200)):
            v = ws.cell(r, 1).value
            if v and str(v).strip() == '主布颜色':
                zhbu_row = r
                break
        a_port = ''
        if zhbu_row and zhbu_row > 1:
            ap = ws.cell(zhbu_row-1, 1).value
            if ap: a_port = str(ap).strip()
        # Only show sheets where a_port has BRAZIL
        if 'BRAZIL' not in a_port.upper():
            continue
        from openpyxl.utils import get_column_letter
        print('FILE: %s  SHEET: %s  PO: %s' % (xlsx.name, sheet_name, po_base))
        print('  主布颜色 row: %s' % zhbu_row)
        print('  a_port (row above): %r' % a_port)
        # Scan ALL non-empty cells in rows 1-5
        for row_idx in range(1, 6):
            for c in range(1, 20):
                v = ws.cell(row_idx, c).value
                if v is not None:
                    print('  row%d %s%d = %r' % (row_idx, get_column_letter(c), row_idx, v))
