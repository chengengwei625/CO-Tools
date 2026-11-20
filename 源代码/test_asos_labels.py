"""Run: python -m unittest discover -s 源代码 -p test_asos_labels.py"""
import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz
import openpyxl
from reportlab.pdfgen import canvas

from asos_labels_app import (assign_marks, export_job, make_job, normalize_asn, parse_asn_mapping, plan_batch,
                             positive_int, read_asos_po, read_packing)


PO_HEADER = [('Option ID', 51), ('Supplier Reference', 55), ('Brand', 54), ('Reporting Colour', 60),
             ('Website Color', 60), ('Barcode', 50), ('SKU ID', 59), ('Brand Size', 41), ('First Dest', 49),
             ('Final Dest', 51), ('Units', 40), ('Unit Cost', 50), ('Total Cost', 60), ('Packing Method', 50),
             ('Commodity Code', 70)]


def write_po(path, po_id, colour, rows, warehouse='FC01'):
    """Fake ASOS PO like the real one: bordered header, borderless rows, Barcode sometimes
    wrapped to two lines or blank, and a footer inside the column ranges."""
    pdf = canvas.Canvas(str(path), pagesize=(842, 595))
    pdf.setFont('Helvetica', 9)
    pdf.drawString(40, 560, 'ASOS Purchase Order')
    pdf.drawString(40, 545, f'PO ID: {po_id}')
    pdf.drawString(40, 530, f'Total Units: {sum(q for _, _, q in rows)}')
    x, top, starts = 40, 480, {}
    for name, width in PO_HEADER:
        pdf.rect(x, top, width, 24)
        pdf.setFont('Helvetica-Bold', 5)
        pdf.drawString(x + 2, top + 10, name)
        starts[name] = x + 2
        x += width
    pdf.setFont('Helvetica', 7)
    for number, (size, sku, qty) in enumerate(rows):
        y = top - 18 - number * 22
        values = {'Option ID': '138466687', 'Supplier Reference': sku, 'Brand': 'Cotton:On',
                  'Reporting Colour': '2001', 'Website Color': colour, 'SKU ID': sku, 'Brand Size': size,
                  'First Dest': warehouse, 'Final Dest': warehouse, 'Units': str(qty), 'Packing Method': 'Flat',
                  'Commodity Code': '6212109000'}
        for name, value in values.items():
            pdf.drawString(starts[name], y, value)
        if number % 2 == 0:  # wrapped barcode on even rows, blank on odd rows
            pdf.drawString(starts['Barcode'], y, '935919765')
            pdf.drawString(starts['Barcode'], y - 8, '6646')
    pdf.drawString(starts['Supplier Reference'], 30, 'Print Date: 08/10/2026')
    pdf.drawString(starts['Packing Method'], 30, 'Page 1 of 1')
    pdf.save()


class AsosTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.po_path = self.folder / 'po.pdf'
        write_po(self.po_path, '500038555238', 'BLACK',
                 [('XS', '138466752', 17), ('S', '138466753', 50), ('M', '138466754', 45),
                  ('L', '138466755', 18), ('XL', '138466756', 10)])
        self.pl_path = self.folder / 'PO159521 6335745 PL-ASOS UK.xlsx'
        w = openpyxl.Workbook()
        s = w.active
        for cell, value in {'A1': '装箱资料表', 'G2': 'ASOS UK -1609', 'D3': 'PO159521',
                            'D5': '6335745', 'B12': '箱号', 'G12': 'XS', 'H12': 'S',
                            'I12': 'M', 'J12': 'L', 'K12': 'XL', 'A14': '黑色',
                            'B14': 1, 'D14': 1, 'G14': 17, 'H14': 50, 'I14': 23,
                            'A15': 'BLACK', 'B15': 2, 'D15': 2, 'I15': 22, 'J15': 18,
                            'K15': 10, 'Y14': '=SUM(F14:X14)', 'A18': '米白色',
                            'B18': 3, 'D18': 4, 'G18': 5, 'A19': 'CREAM', 'D21': '共'}.items():
            s[cell] = value
        w.save(self.pl_path)
        w.close()
        self.po = read_asos_po(self.po_path)
        self.packing = read_packing(self.pl_path)

    def job(self):
        return make_job(self.po, self.packing, 0, '42370000165328', 'COTTON ON', 'Clothing')

    def test_sample_quantities_and_formula_free_reading(self):
        job = self.job()
        self.assertEqual(job['po'], '500038555238')
        self.assertEqual([sum(i['qty'] for i in c['items']) for c in job['cartons']], [90, 50])
        self.assertEqual(job['cartons'][0]['items'][2], dict(sku='138466754', size='M', qty=23))
        self.assertEqual([c['source'] for c in self.packing['groups'][1]['cartons']], [3, 4])
        self.assertIn('CREAM', self.packing['groups'][1]['names'])

    def test_asn_and_integer_validation(self):
        self.assertEqual(normalize_asn('123456789123'), '00123456789123')
        self.assertEqual(normalize_asn('01234567891234'), '01234567891234')
        for value in ['123', '１２３４５６７８９１２３', '123456789123456', '12345678912A']:
            with self.assertRaises(ValueError):
                normalize_asn(value)
        for value in [1.5, -1, True, '=SUM(A1)', 'NaN']:
            with self.assertRaises(ValueError):
                positive_int(value, '数量')

    def test_wrong_colour_quantity_and_destination_blocked(self):
        with self.assertRaisesRegex(ValueError, '颜色'):
            make_job(self.po, self.packing, 1, '42370000165328', 'COTTON ON', 'Clothing')
        self.packing['groups'][0]['cartons'][0]['sizes']['XS'] = 16
        with self.assertRaisesRegex(ValueError, '尺码数量'):
            self.job()
        self.packing['destination'] = 'ASOS EU'
        with self.assertRaisesRegex(ValueError, '目的仓'):
            self.job()

    def test_duplicate_cartons_and_fractional_quantities_blocked(self):
        for cell, value in [('B15', 1), ('G14', 1.5), ('L14', 2)]:
            w = openpyxl.load_workbook(self.pl_path)
            old = w.active[cell].value
            w.active[cell] = value
            w.save(self.pl_path)
            w.close()
            with self.assertRaises(ValueError):
                read_packing(self.pl_path)
            w = openpyxl.load_workbook(self.pl_path)
            w.active[cell] = old
            w.save(self.pl_path)
            w.close()

    def test_asn_mapping_parsing(self):
        # COG PO may sit anywhere on the line (email column order puts it first)
        self.assertEqual(parse_asn_mapping('500038555238  42370000165328  PO159521\n\n'
                                           'po159531,\t500038133504, 423700001322'),
                         {'500038555238': dict(asn='42370000165328', cog='PO159521'),
                          '500038133504': dict(asn='00423700001322', cog='PO159531')})
        for text in ['', '500038555238 42370000165328', '500038555238 PO159521',
                     '500038555238 42370000165328 PO159521 6335745',
                     '1 42370000165328 PO1\n1 42370000165329 PO1', '1 42370000165328 PO1\n2 42370000165328 PO1',
                     '500038555238 123 PO159521']:
            with self.assertRaises(ValueError):
                parse_asn_mapping(text)

    def save_packing_copy(self, name, **cells):
        w = openpyxl.load_workbook(self.pl_path)
        for cell, value in cells.items():
            w.active[cell] = value
        w.save(self.folder / name)
        w.close()
        return self.folder / name

    def test_batch_matches_by_cog_po_in_mixed_folder(self):
        cream = self.folder / 'cream.pdf'
        write_po(cream, '500038133504', 'CREAM', [('XS', '138466800', 10)])
        eu = self.folder / 'eu.pdf'
        write_po(eu, '500038555239', 'BLACK', [('XS', '138466900', 1)], warehouse='FC04')
        mapping = ('500038555238 42370000165328 PO159521\n500038133504 42370000132228 PO159521\n'
                   '500038555239 42370000165329 PO159521\n500038131474 42370000159843 PO159872')
        # Unrelated order with an identical BLACK UK group, a renamed file, a non-packing xlsx, a lock file
        folder = [self.pl_path, self.save_packing_copy('PO888888 other order.xlsx', D3='PO888888'),
                  self.save_packing_copy('PO159521 wrong.xlsx', D3='PO777777'),
                  self.save_packing_copy('PO159521 summary.xlsx', A1='汇总'),
                  self.save_packing_copy('~$PO159521 6335745 PL-ASOS UK.xlsx')]
        plan = plan_batch([self.po_path, cream, eu], folder, mapping, 'COTTON ON', 'Clothing')
        self.assertEqual([(j['po'], j['asn'], j['colour'], j['mark']) for j in plan['jobs']],
                         [('500038133504', '42370000132228', 'CREAM', 'Red'),
                          ('500038555238', '42370000165328', 'BLACK', 'Blue')])
        self.assertEqual([c['source'] for c in plan['jobs'][0]['cartons']], [3, 4])
        self.assertEqual(len(plan['errors']), 1)
        self.assertIn('eu.pdf', plan['errors'][0])  # FC04 PO, only a UK packing list given
        self.assertEqual(plan['unused'], ['500038131474'])
        self.assertEqual(plan['ignored'], 2)
        other = plan_batch([self.po_path, self.folder / 'bad.pdf'] if (self.folder / 'bad.pdf').exists() else [self.po_path], [self.pl_path], mapping, 'COTTON ON', 'Clothing')
        self.assertEqual(other['errors'], [])  # PO888888 file and the ~$ lock file
        self.assertEqual(sorted(s.split('：')[0] for s in plan['skipped']),
                         ['PO159521 summary.xlsx', 'PO159521 wrong.xlsx'])
        ncs = plan_batch([self.po_path], [self.pl_path], mapping, 'COTTON ON', 'Clothing', ncs=True)
        self.assertEqual(ncs['jobs'][0]['mark'], 'NCS - no colour')
        missing = plan_batch([self.po_path], [self.pl_path], '1 42370000165328 PO159521', 'COTTON ON', 'Clothing')
        self.assertIn('没有 ASOS PO 500038555238', missing['errors'][0])
        no_list = plan_batch([self.po_path], [self.pl_path], '500038555238 42370000165328 PO159531',
                             'COTTON ON', 'Clothing')
        self.assertIn('找不到合同号 PO159531', no_list['errors'][0])
        revision = self.save_packing_copy('PO159521 6335745 PL-ASOS UK 修正1.xlsx')
        both = plan_batch([self.po_path], [self.pl_path, revision], mapping, 'COTTON ON', 'Clothing')
        self.assertIn('修正1', both['errors'][0])  # old + revised version: user must keep one

    def test_single_colour_group_used_when_names_differ(self):
        multi = self.folder / 'multi.pdf'
        write_po(multi, '500038131474', 'MULTI',
                 [('XS', '160825628', 17), ('S', '160825629', 50), ('M', '160825630', 45),
                  ('L', '160825631', 18), ('XL', '160825632', 10)])
        # PO159872 packing list with one group named in Chinese / "RED DITSY/HEART", not MULTI
        one_group = self.save_packing_copy('PO159872 6340305 PL-ASOS UK.xlsx', D3='PO159872', A15='RED DITSY/HEART',
                                           A18=None, B18=None, D18=None, G18=None, A19=None)
        plan = plan_batch([multi], [one_group], '500038131474 42370000159843 PO159872', 'COTTON ON', 'Clothing')
        self.assertEqual(plan['errors'], [])
        self.assertIn('颜色名称不同', plan['jobs'][0]['note'])
        # With two colour groups the names must match: no guessing
        plan = plan_batch([multi], [self.save_packing_copy('PO159872 two.xlsx', D3='PO159872')],
                          '500038131474 42370000159843 PO159872', 'COTTON ON', 'Clothing')
        self.assertIn('找不到 MULTI', plan['errors'][0])

    def test_every_asn_in_batch_gets_its_own_mark(self):
        jobs = [dict(destination='FC01' if n % 2 else 'FC04') for n in range(24)]  # UK and EU mixed
        assign_marks(jobs)
        marks = [(j['mark'], j['shape']) for j in jobs]
        self.assertEqual(len(set(marks)), 24)
        self.assertEqual(marks[:7], [('Red', 'block'), ('Blue', 'block'), ('Green', 'block'), ('Yellow', 'block'),
                                     ('Orange', 'block'), ('Purple', 'block'), ('Red', 'circle')])
        with self.assertRaisesRegex(ValueError, '24'):
            assign_marks(jobs + [dict(destination='FC01')])
        assign_marks(jobs + [dict(destination='FC01')], ncs=True)  # NCS: no marks, no limit
        for shape in ['circle', 'triangle', 'star']:
            target = export_job(self.job(), self.folder, 'Purple', shape)
            self.assertTrue((target / 'CartonLabels.pdf').is_file())

    def test_more_than_four_skus_in_carton_blocked(self):
        self.packing['groups'][0]['cartons'] = [dict(source=1, sizes={s: 1 for s in 'ABCDE'})]
        self.po['rows'] = [dict(size=s, sku=str(n), qty=1) for n, s in enumerate('ABCDE')]
        with self.assertRaisesRegex(ValueError, '最多 4'):
            self.job()

    def test_mixed_sku_limit(self):
        self.packing['groups'][0]['cartons'] = [
            dict(source=n, sizes={'XS': 1, 'S': 1}) for n in range(1, 4)]
        self.po['rows'] = [dict(size='XS', sku='111', qty=3), dict(size='S', sku='222', qty=3)]
        with self.assertRaisesRegex(ValueError, '超过两只'):
            self.job()

    def test_pdf_exports_continuations_and_no_overwrite(self):
        job = self.job()
        first = export_job(job, self.folder, 'Red')
        second = export_job(job, self.folder, 'Blue')
        self.assertNotEqual(first, second)
        with fitz.open(first / 'CartonLabels.pdf') as doc:
            self.assertEqual(len(doc), 2)
            text = doc[0].get_text()
            self.assertIn('1 of 2', text)
            self.assertNotIn('SKU D', text)
            # Barcodes are template-font text: *ASN*, *SKU*, *ASN+SKU+QTY*, *NO OF SKUs*
            for code in ['*42370000165328*', '*138466754*', '*4237000016532813846675423*', '*3*']:
                self.assertIn(code, text)
            fonts = {f[3] for f in doc[0].get_fonts()}
            self.assertTrue(fonts - {'Helvetica'}, fonts)  # embedded barcode TTF (internal name "New")
        with fitz.open(first / (job['asn'] + '.pdf')) as doc:
            self.assertIn('140 UNITS / 2 CARTONS', doc[0].get_text())
        extra = copy.deepcopy(job)
        extra['cartons'] = [dict(number=1, source=3, items=[
            dict(sku=str(138466752 + n), size=str(n), qty=1) for n in range(4)])]
        target = export_job(extra, self.folder, 'NCS - no colour')
        with fitz.open(target / 'CartonLabels.pdf') as doc:
            self.assertEqual(len(doc), 1)  # Extra SKUs layout: one label, SKU A–D
            self.assertIn('SKU D', doc[0].get_text())
            self.assertIn('*138466755*', doc[0].get_text())
            self.assertIn('*4*', doc[0].get_text())
            self.assertIn('1 of 1', doc[0].get_text())
        extra['cartons'][0]['items'].append(dict(sku='138466757', size='X', qty=1))
        with self.assertRaisesRegex(ValueError, '超过 4'):
            export_job(extra, self.folder, 'Red')
        extra['cartons'] = [dict(number=n, source=n + 2, items=job['cartons'][0]['items'])
                            for n in range(1, 31)]
        target = export_job(extra, self.folder, 'Green')
        with fitz.open(target / (job['asn'] + '.pdf')) as doc:
            self.assertGreater(len(doc), 1)
            self.assertIn('2700 UNITS / 30 CARTONS', doc[-1].get_text())
            self.assertIn('SOURCE CTN', doc[-1].get_text())

    def test_failed_export_leaves_no_partial_output(self):
        before = set(self.folder.iterdir())
        with patch('asos_labels_app.write_packing_pdf', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                export_job(self.job(), self.folder, 'Red')
        self.assertEqual(set(self.folder.iterdir()), before)

    def test_incomplete_po_rejected(self):
        pdf = canvas.Canvas(str(self.folder / 'bad.pdf'))
        pdf.drawString(10, 10, 'ASOS Purchase Order PO ID: 1234')
        pdf.save()
        with self.assertRaises(ValueError):
            read_asos_po(self.folder / 'bad.pdf')


if __name__ == '__main__':
    unittest.main()
