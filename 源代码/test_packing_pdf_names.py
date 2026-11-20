"""Filename regression checks for both routes through gen_packing_pdfs."""
import tempfile
import unittest
from pathlib import Path

import fitz

from generate_tabs import _format_b_to_a, gen_packing_pdfs, _PACKING_FOOTER


class PackingPdfNamesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.addCleanup(self.temp.cleanup)
        self.output = Path(self.temp.name)
        self.record = {
            '合同号：': 'PO159521', '客款号': '6335745', '厂款号：': 'B54397',
            '客户色号': '01', '颜色': '黑色', '英文颜色': 'BLACK',
            '港口': 'ASOS UK', 'sizes_by_name': {'XS': 17, 'S': 50},
        }

    def test_all_three_codes_and_existing_pages(self):
        cream = dict(self.record, 客户色号='02', 颜色='米白色', 英文颜色='CREAM')
        gen_packing_pdfs([self.record, cream], self.output)
        files = list(self.output.glob('*.pdf'))
        self.assertEqual([p.name for p in files], ['PO159521 6335745 B54397 包装明细.pdf'])
        with fitz.open(files[0]) as pdf:
            self.assertEqual(len(pdf), 2)
            self.assertIn('PO159521', pdf[0].get_text())
            self.assertIn('B54397', pdf[0].get_text())

    def test_multiple_styles_and_set_components_are_deduplicated(self):
        first = dict(self.record, 厂款号2='T54396')
        second = dict(self.record, 客款号='6335746', 客户色号='02',
                      厂款号拆分='B54397|P37178|B54397')
        gen_packing_pdfs([first, first, second], self.output)
        self.assertEqual([p.name for p in self.output.glob('*.pdf')],
                         ['PO159521 6335745+6335746 B54397+T54396+P37178 包装明细.pdf'])

    def test_existing_workbook_route_and_windows_filename_characters(self):
        records = _format_b_to_a([{
            '合同号': 'PO/159521', '客款号': '633:5745', '厂款号': 'B*54397',
            '客户色号': '01', '港口': 'UK', 'size_headers': ['S'], 'sizes': [10],
        }])
        gen_packing_pdfs(records, self.output)
        self.assertEqual([p.name for p in self.output.glob('*.pdf')],
                         ['PO-159521 633-5745 B-54397 包装明细.pdf'])

    def test_missing_styles_do_not_add_none_or_empty_separators(self):
        self.record['客款号'] = ''
        self.record['厂款号：'] = None
        gen_packing_pdfs([self.record], self.output)
        self.assertEqual([p.name for p in self.output.glob('*.pdf')], ['PO159521 包装明细.pdf'])

    def test_summary_and_footer_fit_with_many_ports(self):
        for count in (7, 8, 11, 12, 25):
            with self.subTest(ports=count):
                gen_packing_pdfs([dict(self.record, 港口=f'PORT {i}') for i in range(count)], self.output)
                path = self.output / 'PO159521 6335745 B54397 包装明细.pdf'
                with fitz.open(path) as pdf:
                    self.assertEqual(len(pdf), (count + 10) // 11)
                    all_text = ''
                    for page in pdf:
                        text = page.get_text()
                        all_text += text
                        self.assertIn('汇总', text)
                        self.assertIn(_PACKING_FOOTER, text)
                        for block in page.get_text('dict')['blocks']:
                            for line in block.get('lines', []):
                                for span in line['spans']:
                                    self.assertTrue(page.rect.contains(fitz.Rect(span['bbox'])), span['text'])
                    for i in range(count):
                        self.assertIn(f'PORT {i}\n', all_text)


if __name__ == '__main__':
    unittest.main()
