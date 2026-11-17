"""Run with python -m unittest discover -s tests."""
import io
import unittest

from openpyxl import Workbook
import app


def quotation(sheet, prices, default='YG', variation='YES'):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    headers = ['No.', 'Style No', 'Description', 'D. Pcs', 'Gm/Cts']
    headers += list(prices) + ['DEFAULT METAL COLOR', 'VARIATION']
    ws.append(headers)
    ws.append([1, 'SRR-00001BRN', 'RING', None, None] + list(prices.values()) + [default, variation])
    ws.append([None, None, 'RND - WHSR', 2, 0.5])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class ConverterTests(unittest.TestCase):
    def combined(self, **kwargs):
        return quotation('14K-18K', {'14KDC': 100, '18KDC': 200,
                                   '14KSRP': 300, '18KSRP': 400}, **kwargs)

    def test_combined_prices_and_master(self):
        df, _ = app.convert([self.combined()])
        self.assertEqual(len(df), 6)
        self.assertEqual(set(df['Master Stock']), {'SRR-00001BRN-14KY'})
        for purity, dc, srp in [('14K', '100', '300'), ('18K', '200', '400')]:
            rows = df[df['Stock Number'].str.contains(purity)]
            self.assertEqual(set(rows['Total Price']), {dc})
            self.assertEqual(set(rows['Config Field 17']), {srp})

    def test_combined_wins_in_either_upload_order(self):
        for combined_first in (True, False):
            inputs = [self.combined(), quotation('14K', {'DC': 1, 'SRP': 2}),
                      quotation('18K', {'DC': 3, 'SRP': 4})]
            if not combined_first:
                inputs.reverse()
            df, _ = app.convert(inputs)
            self.assertEqual(set(df['Total Price']), {'100', '200'})
            self.assertEqual(len(df), 6)

    def test_two_tone_and_no_variations(self):
        df, _ = app.convert([self.combined(default='WY', variation='NO')])
        self.assertEqual(list(df['Stock Number']), ['SRR-00001BRN-14KWY', 'SRR-00001BRN-18KWY'])
        self.assertEqual(set(df['Master Stock']), {'SRR-00001BRN-14KWY'})
        self.assertTrue(df['Metal'].str.contains('White Yellow Gold').all())
        self.assertEqual(df.iloc[0]['Available Metals'], '#'.join(df['Metal']))
        df, _ = app.convert([self.combined(default='RG', variation='NO')])
        self.assertEqual(len(df), 2)
        self.assertTrue(df['Stock Number'].str.endswith('R').all())

    def test_legacy_separate_sheets(self):
        df, _ = app.convert([quotation('14', {'DC': 10, 'SRP': 30}, default=''),
                             quotation('18', {'DC': 20, 'SRP': 40}, default='')])
        self.assertEqual(len(df), 6)
        self.assertEqual(set(df['Master Stock']), {'SRR-00001BRN-14KW'})
        self.assertEqual(set(df['Total Price']), {'10', '20'})

    def test_two_tone_image_codes(self):
        images, _ = app.load_image_map(io.StringIO('filename\nSRR-00001BRN_WY1.jpg\n'))
        df, _ = app.convert([self.combined(default='WY')], images)
        self.assertTrue(df['Image Url 1'].str.endswith('_WY1.jpg').all())

    def test_unknown_two_tone_requires_colors(self):
        with self.assertRaisesRegex(ValueError, 'DEFAULT METAL COLOR'):
            app.convert([self.combined(default='Two Tone')])


if __name__ == '__main__':
    unittest.main()
