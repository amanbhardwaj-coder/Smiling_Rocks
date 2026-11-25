import io
import unittest
import zipfile
from copy import deepcopy

import pandas as pd
from openpyxl import Workbook

import app
from inventory_reconcile import reconcile, audit_zip


def item(base='SRR-90001WHT', purity='14K', dc=100, srp=250, color='W', **changes):
    value = dict(base=base, purity=purity, dc=dc, srp=srp, default_color=color,
                 default_color_raw=color, source_collection='Flora', remark='Flora',
                 sheet=purity, source_row=2, combined=False, variation='',
                 price_raw={'dc': str(dc), 'srp': str(srp)}, total_dia=0.5,
                 jtype='Ring', centers=[], sides=[], source_errors=[], unrecognized_stones=[])
    value.update(changes)
    return value


def row(base='SRR-90001WHT', purity='14K', color='W', master=None):
    value = {c: '' for c in app.COLUMNS}
    value.update({'Stock Number': f'{base}-{purity}{color}', 'Master Stock': master or f'{base}-14KW',
                  'Metal': app._metal_name(purity, color), 'Jewelry Sub-Category': 'Ring',
                  'Jewelry Style': 'Flora', 'Config Field 13': 'Flora', 'Config Field 14': 'Fashion',
                  'Tags': 'Fashion,Flora', 'Total Price': '100.00', 'Config Field 17': '250.00',
                  'Weight': '0.50', 'Image Url 1': 'https://example.com/curated.jpg',
                  'Short Title': 'Approved title', 'Description': 'Approved description',
                  'Available Metals': app._metal_name(purity, color),
                  'Center Carat Weight': '0.30', 'Side Carat Weight': '0.20'})
    return value


class ReconciliationTests(unittest.TestCase):
    def test_identical_source_reproduces_every_curated_column(self):
        ref = pd.DataFrame([row()])
        result, report = reconcile([item()], ref)
        pd.testing.assert_frame_equal(result.reset_index(drop=True), ref.reset_index(drop=True))
        self.assertTrue(report['summary']['publish_ready'])

    def test_price_updates_do_not_overwrite_stones_or_styles(self):
        ref = pd.DataFrame([row()])
        result, report = reconcile([item(dc=110, srp=275, total_dia=0.6)], ref)
        self.assertEqual(result.iloc[0]['Total Price'], '110')
        self.assertEqual(result.iloc[0]['Config Field 17'], '275')
        self.assertEqual(result.iloc[0]['Center Carat Weight'], '0.30')
        self.assertEqual(result.iloc[0]['Description'], 'Approved description')
        self.assertIn('Weight', set(report['differences']['field']))

    def test_conflicting_reference_skus_are_not_arbitrarily_chosen(self):
        a, b = row(), row()
        b['Master Stock'] = 'SRR-OTHER-14KW'
        result, report = reconcile([item()], pd.DataFrame([a, b]))
        self.assertTrue(result.empty)
        self.assertEqual(report['summary']['conflicting_reference_skus_excluded'], 1)
        self.assertFalse(report['summary']['publish_ready'])

    def test_conflicting_source_prices_keep_reference_for_review(self):
        result, report = reconcile([item(), item(dc=120)], pd.DataFrame([row()]))
        self.assertEqual(result.iloc[0]['Total Price'], '100.00')
        self.assertIn('price_update_blocked', set(report['issues']['issue']))
        self.assertFalse(report['summary']['publish_ready'])

    def test_combined_priority_is_independent_of_input_order(self):
        combined = item(dc=110, srp=275, combined=True,
                        price_raw={'14kdc': '110', '14ksrp': '275'})
        for records in ([item(), combined], [combined, item()]):
            result, report = reconcile(records, pd.DataFrame([row()]))
            self.assertEqual(result.iloc[0]['Total Price'], '110')
            self.assertNotIn('conflicting_source_prices', set(report['issues']['issue']))

    def test_equivalent_numeric_prices_do_not_create_conflicts(self):
        result, report = reconcile([item(), item(dc='100.00', srp='250.0')], pd.DataFrame([row()]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(result.iloc[0]['Total Price'], '100.00')

    def test_shared_master_stays_one_group(self):
        master = 'SRR-90001WHT-14KW'
        ref = pd.DataFrame([row(color='Y', master=master),
                            row(base='SRR-90002WHT', color='Y', master=master)])
        result, report = reconcile([item(color='Y'), item(base='SRR-90002WHT', color='Y')], ref)
        self.assertEqual(set(result['Master Stock']), {'SRR-90001WHT-14KY'})

    def test_default_not_available_keeps_existing_master(self):
        result, report = reconcile([item(color='Y')], pd.DataFrame([row()]))
        self.assertEqual(result.iloc[0]['Master Stock'], 'SRR-90001WHT-14KW')
        self.assertIn('default_master_variant_missing', set(report['issues']['issue']))

    def test_absence_is_not_deletion_and_zero_is_not_a_price(self):
        result, report = reconcile([], pd.DataFrame([row()]))
        self.assertEqual(len(result), 1)
        self.assertFalse(report['summary']['publish_ready'])
        result, report = reconcile([item(dc=0, price_raw={'dc': '', 'srp': '250'})], pd.DataFrame([row()]))
        self.assertEqual(result.iloc[0]['Total Price'], '100.00')
        self.assertFalse(report['summary']['publish_ready'])

    def test_new_styles_stay_outside_live_export(self):
        result, report = reconcile([item(), item(base='SRR-NEW001WHT')], pd.DataFrame([row()]))
        self.assertEqual(len(result), 1)
        self.assertEqual(report['summary']['new_client_base_styles'], 1)
        self.assertEqual(len(report['new_items']), 1)
        archive = zipfile.ZipFile(io.BytesIO(audit_zip(result, report)))
        self.assertIn('summary.json', archive.namelist())
        self.assertIn('new_items.csv', archive.namelist())

    def test_two_tone_image_fills_blank_exact_color_only(self):
        refrow = row(color='WY')
        refrow['Image Url 1'] = ''
        images = {('SRR-90001WHT', 'WY'): ['https://example.com/WY.jpg'],
                  ('SRR-90001WHT', 'R'): ['https://example.com/R.jpg']}
        result, _ = reconcile([item(color='WY')], pd.DataFrame([refrow]), images)
        self.assertEqual(result.iloc[0]['Image Url 1'], 'https://example.com/WY.jpg')
        self.assertNotIn('https://example.com/R.jpg', list(result.iloc[0]))

    def test_latest_headers_and_wrapped_zip(self):
        wb = Workbook()
        ws = wb.active
        ws.title = '14K'
        ws.append(['No.', 'Style No', 'Description', 'D. Pcs', 'Gm/Cts', 'DC', 'SRP',
                   'BY DEFAULT METAL COLOR', 'COLLECTION NAME'])
        ws.append([1, None, '14K - YG', None, None, 100, 250, '14K-YG', 'FLORA'])
        ws.append([None, 'RING', 'RND - WHSR', 2, 0.5])
        ws.append([None, 'SRR-90001WHT'])
        raw = io.BytesIO()
        wb.save(raw)
        wrapped = io.BytesIO()
        with zipfile.ZipFile(wrapped, 'w') as z:
            z.writestr('inventory.xlsx', raw.getvalue())
        wrapped.seek(0)
        parsed = app.extract(wrapped)
        self.assertEqual(parsed[0]['source_collection'], 'FLORA')
        self.assertEqual(parsed[0]['default_color'], 'Y')
        self.assertEqual(parsed[0]['source_row'], 2)


if __name__ == '__main__':
    unittest.main()
