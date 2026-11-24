import io
import unittest
from collections import defaultdict
from openpyxl import Workbook
import app
from collection_mappings import CLASSIFICATIONS, COLLECTION_GROUPS


class CollectionMappingTests(unittest.TestCase):
    def resolve(self, collection, style):
        return app.resolve_classification({'source_collection': collection, 'remark': style})

    def test_all_reference_collection_style_pairs(self):
        pairs = defaultdict(set)
        for collection, style, cf14 in CLASSIFICATIONS:
            pairs[collection, style].add(cf14)
        self.assertEqual(len(COLLECTION_GROUPS), 61)  # 60 canonical reference collections plus Brown
        for (collection, style), options in pairs.items():
            with self.subTest(collection=collection, style=style):
                c, s, cf14, tags, issue = self.resolve(collection, style)
                self.assertEqual((c, s), (collection, style))
                if len(options) == 1:
                    self.assertEqual(cf14, next(iter(options)))
                    self.assertFalse(issue)
                    self.assertTrue(tags)
                else:
                    self.assertEqual(cf14, '')
                    self.assertIn('Ambiguous', issue)

    def test_fashion_and_men_keep_distinct_fields(self):
        self.assertEqual(self.resolve('Flora', 'Flora')[:4], ('Flora', 'Flora', 'Fashion', 'Fashion,Flora'))
        self.assertEqual(self.resolve('Zulu', 'Zulu')[:4], ('Zulu', 'Zulu', 'Men', 'Zulu,Men'))
        self.assertEqual(self.resolve('Frame', 'Frame')[:4], ('Frame', 'Frame', 'Frame', 'Fashion,Frame'))

    def test_main_group_disambiguates_named_style(self):
        result = app.resolve_classification({'remark': 'Pave Petite'}, 'Bridals')
        self.assertEqual(result[:3], ('Smiling Brides', 'Pave Petite', 'Pave Petite'))
        result = app.resolve_classification({'remark': 'Pave Petite'}, 'Fashion Collections')
        self.assertEqual(result[:3], ('Pave Petite', 'Pave Petite', 'Fashion'))
        result = app.resolve_classification({'remark': 'FLORA'}, 'Fashion Collections')
        self.assertEqual(result[:3], ('Flora', 'Flora', 'Fashion'))
        result = app.resolve_classification({'remark': 'ZULU'}, "Men's")
        self.assertEqual(result[:3], ('Zulu', 'Zulu', 'Men'))

    def test_unknown_requires_review_without_assigned_group(self):
        result = app.resolve_classification({'remark': 'NEW COLLECTION'})
        self.assertEqual(result[2], '')
        self.assertIn('Unmapped', result[-1])
        result = app.resolve_classification({'remark': 'NEW COLLECTION'}, 'Fashion Collections')
        self.assertEqual(result[:3], ('NEW COLLECTION', 'NEW COLLECTION', 'Fashion'))

    def test_legacy_essential_alias_and_singular_bridal(self):
        result = app.resolve_classification({'remark': 'SOLITAIRE STUDS'}, 'Essentials')
        self.assertEqual(result[:3], ('Essential', 'Solitaire Stud', 'Solitaire Stud'))
        self.assertEqual(self.resolve('Fashion Bridal', 'Bridal Double')[:3],
                         ('Fashion Bridals', 'Bridal Double', 'Bridal Double'))

    def test_conversion_retains_separate_collection_and_remark(self):
        wb = Workbook()
        ws = wb.active
        ws.title = '14K'
        ws.append(['No.', 'Style No', 'Description', 'D. Pcs', 'Gm/Cts', 'DC', 'SRP', 'collection', 'Remark'])
        ws.append([1, 'SRR-99999WHT', 'RING', None, None, 100, 250, 'Smiling Brides', 'Pave Petite'])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        df, issues = app.convert([buf])
        self.assertFalse(issues)
        self.assertEqual(set(df['Config Field 13']), {'Smiling Brides'})
        self.assertEqual(set(df['Jewelry Style']), {'Pave Petite'})
        self.assertEqual(set(df['Config Field 14']), {'Pave Petite'})
        self.assertEqual(set(df['Tags']), {'Bridal,Smiling Brides,Pave Petite'})


if __name__ == '__main__':
    unittest.main()
