import io
import unittest
from copy import deepcopy
from unittest.mock import patch
import pandas as pd
from openpyxl import Workbook
import app
from workbook_inventory import build_from_items
from test_inventory_reconcile import item,row


def confirmed(base='SRR-90001WHT',purity='14K',members=None,columns='Metal',**changes):
    value=item(base=base,purity=purity,**changes)
    value['instructions_present']=True
    value['instructions']=dict(collection='Fashion Collections',style='Flora',variation='Yes',
                               styles=','.join(members or [base]),columns=columns)
    return value


def rules(rows):
    return {'profiles':{r['Stock Number']:[r] for r in rows},'source_labels':{}}


class WorkbookOnlyTests(unittest.TestCase):
    def test_file_header_reads_new_collection_not_old_client_label(self):
        w=Workbook();s=w.active;s.title='14K'
        s.append(['No.','Style No','Description','D. Pcs','Gm/Cts','DC','SRP',
                  'BY DEFAULT METAL COLOR','COLLECTION NAME',None,
                  'COLLECTION NAME','Style Name','Variations (Yes / No)','Variation Styles','Variation Columns','Available Metals'])
        s.append([1,'SRR-90001WHT','RING',None,None,100,250,'WG','OLD LABEL',None,
                  'Fashion Collections','Flora','Yes','SRR-90001WHT','Metal','White Yellow, White Rose'])
        b=io.BytesIO();w.save(b);b.seek(0)
        parsed=app.extract(b)
        self.assertEqual(parsed[0]['source_collection'],'OLD LABEL')
        self.assertEqual(parsed[0]['instructions']['collection'],'Fashion Collections')
        self.assertEqual(parsed[0]['instructions']['style'],'Flora')
        self.assertTrue(parsed[0]['available_metals_present'])
        self.assertEqual(parsed[0]['instructions']['available_metals'],'White Yellow, White Rose')

    def test_only_workbook_and_images_create_prices_and_14k_master(self):
        base='SRR-90001WHT'
        rs=[row(purity=p,color=c) for p in ['14K','18K'] for c in ['W','Y']]
        source=[confirmed(),confirmed(purity='18K',dc=130,srp=325)]
        result,report=build_from_items(source,{(base,'W'):['W.jpg'],(base,'Y'):['Y.jpg']},rules=rules(rs))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(len(result),4)
        self.assertEqual(set(result['Master Stock']),{base+'-14KW'})
        self.assertEqual(set(result[result['Stock Number'].str.contains('18K')]['Total Price']),{'130'})
        self.assertTrue(result['Available Metals'].str.contains('18K Yellow Gold').all())

    def test_shared_styles_create_one_master_with_parent_child_options(self):
        a='SRR-90001WHT';b='SRR-90002WHT';rs=[]
        for base,shape in [(a,'Round'),(b,'Oval')]:
            for color in ['W','Y']:
                r=row(base=base,color=color,master=a+'-14KW');r['Shape']=shape;rs.append(r)
        source=[confirmed(base=base,members=[a,b],columns='Metal, Shape') for base in [a,b]]
        result,report=build_from_items(source,rules=rules(rs))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(result['Master Stock'].nunique(),1)
        self.assertEqual(set(result['Supported Shape Variations']),{'Oval#Round'})
        self.assertEqual(report['summary']['masters_created'],1)

    def test_confirmed_group_can_merge_old_masters_in_declared_order(self):
        a='SRR-90001WHT';b='SRR-90002WHT'
        source=[confirmed(base=a,members=[b,a]),confirmed(base=b,members=[b,a])]
        result,report=build_from_items(source,rules=rules([row(),row(base=b)]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(set(result['Master Stock']),{b+'-14KW'})

    def test_missing_group_member_is_deferred(self):
        a='SRR-90001WHT';b='SRR-90002WHT'
        result,report=build_from_items([confirmed(members=[a,b])],rules=rules([row()]))
        self.assertEqual(len(result),1)
        self.assertTrue(report['summary']['publish_ready'])
        self.assertIn('group_members_deferred',set(report['issues']['issue']))

    def test_missing_confirmation_fields_not_silently_inferred(self):
        source=confirmed();source['instructions']['columns']=''
        result,report=build_from_items([source],rules=rules([row()]))
        self.assertTrue(result.empty)
        self.assertFalse(report['summary']['publish_ready'])

    def test_each_group_style_may_have_its_own_default(self):
        a='SRR-90001WHT';b='SRR-90002WHT'
        source=[confirmed(base=a,members=[a,b]),confirmed(base=b,members=[a,b],color='Y')]
        result,report=build_from_items(source,rules=rules([row(),row(base=b,color='Y')]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(set(result['Master Stock']),{a+'-14KW'})

    def test_combined_prices_have_priority(self):
        combined=confirmed(dc=140,srp=350,combined=True)
        result,report=build_from_items([confirmed(),combined],rules=rules([row()]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(set(result['Total Price']),{'140'})

    def test_auto_route_needs_no_reference_csv(self):
        with patch.object(app,'extract',return_value=[confirmed()]),patch('saved_inventory_rules.load_rules',return_value=rules([row()])):
            result,report,unknown=app.convert_with_reports(['client.xlsx'])
        self.assertEqual(report['summary']['mode'],'confirmation_workbook')
        self.assertEqual(len(result),1)
        self.assertFalse(unknown)

    def test_no_saved_prices_or_images_are_used(self):
        r=row();r.pop('Total Price');r.pop('Config Field 17');r.pop('Image Url 1')
        result,report=build_from_items([confirmed(dc=117,srp=299)],rules=rules([r]))
        self.assertEqual(result.iloc[0]['Total Price'],'117')
        self.assertEqual(result.iloc[0]['Config Field 17'],'299')
        self.assertEqual(result.iloc[0]['Image Url 1'],'')


if __name__=='__main__':unittest.main()

class PrivateRuleStorageTests(unittest.TestCase):
    def test_rules_are_read_from_hidden_workbook_sheet(self):
        import base64,json,zlib
        from saved_inventory_rules import load_rules
        expected=rules([row()])
        w=Workbook();w.active.title='14K'
        hidden=w.create_sheet('_VDB_Rules');hidden.sheet_state='hidden'
        hidden['A1']='VDB_RULES_V1'
        hidden['A4']=base64.b64encode(zlib.compress(json.dumps(expected).encode())).decode()
        source=io.BytesIO();w.save(source);source.seek(0)
        self.assertEqual(load_rules([source]),expected)


class ClientColorTests(unittest.TestCase):
    def test_color_names_and_two_tone(self):
        from metal_colors import parse_colors,parse_color
        self.assertEqual(parse_colors('white, yellow, rose, white yellow, white rose'),('W','Y','R','WY','WR'))
        self.assertEqual(parse_color('White/Yellow Gold Two Tone'),'WY')
        self.assertEqual(parse_color('yellow white'),'WY')
        for invalid in ['YY','Yellow Yellow','Blue']:
            with self.assertRaises(ValueError):parse_color(invalid)

    def explicit(self,colors='white, yellow, rose',**changes):
        value=confirmed(**changes)
        value['available_metals_present']=True
        value['instructions']['available_metals']=colors
        return value

    def test_explicit_colors_create_both_karats_and_prices(self):
        base='SRR-90001WHT'
        source=[self.explicit(),self.explicit(purity='18K',dc=130,srp=325)]
        result,report=build_from_items(source,{(base,'WY'):['extra.jpg']},rules=rules([row()]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(len(result),6)
        self.assertEqual(set(result['Master Stock']),{base+'-14KW'})
        self.assertFalse(result['Stock Number'].str.endswith('WY').any())
        self.assertEqual(set(result[result['Stock Number'].str.contains('18K')]['Config Field 17']),{'325'})
        self.assertTrue(result['Image Url 1'].eq('').all())

    def test_two_tone_is_one_choice(self):
        source=[self.explicit('White Yellow, White Rose',color='WY'),self.explicit('White Yellow, White Rose',purity='18K',color='WY')]
        result,report=build_from_items(source,rules=rules([row(color='WY')]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(set(result['Metal']),{'14K White Yellow Gold','18K White Yellow Gold','14K White Rose Gold','18K White Rose Gold'})

    def test_explicit_list_requires_both_prices_and_default(self):
        for source,kind in [([self.explicit()],'missing_karat_prices'),([self.explicit('Yellow'),self.explicit('Yellow',purity='18K')],'default_color_not_available')]:
            _,report=build_from_items(source,rules=rules([row()]))
            self.assertFalse(report['summary']['publish_ready'])
            self.assertIn(kind,set(report['issues']['issue']))

    def test_group_uses_first_listed_styles_own_default(self):
        a='SRR-90001WHT';b='SRR-90002WHT'
        source=[self.explicit(base=base,purity=purity,color=color,members=[b,a]) for base,color in [(a,'W'),(b,'Y')] for purity in ['14K','18K']]
        result,report=build_from_items(source,rules=rules([row(),row(base=b,color='Y')]))
        self.assertTrue(report['summary']['publish_ready'])
        self.assertEqual(set(result['Master Stock']),{b+'-14KY'})

    def test_image_code_reversed_tone_and_repeated_color(self):
        mapping,stats=app.load_image_map(io.StringIO('SRR-90001WHT_YW1.jpg\nSRR-90001WHT_YY1.jpg\n'))
        self.assertIn(('SRR-90001WHT','WY'),mapping)
        self.assertEqual(stats['accepted'],1)
        self.assertEqual(stats['skipped'],1)

    def test_present_members_with_conflicting_lists_are_blocked(self):
        a='SRR-90001WHT';b='SRR-90002WHT'
        _,report=build_from_items([confirmed(members=[a,b]),confirmed(base=b)],rules=rules([row(),row(base=b)]))
        self.assertFalse(report['summary']['publish_ready'])
        self.assertIn('conflicting_group_membership',set(report['issues']['issue']))
