import unittest
from copy import deepcopy

import pandas as pd

from master_validation import validate_masters
from test_inventory_reconcile import row, item


class MasterValidationTests(unittest.TestCase):
    def group(self):
        white = row()
        yellow = row(color='Y')
        for r in [white, yellow]:
            r['Available Metals'] = '14K White Gold#14K Yellow Gold'
        return white, yellow

    def test_every_column_for_every_master_is_profiled(self):
        a, b = self.group()
        ref = pd.DataFrame([a,b])
        _, reports = validate_masters(ref, ref, [item()])
        self.assertEqual(reports['master_validation_summary']['master_column_checks'],len(ref.columns))
        self.assertEqual(set(reports['master_column_audit']['column']),set(ref.columns))
        self.assertEqual(reports['master_trace'].iloc[0]['base_style_count'],1)

    def test_available_metals_reflect_exported_members(self):
        a,b = self.group()
        result, reports = validate_masters(pd.DataFrame([a,b]),pd.DataFrame([a]),[item()])
        self.assertEqual(result.iloc[0]['Available Metals'],'14K White Gold')
        self.assertEqual(reports['availability_repairs'].iloc[0]['field'],'Available Metals')

    def test_conditional_metals_use_shape_not_the_whole_master(self):
        a,b = self.group()
        oval = row(base='SRR-90002WHT',master=a['Master Stock'])
        a['Shape'] = b['Shape'] = 'Round'
        oval['Shape']='Oval'
        oval['Available Metals']='14K White Gold'
        new_oval = deepcopy(oval)
        new_oval['Stock Number']='SRR-90002WHT-14KY'
        new_oval['Metal']='14K Yellow Gold'
        ref=pd.DataFrame([a,b,oval])
        out=pd.DataFrame([a,b,oval,new_oval])
        result,reports=validate_masters(ref,out,[item()])
        records=reports['availability_audit']
        metals=records[records['column']=='Available Metals']
        self.assertEqual(set(metals['dependency_columns']),{'Shape'})
        self.assertTrue(result['Available Metals'].eq('14K White Gold#14K Yellow Gold').all())

    def test_ambiguous_dependency_is_flagged_not_guessed(self):
        a,b = self.group()
        a['Shape']='Round';b['Shape']='Oval'
        a['Center Carat Weight']='0.3';b['Center Carat Weight']='0.5'
        a['Available Metals']='14K White Gold';b['Available Metals']='14K Yellow Gold'
        ref=pd.DataFrame([a,b])
        result,reports=validate_masters(ref,ref,[item()])
        self.assertEqual(list(result['Available Metals']),list(ref['Available Metals']))
        self.assertIn('availability_dependency_unresolved',set(reports['master_validation_issues']['issue']))

    def test_numeric_options_are_sorted_and_formats_are_equivalent(self):
        a,b=self.group()
        a['Center Carat Weight']='0.30';b['Center Carat Weight']='0.5'
        a['Available Center Carat Weight']=b['Available Center Carat Weight']='0.50#0.3'
        ref=pd.DataFrame([a,b])
        result,_=validate_masters(ref,ref,[item()])
        self.assertEqual(set(result['Available Center Carat Weight']),{'0.30#0.5'})

    def test_ring_and_length_policies_are_not_invented(self):
        a=row();a['Available Ring Size']='3#3.5#11.60';a['Available Config Field 7']='16"#18"'
        ref=pd.DataFrame([a])
        result,reports=validate_masters(ref,ref,[item()])
        self.assertEqual(result.iloc[0]['Available Ring Size'],a['Available Ring Size'])
        self.assertIn('invalid_ring_size_increment',set(reports['master_validation_issues']['issue']))
        self.assertIn('option_policy_not_in_source_variants',set(reports['master_validation_issues']['issue']))

    def test_shared_styles_are_not_split_into_new_masters(self):
        a=row();b=row(base='SRR-90002WHT',master=a['Master Stock'])
        ref=pd.DataFrame([a,b]);out=ref.copy();out.loc[1,'Master Stock']=out.loc[1,'Stock Number']
        _,reports=validate_masters(ref,out,[item(),item(base='SRR-90002WHT')])
        self.assertIn('master_group_split',set(reports['master_validation_issues']['issue']))


if __name__=='__main__':
    unittest.main()
