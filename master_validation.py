"""Per-master, per-column lineage and conditional availability validation.

Existing grouping is learned from the approved CSV, not inferred from SKU
proximity. Available options are derived from real variation values only when
one reference-supported dependency rule can explain the existing option lists.
Sizing fields without scalar variant columns stay explicit policy inputs.
"""
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from itertools import combinations
import json
import re
from urllib.parse import urlparse

import pandas as pd


AVAILABLE = {
    'Available Metals': 'Metal',
    'Available Config field 11': 'Config field 11',
    'Available Center Carat Weight': 'Center Carat Weight',
    'Supported Shape Variations': 'Shape',
    'Available Config Field 6': 'Config Field 6',
}
POLICY_OPTIONS = {'Available Config Field 7': 'Length', 'Available Ring Size': 'Ring Size'}
DIMENSIONS = tuple(AVAILABLE.values())
NUMERIC = {'Config field 11', 'Center Carat Weight', 'Weight', 'Side Carat Weight',
           'Side Total Stones', 'Total Price', 'Config Field 17'}
SKU = re.compile(r'^(.*)-(14K|18K)([WYR]{1,3})$', re.I)
IMAGE_CODE = re.compile(r'_([WYR]{1,3})(\d*)\.[a-z0-9]+$', re.I)


def _txt(value):
    return '' if value is None else str(value).strip()


def _key(value, field):
    value = _txt(value)
    if field in NUMERIC:
        try:
            n = Decimal(value.replace(',', ''))
            if n.is_finite():
                return format(n.normalize(), 'f')
        except InvalidOperation:
            pass
    return re.sub(r'\s+', ' ', value).casefold()


def _parts(value, field):
    return {_key(v, field) for v in _txt(value).split('#') if _txt(v)}


def _sort(values, field):
    def order(v):
        if field in NUMERIC:
            try:
                return (0, Decimal(v.replace(',', '')))
            except InvalidOperation:
                return (1, v.casefold())
        if field == 'Metal':
            m = re.match(r'(\d+)K\s+(.*)', v, re.I)
            if m:
                return (int(m[1]), m[2].casefold())
        return (0, v.casefold())
    return sorted(values, key=order)


def _domain(rows, target, context, current):
    values = {}
    for row in rows:
        if all(_key(row.get(c), c) == _key(current.get(c), c) for c in context):
            value = _txt(row.get(target))
            if value:
                # Scalar source fields must not contain an option list.
                if '#' in value:
                    continue
                values.setdefault(_key(value, target), value)
    return values


def _infer_rule(rows, available, target):
    dimensions = [c for c in DIMENSIONS if c != target and any(_txt(r.get(c)) for r in rows)]
    declared = [r for r in rows if _txt(r.get(available))]
    if not declared:
        return (), 'derived_from_variants_no_declared_list'
    candidates = []
    for size in range(len(dimensions) + 1):
        for context in combinations(dimensions, size):
            if all(_parts(r[available], target) == set(_domain(rows, target, context, r)) for r in declared):
                candidates.append(context)
        if candidates:
            # Equivalent minimum rules can coincide on existing rows but differ
            # after adding/removing variants. Do not pick one arbitrarily.
            return (candidates[0], 'reference_supported') if len(candidates) == 1 else (None, 'ambiguous_dependencies')
    return None, 'declared_options_not_explained_by_variants'


def _lineage(column):
    if column in AVAILABLE:
        return f'Derive from {AVAILABLE[column]} using reference-supported master/context grouping.'
    if column in POLICY_OPTIONS:
        return 'Approved option policy; scalar ' + POLICY_OPTIONS[column] + ' column is absent. Validate list, do not invent options.'
    if column == 'Master Stock':
        return 'One approved reference group across all linked base styles; existing 14K variant in confirmed default color.'
    if column == 'Stock Number':
        return 'Preserve approved variant SKU; validate unique base-purity-color suffix.'
    if column in ('Total Price', 'Config Field 17'):
        return 'Client DC/SRP per base and purity; combined sheet wins; conflicting/missing prices block update.'
    if column == 'Default Metal (MASTER STOCK)':
        return 'Confirmed client default color applied once to the whole existing master group.'
    if column.startswith('Image Url '):
        return 'Approved image preserved; fill blank slots by exact base/color, including WY/WR/YR. Validate URL/color/duplicates.'
    if column in ('Config Field 13', 'Config Field 14', 'Jewelry Style', 'Tags'):
        return 'Approved reference taxonomy, with client collection label crosswalk; keep the four fields distinct.'
    if column in NUMERIC or column.startswith(('Center ', 'Side ')) or column in ('Shape', 'Comments'):
        return 'Approved stone assignment preserved; validate numeric/shape/quality consistency and audit source differences.'
    if column in ('Short Title', 'Description'):
        return 'Approved text preserved; validate embedded metal and CT weight when present.'
    if column == 'barcode No':
        return 'Preserve text, including multiple barcodes; flag scientific notation.'
    return 'Preserve approved reference value; record populated/blank/distinct values for each master.'


def validate_masters(reference, generated, source_items, repair=True):
    """Return repaired copy and comprehensive per-master reports.

    Repairs only supported availability fields. Unexplained dependency rules,
    missing source policies, reference conflicts and all other discrepancies
    remain visible. No new master grouping or size/length options are guessed.
    """
    reference = reference.fillna('').astype(str)
    result = generated.copy().fillna('').astype(str)
    reference = reference[reference['Stock Number'].str.strip().ne('')]
    reference_groups = {m: g for m, g in reference.groupby('Master Stock', sort=False)}
    result['Stock Number'] = result['Stock Number'].str.strip()
    sku_to_group = dict(zip(result['Stock Number'], result['Master Stock']))
    sku_positions = {sku: idx for idx, sku in result['Stock Number'].items()}
    source = defaultdict(list)
    for item in source_items:
        source[item['base'].upper()].append(item)
    trace, column_rows, availability, issues, repairs = [], [], [], [], []
    dictionary = pd.DataFrame([{'column': c, 'rule': _lineage(c)} for c in reference])

    def issue(severity, kind, master, sku='', column='', details=''):
        issues.append(dict(severity=severity, issue=kind, master_stock=master,
                           stock_number=sku, column=column, details=details))

    for old_master, ref_group in reference_groups.items():
        old_skus = set(ref_group['Stock Number'].str.strip())
        targets = {sku_to_group[s] for s in old_skus if s in sku_to_group}
        selected = result[result['Master Stock'].eq(next(iter(targets)))] if len(targets)==1 else result[result['Stock Number'].isin(old_skus)]
        bases = {m[1].upper() for s in old_skus if (m := SKU.fullmatch(s))}
        records = [i for b in sorted(bases) for i in source.get(b, [])]
        combined_bases = {(i['base'], i['purity']) for i in records if i.get('combined')}
        records = [i for i in records if i.get('combined') or (i['base'], i['purity']) not in combined_bases]
        defaults = sorted({i['default_color'] for i in records if i.get('default_color_raw')})
        eligible = sorted({s for s in selected['Stock Number']
                           if (m := SKU.fullmatch(s)) and m[2].upper() == '14K'
                           and len(defaults) == 1 and m[3].upper() == defaults[0]})
        selected_master = next(iter(targets)) if len(targets) == 1 else ''
        old_match = SKU.fullmatch(old_master)
        preferred = f'{old_match[1]}-14K{defaults[0]}' if old_match and len(defaults)==1 else ''
        ordered_candidates = [s for s in selected['Stock Number'] if s in eligible]
        expected_master = preferred if preferred in eligible else ordered_candidates[0] if ordered_candidates else old_master
        selection_rule = 'retain_existing_on_conflict_or_missing_variant' if len(defaults)!=1 or not eligible else 'same_base_14k_default_color' if preferred in eligible else 'first_existing_group_14k_default_color'
        if len(targets) > 1:
            issue('error', 'master_group_split', old_master, details=json.dumps(sorted(targets)))
        if selected_master and selected_master not in set(selected['Stock Number']):
            issue('error', 'master_outside_own_group', old_master, selected_master)
        if selected_master and len(defaults) == 1 and eligible and selected_master not in eligible:
            issue('error', 'incorrect_default_master', old_master, selected_master,
                  details='Expected one of ' + ', '.join(eligible))
        if selected_master and selected_master != expected_master:
            issue('error', 'master_selection_rule_mismatch', old_master, selected_master,
                  details='Expected ' + expected_master + ' using ' + selection_rule)
        trace.append(dict(reference_master=old_master, generated_master=selected_master,
                          expected_master=expected_master, selection_rule=selection_rule,
                          reference_rows=len(ref_group), reference_unique_skus=len(old_skus),
                          generated_rows=len(selected), base_style_count=len(bases),
                          linked_base_styles='#'.join(sorted(bases)),
                          client_default_colors='#'.join(defaults),
                          eligible_14k_masters='#'.join(eligible),
                          reference_master_present=old_master in old_skus,
                          excluded_or_missing_skus='#'.join(sorted(old_skus - set(selected['Stock Number']))),
                          source_rows=json.dumps(sorted({(str(i.get('sheet', '')), str(i.get('source_row', ''))) for i in records})),
                          decision='group_split_error' if len(targets)>1 else 'no_exported_rows' if selected.empty else
                                   'default_color_conflict_retained' if len(defaults)>1 else
                                   'default_variant_missing_retained' if defaults and not eligible else
                                   'updated_to_default_14k' if selected_master != old_master else 'existing_master_retained'))

        ref_records = ref_group.to_dict('records')
        out_records = selected.to_dict('records')
        for available, target in AVAILABLE.items():
            if available not in reference or target not in reference:
                continue
            context, reason = _infer_rule(ref_records, available, target)
            if context is None:
                issue('error', 'availability_dependency_unresolved', old_master, column=available, details=reason)
            aggregates = {}
            for row in out_records:
                expected_values = _domain(out_records, target, context or (), row)
                expected = '#'.join(_sort(expected_values.values(), target))
                current = row.get(available, '')
                actual_set = _parts(current, target)
                semantic_match = actual_set == set(expected_values)
                implicit_singleton = not current and len(expected_values) <= 1
                current_tokens = [_key(t, target) for t in current.split('#') if t.strip()]
                has_duplicates = len(current_tokens) != len(set(current_tokens))
                expected_order = [_key(v, target) for v in _sort(expected_values.values(), target)]
                order_matches = current_tokens == expected_order
                status = 'unresolved_rule' if context is None else 'implicit_singleton' if implicit_singleton else 'valid' if semantic_match and not has_duplicates and order_matches else 'rebuild'
                # Empty current scalar means this option is inapplicable on the row.
                if not _txt(row.get(target)):
                    status = 'not_applicable' if not current else status
                key = (current, expected, json.dumps({c: row.get(c, '') for c in context or ()}, sort_keys=True), status)
                if key not in aggregates:
                    aggregates[key] = {'rows': 0, 'sample_sku': row['Stock Number']}
                aggregates[key]['rows'] += 1
                if context is not None and status == 'rebuild' and repair:
                    result.at[sku_positions[row['Stock Number']], available] = expected
                    repairs.append(dict(master_stock=selected_master, stock_number=row['Stock Number'],
                                        field=available, previous=current, updated=expected,
                                        dependency_columns='#'.join(context), reason='observed_variations'))
                elif status == 'rebuild':
                    issue('error', 'available_options_mismatch', old_master, row['Stock Number'], available,
                          'Expected ' + expected + '; found ' + current)
            for (current, expected, context_values, status), info in aggregates.items():
                availability.append(dict(reference_master=old_master, generated_master=selected_master,
                    column=available, source_column=target, dependency_columns='#'.join(context or ()),
                    inference=reason, context_values=context_values, existing=current,
                    expected=expected, status=status, **info))

        # All business columns are audited, not only those with differences.
        for column in reference:
            values = list(dict.fromkeys(_txt(v) for v in ref_group[column]))
            output_values = list(dict.fromkeys(_txt(v) for v in selected[column])) if column in selected else []
            status = 'availability_validated' if column in AVAILABLE else 'policy_options_checked' if column in POLICY_OPTIONS else 'lineage_checked'
            column_rows.append(dict(reference_master=old_master, generated_master=selected_master,
                column=column, rule=_lineage(column), reference_rows=len(ref_group),
                reference_populated=int(ref_group[column].str.strip().ne('').sum()),
                reference_distinct=len(values), reference_values=json.dumps(values),
                generated_rows=len(selected), generated_populated=int(selected[column].str.strip().ne('').sum()) if column in selected else 0,
                generated_distinct=len(output_values), generated_values=json.dumps(output_values), status=status))
        for column, label in POLICY_OPTIONS.items():
            if column not in selected:
                continue
            policy_values = set(selected[column]) - {''}
            if policy_values:
                issue('warning', 'option_policy_not_in_source_variants', old_master, column=column,
                      details=f'No scalar {label} column; preserve the approved list. Client single-item size cannot prove the complete option range.')
            if len(policy_values)>1:
                issue('warning', 'option_policy_varies_within_master', old_master, column=column,
                      details=json.dumps(sorted(policy_values)))
            for text in policy_values:
                numbers = []
                for part in text.split('#'):
                    token = part.strip().replace('"','').replace('US','').strip()
                    try:
                        value = Decimal(token)
                        if not value.is_finite() or value <= 0:
                            raise InvalidOperation
                        numbers.append(value)
                        if column == 'Available Ring Size' and value * 4 != (value * 4).to_integral_value():
                            issue('error', 'invalid_ring_size_increment', old_master, column=column, details=part)
                    except InvalidOperation:
                        issue('error', 'invalid_option_policy_value', old_master, column=column, details=part)
                if numbers and (numbers != sorted(set(numbers))):
                    issue('warning', 'unsorted_or_duplicate_policy_options', old_master, column=column, details=text)

    # Validate each generated SKU's remaining columns as individual variations.
    for row in result.to_dict('records'):
        master, sku = row['Master Stock'], row['Stock Number']
        match = SKU.fullmatch(sku)
        if match:
            import app as core
            expected_metal = core._metal_name(match[2].upper(), match[3].upper())
            if _key(row.get('Metal'), 'Metal') != _key(expected_metal, 'Metal'):
                issue('error', 'sku_metal_mismatch', master, sku, 'Metal', expected_metal)
        for column in NUMERIC:
            value = _txt(row.get(column))
            if not value:
                continue
            try:
                n = Decimal(value.replace(',', ''))
                if not n.is_finite() or n < 0:
                    raise InvalidOperation
            except InvalidOperation:
                issue('error', 'invalid_numeric_value', master, sku, column, value)
        if '#' in _txt(row.get('Config field 11')):
            issue('error', 'multiple_values_in_scalar_field', master, sku, 'Config field 11', row['Config field 11'])
        for column in ('Short Title', 'Description'):
            text = row.get(column, '')
            m = re.search(r'(\d+(?:\.\d+)?)\s*CT\b', text, re.I)
            if m and _txt(row.get('Weight')):
                if _key(m[1], 'Weight') != _key(row['Weight'], 'Weight'):
                    issue('warning', 'title_weight_mismatch', master, sku, column,
                          f"Text CT={m[1]}; Weight={row['Weight']}")
        def number(column):
            try:
                n = Decimal(_txt(row.get(column)).replace(',', ''))
                return n if n.is_finite() else None
            except InvalidOperation:
                return None
        center, side, total = number('Center Carat Weight'), number('Side Carat Weight'), number('Weight')
        if total is not None and (center is not None or side is not None):
            if abs((center or Decimal(0)) + (side or Decimal(0)) - total) > Decimal('0.02'):
                issue('warning', 'stone_weights_do_not_sum_to_total', master, sku, 'Weight',
                      f"Center={center}; Side={side}; Total={total}")
        for prefix, weight in [('Center', center), ('Side', side)]:
            if weight and weight > 0:
                for suffix in ('Color','Clarity'):
                    column = prefix + ' ' + suffix
                    if not _txt(row.get(column)):
                        issue('warning','missing_stone_quality',master,sku,column,'Positive stone weight but missing quality.')
        tag_values = [_txt(t) for t in row.get('Tags','').split(',') if _txt(t)]
        if len(tag_values) != len(set(v.casefold() for v in tag_values)):
            issue('warning','duplicate_tags',master,sku,'Tags',row.get('Tags',''))
        collection = _txt(row.get('Config Field 13'))
        if collection and collection.casefold() not in {t.casefold() for t in tag_values}:
            issue('warning','collection_not_in_tags',master,sku,'Tags',collection)
        if match:
            source_rows = [i for i in source.get(match[1].upper(),[]) if i['purity'].upper()==match[2].upper()]
            source_types = {_txt(i.get('jtype')).casefold() for i in source_rows if _txt(i.get('jtype'))}
            related_types = {'ring': {'ring','band'}, 'band': {'ring','band'},
                             'bracelet': {'bracelet','bangle'}, 'bangle': {'bracelet','bangle'},
                             'necklace': {'necklace','pendant'}, 'pendant': {'necklace','pendant'}}
            allowed_types = set().union(*(related_types.get(t, {t}) for t in source_types)) if source_types else set()
            if len(source_types)==1 and row.get('Jewelry Sub-Category','').casefold() not in allowed_types:
                issue('warning','source_product_type_differs',master,sku,'Jewelry Sub-Category',
                      'Client=' + next(iter(source_types)) + '; reference=' + row.get('Jewelry Sub-Category',''))
            counts = [Decimal(p) for p in re.findall(r'(\d+(?:\.\d+)?)\s+Pcs\s*\|',row.get('Comments',''),re.I)]
            source_totals = {sum(Decimal(str(s['pcs'])) for s in i.get('centers',[])+i.get('sides',[])) for i in source_rows}
            if counts and len(source_totals)==1 and sum(counts)!=next(iter(source_totals)):
                issue('warning','comment_stone_count_differs_from_source',master,sku,'Comments',
                      f"Comment count={sum(counts)}; client count={next(iter(source_totals))}")
        if re.search(r'\d(?:\.\d+)?E[+-]?\d+', row.get('barcode No', ''), re.I):
            issue('warning', 'barcode_scientific_notation', master, sku, 'barcode No', row['barcode No'])
        urls = [row.get(f'Image Url {i}', '') for i in range(1,9)]
        nonblank = [u for u in urls if u]
        if len(nonblank) != len(set(nonblank)):
            issue('warning', 'duplicate_images_within_variant', master, sku, 'Image Url 1–8')
        for i, url in enumerate(urls, 1):
            if not url:
                continue
            parsed = urlparse(url)
            if parsed.scheme not in ('http','https') or not parsed.netloc:
                issue('error', 'invalid_image_url', master, sku, f'Image Url {i}', url)
            code = IMAGE_CODE.search(parsed.path)
            if code and match and code[1].upper() != match[3].upper():
                issue('warning', 'image_metal_color_mismatch', master, sku, f'Image Url {i}',
                      f"Image={code[1].upper()}; variant={match[3].upper()}")

    # Refresh output column profiles after supported availability repairs.
    repaired_groups = {m:g for m,g in result.groupby('Master Stock',sort=False)}
    for record in column_rows:
        group = repaired_groups.get(record['generated_master'])
        if group is not None and record['column'] in AVAILABLE:
            values=list(dict.fromkeys(group[record['column']]))
            record['generated_values']=json.dumps(values)
            record['generated_distinct']=len(values)
            record['generated_populated']=int(group[record['column']].str.strip().ne('').sum())
    reports = dict(master_trace=pd.DataFrame(trace), master_column_audit=pd.DataFrame(column_rows),
                   availability_audit=pd.DataFrame(availability), availability_repairs=pd.DataFrame(repairs),
                   master_validation_issues=pd.DataFrame(issues, columns=['severity','issue','master_stock','stock_number','column','details']),
                   column_rules=dictionary)
    reports['master_validation_summary'] = {
        'masters_checked':len(trace), 'columns_checked':len(reference.columns),
        'master_column_checks':len(column_rows), 'availability_rows_rebuilt':len(repairs),
        'blocking_master_validation_issues':sum(i['severity']=='error' for i in issues),
        'warning_master_validation_issues':sum(i['severity']=='warning' for i in issues),
        'size_length_policy': 'Preserved: complete option ranges cannot be inferred without scalar size/length variants.',
    }
    return result, reports
