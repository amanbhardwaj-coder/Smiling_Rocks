"""Reference-based full inventory conversion with an explicit reconciliation audit.

The approved VDB CSV is the source of curated master grouping, variation rows,
classification, center/side assignment, titles, images and available options.
Client data supplies per-purity DC/SRP and default metal. It cannot recreate
curated values absent from the quotation. Conflicts are reported, never resolved
by upload order or by choosing whichever price is closest to the reference.
"""
from collections import defaultdict
from decimal import Decimal, InvalidOperation
import io
import json
import re
import zipfile

import pandas as pd


SKU_RE = re.compile(r'^(.*)-(14K|18K)([WYR]{1,3})$', re.I)


def _number(value):
    try:
        number = Decimal(str(value).replace(',', '').strip())
        return number if number.is_finite() else None
    except (InvalidOperation, ValueError):
        return None


def _text(value):
    return '' if value is None else str(value).strip()


def load_reference(source):
    if isinstance(source, pd.DataFrame):
        df = source.copy().fillna('')
    else:
        df = pd.read_csv(source, dtype=str, keep_default_na=False)
    required = {'Stock Number', 'Master Stock', 'Metal', 'Total Price', 'Config Field 17',
                'Config Field 13', 'Config Field 14', 'Jewelry Style'}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError('Reference inventory is missing columns: ' + ', '.join(missing))
    # Ignore spreadsheet padding columns. Preserve every named business column.
    df = df.loc[:, [c for c in df if not c.startswith('Unnamed:')]].astype(str)
    return df


def reconcile(items, reference, image_map=None, diagnostics=None):
    import app as core
    ref = load_reference(reference)
    image_map = image_map or {}
    events, differences, source_rows, new_items = [], [], [], []
    invalid = ref['Stock Number'].str.strip().eq('')
    if invalid.any():
        events.append(dict(severity='error', issue='empty_reference_sku', stock_number='',
                           details=f'{int(invalid.sum())} reference rows have no Stock Number.'))
    ref = ref.loc[~invalid].copy()
    ref['Stock Number'] = ref['Stock Number'].str.strip()
    ref['Master Stock'] = ref['Master Stock'].str.strip()

    def event(severity, issue, sku='', details='', **extra):
        events.append(dict(severity=severity, issue=issue, stock_number=sku, details=details, **extra))

    for diagnostic in diagnostics or []:
        entry = dict(diagnostic)
        event(entry.pop('severity', 'error'), entry.pop('issue', 'parse_error'), **entry)

    # Identical business rows can be deduplicated. A conflicting SKU has no
    # safe representative; exclude all its versions from the candidate export.
    unique = ref.drop_duplicates()
    exact_removed = len(ref) - len(unique)
    if exact_removed:
        event('info', 'identical_reference_duplicates_removed', details=str(exact_removed))
    conflict_mask = unique.duplicated('Stock Number', keep=False)
    conflict_rows = unique.loc[conflict_mask]
    for sku, group in conflict_rows.groupby('Stock Number', sort=False):
        differing = [c for c in group if group[c].nunique() > 1]
        event('error', 'conflicting_reference_sku', sku, 'Conflicting columns: ' + ', '.join(differing))
    clean = unique.loc[~conflict_mask].copy()

    # Prefer combined pricing, but retain conflicting duplicates within the
    # preferred source. Do not silently overwrite one quotation with another.
    source = defaultdict(list)
    for item in items:
        source[item['base'].upper(), item['purity'].upper()].append(item)
    for key, records in list(source.items()):
        combined = [i for i in records if i.get('combined')]
        if combined:
            source[key] = combined
    client_bases = {base for base, _ in source}
    ref_bases = set()
    for sku in ref['Stock Number']:
        match = SKU_RE.fullmatch(sku)
        if match:
            ref_bases.add(match[1].upper())

    for (base, purity), records in source.items():
        prices = {(_number(i['dc']), _number(i['srp'])) for i in records}
        defaults = {i.get('default_color', '') for i in records if i.get('default_color_raw')}
        if len(prices) > 1:
            event('error', 'conflicting_source_prices', f'{base}-{purity}',
                  json.dumps([{'sheet': i.get('sheet'), 'row': i.get('source_row'),
                               'dc': i['dc'], 'srp': i['srp'], 'collection': i.get('source_collection')}
                              for i in records]))
        if len(defaults) > 1:
            event('error', 'conflicting_source_default_colors', f'{base}-{purity}', ', '.join(sorted(defaults)))
        for i in records:
            source_rows.append(dict(base=base, purity=purity, sheet=i.get('sheet'), row=i.get('source_row'),
                                    collection=i.get('source_collection'), default_color=i.get('default_color_raw'),
                                    dc=i['dc'], srp=i['srp'], matched_reference=base in ref_bases))
            for error in i.get('source_errors', []):
                # Reference classification safely covers invalid source collection cells.
                event('warning' if error.startswith('collection:') and base in ref_bases else 'error',
                      'source_excel_error', f'{base}-{purity}', error,
                      sheet=i.get('sheet'), source_row=i.get('source_row'))
            for stone in i.get('unrecognized_stones', []):
                event('warning', 'unrecognized_source_stone', f'{base}-{purity}', json.dumps(stone),
                      sheet=i.get('sheet'), source_row=i.get('source_row'))

    output = []
    group_sources = defaultdict(list)
    for row in clean.to_dict('records'):
        sku = row['Stock Number']
        match = SKU_RE.fullmatch(sku)
        if not match:
            event('error', 'unrecognized_reference_sku', sku, 'Expected base-14K/18K-metal suffix.')
            output.append(row)
            continue
        base, purity, color = match[1].upper(), match[2].upper(), match[3].upper()
        records = source.get((base, purity), [])
        if not records:
            event('error', 'reference_variant_missing_in_source', sku,
                  'Reference row retained; source absence is not treated as deletion.')
        else:
            group_sources[row['Master Stock']].extend(records)
            prices = {(_number(i['dc']), _number(i['srp'])) for i in records}
            item = records[0]
            if len(prices) == 1:
                for column, field in [('Total Price', 'dc'), ('Config Field 17', 'srp')]:
                    value = _number(item[field])
                    raw_field = purity.lower() + field if item.get('combined') else field
                    raw = item.get('price_raw', {}).get(raw_field, item[field])
                    if _number(raw) is None or value is None or value <= 0:
                        event('error', 'missing_or_invalid_source_price', sku, f'{column}: {raw!r}')
                        continue
                    previous = _number(row[column])
                    if previous != value:
                        differences.append(dict(stock_number=sku, field=column, reference=row[column],
                                                client=str(value), action='updated_from_client'))
                        row[column] = core._fmt(value)
            else:
                event('error', 'price_update_blocked', sku, 'Multiple client prices; reference price retained for review.')
            if _number(row['Total Price']) is not None and _number(row['Config Field 17']) is not None:
                if _number(row['Config Field 17']) < _number(row['Total Price']):
                    event('error', 'srp_below_dc', sku, f"DC={row['Total Price']}; SRP={row['Config Field 17']}")
            # Audit source-vs-curated details, without destroying an approved
            # center/side split or commercial rounded carat variation.
            weight = _number(row.get('Weight'))
            actual = _number(item.get('total_dia'))
            if weight is not None and actual is not None and actual > 0 and abs(weight - actual) > Decimal('0.02'):
                differences.append(dict(stock_number=sku, field='Weight', reference=row.get('Weight'),
                                        client=str(actual), action='preserved_reference_review'))
            declared = _text(item.get('source_collection'))
            known = {core._category_key(row.get(c, '')) for c in ['Config Field 13', 'Config Field 14', 'Jewelry Style']}
            if declared and not declared.startswith('#') and core._category_key(declared) not in known:
                differences.append(dict(stock_number=sku, field='Source collection label',
                                        reference=' / '.join(row.get(c, '') for c in ['Config Field 13', 'Jewelry Style', 'Config Field 14']),
                                        client=declared, action='preserved_reference_mapping'))
            if item.get('variation', '').upper() in ('NO', 'N', 'FALSE', '0') and color != item.get('default_color'):
                event('warning', 'reference_variation_exceeds_client_no', sku,
                      'Approved reference variation retained; client VARIATION=NO disagrees.')
        # Keep curated images, supplement blanks only from the exact base/color.
        supplied = list(dict.fromkeys(image_map.get((base, color), [])))
        existing = [row.get(f'Image Url {i}', '') for i in range(1, 9)]
        for url in supplied:
            if url in existing:
                continue
            blank = next((i for i, value in enumerate(existing) if not value), None)
            if blank is None:
                break
            existing[blank] = url
        for i, url in enumerate(existing, 1):
            row[f'Image Url {i}'] = url
        if not any(existing):
            event('warning', 'no_image', sku, 'No image for the exact base and metal color.')
        for column in ['Config Field 13', 'Config Field 14', 'Jewelry Style', 'Tags']:
            if not row.get(column, '').strip():
                event('error', 'missing_reference_classification', sku, column)
        output.append(row)

    result = pd.DataFrame(output, columns=clean.columns).fillna('')
    # One master per existing reference group, even when it contains many base
    # styles. Choose an existing 14K variant; otherwise retain the old master.
    for master, indices in result.groupby('Master Stock', sort=False).groups.items():
        if not master:
            event('error', 'missing_master', details=f'{len(indices)} rows')
            continue
        records = group_sources.get(master, [])
        colors = {i['default_color'] for i in records if i.get('default_color_raw')}
        group = result.loc[indices]
        if len(colors) > 1:
            event('warning', 'shared_master_default_color_conflict', master,
                  'Master retained; source colors: ' + ', '.join(sorted(colors)))
            continue
        if not colors:
            continue
        color = next(iter(colors))
        candidates = list(group[group['Stock Number'].str.endswith('-14K' + color)]['Stock Number'])
        if not candidates:
            event('warning', 'default_master_variant_missing', master,
                  f'No existing 14K {color} variant in this master group; existing master retained.')
            continue
        match = SKU_RE.fullmatch(master)
        preferred = f'{match[1]}-14K{color}' if match else ''
        replacement = preferred if preferred in candidates else candidates[0]
        default_column = 'Default Metal (MASTER STOCK)'
        if default_column in result:
            default_name = ' '.join(core.COLOR_NAME[c] for c in color) + ' Gold'
            changed = result.loc[indices, default_column].ne(default_name)
            for position in result.loc[indices].index[changed]:
                differences.append(dict(stock_number=result.at[position, 'Stock Number'],
                    field=default_column, reference=result.at[position, default_column],
                    client=default_name, action='updated_group_default_metal'))
            result.loc[indices, default_column] = default_name
        if replacement != master:
            result.loc[indices, 'Master Stock'] = replacement
            differences.append(dict(stock_number=master, field='Master Stock', reference=master,
                                    client=replacement, action='updated_group_master'))

    exported_skus = set(result['Stock Number'])
    for master in result['Master Stock'].unique():
        if master and master not in exported_skus:
            event('error', 'master_not_in_export', master,
                  'Master variant is absent or quarantined. Restore/resolve it before publishing.')

    # Report additions as draft candidates, never mix unapproved
    # new products into the reconstructed live inventory.
    for base in sorted(client_bases - ref_bases):
        candidates = [records for (b, _), records in source.items() if b == base]
        for records in candidates:
            if len({(i['dc'], i['srp']) for i in records}) != 1:
                continue
            item = dict(records[0])
            item['variation'] = item.get('variation') or 'NO'
            new_items.append(item)
        event('warning', 'new_client_style', base,
              'Not in approved reference. Default-color draft exported separately for review.')
    if new_items:
        drafts, messages = core.convert_items(new_items, image_map)
        for message in sorted(messages):
            event('warning', 'new_style_classification_review', details=message)
        for row in drafts.to_dict('records'):
            if not row.get('Image Url 1'):
                event('warning', 'new_style_missing_image', row['Stock Number'],
                      'Default-color draft has no exact-color image in the supplied list.')
    else:
        drafts = pd.DataFrame(columns=core.COLUMNS)

    issues = pd.DataFrame(events, columns=['severity', 'issue', 'stock_number', 'details', 'sheet', 'source_row']).fillna('')
    actions = {
        'conflicting_reference_sku': 'Confirm one approved row per SKU, including master, taxonomy and stone details.',
        'conflicting_source_prices': 'Client must confirm the correct DC and SRP for the listed style/purity and source rows.',
        'price_update_blocked': 'Resolve the related conflicting source prices; retained reference price is for review only.',
        'missing_or_invalid_source_price': 'Client must provide a positive numeric price in the stated source cell.',
        'reference_variant_missing_in_source': 'Confirm whether the reference item is still sold; provide its source row or approve removal.',
        'missing_source_sku': 'Client must add the stock number to the numbered item block at the stated sheet/row.',
        'source_excel_error': 'Client must replace the Excel error with the intended collection or price.',
        'conflicting_source_default_colors': 'Confirm one default metal color for this style/purity.',
        'shared_master_default_color_conflict': 'Confirm one default color for the shared master group; existing master was retained.',
        'default_master_variant_missing': 'Confirm whether to add the requested 14K variant; existing master was retained.',
        'master_not_in_export': 'Resolve/restore the master variant before publishing the inventory.',
        'new_client_style': 'Approve new product classification, color variations and master before importing the draft.',
        'new_style_classification_review': 'Confirm main collection, collection name, Jewelry Style and Config Field 14.',
        'new_style_missing_image': 'Provide an image for this exact style/color or approve the correct image mapping.',
        'unrecognized_source_stone': 'Confirm the stone shape/color abbreviation; no stone mapping was guessed.',
    }
    issues['action'] = issues['issue'].map(actions).fillna('Review the reported source/reference discrepancy.')
    action_summary = issues.groupby(['severity', 'issue', 'action'], sort=False).size().reset_index(name='occurrences')
    diffs = pd.DataFrame(differences, columns=['stock_number', 'field', 'reference', 'client', 'action']).fillna('')
    summary = {
        'reference_rows': len(ref), 'reference_unique_skus': ref['Stock Number'].nunique(),
        'client_item_records': len(items), 'client_base_styles': len(client_bases),
        'reference_base_styles': len(ref_bases), 'matching_base_styles': len(client_bases & ref_bases),
        'new_client_base_styles': len(client_bases - ref_bases),
        'reference_base_styles_missing_in_client': len(ref_bases - client_bases),
        'conflicting_reference_skus_excluded': conflict_rows['Stock Number'].nunique(),
        'identical_reference_duplicates_removed': exact_removed,
        'candidate_rows': len(result), 'new_draft_rows': len(drafts),
        'blocking_issues': int((issues['severity'] == 'error').sum()),
        'warnings': int((issues['severity'] == 'warning').sum()),
        'publish_ready': not bool((issues['severity'] == 'error').any()),
        'image_check_scope': 'Filename/color mapping only; remote image URLs were not fetched.',
        'curated_fields_policy': 'Preserve approved reference fields; update unambiguous DC/SRP and default masters.',
    }
    crosswalk = []
    reference_classifications = defaultdict(set)
    for row in ref.to_dict('records'):
        match = SKU_RE.fullmatch(row['Stock Number'])
        if match:
            reference_classifications[match[1].upper()].add(tuple(row.get(c, '') for c in
                ['Config Field 13', 'Jewelry Style', 'Config Field 14']))
    for (base, purity), records in source.items():
        for record in records:
            label = record.get('source_collection', '')
            for classification in reference_classifications.get(base, {('', '', '')}):
                crosswalk.append((label, *classification, base, purity))
    crosswalk = pd.DataFrame(crosswalk, columns=['Client collection label', 'Config Field 13',
        'Jewelry Style', 'Config Field 14', 'Base Stock', 'Purity']).drop_duplicates()
    return result, dict(issues=issues, client_actions=action_summary, differences=diffs, collection_crosswalk=crosswalk,
                       source_items=pd.DataFrame(source_rows), new_items=drafts,
                       conflicting_reference_rows=conflict_rows, summary=summary)


def build_inventory(inputs, reference, image_map=None):
    import app as core
    items, diagnostics = [], []
    for source in inputs:
        items.extend(core.extract(source, diagnostics))
    return reconcile(items, reference, image_map, diagnostics)


def audit_zip(inventory, reports):
    """Download includes a review candidate and every reconciliation report."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('inventory_review.csv', inventory.to_csv(index=False))
        for name, value in reports.items():
            if isinstance(value, pd.DataFrame):
                archive.writestr(name + '.csv', value.to_csv(index=False))
        archive.writestr('summary.json', json.dumps(reports['summary'], indent=2))
        archive.writestr('READ_ME.txt',
            'REVIEW EXPORT: read summary.json and client_actions.csv before importing.\n'
            'inventory_review.csv preserves curated reference metadata and applies only unambiguous client prices/default masters.\n'
            'Conflicting reference SKUs are excluded. Missing source items and ambiguous source prices retain reference values for review.\n'
            'new_items.csv contains unapproved default-color drafts, not live inventory additions.\n'
            'collection_crosswalk.csv traces client labels to the existing collection/style/filter values.\n'
            'Image checks cover filename/color matching only; remote URLs have not been fetched.\n')
    return buffer.getvalue()


def render_audit(st, inventory, reports, key='reconciliation'):
    summary = reports['summary']
    st.subheader('Inventory reconciliation')
    a, b, c = st.columns(3)
    a.metric('Matched base styles', summary['matching_base_styles'])
    b.metric('New styles for review', summary['new_client_base_styles'])
    c.metric('Blocking issues', summary['blocking_issues'])
    if summary['blocking_issues']:
        st.error('Review required. Conflicting reference SKUs are excluded; unresolved source prices '
                 'and absent source rows retain reference values. This is a review export, not a publish-ready file.')
    else:
        st.success('Reconciliation passed without blocking issues. Review any warnings and new-item drafts.')
    with st.expander('Issues and client actions', expanded=bool(summary['blocking_issues'])):
        st.dataframe(reports['issues'], use_container_width=True, hide_index=True)
    with st.expander('Price changes, master changes and source differences'):
        st.dataframe(reports['differences'], use_container_width=True, hide_index=True)
    st.download_button('Download inventory and audit reports', data=audit_zip(inventory, reports),
                       file_name='Smiling_Rocks_Inventory_Audit.zip', mime='application/zip', key=key)
