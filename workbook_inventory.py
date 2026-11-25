"""Build VDB inventory from the five confirmation fields plus an image list.

Saved approved rules supply previously curated stone/taxonomy/size conventions;
no reference CSV, prices or image URLs are required from that saved catalogue.
Uploaded grouping instructions and per-purity prices remain authoritative.
"""
from collections import defaultdict
from decimal import Decimal, InvalidOperation
import re
import pandas as pd
from saved_inventory_rules import load_rules

DIMENSIONS = {
    'metal': ('Metal','Available Metals'),
    'shape': ('Shape','Supported Shape Variations'),
    'caratweight': ('Config field 11','Available Config field 11'),
    'centercaratweight': ('Center Carat Weight','Available Center Carat Weight'),
    'stonecolor': ('Config Field 6','Available Config Field 6'),
    'ringsize': (None,'Available Ring Size'),
    'length': (None,'Available Config Field 7'),
}
ALIASES={'metals':'metal','color':'stonecolor','gemstone':'stonecolor','stonecolour':'stonecolor',
         'totalcaratweight':'caratweight','weight':'caratweight','centerweight':'centercaratweight',
         'centrecaratweight':'centercaratweight','size':'ringsize','necklacelength':'length'}


def _key(value):return re.sub('[^a-z0-9]','',str(value or '').lower())
def _sku_base(value):return re.sub(r'-(?:14K|18K)[WYR]{1,3}$','',value,flags=re.I)
def _number(value):
    try:
        n=Decimal(str(value).replace(',','').strip());return n if n.is_finite() else None
    except InvalidOperation:return None

def _option_key(value):
    n=_number(value)
    return format(n.normalize(),'f') if n is not None else str(value or '').strip().casefold()

def _group_name(value):
    k=_key(value)
    return {'essential':'Essentials','essentials':'Essentials','bridal':'Bridals','bridals':'Bridals',
            'hautecouture':'Haute Couture','fashion':'Fashion Collections','fashioncollections':'Fashion Collections',
            'men':'Men\'s','mens':'Men\'s','zulu':'Men\'s'}.get(k,'')


def build_from_items(items, image_map=None, diagnostics=None, rules=None):
    import app as core
    from collection_mappings import COLLECTION_GROUPS
    rules=rules if rules is not None else load_rules();profiles=rules['profiles'];image_map=image_map or {}
    events=[];traces=[];availability_trace=[]
    def issue(severity,kind,sku,details,**extra):
        events.append(dict(severity=severity,issue=kind,stock_number=sku,details=details,**extra))
    if not profiles:issue('error','saved_rules_missing','','Use the supplied client template with its hidden saved-rule sheet. No separate inventory CSV is needed.')
    for d in diagnostics or []:issue(d.get('severity','error'),d.get('issue','parse_error'),'',d.get('details',''),sheet=d.get('sheet'),source_row=d.get('source_row'))
    profiles_by_base=defaultdict(list)
    for sku,entries in profiles.items():profiles_by_base[_sku_base(sku)].extend(entries)
    indexed=defaultdict(list)
    for i in items:indexed[i['base'],i['purity']].append(i)
    selected={}
    for key,rows in indexed.items():
        preferred=[i for i in rows if i.get('combined')] or rows
        signatures={(i['dc'],i['srp'],i.get('default_color'),str(i.get('instructions'))) for i in preferred}
        if len(signatures)>1:
            issue('error','conflicting_client_item','-'.join(key),'Different prices, defaults or instructions for this style/purity.',sheet=preferred[0].get('sheet'),source_row=preferred[0].get('source_row'));continue
        selected[key]=preferred[0]
    by_base=defaultdict(list)
    for (base,purity),item in selected.items():by_base[base].append(item)
    groups=defaultdict(list);declared={};member_orders={}
    for base,records in by_base.items():
        signatures=[]
        for item in records:
            raw=item.get('instructions',{})
            collection=str(raw.get('collection','')).strip();style=str(raw.get('style','')).strip()
            known_labels={_key(v) for v in rules.get('source_labels',{}).get(base,[])}
            styles={p.get('Jewelry Style','') for p in profiles_by_base.get(base,[])}-{''}
            if _key(style) in known_labels and len(styles)==1:style=next(iter(styles))
            flag=str(raw.get('variation','')).strip().lower()
            members=[s.strip().upper() for s in re.split('[,#;\n]+',str(raw.get('styles',''))) if s.strip()]
            columns=[ALIASES.get(_key(c),_key(c)) for c in re.split('[,#;\n]+',str(raw.get('columns',''))) if c.strip()]
            missing=[name for name,val in [('COLLECTION NAME',collection),('Style Name',style),('Variation Styles',members)] if not val]
            if flag not in ('yes','no'):missing.append('Variations (Yes / No)')
            if flag=='yes' and not columns:missing.append('Variation Columns')
            if flag=='no' and columns:issue('error','no_with_variation_columns',base,'No requires empty Variation Columns.')
            if missing:
                issue('error','incomplete_workbook_instructions',base,', '.join(missing),sheet=item.get('sheet'),source_row=item.get('source_row'));continue
            if base not in members:issue('error','group_does_not_include_current_style',base,','.join(members));continue
            unknown=set(columns)-set(DIMENSIONS)
            if unknown:issue('error','unknown_variation_columns',base,', '.join(sorted(unknown)));continue
            if len(members)!=len(set(members)):issue('error','duplicate_variation_styles',base,','.join(members));continue
            member_orders.setdefault(base,tuple(members))
            signature=(collection,style,flag,tuple(sorted(members)),tuple(columns))
            signatures.append(signature)
        if len(signatures)!=len(records) or len(set(signatures))!=1:
            if signatures:issue('error','inconsistent_instructions_between_purities',base,'Use the same five fields on 14K and 18K records.')
            continue
        signature=signatures[0];declared[base]=signature;groups[signature[3]].append(base)
    output=[]
    for members,bases in groups.items():
        if set(members)!=set(bases):
            issue('error','incomplete_or_overlapping_variation_group',','.join(bases),'Missing/inconsistent instructions for: '+','.join(sorted(set(members)-set(bases))));continue
        definitions={declared[b] for b in bases}
        if len(definitions)!=1:
            issue('error','group_metadata_conflict',','.join(bases),'All linked styles must use identical collection, style, Yes/No and ordered variation columns.');continue
        collection,style,flag,_,columns=next(iter(definitions))
        members=member_orders[bases[0]]
        records=[i for b in members for i in by_base[b]]
        defaults={i.get('default_color') for i in records if i.get('default_color_raw')}
        if len(defaults)!=1:
            issue('error','group_default_color_conflict',','.join(bases),'Provide one consistent default metal color for the shared product group.');continue
        default=next(iter(defaults))
        if any(_number(i.get(f)) is None or _number(i.get(f))<=0 for i in records for f in ['dc','srp']):
            issue('error','invalid_client_prices',','.join(bases),'DC and SRP must be positive numeric values for every style/purity.');continue
        group_rows=[]
        for base in members:
            saved_colors={m[1] for p in profiles_by_base.get(base,[]) if (m:=re.search(r'-(?:14K|18K)([WYR]{1,3})$',p['Stock Number']))}
            supplied_colors={color for b,color in image_map if b==base}
            colors=sorted(saved_colors|supplied_colors|{default}) if 'metal' in columns else [default]
            for item in sorted(by_base[base],key=lambda x:x['purity']):
                draft_item=dict(item,variation='NO')
                draft,_=core.convert_items([draft_item],image_map,collection='Auto')
                fallback=draft.iloc[0].to_dict()
                for color in colors:
                    sku=f"{base}-{item['purity']}{color}";row=dict(fallback)
                    row['Stock Number']=sku;row['Metal']=core._metal_name(item['purity'],color)
                    candidates=profiles.get(sku,[])
                    exact_profile=bool(candidates)
                    if not candidates:
                        candidates=[p for p in profiles_by_base.get(base,[]) if f"-{item['purity']}" in p['Stock Number']] or profiles_by_base.get(base,[])
                    context=[p for p in candidates if _sku_base(p.get('Master Stock','')) in members]
                    candidates=context or candidates
                    if candidates:
                        for field in candidates[0]:
                            if field in ('Stock Number','Master Stock','Metal') or field.startswith('Available ') or field=='Supported Shape Variations':continue
                            if not exact_profile and field in ('Short Title','Description','barcode No'):continue
                            values={p.get(field,'') for p in candidates}
                            if len(values)==1:row[field]=next(iter(values))
                            else:issue('error','saved_mapping_conflict',sku,'Client instructions do not resolve '+field)
                        # Keep explicit size/length policies; other availability is rebuilt below.
                        for field in ['Available Ring Size','Available Config Field 7']:
                            values={p.get(field,'') for p in candidates}
                            if len(values)==1:row[field]=next(iter(values))
                            else:issue('error','size_length_policy_conflict',sku,field)
                    current_group=COLLECTION_GROUPS.get(row.get('Config Field 13',''),'')
                    known_label=_key(style) in {_key(s) for s in rules.get('source_labels',{}).get(base,[])}
                    matching_style=_key(style)==_key(row.get('Jewelry Style'))
                    if not candidates or current_group!=_group_name(collection) or not (known_label or matching_style):
                        classification=core.resolve_classification({'base':base,'remark':style,'source_collection':'','source_style':style},collection)
                        c,st,cf,tags,message=classification
                        if not message:
                            row.update({'Config Field 13':c,'Jewelry Style':st,'Config Field 14':cf,'Tags':tags})
                        else:issue('error','unresolved_collection_style',sku,message)
                    row['Total Price']=core._fmt(item['dc']);row['Config Field 17']=core._fmt(item['srp'])
                    if item['srp']<item['dc']:issue('error','srp_below_dc',sku,'SRP must not be below DC.')
                    row['Default Metal (MASTER STOCK)']=' '.join(core.COLOR_NAME[c] for c in default)+' Gold'
                    urls=list(dict.fromkeys(image_map.get((base,color),[])))[:8]
                    for n in range(1,9):row[f'Image Url {n}']=urls[n-1] if n<=len(urls) else ''
                    if not urls:issue('warning','missing_exact_color_image',sku,'No image for this base/color in the supplied CSV.')
                    for field in ['Available Metals','Available Config field 11','Available Center Carat Weight','Supported Shape Variations','Available Config Field 6']:row[field]=''
                    if exact_profile and candidates:
                        previous_metals={p.get('Metal','') for p in candidates}
                        if len(previous_metals)==1:
                            old_metal=next(iter(previous_metals))
                            if old_metal and old_metal!=row['Metal']:
                                for text_field in ['Short Title','Description']:row[text_field]=row.get(text_field,'').replace(old_metal,row['Metal'])
                    if not exact_profile:
                        if not candidates and not item.get('centers'):issue('warning','new_style_center_assignment_review',sku,'No explicit CS marker; parsed stones are side stones. Confirm center/side assignment.')
                        if 'caratweight' in columns and not row.get('Config field 11'):row['Config field 11']=row.get('Weight','')
                        row['Short Title']=re.sub(r'\s+',' ',f"{row.get('Config Field 13','')} {row.get('Jewelry Style','')} {row.get('Weight','')}CT {row['Metal']} {row.get('Jewelry Sub-Category','')}").strip()
                        row['Description']=row['Short Title']
                    for field in ['Config Field 13','Jewelry Style','Config Field 14','Tags']:
                        if not row.get(field):issue('error','missing_classification',sku,field)
                    group_rows.append(row)
        candidates=[r['Stock Number'] for r in group_rows if r['Stock Number'].endswith('-14K'+default)]
        approved={p.get('Master Stock','') for b in members for p in profiles_by_base.get(b,[])}
        preferred_bases={_sku_base(s) for s in approved}
        preferred=[s for s in candidates if _sku_base(s) in preferred_bases]
        master=sorted(preferred)[0] if len(preferred)==1 else candidates[0] if candidates else group_rows[0]['Stock Number']
        if not candidates:issue('warning','no_14k_master_available',master,'No uploaded 14K variant exists; first available variant used.')
        for row in group_rows:row['Master Stock']=master
        # Ordered columns define parent-to-child option dependencies.
        parents=[]
        for name in columns:
            field,available=DIMENSIONS[name]
            if field is None:
                values={r.get(available,'') for r in group_rows}
                if not any(values):issue('error','missing_option_policy',master,'Provide an approved '+name+' range; it cannot be inferred from one sample size.')
                continue
            for row in group_rows:
                matching=[r for r in group_rows if all(_option_key(r.get(p,''))==_option_key(row.get(p,'')) for p in parents)]
                unique={}
                for r in matching:
                    value=str(r.get(field,'')).strip()
                    if value:unique.setdefault(_option_key(value),value)
                values=list(unique.values())
                def sort(v):
                    n=_number(v);return (0,n) if n is not None else (1,v.casefold())
                row[available]='#'.join(sorted(values,key=sort))
            parents.append(field)
        for name,(field,available) in DIMENSIONS.items():
            if field and name not in columns and len({_option_key(r.get(field,'')) for r in group_rows if r.get(field,'')})>1:
                issue('error','undeclared_variation_attribute',master,'Add '+name+' to Variation Columns or correct grouping.')
        for name,(field,available) in DIMENSIONS.items():
            option_groups={}
            for row in group_rows:
                value=row.get(available,'')
                if value:option_groups[value]=option_groups.get(value,0)+1
            for value,count in option_groups.items():
                availability_trace.append(dict(master_stock=master,column=available,source_column=field or 'saved size/length policy',
                                               options=value,rows=count,declared_order=','.join(columns)))
        traces.append(dict(master_stock=master,variation_styles=','.join(members),collection=collection,style=style,
                           variation_columns=','.join(columns),default_color=default,variants=len(group_rows),
                           rule='Confirmed group; approved parent when unique, otherwise first declared available 14K default-color variant.'))
        output.extend(group_rows)
    columns=list(core.COLUMNS)+['Default Metal (MASTER STOCK)','Best Selling Items ']
    df=pd.DataFrame(output).reindex(columns=columns).fillna('')
    issue_df=pd.DataFrame(events,columns=['severity','issue','stock_number','details','sheet','source_row']).fillna('')
    issue_df['action']='Correct the indicated workbook fields and upload again; no separate inventory CSV is required.'
    errors=int(issue_df.severity.eq('error').sum());warnings=int(issue_df.severity.eq('warning').sum())
    summary=dict(mode='confirmation_workbook',candidate_rows=len(df),client_base_styles=len(by_base),
                 matching_base_styles=sum(b in profiles_by_base for b in by_base),
                 new_client_base_styles=sum(b not in profiles_by_base for b in by_base),
                 blocking_issues=errors,warnings=warnings,publish_ready=errors==0,masters_created=len(traces))
    empty=pd.DataFrame()
    reports=dict(summary=summary,issues=issue_df,client_actions=issue_df.groupby(['severity','issue','action']).size().reset_index(name='occurrences'),
                 differences=empty,master_trace=pd.DataFrame(traces),availability_audit=pd.DataFrame(availability_trace),new_items=empty)
    return df,reports
