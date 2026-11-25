# Smiling Rocks converter

Run `streamlit run app.py` or `streamlit run Dashboard.py`.

## Full inventory updates

1. Upload the client XLSX or XLSX ZIP (including the collection-wise full workbook).
2. Optionally upload the image filename CSV. Files may have a header or not; W/Y/R and two tone codes such as WY/WY1 are supported.
3. Upload the approved VDB inventory CSV in the reference input.
4. Convert and download the inventory/audit ZIP. Review the issues before publishing.

The reference CSV preserves existing master groups across multiple styles, variation rows, taxonomy, stone assignments, titles, image mappings, barcodes and availability options. The quotation alone does not contain enough information to reproduce those curated values. Client DC/SRP update Total Price/Config Field 17 only where the source has one valid price pair. Combined 14K-18K sheets take priority. Default colors update a master once per existing group using an existing 14K variant. Missing defaults/variants retain the current master and are reported.

Conflicting duplicate reference SKUs are quarantined rather than arbitrarily chosen. Conflicting client prices retain the reference price for review. Missing source items do not cause deletion. New styles are separate default-color drafts and require classification/variation review. A blocking issue marks the result as a review export. Warnings also require attention; a review export must not be treated as an automatic go-live approval.

Reports include source sheet/row provenance, price/master changes, source collection crosswalk, missing styles, new drafts and conflicting reference rows. Image checks cover filename/color mapping; remote URLs are not fetched.

## Batch usage

```bash
python app.py CLIENT.xlsx.zip --reference APPROVED.csv -i IMAGES.csv -o inventory_review.csv
```

Writes the review CSV plus `inventory_review.audit.zip`. Exit code 2 means blocking reconciliation issues remain. Normal quotation-only conversion remains available without `--reference`, but cannot reproduce fields absent from the client workbook. Upload limit is configured to 800 MB for photo-heavy files; the hosting service must have sufficient memory and disk.

## Validation

```bash
python -m unittest discover -s tests
```

Regression coverage includes header aliases, nested XLSX ZIPs, reference field preservation, combined-sheet priority, conflicting prices/SKUs, shared masters, missing source items, invalid prices, new-style isolation and exact-color image matching.

## Every-master, every-column validation

Full-inventory mode audits every named reference column for every master group. Reports include:

- `master_trace.csv`: original/generated/expected master, linked styles, source rows/default colors, eligible 14K variants, missing members, and exact selection rule (same base first; then an existing group member; otherwise retain the original).
- `master_column_audit.csv`: every column for every master, with lineage, populated counts and complete distinct reference/generated values.
- `availability_audit.csv`: observed options, expected options, supported dependency columns, and unresolved rules.
- `availability_repairs.csv`: supported option-list corrections, SKU by SKU.
- `master_validation_issues.csv`: structural, numeric, image, title/carat, stone, barcode and option-policy checks.
- `column_rules.csv`: source/generation rules for every business column.

Available Metals, Config field 11, Center Carat Weight, Shape and Config Field 6 are derived from actual master members. If the reference proves that a list depends on another selected field, that dependency is used. An unexplained or ambiguous rule blocks readiness instead of being replaced by a guessed global list. Numeric lists are deduplicated and sorted in ascending order when repaired; equivalent formats such as `1.0`/`1.00` compare numerically.

Available Ring Size and Available Config Field 7 are validated as explicit policies because the file contains no corresponding scalar size/length variation columns. Their full ranges cannot be inferred from one quoted sample size. The tool preserves these lists, reports invalid increments/order, and requests an approved policy rather than inventing sizes or lengths.
