# Smiling Rocks converter

Run `streamlit run app.py` or `streamlit run Dashboard.py`.

## Confirmation workbook + image CSV (normal workflow)

Upload the client XLSX or XLSX ZIP and the image filename CSV. No reference inventory upload is required when the workbook contains the confirmation columns:

- COLLECTION NAME (the main group)
- Style Name
- Variations (Yes / No)
- Variation Styles (comma-separated base stock numbers belonging to one product)
- Variation Columns (ordered selectable attributes)
- Available Metals (comma-separated colors, e.g. White, Yellow, Rose or White Yellow, White Rose)

The confirmation columns are auto-detected even when the original COLLECTION NAME column is also present. All present member styles must contain the same Variation Styles membership and ordered Variation Columns. Each style may have its own default metal color. Prices come from the uploaded DC/SRP columns; combined 14K-18K pricing takes priority.

The tool creates both 14K and 18K variants for each explicitly listed color, using uploaded prices for each karat. Two-tone names such as White Yellow represent a single color option. Image filenames do not determine available colors. Both karat prices must be supplied. Old templates without Available Metals use approved saved metal names as a compatibility fallback.

The tool creates one master per Variation Styles group. It selects the first listed eligible style and its 14K variant matching that style’s existing default metal color. Different defaults across member styles are allowed. Missing group members are reported for later addition; conflicting membership among present styles is an error. Available options are derived from actual generated variants in the order specified by Variation Columns (parent to child).

Supported names include Metal, Shape, Carat Weight, Center Carat Weight, Stone Color, Ring Size and Length. Metal includes karat and color. A No instruction must have no varying attributes. Existing size/length policies and curated stone/taxonomy conventions are stored privately in the client template's hidden _VDB_Rules worksheet, without prices or image URLs. The repository contains only a loader, never the private inventory catalogue. Keep that hidden sheet when editing the supplied template. An optional server-only SMILING_ROCKS_RULES_FILE setting can point to a private rule file. Exact-color images come from the supplied CSV, including WY/WR/YR codes with or without numeric suffixes.

Saved catalogue rules retain known classification, stone assignments, titles, barcodes and sizing policies. The client workbook overrides grouping and always supplies prices. New or conflicting metadata, missing color lists, missing karat prices or missing size policies generate actionable review reports. Missing images remain blank with a warning. The tool does not silently invent missing instructions. New designs without explicit center markers may need center/side review.

The legacy reference CSV input remains optional for old workbooks without the five columns.

## Batch usage

```bash
python app.py CONFIRMED_CLIENT.xlsx.zip -i IMAGES.csv -o inventory_review.csv
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
