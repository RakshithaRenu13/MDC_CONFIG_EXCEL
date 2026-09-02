import streamlit as st
import pandas as pd
from io import BytesIO
import re
import math

st.set_page_config(
    page_title="Eaton MDC Configuration & BOM Generator",
    page_icon="⚡",
    layout="wide",
)

# ============================================================
# HELPERS
# ============================================================

CANONICAL_FIELDS = [
    "MDC Type",
    "Configuration",
    "Category",
    "Part Number",
    "Description",
    "Quantity",
    "Unit Price",
    "Optional",
    "Component Type",
]

ALIASES = {
    "MDC Type": [
        "mdc type", "type", "mdc", "product type", "configuration type",
        "single/multi", "single rack/multi rack"
    ],
    "Configuration": [
        "configuration", "config", "config no", "config number",
        "configuration no", "configuration number", "model", "variant"
    ],
    "Category": [
        "category", "cat", "section", "bom category", "item category"
    ],
    "Part Number": [
        "part number", "part no", "part#", "part", "pn", "item number",
        "item no", "catalog number", "catalog no", "product number"
    ],
    "Description": [
        "description", "product description", "item description",
        "component", "component description", "product"
    ],
    "Quantity": [
        "quantity", "qty", "required qty", "required quantity", "count"
    ],
    "Unit Price": [
        "unit price", "price", "unit cost", "cost", "selling price",
        "standard price", "rate", "amount"
    ],
    "Optional": [
        "optional", "option", "optional component", "is optional",
        "optional item", "mandatory/optional"
    ],
    "Component Type": [
        "component type", "item type", "type of component", "component category"
    ],
}

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def normalize_name(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())

def auto_map_columns(columns):
    normalized = {normalize_name(c): c for c in columns}
    mapping = {}
    for canonical, aliases in ALIASES.items():
        found = None
        alias_norms = [normalize_name(a) for a in aliases]

        # Exact normalized match first
        for a in alias_norms:
            if a in normalized:
                found = normalized[a]
                break

        # Then partial match
        if found is None:
            for n, original in normalized.items():
                if any(a and (a in n or n in a) for a in alias_norms):
                    found = original
                    break

        mapping[canonical] = found
    return mapping

def normalize_mdc_type(value):
    s = clean_text(value).lower()
    if "multi" in s:
        return "Multi Rack MDC"
    if "single" in s:
        return "Single Rack MDC"
    return clean_text(value)

def normalize_config(value):
    s = clean_text(value)
    if not s:
        return s

    m = re.search(r"(?:config(?:uration)?[\s._-]*)(\d+)", s, re.I)
    if m:
        return f"Config {m.group(1)}"

    m = re.fullmatch(r"(\d+)", s)
    if m:
        return f"Config {m.group(1)}"

    return s

def truthy_optional(value):
    s = clean_text(value).lower()
    return s in {
        "yes", "y", "true", "1", "optional", "option",
        "optional component", "o"
    }

def numeric_value(value):
    if pd.isna(value):
        return math.nan
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    s = str(value).strip().replace(",", "")
    if not s or s.upper() in {"XXX", "N/A", "NA", "-", "TBD", "TBC"}:
        return math.nan

    # Remove currency symbols and spaces
    s = re.sub(r"[₹$€£]", "", s).strip()
    try:
        return float(s)
    except ValueError:
        return math.nan

def format_money(value, currency="INR"):
    if pd.isna(value):
        return "XXX"
    if currency == "INR":
        return f"₹{value:,.2f}"
    return f"{currency} {value:,.2f}"

def safe_sheet_name(name):
    name = re.sub(r"[:\\/?*\[\]]", "_", str(name))
    return name[:31] or "Sheet"

def style_workbook(writer):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = writer.book
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E1F2")

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=thin)

        for column_cells in ws.columns:
            max_len = 0
            for cell in column_cells:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))
            width = min(max(max_len + 2, 10), 45)
            ws.column_dimensions[get_column_letter(column_cells[0].column)].width = width

def create_template():
    single = pd.DataFrame([
        ["Single Rack MDC", "Config 1", "A.1", "801029209", "MDC Base Assembly", 1, "XXX", "No", "Standard"],
        ["Single Rack MDC", "Config 1", "A.2", "COOL-35", "3.5 kW Cooling", 1, "XXX", "No", "Cooling"],
        ["Single Rack MDC", "Config 1", "A.3", "OPT-001", "Example Optional Component", 1, "XXX", "Yes", "Optional"],
        ["Single Rack MDC", "Config 2", "A.2", "COOL-70", "7 kW Cooling", 1, "XXX", "No", "Cooling"],
    ], columns=CANONICAL_FIELDS)

    multi = pd.DataFrame([
        ["Multi Rack MDC", "Config 1", "A.1", "MR-BASE", "Multi Rack Base Assembly", 1, "XXX", "No", "Standard"],
        ["Multi Rack MDC", "Config 1", "A.2", "MR-COOL", "Multi Rack Cooling", 1, "XXX", "No", "Cooling"],
        ["Multi Rack MDC", "Config 1", "A.4", "MR-OPT-001", "Example Optional Component", 1, "XXX", "Yes", "Optional"],
    ], columns=CANONICAL_FIELDS)

    optional = pd.DataFrame([
        ["Single Rack MDC", "Config 1", "A.3", "OPT-001", "Example Optional Component", 1, "XXX", "Yes", "Optional"],
        ["Multi Rack MDC", "Config 1", "A.4", "MR-OPT-001", "Example Optional Component", 1, "XXX", "Yes", "Optional"],
    ], columns=CANONICAL_FIELDS)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        single.to_excel(writer, index=False, sheet_name="Single Rack")
        multi.to_excel(writer, index=False, sheet_name="Multi Rack")
        optional.to_excel(writer, index=False, sheet_name="Optional Components")
        style_workbook(writer)
    output.seek(0)
    return output.getvalue()

def prepare_dataframe(raw_df, mapping, sheet_name):
    df = raw_df.copy()
    output = pd.DataFrame()

    for canonical in CANONICAL_FIELDS:
        source = mapping.get(canonical)
        if source and source in df.columns:
            output[canonical] = df[source]
        else:
            output[canonical] = ""

    # Useful fallbacks
    if not output["MDC Type"].astype(str).str.strip().any():
        output["MDC Type"] = normalize_mdc_type(sheet_name)

    output["MDC Type"] = output["MDC Type"].apply(normalize_mdc_type)
    output["Configuration"] = output["Configuration"].apply(normalize_config)

    if not output["Quantity"].astype(str).str.strip().any():
        output["Quantity"] = 1

    if not output["Optional"].astype(str).str.strip().any():
        output["Optional"] = "No"

    output["Quantity"] = pd.to_numeric(output["Quantity"], errors="coerce").fillna(1)
    output["Quantity"] = output["Quantity"].clip(lower=0)

    output["Part Number"] = output["Part Number"].apply(clean_text)
    output["Description"] = output["Description"].apply(clean_text)
    output["Category"] = output["Category"].apply(clean_text)
    output["Component Type"] = output["Component Type"].apply(clean_text)

    return output

def infer_optional(df):
    result = df["Optional"].apply(truthy_optional)

    if not result.any():
        combined = (
            df["Category"].astype(str) + " " +
            df["Component Type"].astype(str) + " " +
            df["Description"].astype(str)
        ).str.lower()

        result = combined.str.contains(
            r"\boptional\b|\boption\b|\badd[- ]?on\b",
            regex=True,
            na=False,
        )

    return result

def build_catalog(df):
    df = df.copy()
    df["_OptionalFlag"] = infer_optional(df)
    return df

def get_types(df):
    vals = [clean_text(v) for v in df["MDC Type"].dropna().unique()]
    vals = [v for v in vals if v]
    if not vals:
        return ["Single Rack MDC", "Multi Rack MDC"]
    return vals

def get_configs(df, mdc_type):
    subset = df[df["MDC Type"].astype(str) == str(mdc_type)]
    vals = [clean_text(v) for v in subset["Configuration"].dropna().unique()]
    vals = [v for v in vals if v]

    def sort_key(x):
        m = re.search(r"(\d+)", x)
        return (int(m.group(1)) if m else 9999, x)

    return sorted(vals, key=sort_key)

def rows_for_selection(df, mdc_type, config):
    mask = (
        (df["MDC Type"].astype(str) == str(mdc_type)) &
        (df["Configuration"].astype(str) == str(config))
    )
    return df[mask].copy()

def render_component_quantity_table(options, prefix):
    selected = []

    if options.empty:
        st.info("No optional components were found for this selection.")
        return selected

    for idx, row in options.reset_index(drop=True).iterrows():
        part = row["Part Number"] or f"ITEM-{idx+1}"
        desc = row["Description"] or "Unnamed component"
        default_qty = int(row["Quantity"]) if float(row["Quantity"]).is_integer() else float(row["Quantity"])

        c1, c2, c3, c4 = st.columns([0.7, 3.8, 1.2, 1.5])
        with c1:
            checked = st.checkbox(
                "Select",
                key=f"{prefix}_check_{idx}",
                label_visibility="collapsed",
            )
        with c2:
            st.write(f"**{desc}**")
            st.caption(f"Part No: {part}")
        with c3:
            qty = st.number_input(
                "Qty",
                min_value=1.0,
                value=float(default_qty) if default_qty > 0 else 1.0,
                step=1.0,
                key=f"{prefix}_qty_{idx}",
                disabled=not checked,
            )
        with c4:
            price = numeric_value(row["Unit Price"])
            st.write("XXX" if pd.isna(price) else format_money(price))

        if checked:
            item = row.to_dict()
            item["_SelectedQty"] = qty
            selected.append(item)

    return selected

def build_bom(records):
    if not records:
        return pd.DataFrame(columns=[
            "MDC No", "MDC Type", "Configuration", "Category",
            "Part Number", "Description", "Quantity", "Unit Price",
            "Extended Price", "Optional"
        ])

    rows = []
    for r in records:
        qty = float(r.get("_SelectedQty", r.get("Quantity", 1)))
        unit = numeric_value(r.get("Unit Price"))
        extended = qty * unit if not pd.isna(unit) else math.nan

        rows.append({
            "MDC No": r.get("_MDCNo", ""),
            "MDC Type": r.get("MDC Type", ""),
            "Configuration": r.get("Configuration", ""),
            "Category": r.get("Category", ""),
            "Part Number": r.get("Part Number", ""),
            "Description": r.get("Description", ""),
            "Quantity": qty,
            "Unit Price": unit if not pd.isna(unit) else "XXX",
            "Extended Price": extended if not pd.isna(extended) else "XXX",
            "Optional": "Yes" if r.get("_OptionalFlag", False) else "No",
        })

    bom = pd.DataFrame(rows)

    # Aggregate identical BOM lines
    if not bom.empty:
        grouped_rows = []
        group_cols = [
            "MDC Type", "Configuration", "Category",
            "Part Number", "Description", "Unit Price", "Optional"
        ]

        for keys, group in bom.groupby(group_cols, dropna=False, sort=False):
            rec = dict(zip(group_cols, keys))
            rec["MDC No"] = ", ".join(sorted(set(group["MDC No"].astype(str))))
            rec["Quantity"] = group["Quantity"].sum()

            unit = numeric_value(rec["Unit Price"])
            rec["Extended Price"] = (
                rec["Quantity"] * unit if not pd.isna(unit) else "XXX"
            )
            grouped_rows.append(rec)

        bom = pd.DataFrame(grouped_rows)

    return bom

# ============================================================
# HEADER
# ============================================================

st.title("⚡ Eaton MDC Configuration & BOM Generator")
st.caption(
    "Excel-driven configuration system — upload the current Excel workbook, "
    "map its columns once, configure each MDC, and download the BOM."
)

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Project Controls")
    currency = st.selectbox("Display Currency", ["INR", "USD"], index=0)

    if currency == "USD":
        inr_to_usd = st.number_input(
            "INR → USD conversion rate",
            min_value=0.000001,
            value=0.012,
            step=0.001,
            format="%.6f",
            help="Manual rate so the application remains fully offline.",
        )
    else:
        inr_to_usd = 1.0

    st.divider()
    st.download_button(
        "⬇️ Download Excel Template",
        data=create_template(),
        file_name="MDC_Excel_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ============================================================
# FILE UPLOAD
# ============================================================

st.header("1. Upload Current MDC Excel")

uploaded = st.file_uploader(
    "Upload the Excel workbook used by your team",
    type=["xlsx", "xls"],
    help="The workbook is read directly in this Streamlit session. "
         "It does not need to be stored in GitHub.",
)

if uploaded is None:
    st.info(
        "Upload the mentor/team Excel file to start. "
        "The application is intentionally designed so the Excel file is supplied "
        "by the user through the UI."
    )

    st.markdown(
        """
### How this version works

1. Upload the current Excel workbook.
2. Select the sheet containing the MDC data.
3. The application automatically detects common column names.
4. Correct the mapping if required.
5. Review/edit the imported data.
6. Select the number of MDCs.
7. Configure each MDC separately.
8. A.2 is one cooling selection per MDC.
9. A.3/A.4/A.5 can contain multiple selections.
10. Optional components have their own quantity.
11. Standard + optional pricing is calculated.
12. Download the final BOM with and without pricing.
        """
    )
    st.stop()

# ============================================================
# READ WORKBOOK
# ============================================================

try:
    file_bytes = uploaded.getvalue()
    excel = pd.ExcelFile(BytesIO(file_bytes))
    sheet_names = excel.sheet_names

    st.success(f"Excel loaded successfully: {uploaded.name}")
    st.write(f"**Sheets found:** {len(sheet_names)}")

except Exception as e:
    st.error(f"Could not read the Excel workbook: {e}")
    st.stop()

# ============================================================
# SHEET SELECTION
# ============================================================

st.header("2. Select Excel Sheet")

selected_sheet = st.selectbox(
    "Choose the sheet containing the configuration/BOM data",
    sheet_names,
)

try:
    raw_df = pd.read_excel(BytesIO(file_bytes), sheet_name=selected_sheet)
except Exception as e:
    st.error(f"Could not read sheet '{selected_sheet}': {e}")
    st.stop()

if raw_df.empty:
    st.warning("The selected sheet is empty.")
    st.stop()

st.write(f"Rows: **{len(raw_df)}** | Columns: **{len(raw_df.columns)}**")

with st.expander("Preview original Excel data"):
    st.dataframe(raw_df.head(20), use_container_width=True)

# ============================================================
# COLUMN MAPPING
# ============================================================

st.header("3. Map Excel Columns")

auto_mapping = auto_map_columns(raw_df.columns)
column_options = ["— Not available —"] + list(raw_df.columns)

mapping = {}

cols = st.columns(3)
for i, canonical in enumerate(CANONICAL_FIELDS):
    with cols[i % 3]:
        suggested = auto_mapping.get(canonical)
        default_index = (
            column_options.index(suggested)
            if suggested in column_options
            else 0
        )

        selected = st.selectbox(
            canonical,
            column_options,
            index=default_index,
            key=f"map_{canonical}",
        )
        mapping[canonical] = None if selected == "— Not available —" else selected

if mapping.get("Part Number") is None or mapping.get("Description") is None:
    st.warning(
        "Part Number and Description are the most important BOM fields. "
        "Please map them if they exist in the workbook."
    )

# ============================================================
# PREPARE DATA
# ============================================================

df = prepare_dataframe(raw_df, mapping, selected_sheet)
df = build_catalog(df)

# If there is no useful MDC Type information, allow a manual default.
if not df["MDC Type"].astype(str).str.strip().any():
    manual_type = st.selectbox(
        "Default MDC Type for this sheet",
        ["Single Rack MDC", "Multi Rack MDC"],
    )
    df["MDC Type"] = manual_type

# If no configuration information exists, create a default configuration.
if not df["Configuration"].astype(str).str.strip().any():
    manual_config = st.text_input("Default Configuration for this sheet", "Config 1")
    df["Configuration"] = manual_config

# ============================================================
# DATA CORRECTION
# ============================================================

st.header("4. Review / Correct Imported Data")

st.caption(
    "You can correct missing values here before configuring the MDC. "
    "This does not modify the original Excel file."
)

editable_columns = CANONICAL_FIELDS
edited_df = st.data_editor(
    df[editable_columns],
    use_container_width=True,
    num_rows="dynamic",
    key="mdc_data_editor",
)

df = build_catalog(edited_df)

# ============================================================
# MDC CONFIGURATION
# ============================================================

st.header("5. MDC Configuration")

types = get_types(df)
mdc_count = st.number_input(
    "Number of MDCs required",
    min_value=1,
    max_value=50,
    value=1,
    step=1,
    help="A.1 quantity is controlled here. You do not need to enter A.1 quantity again for MDC 1, MDC 2, etc.",
)

st.caption(
    "Each MDC is configured independently. Standard configuration data is taken "
    "from the uploaded Excel workbook."
)

all_records = []
configuration_summary = []
cost_rows = []

for mdc_no in range(1, int(mdc_count) + 1):
    st.subheader(f"MDC {mdc_no}")

    type_key = f"mdc_type_{mdc_no}"
    config_key = f"mdc_config_{mdc_no}"

    selected_type = st.selectbox(
        "MDC Type",
        types,
        key=type_key,
    )

    configs = get_configs(df, selected_type)

    if not configs:
        st.warning(
            f"No configurations were found for {selected_type}. "
            "Check the Configuration column mapping/data."
        )
        continue

    selected_config = st.selectbox(
        "Configuration",
        configs,
        key=config_key,
    )

    selected_rows = rows_for_selection(df, selected_type, selected_config)

    if selected_rows.empty:
        st.warning("No BOM rows found for this MDC/configuration.")
        continue

    standard_rows = selected_rows[
        ~selected_rows["_OptionalFlag"]
    ].copy()

    optional_rows = selected_rows[
        selected_rows["_OptionalFlag"]
    ].copy()

    # --------------------------------------------------------
    # STANDARD BOM
    # --------------------------------------------------------

    with st.expander(f"MDC {mdc_no} — Standard Components", expanded=True):
        if standard_rows.empty:
            st.info("No standard components found.")
        else:
            display_standard = standard_rows[
                [
                    "Category", "Part Number", "Description",
                    "Quantity", "Unit Price", "Component Type"
                ]
            ].copy()

            display_standard["Unit Price"] = display_standard["Unit Price"].apply(
                lambda x: "XXX" if pd.isna(numeric_value(x)) else format_money(numeric_value(x))
            )

            st.dataframe(display_standard, use_container_width=True)

    # --------------------------------------------------------
    # CATEGORY-BASED CONFIGURATION
    # --------------------------------------------------------

    category_values = [
        c for c in selected_rows["Category"].dropna().astype(str).unique()
        if c.strip()
    ]

    if category_values:
        st.markdown("#### Category Selections")

        for category in category_values:
            category_rows = selected_rows[
                selected_rows["Category"].astype(str) == category
            ].copy()

            # A.2 = exactly one cooling choice
            if category.strip().upper().startswith("A.2"):
                cooling_rows = category_rows[
                    ~category_rows["_OptionalFlag"]
                ].copy()

                if len(cooling_rows) > 1:
                    options = [
                        f"{r['Description']} | {r['Part Number']}"
                        for _, r in cooling_rows.iterrows()
                    ]

                    chosen = st.radio(
                        f"{category} — Cooling (choose one)",
                        options,
                        key=f"cooling_{mdc_no}_{re.sub(r'[^a-zA-Z0-9]', '_', category)}",
                    )

                    chosen_idx = options.index(chosen)
                    chosen_row = cooling_rows.iloc[chosen_idx].to_dict()
                    chosen_row["_MDCNo"] = mdc_no
                    chosen_row["_SelectedQty"] = chosen_row.get("Quantity", 1)
                    chosen_row["_OptionalFlag"] = False
                    all_records.append(chosen_row)

                    # Add any non-cooling standard rows from this category only if
                    # there was no multiple-choice cooling competition.
                    other_rows = cooling_rows.drop(cooling_rows.index[chosen_idx])
                    if not other_rows.empty:
                        # intentionally not added because A.2 is one choice
                        pass
                elif len(cooling_rows) == 1:
                    r = cooling_rows.iloc[0].to_dict()
                    r["_MDCNo"] = mdc_no
                    r["_SelectedQty"] = r.get("Quantity", 1)
                    r["_OptionalFlag"] = False
                    all_records.append(r)

            # A.3/A.4/A.5 = multiple selections
            elif re.match(r"^A\.[345]", category.strip(), re.I):
                st.markdown(f"**{category} — Multiple selections allowed**")
                candidates = category_rows[category_rows["_OptionalFlag"]].copy()

                if candidates.empty:
                    candidates = category_rows.copy()

                chosen_records = render_component_quantity_table(
                    candidates,
                    prefix=f"mdc_{mdc_no}_{re.sub(r'[^a-zA-Z0-9]', '_', category)}",
                )

                for r in chosen_records:
                    r["_MDCNo"] = mdc_no
                    all_records.append(r)

            else:
                # Other categories are informational/standard.
                pass

    # --------------------------------------------------------
    # OPTIONAL COMPONENTS
    # --------------------------------------------------------

    with st.expander(f"MDC {mdc_no} — Optional Components", expanded=True):
        optional_selected = render_component_quantity_table(
            optional_rows,
            prefix=f"optional_{mdc_no}",
        )

        for r in optional_selected:
            r["_MDCNo"] = mdc_no
            r["_OptionalFlag"] = True
            all_records.append(r)

    # --------------------------------------------------------
    # STANDARD ROWS NOT COVERED BY A.2 RADIO
    # --------------------------------------------------------

    # Add standard rows from categories other than A.2/A.3/A.4/A.5.
    for _, r in standard_rows.iterrows():
        category = clean_text(r["Category"]).upper()

        if category.startswith("A.2"):
            # Already handled by A.2 selector above.
            continue

        if re.match(r"^A\.[345]", category, re.I):
            # These are explicitly selectable if optional; do not automatically
            # add them because the user may choose a subset.
            continue

        rec = r.to_dict()
        rec["_MDCNo"] = mdc_no
        rec["_SelectedQty"] = rec.get("Quantity", 1)
        rec["_OptionalFlag"] = False
        all_records.append(rec)

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    configuration_summary.append({
        "MDC No": mdc_no,
        "MDC Type": selected_type,
        "Configuration": selected_config,
    })

# ============================================================
# BOM + PRICING
# ============================================================

st.header("6. BOM & Pricing")

bom = build_bom(all_records)

if bom.empty:
    st.info("Configure at least one MDC to generate the BOM.")
    st.stop()

bom_with_pricing = bom.copy()

# Numeric totals
numeric_extended = pd.to_numeric(
    bom_with_pricing["Extended Price"],
    errors="coerce"
)

known_total = numeric_extended.sum(min_count=1)
unknown_price_rows = numeric_extended.isna().sum()

if pd.isna(known_total):
    known_total = math.nan

optional_bom = bom_with_pricing[
    bom_with_pricing["Optional"].astype(str).str.lower() == "yes"
].copy()

standard_bom = bom_with_pricing[
    bom_with_pricing["Optional"].astype(str).str.lower() != "yes"
].copy()

standard_total = pd.to_numeric(
    standard_bom["Extended Price"], errors="coerce"
).sum(min_count=1)

optional_total = pd.to_numeric(
    optional_bom["Extended Price"], errors="coerce"
).sum(min_count=1)

if pd.isna(standard_total):
    standard_total = math.nan
if pd.isna(optional_total):
    optional_total = math.nan

# USD display only
def convert_value(value):
    if pd.isna(value):
        return math.nan
    return value * inr_to_usd

display_standard_total = convert_value(standard_total)
display_optional_total = convert_value(optional_total)
display_total = convert_value(known_total)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Standard Cost",
        "XXX" if pd.isna(display_standard_total)
        else format_money(display_standard_total, currency),
    )

with c2:
    st.metric(
        "Optional Cost",
        "XXX" if pd.isna(display_optional_total)
        else format_money(display_optional_total, currency),
    )

with c3:
    st.metric(
        "Total Cost",
        "XXX" if pd.isna(display_total)
        else format_money(display_total, currency),
    )

with c4:
    st.metric("BOM Lines", len(bom_with_pricing))

if unknown_price_rows > 0:
    st.warning(
        f"{unknown_price_rows} BOM line(s) have non-numeric prices such as "
        "`XXX`. Those lines are included in the BOM but are not included in "
        "the numeric total."
    )

st.subheader("Generated BOM")

display_bom = bom_with_pricing.copy()

# Display converted currency without changing export values
if currency == "USD":
    for col in ["Unit Price", "Extended Price"]:
        display_bom[col] = display_bom[col].apply(
            lambda x: "XXX"
            if pd.isna(numeric_value(x))
            else format_money(convert_value(numeric_value(x)), currency)
        )
else:
    for col in ["Unit Price", "Extended Price"]:
        display_bom[col] = display_bom[col].apply(
            lambda x: "XXX"
            if pd.isna(numeric_value(x))
            else format_money(numeric_value(x), currency)
        )

st.dataframe(display_bom, use_container_width=True)

# ============================================================
# EXPORT
# ============================================================

st.header("7. Download Excel BOM")

bom_without_pricing = bom_with_pricing[
    [
        "MDC No", "MDC Type", "Configuration", "Category",
        "Part Number", "Description", "Quantity", "Optional"
    ]
].copy()

optional_export = bom_with_pricing[
    bom_with_pricing["Optional"].astype(str).str.lower() == "yes"
].copy()

summary_df = pd.DataFrame(configuration_summary)

cost_summary = pd.DataFrame([
    {
        "Cost Type": "Standard Components",
        "Amount INR": standard_total if not pd.isna(standard_total) else "XXX",
        "Amount Display Currency": (
            display_standard_total if not pd.isna(display_standard_total) else "XXX"
        ),
    },
    {
        "Cost Type": "Optional Components",
        "Amount INR": optional_total if not pd.isna(optional_total) else "XXX",
        "Amount Display Currency": (
            display_optional_total if not pd.isna(display_optional_total) else "XXX"
        ),
    },
    {
        "Cost Type": "Total",
        "Amount INR": known_total if not pd.isna(known_total) else "XXX",
        "Amount Display Currency": (
            display_total if not pd.isna(display_total) else "XXX"
        ),
    },
])

source_data = df.copy()

output = BytesIO()

try:
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Configuration Summary",
        )

        bom_with_pricing.to_excel(
            writer,
            index=False,
            sheet_name="BOM_With_Pricing",
        )

        bom_without_pricing.to_excel(
            writer,
            index=False,
            sheet_name="BOM_Without_Pricing",
        )

        optional_export.to_excel(
            writer,
            index=False,
            sheet_name="Optional Components",
        )

        cost_summary.to_excel(
            writer,
            index=False,
            sheet_name="Cost Summary",
        )

        source_data.to_excel(
            writer,
            index=False,
            sheet_name="Imported Data",
        )

        style_workbook(writer)

    output.seek(0)

    st.download_button(
        "⬇️ Download Complete BOM Excel",
        data=output.getvalue(),
        file_name="Eaton_MDC_BOM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

except Exception as e:
    st.error(f"Could not create the Excel output: {e}")

# ============================================================
# DEBUG / DATA INSPECTION
# ============================================================

with st.expander("Technical: detected data structure"):
    st.write("Automatic column mapping:")
    st.json({k: v for k, v in mapping.items()})

    st.write("Detected MDC types:")
    st.write(types)

    st.write("Detected configurations:")
    for t in types:
        st.write(f"**{t}:** {get_configs(df, t)}")

st.divider()
st.caption(
    "The Excel workbook uploaded through the UI is the data source. "
    "No online currency API is used, so the application can also operate "
    "without internet access when Streamlit itself is available."
)
