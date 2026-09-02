
import streamlit as st
import pandas as pd
from io import BytesIO
import re

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="MDC Configuration & BOM Generator",
    page_icon="🏭",
    layout="wide"
)

# ============================================================
# CONSTANTS
# ============================================================
UNKNOWN = "XXX"

ALIASES = {
    "part_number": [
        "part number", "part no", "part no.", "part_number", "partnumber",
        "pn", "item code", "item no", "item number", "material number",
        "material no", "catalog number", "catalog no"
    ],
    "description": [
        "description", "product description", "component", "component name",
        "item description", "product", "product name", "material description"
    ],
    "quantity": [
        "quantity", "qty", "qnty", "required quantity", "required qty"
    ],
    "unit_price": [
        "unit price", "price", "unit cost", "cost", "standard price",
        "selling price", "base price", "price per unit"
    ],
    "configuration": [
        "configuration", "config", "config no", "config number",
        "configuration no", "configuration number"
    ],
    "mdc_type": [
        "mdc type", "type", "product type", "mdc"
    ],
    "optional": [
        "optional", "option", "optional component", "component type",
        "category", "selection type", "status"
    ],
    "specification": [
        "specification", "spec", "value", "details", "feature",
        "configuration specification"
    ]
}

DISPLAY_NAMES = {
    "part_number": "Part Number",
    "description": "Description",
    "quantity": "Quantity",
    "unit_price": "Unit Price",
    "configuration": "Configuration",
    "mdc_type": "MDC Type",
    "optional": "Optional / Category",
    "specification": "Specification"
}

# ============================================================
# HELPERS
# ============================================================
def clean_col_name(value):
    return re.sub(r"\s+", " ", str(value).strip().lower().replace("_", " "))


def normalise_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def is_numeric_price(value):
    if value is None or pd.isna(value):
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip().replace(",", "").replace("₹", "").replace("$", "")
    if text.upper() in {"", "XXX", "NA", "N/A", "NONE", "NAN", "-"}:
        return False
    try:
        float(text)
        return True
    except ValueError:
        return False


def numeric_price(value):
    if not is_numeric_price(value):
        return None
    text = str(value).strip().replace(",", "").replace("₹", "").replace("$", "")
    return float(text)


def format_price(value, currency="INR", exchange_rate=1.0):
    amount = numeric_price(value)
    if amount is None:
        return UNKNOWN

    converted = amount * exchange_rate
    symbol = "₹" if currency == "INR" else "$"
    return f"{symbol} {converted:,.2f}"


def parse_quantity(value, default=1):
    if value is None or pd.isna(value):
        return default
    try:
        number = float(str(value).strip().replace(",", ""))
        if number.is_integer():
            return max(1, int(number))
        return max(1, int(round(number)))
    except Exception:
        return default


def safe_sheet_name(name, used_names):
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name))[:31] or "Sheet"
    original = name
    counter = 1
    while name in used_names:
        suffix = f"_{counter}"
        name = (original[:31-len(suffix)] + suffix)
        counter += 1
    used_names.add(name)
    return name


def create_excel_file(dataframes):
    output = BytesIO()
    used_names = set()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, dataframe in dataframes.items():
            final_name = safe_sheet_name(sheet_name, used_names)
            dataframe.to_excel(writer, sheet_name=final_name, index=False)

    output.seek(0)
    return output


def best_alias_match(columns, aliases):
    cleaned = {col: clean_col_name(col) for col in columns}

    # Exact match first
    for col, cleaned_name in cleaned.items():
        if cleaned_name in aliases:
            return col

    # Then partial match
    for alias in aliases:
        for col, cleaned_name in cleaned.items():
            if alias in cleaned_name or cleaned_name in alias:
                return col

    return ""


def auto_mapping(columns):
    mapping = {}
    for key, aliases in ALIASES.items():
        mapping[key] = best_alias_match(columns, aliases)
    return mapping


def value_is_optional(value):
    text = normalise_text(value).lower()
    optional_words = {
        "optional", "option", "yes", "y", "true", "1",
        "optional component", "customer option", "customer selectable"
    }
    return text in optional_words


def value_is_na(value):
    text = normalise_text(value).lower()
    return text in {"na", "n/a", "not applicable", "not available", "-", ""}


def infer_mdc_type(config_name, row_value=""):
    text = f"{config_name} {row_value}".lower()
    if "multi" in text:
        return "Multi Rack MDC"
    if "single" in text:
        return "Single Rack MDC"
    return ""


def infer_configuration_from_sheet(sheet_name):
    match = re.search(r"(config(?:uration)?\s*[\._-]?\s*\d+)", str(sheet_name), re.I)
    if match:
        raw = match.group(1)
        number = re.search(r"\d+", raw)
        if number:
            return f"Config {number.group()}"
    return str(sheet_name).strip()


def prepare_source_data(uploaded_file):
    excel = pd.ExcelFile(uploaded_file)
    sheets = {}
    combined = []

    for sheet in excel.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        df = df.dropna(how="all").copy()
        if df.empty:
            continue

        # Remove completely unnamed empty columns
        df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
        df["__Source Sheet"] = sheet

        if "Configuration" not in [str(c) for c in df.columns]:
            df["__Inferred Configuration"] = infer_configuration_from_sheet(sheet)

        sheets[sheet] = df.copy()
        combined.append(df)

    if combined:
        combined_df = pd.concat(combined, ignore_index=True, sort=False)
    else:
        combined_df = pd.DataFrame()

    return excel.sheet_names, sheets, combined_df


def get_config_values(df, config_col, sheet_col="__Source Sheet"):
    if df.empty:
        return []

    if config_col and config_col in df.columns:
        values = [
            normalise_text(v) for v in df[config_col].dropna().unique()
            if normalise_text(v)
        ]
        if values:
            return values

    if "__Inferred Configuration" in df.columns:
        values = [
            normalise_text(v) for v in df["__Inferred Configuration"].dropna().unique()
            if normalise_text(v)
        ]
        if values:
            return values

    if sheet_col in df.columns:
        return [
            normalise_text(v) for v in df[sheet_col].dropna().unique()
            if normalise_text(v)
        ]

    return []


def get_mdc_types(df, mdc_col):
    if mdc_col and mdc_col in df.columns:
        values = [
            normalise_text(v) for v in df[mdc_col].dropna().unique()
            if normalise_text(v)
        ]
        if values:
            return values
    return []


def filter_for_selection(df, mapping, selected_config, selected_mdc_type):
    result = df.copy()

    config_col = mapping.get("configuration", "")
    mdc_col = mapping.get("mdc_type", "")

    if config_col and config_col in result.columns:
        result = result[
            result[config_col].astype(str).str.strip() == str(selected_config).strip()
        ]
    elif "__Inferred Configuration" in result.columns:
        result = result[
            result["__Inferred Configuration"].astype(str).str.strip()
            == str(selected_config).strip()
        ]

    if selected_mdc_type and mdc_col and mdc_col in result.columns:
        result = result[
            result[mdc_col].astype(str).str.strip() == str(selected_mdc_type).strip()
        ]

    return result.copy()


def build_standard_bom(config_df, mapping):
    pn_col = mapping.get("part_number", "")
    desc_col = mapping.get("description", "")
    qty_col = mapping.get("quantity", "")
    price_col = mapping.get("unit_price", "")
    spec_col = mapping.get("specification", "")
    optional_col = mapping.get("optional", "")

    rows = []

    for _, row in config_df.iterrows():
        if optional_col and optional_col in config_df.columns:
            if value_is_optional(row.get(optional_col)):
                continue

        description = normalise_text(row.get(desc_col)) if desc_col else ""
        if not description:
            continue

        part_number = normalise_text(row.get(pn_col)) if pn_col else UNKNOWN
        specification = normalise_text(row.get(spec_col)) if spec_col else ""

        quantity = parse_quantity(row.get(qty_col), 1) if qty_col else 1
        unit_price = row.get(price_col) if price_col else UNKNOWN

        rows.append({
            "Category": "Standard",
            "Part Number": part_number or UNKNOWN,
            "Component": description,
            "Specification": specification,
            "Quantity": quantity,
            "Unit Price": numeric_price(unit_price) if is_numeric_price(unit_price) else UNKNOWN
        })

    return pd.DataFrame(rows)


def build_optional_catalog(config_df, all_df, mapping):
    desc_col = mapping.get("description", "")
    pn_col = mapping.get("part_number", "")
    qty_col = mapping.get("quantity", "")
    price_col = mapping.get("unit_price", "")
    optional_col = mapping.get("optional", "")

    source = config_df.copy()

    if optional_col and optional_col in source.columns:
        optional_mask = source[optional_col].apply(value_is_optional)
        source = source[optional_mask]

    # If no explicit optional column exists, use rows whose description contains
    # common optional wording.
    if source.empty and desc_col:
        source = config_df[
            config_df[desc_col].astype(str).str.contains(
                r"optional|option|vesda|cctv|ups|pdu|fire suppression|dcim|branch circuit",
                case=False,
                na=False
            )
        ].copy()

    # If the selected config contains no optional rows, look through the source
    # workbook for a dedicated optional table.
    if source.empty and optional_col and optional_col in all_df.columns:
        source = all_df[
            all_df[optional_col].apply(value_is_optional)
        ].copy()

    catalog = []

    if desc_col:
        seen = set()
        for _, row in source.iterrows():
            description = normalise_text(row.get(desc_col))
            if not description or description in seen:
                continue

            seen.add(description)

            catalog.append({
                "Part Number": normalise_text(row.get(pn_col)) if pn_col else UNKNOWN,
                "Component": description,
                "Default Quantity": parse_quantity(row.get(qty_col), 1) if qty_col else 1,
                "Unit Price": numeric_price(row.get(price_col)) if price_col and is_numeric_price(row.get(price_col)) else UNKNOWN
            })

    return pd.DataFrame(catalog)


def add_selected_optional_to_bom(bom_df, selected_optional):
    rows = []
    if bom_df is not None and not bom_df.empty:
        rows.extend(bom_df.to_dict("records"))

    for item in selected_optional:
        rows.append({
            "Category": "Optional",
            "Part Number": item["part_number"],
            "Component": item["component"],
            "Specification": "Selected by User",
            "Quantity": item["quantity"],
            "Unit Price": item["unit_price"]
        })

    return pd.DataFrame(rows, columns=[
        "Category", "Part Number", "Component",
        "Specification", "Quantity", "Unit Price"
    ])


def calculate_bom_totals(bom_df):
    standard_total = 0.0
    optional_total = 0.0
    unknown = False

    if bom_df is None or bom_df.empty:
        return standard_total, optional_total, None, unknown

    for _, row in bom_df.iterrows():
        price = numeric_price(row.get("Unit Price"))
        qty = parse_quantity(row.get("Quantity"), 1)

        if price is None:
            unknown = True
            continue

        amount = price * qty

        if str(row.get("Category", "")).lower() == "optional":
            optional_total += amount
        else:
            standard_total += amount

    total = None if unknown else standard_total + optional_total
    return standard_total, optional_total, total, unknown


# ============================================================
# TITLE
# ============================================================
st.title("🏭 MDC Configuration & BOM Generator")
st.markdown(
    """
    Upload the latest mentor Excel workbook. The application reads the workbook,
    lets you map the columns, dynamically creates the MDC configuration UI,
    supports optional components with quantities, calculates pricing, and exports
    BOM files. No configuration data needs to be hard-coded in this Python file.
    """
)

# ============================================================
# STEP 1 — EXCEL UPLOAD
# ============================================================
st.header("1️⃣ Upload MDC Excel Data")

uploaded_file = st.file_uploader(
    "Upload the Excel file received from your mentor",
    type=["xlsx", "xls"],
    help="The uploaded workbook becomes the source of truth for configurations, BOM data and prices."
)

if uploaded_file is None:
    st.info("Please upload the current MDC Excel workbook to start.")
    st.stop()

try:
    sheet_names, sheets, source_df = prepare_source_data(uploaded_file)
except Exception as exc:
    st.error(f"Unable to read the Excel workbook: {exc}")
    st.stop()

if source_df.empty:
    st.error("The workbook does not contain readable data.")
    st.stop()

st.success(
    f"Excel loaded successfully: {len(sheet_names)} sheet(s), "
    f"{len(source_df):,} data row(s)."
)

with st.expander("Preview Excel data"):
    st.write("Sheets found:", sheet_names)
    st.dataframe(source_df.head(50), use_container_width=True, hide_index=True)

# ============================================================
# STEP 2 — COLUMN MAPPING
# ============================================================
st.header("2️⃣ Map Excel Columns")

st.caption(
    "The program automatically detects common column names. "
    "If the mentor uses different names, select the correct column manually."
)

# Ignore internal helper columns from user mapping
available_columns = [
    c for c in source_df.columns
    if not str(c).startswith("__")
]

auto = auto_mapping(available_columns)

def mapping_selectbox(label, key_name, required=False):
    options = [""] + available_columns
    auto_value = auto.get(key_name, "")
    default_index = options.index(auto_value) if auto_value in options else 0

    return st.selectbox(
        f"{label}{' *' if required else ''}",
        options,
        index=default_index,
        key=f"map_{key_name}"
    )

m1, m2, m3 = st.columns(3)

with m1:
    part_number_col = mapping_selectbox("Part Number", "part_number", True)
    description_col = mapping_selectbox("Description / Component", "description", True)
    quantity_col = mapping_selectbox("Quantity", "quantity", False)

with m2:
    unit_price_col = mapping_selectbox("Unit Price / Price", "unit_price", False)
    configuration_col = mapping_selectbox("Configuration / Config", "configuration", False)
    mdc_type_col = mapping_selectbox("MDC Type", "mdc_type", False)

with m3:
    optional_col = mapping_selectbox("Optional / Category", "optional", False)
    specification_col = mapping_selectbox("Specification", "specification", False)

mapping = {
    "part_number": part_number_col,
    "description": description_col,
    "quantity": quantity_col,
    "unit_price": unit_price_col,
    "configuration": configuration_col,
    "mdc_type": mdc_type_col,
    "optional": optional_col,
    "specification": specification_col
}

if not part_number_col or not description_col:
    st.error("Part Number and Description / Component are required.")
    st.stop()

# ============================================================
# STEP 3 — DATA CORRECTION
# ============================================================
st.header("3️⃣ Review / Correct Excel Data")

st.caption(
    "The original Excel file is not modified. Use this table to correct missing "
    "or incorrect values for the current configuration session."
)

editable_df = source_df[available_columns].copy()

edited_df = st.data_editor(
    editable_df,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    key="editable_source_data"
)

# Re-add helper fields
working_df = edited_df.copy()

if "__Inferred Configuration" in source_df.columns:
    inferred_values = source_df["__Inferred Configuration"].reindex(
        working_df.index
    ).fillna("")
    working_df["__Inferred Configuration"] = inferred_values

if "__Source Sheet" in source_df.columns:
    source_values = source_df["__Source Sheet"].reindex(
        working_df.index
    ).fillna("")
    working_df["__Source Sheet"] = source_values

# ============================================================
# STEP 4 — MDC TYPE
# ============================================================
st.header("4️⃣ Select MDC Type")

mdc_values = get_mdc_types(working_df, mdc_type_col)

if mdc_values:
    selected_mdc_type = st.selectbox(
        "MDC Type",
        mdc_values,
        key="selected_mdc_type"
    )
else:
    selected_mdc_type = st.selectbox(
        "MDC Type",
        ["Single Rack MDC", "Multi Rack MDC"],
        key="selected_mdc_type_fallback"
    )
    st.caption(
        "No MDC Type column was detected. The type above is used only for the current BOM."
    )

# ============================================================
# STEP 5 — CONFIGURATION
# ============================================================
st.header("5️⃣ Select Configuration")

config_values = get_config_values(working_df, configuration_col)

if not config_values:
    st.error(
        "No configuration values were found. "
        "Map the Configuration / Config column, or make sure the workbook sheet names "
        "represent the configurations."
    )
    st.stop()

selected_config = st.selectbox(
    "Configuration",
    config_values,
    key="selected_configuration"
)

config_df = filter_for_selection(
    working_df,
    mapping,
    selected_config,
    selected_mdc_type
)

if config_df.empty:
    st.warning(
        "No rows matched the selected MDC type and configuration. "
        "The application will try the configuration without applying MDC type filtering."
    )
    config_df = filter_for_selection(
        working_df,
        {**mapping, "mdc_type": ""},
        selected_config,
        ""
    )

if config_df.empty:
    st.error("No data is available for the selected configuration.")
    st.stop()

# ============================================================
# STEP 6 — CONFIGURATION DATA
# ============================================================
st.header("6️⃣ Configuration Details")

display_columns = [
    c for c in [
        configuration_col,
        mdc_type_col,
        part_number_col,
        description_col,
        specification_col,
        quantity_col,
        unit_price_col,
        optional_col
    ]
    if c and c in config_df.columns
]

# Remove duplicate columns while preserving order
display_columns = list(dict.fromkeys(display_columns))

st.dataframe(
    config_df[display_columns],
    use_container_width=True,
    hide_index=True
)

# ============================================================
# STEP 7 — STANDARD BOM
# ============================================================
st.header("7️⃣ Standard BOM")

standard_bom = build_standard_bom(config_df, mapping)

if standard_bom.empty:
    st.warning(
        "No standard BOM rows were identified. Check the Description, Optional and Quantity mappings."
    )
else:
    standard_bom_display = standard_bom.copy()
    standard_bom_display["Unit Price"] = standard_bom_display["Unit Price"].apply(
        lambda x: format_price(x, "INR", 1.0)
    )
    st.dataframe(
        standard_bom_display,
        use_container_width=True,
        hide_index=True
    )

# ============================================================
# STEP 8 — OPTIONAL COMPONENTS
# ============================================================
st.header("8️⃣ Optional Components")

optional_catalog = build_optional_catalog(
    config_df,
    working_df,
    mapping
)

selected_optional = []

if optional_catalog.empty:
    st.info(
        "No optional components were detected from this configuration. "
        "If the workbook has a separate optional-component table, ensure its "
        "Optional / Category column contains values such as Optional or Option."
    )
else:
    st.info(
        "Tick the components required by the customer and enter the required quantity. "
        "Optional amount = Quantity × Unit Price."
    )

    for idx, item in optional_catalog.iterrows():
        component = normalise_text(item["Component"])
        part_number = normalise_text(item["Part Number"]) or UNKNOWN
        unit_price = item["Unit Price"]

        c1, c2, c3, c4 = st.columns([0.7, 3.6, 1.5, 1.5])

        with c1:
            checked = st.checkbox(
                "",
                key=f"opt_check_{selected_mdc_type}_{selected_config}_{idx}"
            )

        with c2:
            st.write(f"**{component}**")
            st.caption(f"Part Number: {part_number}")

        with c3:
            st.write(f"Unit Price: **{format_price(unit_price)}**")

        with c4:
            if checked:
                quantity = st.number_input(
                    "Quantity",
                    min_value=1,
                    value=parse_quantity(item["Default Quantity"], 1),
                    step=1,
                    key=f"opt_qty_{selected_mdc_type}_{selected_config}_{idx}"
                )
            else:
                quantity = 0

        if checked:
            selected_optional.append({
                "part_number": part_number,
                "component": component,
                "quantity": int(quantity),
                "unit_price": unit_price
            })

if selected_optional:
    optional_display = pd.DataFrame([
        {
            "Part Number": x["part_number"],
            "Component": x["component"],
            "Quantity": x["quantity"],
            "Unit Price": format_price(x["unit_price"]),
            "Amount": (
                format_price(numeric_price(x["unit_price"]) * x["quantity"])
                if numeric_price(x["unit_price"]) is not None
                else UNKNOWN
            )
        }
        for x in selected_optional
    ])

    st.subheader("Selected Optional Components")
    st.dataframe(
        optional_display,
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("No optional components selected.")

# ============================================================
# STEP 9 — FINAL BOM
# ============================================================
st.header("9️⃣ Final BOM")

bom_df = add_selected_optional_to_bom(
    standard_bom,
    selected_optional
)

if bom_df.empty:
    st.warning("The final BOM is empty.")
else:
    bom_ui = bom_df.copy()
    bom_ui["Unit Price"] = bom_ui["Unit Price"].apply(
        lambda x: format_price(x, "INR", 1.0)
    )
    bom_ui["Amount"] = bom_df.apply(
        lambda r: (
            format_price(
                numeric_price(r["Unit Price"]) * parse_quantity(r["Quantity"], 1)
            )
            if numeric_price(r["Unit Price"]) is not None
            else UNKNOWN
        ),
        axis=1
    )

    st.dataframe(
        bom_ui,
        use_container_width=True,
        hide_index=True
    )

# ============================================================
# STEP 10 — CURRENCY
# ============================================================
st.header("🔟 Currency & Pricing")

currency = st.radio(
    "Display Currency",
    ["INR", "USD"],
    horizontal=True,
    key="display_currency"
)

if currency == "INR":
    exchange_rate = 1.0
else:
    st.caption(
        "Offline mode: enter the INR → USD rate manually. "
        "No internet connection is required."
    )
    exchange_rate = st.number_input(
        "INR to USD exchange rate (1 INR = USD)",
        min_value=0.000001,
        value=0.011,
        step=0.0001,
        format="%.6f",
        key="manual_exchange_rate"
    )

standard_total, optional_total, base_total, has_unknown = calculate_bom_totals(bom_df)

st.subheader("💵 Cost Summary")

summary_rows = [
    {
        "Cost Item": "Standard Components",
        "Amount": format_price(standard_total, currency, exchange_rate)
    },
    {
        "Cost Item": "Optional Components",
        "Amount": format_price(optional_total, currency, exchange_rate)
        if not has_unknown else UNKNOWN
    },
    {
        "Cost Item": "TOTAL COST",
        "Amount": format_price(base_total, currency, exchange_rate)
        if base_total is not None else UNKNOWN
    }
]

st.dataframe(
    pd.DataFrame(summary_rows),
    use_container_width=True,
    hide_index=True
)

# ============================================================
# STEP 11 — COST TO PRICE
# ============================================================
st.header("1️⃣1️⃣ Cost → Selling Price")

st.caption(
    "Each percentage is applied to the running amount. "
    "All names and percentages can be changed for the current quotation."
)

default_factors = [
    ("Factory Cost (COGS)", 0.0),
    ("Admin & R&D Overhead", 15.0),
    ("Marketing & Sales", 20.0),
    ("Manufacturer Profit", 15.0),
    ("Distribution & Retail", 45.0)
]

pricing_rows = []

for idx, (default_name, default_pct) in enumerate(default_factors, start=1):
    c1, c2 = st.columns([3, 1])

    with c1:
        factor_name = st.text_input(
            f"Layer {idx} — Name",
            value=default_name,
            key=f"factor_name_{idx}"
        )

    with c2:
        factor_pct = st.number_input(
            f"Layer {idx} — Percentage",
            min_value=0.0,
            max_value=1000.0,
            value=default_pct,
            step=0.5,
            key=f"factor_pct_{idx}"
        )

    pricing_rows.append({
        "Layer": idx,
        "Name": factor_name.strip(),
        "Percentage": float(factor_pct)
    })

if base_total is None:
    selling_price = None
    st.warning(
        "Selling price cannot be calculated because one or more selected BOM items "
        "have an unknown price (XXX)."
    )
else:
    running = base_total
    buildup = []

    for row in pricing_rows:
        if row["Layer"] == 1:
            previous = running
            added = 0.0
        else:
            previous = running
            added = running * row["Percentage"] / 100.0
            running += added

        buildup.append({
            "Layer": row["Layer"],
            "Name": row["Name"],
            "Percentage": (
                "Baseline"
                if row["Layer"] == 1
                else f"{row['Percentage']:.2f}%"
            ),
            "Previous Amount": format_price(previous, currency, exchange_rate),
            "Added Amount": (
                "—"
                if row["Layer"] == 1
                else format_price(added, currency, exchange_rate)
            ),
            "Cumulative Price": format_price(running, currency, exchange_rate)
        })

    selling_price = running

    st.dataframe(
        pd.DataFrame(buildup),
        use_container_width=True,
        hide_index=True
    )

st.subheader("🏷️ Final Selling Price")

if selling_price is None:
    st.error("XXX")
else:
    st.success(
        f"### {format_price(selling_price, currency, exchange_rate)}"
    )

# ============================================================
# STEP 12 — EXCEL EXPORT
# ============================================================
st.header("1️⃣2️⃣ Excel Export")

customer_col1, customer_col2 = st.columns(2)

with customer_col1:
    customer_name = st.text_input(
        "Customer Name *",
        placeholder="Enter customer / company name",
        key="customer_name_export"
    )

with customer_col2:
    customer_place = st.text_input(
        "Customer Place *",
        placeholder="Enter city / location",
        key="customer_place_export"
    )

problem_statement = st.text_area(
    "Problem / Requirement",
    placeholder="Describe the customer's requirement...",
    key="problem_export"
)

proposed_solution = st.text_area(
    "Proposed Solution",
    placeholder="Describe the proposed MDC solution...",
    key="solution_export"
)

# ============================================================
# EXPORT DATAFRAMES
# ============================================================
bom_without_price_df = bom_df[
    [
        "Category",
        "Part Number",
        "Component",
        "Specification",
        "Quantity"
    ]
].copy() if not bom_df.empty else pd.DataFrame(
    columns=[
        "Category",
        "Part Number",
        "Component",
        "Specification",
        "Quantity"
    ]
)

bom_with_price_df = bom_df.copy() if not bom_df.empty else pd.DataFrame(
    columns=[
        "Category",
        "Part Number",
        "Component",
        "Specification",
        "Quantity",
        "Unit Price"
    ]
)

if not bom_with_price_df.empty:
    bom_with_price_df["Amount"] = bom_with_price_df.apply(
        lambda r: (
            numeric_price(r["Unit Price"]) * parse_quantity(r["Quantity"], 1)
            if numeric_price(r["Unit Price"]) is not None
            else None
        ),
        axis=1
    )

    bom_with_price_df["Unit Price Display"] = bom_with_price_df[
        "Unit Price"
    ].apply(
        lambda x: format_price(x, currency, exchange_rate)
    )

    bom_with_price_df["Amount Display"] = bom_with_price_df[
        "Amount"
    ].apply(
        lambda x: format_price(x, currency, exchange_rate)
        if x is not None else UNKNOWN
    )

optional_export_df = pd.DataFrame([
    {
        "Part Number": x["part_number"],
        "Component": x["component"],
        "Quantity": x["quantity"],
        "Unit Price": (
            numeric_price(x["unit_price"])
            if numeric_price(x["unit_price"]) is not None
            else UNKNOWN
        ),
        "Amount": (
            numeric_price(x["unit_price"]) * x["quantity"]
            if numeric_price(x["unit_price"]) is not None
            else UNKNOWN
        )
    }
    for x in selected_optional
])

if optional_export_df.empty:
    optional_export_df = pd.DataFrame(
        columns=[
            "Part Number",
            "Component",
            "Quantity",
            "Unit Price",
            "Amount"
        ]
    )

customer_details_df = pd.DataFrame([
    {"Field": "Customer Name", "Value": customer_name},
    {"Field": "Customer Place", "Value": customer_place},
    {"Field": "MDC Type", "Value": selected_mdc_type},
    {"Field": "Configuration", "Value": selected_config},
    {"Field": "Problem / Requirement", "Value": problem_statement},
    {"Field": "Proposed Solution", "Value": proposed_solution}
])

cost_summary_df = pd.DataFrame([
    {
        "Metric": "Standard Components Cost",
        "Value": format_price(standard_total, currency, exchange_rate)
    },
    {
        "Metric": "Optional Components Cost",
        "Value": (
            format_price(optional_total, currency, exchange_rate)
            if not has_unknown else UNKNOWN
        )
    },
    {
        "Metric": "TOTAL COST",
        "Value": (
            format_price(base_total, currency, exchange_rate)
            if base_total is not None else UNKNOWN
        )
    },
    {
        "Metric": "FINAL SELLING PRICE",
        "Value": (
            format_price(selling_price, currency, exchange_rate)
            if selling_price is not None else UNKNOWN
        )
    }
])

mapping_export_df = pd.DataFrame([
    {
        "Application Field": DISPLAY_NAMES[key],
        "Excel Column": value or "(not mapped)"
    }
    for key, value in mapping.items()
])

source_export_df = working_df.copy()

price_build_up_export = pd.DataFrame(
    buildup if base_total is not None else []
)

# ============================================================
# VALIDATION
# ============================================================
if not customer_name.strip() or not customer_place.strip():
    st.warning("Customer Name and Customer Place are compulsory for the final quotation.")

# ============================================================
# DOWNLOAD 1 — BOM WITHOUT PRICE
# ============================================================
excel_without_price = create_excel_file({
    "Customer Details": customer_details_df,
    "BOM": bom_without_price_df,
    "Source Mapping": mapping_export_df
})

st.download_button(
    "📥 Download BOM — Without Price",
    data=excel_without_price,
    file_name=(
        f"{selected_mdc_type.replace(' ', '_')}_"
        f"{str(selected_config).replace(' ', '_')}_"
        f"BOM_Without_Price.xlsx"
    ),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_without_price"
)

# ============================================================
# DOWNLOAD 2 — BOM WITH PRICE
# ============================================================
excel_with_price = create_excel_file({
    "Customer Details": customer_details_df,
    "BOM With Price": bom_with_price_df,
    "Optional Components": optional_export_df,
    "Cost Summary": cost_summary_df,
    "Price Build-up": price_build_up_export,
    "Source Mapping": mapping_export_df,
    "Source Data": source_export_df
})

st.download_button(
    "📥 Download BOM — With Price",
    data=excel_with_price,
    file_name=(
        f"{selected_mdc_type.replace(' ', '_')}_"
        f"{str(selected_config).replace(' ', '_')}_"
        f"BOM_With_Price.xlsx"
    ),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="download_with_price"
)

# ============================================================
# FOOTER
# ============================================================
st.divider()
st.caption(
    "MDC Configuration & BOM Generator | "
    "Excel is the source of truth. Configuration, BOM, part numbers, quantities "
    "and prices are read dynamically from the uploaded workbook. "
    "The application does not require internet access."
)
