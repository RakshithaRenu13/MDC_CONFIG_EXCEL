import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="MDC BOM Generator",
    page_icon="🖥️",
    layout="wide"
)


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = "1 Rack SKU'S - MDC BOQ (01.09.2026).xlsx"
SHEET_NAME = "1R MDC (4 Configs) BOM"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """Convert Excel values into clean strings."""
    if pd.isna(value):
        return ""
    return str(value).replace("\n", " ").replace("\xa0", " ").strip()


def to_number(value):
    """Safely convert Excel value to float."""
    if pd.isna(value) or value == "":
        return 0.0

    try:
        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).replace(",", "").strip()

        return float(text)

    except Exception:
        return 0.0


def is_solution_header(value):
    """Check whether a cell contains a solution heading."""
    text = clean_text(value).upper()
    return "SOLUTION" in text


def normalize_solution_name(text):
    """
    Convert:
    SOLUTION 1
    (3.5 kw Cooling W/o Dehumidifier)

    into a clean display name.
    """
    text = clean_text(text)

    text = re.sub(r"\s+", " ", text)

    return text


def is_part_number(value):
    """
    Identify a probable part number.

    Allows:
    801029209
    CTO3M002
    HRD-XH1C
    etc.
    """
    value = clean_text(value)

    if not value:
        return False

    # Typical SKU / part number patterns
    return bool(
        re.match(r"^[A-Za-z0-9][A-Za-z0-9._/-]{2,30}$", value)
    )


# ============================================================
# LOAD EXCEL
# ============================================================

@st.cache_data
def load_excel():

    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        header=None
    )

    return df


# ============================================================
# EXTRACT SOLUTIONS
# ============================================================

def extract_solutions(df):

    solutions = {}

    # --------------------------------------------------------
    # Find all cells containing "SOLUTION"
    # --------------------------------------------------------

    solution_locations = []

    for row_idx in range(len(df)):

        for col_idx in range(len(df.columns)):

            value = clean_text(df.iloc[row_idx, col_idx])

            if is_solution_header(value):

                solution_locations.append(
                    (row_idx, col_idx, value)
                )

    # --------------------------------------------------------
    # Each solution occupies a block.
    # The current Excel has two blocks:
    #
    # LEFT  -> columns 0-4
    # RIGHT -> columns 5-9
    #
    # This function detects the starting column dynamically.
    # --------------------------------------------------------

    for index, (header_row, header_col, header_text) in enumerate(
        solution_locations
    ):

        # Determine block start.
        # Header is normally at first column of block.
        block_start = header_col

        # Expected columns:
        # Part Number, Description, Qty, UOM, Price
        part_col = block_start
        desc_col = block_start + 1
        qty_col = block_start + 2
        uom_col = block_start + 3
        price_col = block_start + 4

        # Check whether columns exist
        if price_col >= len(df.columns):
            continue

        # ----------------------------------------------------
        # Find where this solution ends.
        # It ends before:
        #   - next SOLUTION
        #   - OPTIONAL section
        #   - blank boundary
        # ----------------------------------------------------

        next_boundary = len(df)

        for r in range(header_row + 1, len(df)):

            row_values = [
                clean_text(df.iloc[r, c])
                for c in range(len(df.columns))
            ]

            joined = " ".join(row_values).upper()

            if r > header_row and (
                "SOLUTION" in joined
                or "OTHER OPTIONAL ITEMS" in joined
                or "SINGLE PHASE PDU" in joined
            ):
                next_boundary = r
                break

        # ----------------------------------------------------
        # Extract rows
        # ----------------------------------------------------

        items = []

        for r in range(header_row + 1, next_boundary):

            part_number = clean_text(df.iloc[r, part_col])
            description = clean_text(df.iloc[r, desc_col])
            qty = to_number(df.iloc[r, qty_col])
            uom = clean_text(df.iloc[r, uom_col])
            price = to_number(df.iloc[r, price_col])

            # Ignore header rows
            if part_number.upper() == "PART NUMBER":
                continue

            # Ignore completely empty rows
            if not part_number and not description:
                continue

            # Need at least a description
            if not description:
                continue

            # Avoid accidentally reading another section
            if "OPTIONAL ITEMS" in description.upper():
                continue

            if "SINGLE PHASE PDU" in description.upper():
                continue

            items.append({
                "Part Number": part_number,
                "Description": description,
                "Qty": qty,
                "UOM": uom,
                "Unit Price": price,
                "Amount": qty * price
            })

        # ----------------------------------------------------
        # Store only valid solution
        # ----------------------------------------------------

        if items:

            # Make key unique
            solution_key = normalize_solution_name(header_text)

            solutions[solution_key] = items

    return solutions


# ============================================================
# EXTRACT OPTIONAL ITEMS
# ============================================================

def extract_optional_items(df):

    optional_items = []

    optional_row = None

    # Find optional section
    for r in range(len(df)):

        for c in range(len(df.columns)):

            value = clean_text(df.iloc[r, c]).upper()

            if "OTHER OPTIONAL ITEMS" in value:
                optional_row = r
                break

        if optional_row is not None:
            break

    if optional_row is None:
        return optional_items

    # Optional section is normally first 5 columns
    part_col = 0
    desc_col = 1
    qty_col = 2
    uom_col = 3
    price_col = 4

    for r in range(optional_row + 1, len(df)):

        # Stop when PDU section starts
        row_text = " ".join(
            clean_text(df.iloc[r, c])
            for c in range(len(df.columns))
        ).upper()

        if "SINGLE PHASE PDU" in row_text:
            break

        part_number = clean_text(df.iloc[r, part_col])
        description = clean_text(df.iloc[r, desc_col])
        qty = to_number(df.iloc[r, qty_col])
        uom = clean_text(df.iloc[r, uom_col])
        price = to_number(df.iloc[r, price_col])

        if not description:
            continue

        optional_items.append({
            "Part Number": part_number,
            "Description": description,
            "Qty": qty,
            "UOM": uom,
            "Unit Price": price,
            "Amount": qty * price
        })

    return optional_items


# ============================================================
# EXTRACT PDU ITEMS
# ============================================================

def extract_pdu_items(df):

    pdu_items = []

    pdu_row = None

    # --------------------------------------------------------
    # Find PDU header
    # --------------------------------------------------------

    for r in range(len(df)):

        row_text = " ".join(
            clean_text(df.iloc[r, c])
            for c in range(len(df.columns))
        ).upper()

        if "SINGLE PHASE PDU" in row_text:

            pdu_row = r
            break

    if pdu_row is None:
        return pdu_items

    # --------------------------------------------------------
    # PDU structure:
    #
    # Part Number
    # Description
    # C13
    # C19
    # TYPE
    # Price
    # --------------------------------------------------------

    for r in range(pdu_row + 1, len(df)):

        part_number = clean_text(df.iloc[r, 0])
        description = clean_text(df.iloc[r, 1])
        c13 = to_number(df.iloc[r, 2])
        c19 = to_number(df.iloc[r, 3])
        pdu_type = clean_text(df.iloc[r, 4])
        price = to_number(df.iloc[r, 5])

        if not part_number or not description:
            continue

        # Ignore accidental rows
        if part_number.upper() in [
            "PART NUMBER",
            "DESCRIPTION"
        ]:
            continue

        pdu_items.append({
            "Part Number": part_number,
            "Description": description,
            "C13": int(c13) if c13.is_integer() else c13,
            "C19": int(c19) if c19.is_integer() else c19,
            "Type": pdu_type,
            "Unit Price": price,
            "Amount": price
        })

    # --------------------------------------------------------
    # Fill down PDU type where Excel has blank merged cells
    # --------------------------------------------------------

    previous_type = ""

    for item in pdu_items:

        if item["Type"]:
            previous_type = item["Type"]

        else:
            item["Type"] = previous_type

    return pdu_items


# ============================================================
# FORMAT CURRENCY
# ============================================================

def currency(value, symbol="₹"):

    try:
        return f"{symbol}{value:,.2f}"
    except Exception:
        return f"{symbol}0.00"


# ============================================================
# EXCEL EXPORT
# ============================================================

def create_excel_file(bom_df, summary_df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        bom_df.to_excel(
            writer,
            sheet_name="BOM",
            index=False
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

    output.seek(0)

    return output


# ============================================================
# APP
# ============================================================

st.title("🖥️ MDC Rack BOM Generator")

st.caption(
    "BOM and pricing are extracted directly from the uploaded Excel configuration."
)


# ============================================================
# LOAD DATA
# ============================================================

try:

    df = load_excel()

    solutions = extract_solutions(df)

    optional_items = extract_optional_items(df)

    pdu_items = extract_pdu_items(df)

except Exception as e:

    st.error(f"Unable to load Excel data: {e}")

    st.stop()


# ============================================================
# DATA VALIDATION
# ============================================================

if not solutions:

    st.error(
        "No solution configurations were detected in the Excel file."
    )

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Configuration")


# ------------------------------------------------------------
# Customer
# ------------------------------------------------------------

customer_name = st.sidebar.text_input(
    "Customer Name",
    placeholder="Enter customer name"
)


project_name = st.sidebar.text_input(
    "Project Name",
    placeholder="Enter project name"
)


# ------------------------------------------------------------
# Currency
# ------------------------------------------------------------

currency_option = st.sidebar.selectbox(
    "Currency",
    ["INR", "USD"]
)


usd_rate = 1.0

if currency_option == "USD":

    usd_rate = st.sidebar.number_input(
        "INR → USD Conversion Rate",
        min_value=0.0001,
        value=0.0119,
        step=0.0001,
        format="%.4f"
    )


def convert_price(value):

    if currency_option == "INR":
        return value

    return value * usd_rate


currency_symbol = "₹" if currency_option == "INR" else "$"


# ============================================================
# SOLUTION SELECTION
# ============================================================

st.header("1️⃣ Select MDC Solution")

solution_names = list(solutions.keys())

selected_solution = st.selectbox(
    "Available Solutions",
    solution_names
)


selected_items = solutions[selected_solution]


# ============================================================
# SOLUTION DETAILS
# ============================================================

st.subheader("Selected Solution")

solution_df = pd.DataFrame(selected_items)

display_solution_df = solution_df.copy()

display_solution_df["Unit Price"] = display_solution_df[
    "Unit Price"
].apply(convert_price)

display_solution_df["Amount"] = display_solution_df[
    "Amount"
].apply(convert_price)


st.dataframe(
    display_solution_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# BASE COST
# ============================================================

base_cost_inr = sum(
    item["Amount"]
    for item in selected_items
)

base_cost = convert_price(base_cost_inr)


# ============================================================
# OPTIONAL ITEMS
# ============================================================

st.header("2️⃣ Optional Components")

selected_optional = []

if optional_items:

    for index, item in enumerate(optional_items):

        label = (
            f"{item['Part Number']} — "
            f"{item['Description']} "
            f"({currency(convert_price(item['Unit Price']), currency_symbol)})"
        )

        checked = st.checkbox(
            label,
            key=f"optional_{index}"
        )

        if checked:

            selected_optional.append(item)

else:

    st.info("No optional components were found in the Excel file.")


optional_cost_inr = sum(
    item["Amount"]
    for item in selected_optional
)

optional_cost = convert_price(optional_cost_inr)


# ============================================================
# PDU SELECTION
# ============================================================

st.header("3️⃣ PDU Selection")

selected_pdu = None

if pdu_items:

    pdu_options = [
        (
            f"{item['Part Number']} — "
            f"{item['Description']} — "
            f"{item['Type']} — "
            f"C13: {item['C13']} / "
            f"C19: {item['C19']} — "
            f"{currency(convert_price(item['Unit Price']), currency_symbol)}"
        )
        for item in pdu_items
    ]

    selected_pdu_index = st.selectbox(
        "Select Single Phase PDU",
        range(len(pdu_items)),
        format_func=lambda i: pdu_options[i]
    )

    selected_pdu = pdu_items[selected_pdu_index]

else:

    st.warning(
        "No PDU items were detected in the Excel file."
    )


pdu_cost_inr = (
    selected_pdu["Amount"]
    if selected_pdu
    else 0
)

pdu_cost = convert_price(pdu_cost_inr)


# ============================================================
# TOTAL COST
# ============================================================

total_cost_inr = (
    base_cost_inr
    + optional_cost_inr
    + pdu_cost_inr
)

total_cost = convert_price(total_cost_inr)


# ============================================================
# COST SUMMARY
# ============================================================

st.header("4️⃣ Cost Summary")


col1, col2, col3, col4 = st.columns(4)


with col1:

    st.metric(
        "Base Solution",
        currency(base_cost, currency_symbol)
    )


with col2:

    st.metric(
        "Optional Items",
        currency(optional_cost, currency_symbol)
    )


with col3:

    st.metric(
        "PDU",
        currency(pdu_cost, currency_symbol)
    )


with col4:

    st.metric(
        "Total BOM Cost",
        currency(total_cost, currency_symbol)
    )


# ============================================================
# BUILD FINAL BOM
# ============================================================

final_bom = []


# ------------------------------------------------------------
# Base solution
# ------------------------------------------------------------

for item in selected_items:

    final_bom.append({
        "Category": "Base Solution",
        "Part Number": item["Part Number"],
        "Description": item["Description"],
        "Qty": item["Qty"],
        "UOM": item["UOM"],
        "Unit Price (INR)": item["Unit Price"],
        "Amount (INR)": item["Amount"]
    })


# ------------------------------------------------------------
# Optional components
# ------------------------------------------------------------

for item in selected_optional:

    final_bom.append({
        "Category": "Optional",
        "Part Number": item["Part Number"],
        "Description": item["Description"],
        "Qty": item["Qty"],
        "UOM": item["UOM"],
        "Unit Price (INR)": item["Unit Price"],
        "Amount (INR)": item["Amount"]
    })


# ------------------------------------------------------------
# PDU
# ------------------------------------------------------------

if selected_pdu:

    final_bom.append({
        "Category": "PDU",
        "Part Number": selected_pdu["Part Number"],
        "Description": selected_pdu["Description"],
        "Qty": 1,
        "UOM": "EA",
        "Unit Price (INR)": selected_pdu["Unit Price"],
        "Amount (INR)": selected_pdu["Amount"]
    })


bom_df = pd.DataFrame(final_bom)


# ============================================================
# DISPLAY FINAL BOM
# ============================================================

st.header("5️⃣ Final BOM")


if not bom_df.empty:

    display_bom = bom_df.copy()

    if currency_option == "USD":

        display_bom["Unit Price (USD)"] = (
            display_bom["Unit Price (INR)"] * usd_rate
        )

        display_bom["Amount (USD)"] = (
            display_bom["Amount (INR)"] * usd_rate
        )

        display_bom = display_bom[
            [
                "Category",
                "Part Number",
                "Description",
                "Qty",
                "UOM",
                "Unit Price (USD)",
                "Amount (USD)"
            ]
        ]

    else:

        display_bom["Unit Price (INR)"] = (
            display_bom["Unit Price (INR)"].round(2)
        )

        display_bom["Amount (INR)"] = (
            display_bom["Amount (INR)"].round(2)
        )

    st.dataframe(
        display_bom,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# SUMMARY TABLE
# ============================================================

summary_data = [

    {
        "Component": "Base Solution",
        "Cost": base_cost
    },

    {
        "Component": "Optional Components",
        "Cost": optional_cost
    },

    {
        "Component": "PDU",
        "Cost": pdu_cost
    },

    {
        "Component": "TOTAL",
        "Cost": total_cost
    }
]


summary_df = pd.DataFrame(summary_data)


st.subheader("Cost Breakdown")

summary_display = summary_df.copy()

summary_display["Cost"] = summary_display["Cost"].apply(
    lambda x: currency(x, currency_symbol)
)

st.table(summary_display)


# ============================================================
# CUSTOMER / PROJECT DETAILS
# ============================================================

st.header("6️⃣ Project Information")

info_col1, info_col2 = st.columns(2)


with info_col1:

    st.write("**Customer:**")

    st.write(
        customer_name if customer_name else "Not specified"
    )


with info_col2:

    st.write("**Project:**")

    st.write(
        project_name if project_name else "Not specified"
    )


st.write("**Selected Solution:**")

st.write(selected_solution)


# ============================================================
# EXPORT
# ============================================================

st.header("7️⃣ Export BOM")


# Create Excel export
excel_file = create_excel_file(
    bom_df,
    summary_df
)


st.download_button(
    label="📥 Download BOM Excel",
    data=excel_file,
    file_name="MDC_BOM_Generated.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )
)


# ============================================================
# DATA SOURCE INFORMATION
# ============================================================

with st.expander("📊 Excel Data Source"):

    st.write(
        f"**Workbook:** `{EXCEL_FILE}`"
    )

    st.write(
        f"**Sheet:** `{SHEET_NAME}`"
    )

    st.write(
        f"**Detected Solutions:** {len(solutions)}"
    )

    st.write(
        f"**Optional Items:** {len(optional_items)}"
    )

    st.write(
        f"**PDU Items:** {len(pdu_items)}"
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "MDC Rack BOM Generator • Data extracted from the supplied Excel configuration"
)
