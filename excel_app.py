import streamlit as st
import pandas as pd
import os
import re
import requests
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
# FILE CONFIGURATION
# ============================================================

EXCEL_FILE = "1 Rack SKU'S - MDC BOQ (01.09.2026).xlsx"
SHEET_NAME = "1R MDC (4 Configs) BOM"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """Clean Excel cell value."""

    if pd.isna(value):
        return ""

    return (
        str(value)
        .replace("\n", " ")
        .replace("\xa0", " ")
        .strip()
    )


def is_numeric_price(value):
    """Return True if value can be converted to a number."""

    if isinstance(value, bool):
        return False

    if isinstance(value, (int, float)):
        return True

    if isinstance(value, str):

        try:
            float(
                value
                .strip()
                .replace(",", "")
            )

            return True

        except (ValueError, TypeError):
            return False

    return False


def numeric_price(value):
    """Convert numeric price to float."""

    if not is_numeric_price(value):
        return None

    return float(
        str(value)
        .strip()
        .replace(",", "")
    )


def to_number(value):
    """Safely convert Excel value to number."""

    if pd.isna(value) or value == "":
        return 0.0

    try:

        if isinstance(value, (int, float)):
            return float(value)

        return float(
            str(value)
            .replace(",", "")
            .strip()
        )

    except Exception:

        return 0.0


def calculate_amount(unit_price, quantity):
    """
    Amount = Quantity × Unit Price
    """

    price = numeric_price(unit_price)

    if price is None:
        return None

    try:
        qty = float(quantity)
    except Exception:
        qty = 0

    return price * qty


# ============================================================
# LIVE CURRENCY
# ============================================================

@st.cache_data(ttl=3600)
def get_live_exchange_rate(from_currency, to_currency):

    if from_currency == to_currency:
        return 1.0

    try:

        url = (
            f"https://api.frankfurter.dev/v2/rate/"
            f"{from_currency}/{to_currency}"
        )

        response = requests.get(
            url,
            timeout=10
        )

        response.raise_for_status()

        data = response.json()

        return float(data["rate"])

    except Exception:

        return None


# ============================================================
# LOAD EXCEL
# ============================================================

@st.cache_data
def load_excel():

    if not os.path.exists(EXCEL_FILE):

        raise FileNotFoundError(
            f"Excel file not found:\n{EXCEL_FILE}"
        )

    return pd.read_excel(
        EXCEL_FILE,
        sheet_name=SHEET_NAME,
        header=None
    )


# ============================================================
# DETECT SOLUTION HEADERS
# ============================================================

def find_solution_headers(df):

    headers = []

    for row in range(len(df)):

        for col in range(len(df.columns)):

            value = clean_text(
                df.iloc[row, col]
            )

            if "SOLUTION" in value.upper():

                headers.append(
                    {
                        "row": row,
                        "col": col,
                        "text": value
                    }
                )

    return headers


# ============================================================
# EXTRACT SOLUTIONS
# ============================================================

def extract_solutions(df):

    solutions = {}

    headers = find_solution_headers(df)

    if not headers:
        return solutions

    # --------------------------------------------------------
    # Remove duplicate headers occurring in same block
    # --------------------------------------------------------

    processed = []

    for header in headers:

        row = header["row"]
        col = header["col"]

        duplicate = False

        for existing in processed:

            if (
                existing["row"] == row
                and abs(existing["col"] - col) <= 4
            ):
                duplicate = True
                break

        if not duplicate:

            processed.append(header)

    # --------------------------------------------------------
    # Process each solution
    # --------------------------------------------------------

    for index, header in enumerate(processed):

        start_row = header["row"]
        start_col = header["col"]

        # ----------------------------------------------------
        # Expected structure:
        #
        # Part Number
        # Description
        # QTY
        # UOM
        # Price
        # ----------------------------------------------------

        part_col = start_col
        desc_col = start_col + 1
        qty_col = start_col + 2
        uom_col = start_col + 3
        price_col = start_col + 4

        if price_col >= len(df.columns):
            continue

        # ----------------------------------------------------
        # Determine end row
        # ----------------------------------------------------

        end_row = len(df)

        for r in range(start_row + 1, len(df)):

            row_text = " ".join(
                clean_text(df.iloc[r, c])
                for c in range(len(df.columns))
            ).upper()

            # Stop at major next section
            if (
                "OTHER OPTIONAL ITEMS" in row_text
                or "SINGLE PHASE PDU" in row_text
            ):
                end_row = r
                break

            # Stop at next solution that belongs to same block
            if r > start_row:

                solution_found = False

                for c in range(len(df.columns)):

                    cell = clean_text(
                        df.iloc[r, c]
                    ).upper()

                    if "SOLUTION" in cell:

                        # Only stop if it is another header
                        if re.search(
                            r"SOLUTION\s*\d+",
                            cell
                        ):

                            solution_found = True
                            break

                if solution_found:

                    end_row = r
                    break

        # ----------------------------------------------------
        # Extract BOM
        # ----------------------------------------------------

        items = []

        for r in range(
            start_row + 1,
            end_row
        ):

            part_number = clean_text(
                df.iloc[r, part_col]
            )

            description = clean_text(
                df.iloc[r, desc_col]
            )

            qty = to_number(
                df.iloc[r, qty_col]
            )

            uom = clean_text(
                df.iloc[r, uom_col]
            )

            price = to_number(
                df.iloc[r, price_col]
            )

            # Skip empty rows
            if (
                not part_number
                and not description
            ):
                continue

            # Skip column headers
            if (
                part_number.upper()
                in [
                    "PART NUMBER",
                    "PART NO",
                    "SKU"
                ]
            ):
                continue

            if (
                description.upper()
                in [
                    "DESCRIPTION",
                    "ITEM DESCRIPTION"
                ]
            ):
                continue

            # Skip section headings
            if (
                "OPTIONAL ITEMS" in description.upper()
                or "SINGLE PHASE PDU" in description.upper()
            ):
                continue

            # Description must exist
            if not description:
                continue

            items.append(
                {
                    "Part Number": part_number,
                    "Description": description,
                    "Qty": qty,
                    "UOM": uom,
                    "Unit Price": price,
                    "Amount": calculate_amount(
                        price,
                        qty
                    )
                }
            )

        if items:

            solution_name = clean_text(
                header["text"]
            )

            solutions[
                solution_name
            ] = items

    return solutions


# ============================================================
# EXTRACT OPTIONAL COMPONENTS
# ============================================================

def extract_optional_items(df):

    optional_items = []

    optional_row = None

    # --------------------------------------------------------
    # Find optional section
    # --------------------------------------------------------

    for r in range(len(df)):

        row_text = " ".join(
            clean_text(df.iloc[r, c])
            for c in range(len(df.columns))
        ).upper()

        if "OTHER OPTIONAL ITEMS" in row_text:

            optional_row = r
            break

    if optional_row is None:

        return optional_items

    # --------------------------------------------------------
    # Standard columns
    # --------------------------------------------------------

    part_col = 0
    desc_col = 1
    qty_col = 2
    uom_col = 3
    price_col = 4

    # --------------------------------------------------------
    # Read optional rows
    # --------------------------------------------------------

    for r in range(
        optional_row + 1,
        len(df)
    ):

        row_text = " ".join(
            clean_text(df.iloc[r, c])
            for c in range(len(df.columns))
        ).upper()

        # Stop at PDU section
        if "SINGLE PHASE PDU" in row_text:
            break

        part_number = clean_text(
            df.iloc[r, part_col]
        )

        description = clean_text(
            df.iloc[r, desc_col]
        )

        excel_qty = to_number(
            df.iloc[r, qty_col]
        )

        uom = clean_text(
            df.iloc[r, uom_col]
        )

        price = to_number(
            df.iloc[r, price_col]
        )

        if not description:
            continue

        if (
            description.upper()
            in [
                "DESCRIPTION",
                "ITEM DESCRIPTION"
            ]
        ):
            continue

        optional_items.append(
            {
                "Part Number": part_number,
                "Description": description,
                "Excel Qty": excel_qty,
                "UOM": uom,
                "Unit Price": price
            }
        )

    return optional_items


# ============================================================
# EXTRACT PDU ITEMS
# ============================================================

def extract_pdu_items(df):

    pdu_items = []

    pdu_row = None

    # --------------------------------------------------------
    # Find PDU section
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
    # Read PDU data
    # --------------------------------------------------------

    previous_type = ""

    for r in range(
        pdu_row + 1,
        len(df)
    ):

        part_number = clean_text(
            df.iloc[r, 0]
        )

        description = clean_text(
            df.iloc[r, 1]
        )

        c13 = to_number(
            df.iloc[r, 2]
        )

        c19 = to_number(
            df.iloc[r, 3]
        )

        pdu_type = clean_text(
            df.iloc[r, 4]
        )

        price = to_number(
            df.iloc[r, 5]
        )

        if not part_number:
            continue

        if not description:
            continue

        if (
            part_number.upper()
            == "PART NUMBER"
        ):
            continue

        # Excel may use merged cells for type
        if pdu_type:

            previous_type = pdu_type

        else:

            pdu_type = previous_type

        pdu_items.append(
            {
                "Part Number": part_number,
                "Description": description,
                "C13": int(c13)
                if float(c13).is_integer()
                else c13,
                "C19": int(c19)
                if float(c19).is_integer()
                else c19,
                "Type": pdu_type,
                "Unit Price": price
            }
        )

    return pdu_items


# ============================================================
# CREATE EXCEL WORKBOOK
# ============================================================

def create_excel_file(dataframes):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for sheet_name, dataframe in dataframes.items():

            dataframe.to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False
            )

    output.seek(0)

    return output


# ============================================================
# LOAD DATA
# ============================================================

try:

    df = load_excel()

    solutions = extract_solutions(df)

    optional_items = extract_optional_items(df)

    pdu_items = extract_pdu_items(df)

except Exception as e:

    st.error(
        f"❌ Error loading Excel:\n\n{e}"
    )

    st.stop()


# ============================================================
# VALIDATION
# ============================================================

if not solutions:

    st.error(
        "❌ No solution configurations were detected."
    )

    st.stop()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("⚙️ Configuration")


# ============================================================
# CUSTOMER DETAILS
# ============================================================

customer_name = st.sidebar.text_input(
    "Customer Name",
    placeholder="Enter customer name"
)

customer_place = st.sidebar.text_input(
    "Customer Place",
    placeholder="Enter customer location"
)

problem_statement = st.sidebar.text_area(
    "Problem / Requirement",
    placeholder="Enter customer requirement"
)

proposed_solution = st.sidebar.text_area(
    "Proposed Solution",
    placeholder="Enter proposed solution"
)


# ============================================================
# CURRENCY
# ============================================================

selected_currency = st.sidebar.selectbox(
    "Currency",
    [
        "INR",
        "USD"
    ]
)


# ------------------------------------------------------------
# Exchange rate
# ------------------------------------------------------------

if selected_currency == "USD":

    live_rate = get_live_exchange_rate(
        "INR",
        "USD"
    )

    if live_rate is not None:

        exchange_rate = st.sidebar.number_input(
            "INR → USD Exchange Rate",
            min_value=0.000001,
            value=float(live_rate),
            step=0.0001,
            format="%.6f"
        )

        st.sidebar.caption(
            f"Live rate loaded: 1 INR ≈ {live_rate:.6f} USD"
        )

    else:

        exchange_rate = st.sidebar.number_input(
            "INR → USD Exchange Rate",
            min_value=0.000001,
            value=0.0119,
            step=0.0001,
            format="%.6f"
        )

else:

    exchange_rate = 1.0


def convert_currency(amount):

    if amount is None:
        return None

    return float(amount) * exchange_rate


def currency_symbol():

    if selected_currency == "USD":
        return "$"

    return "₹"


def format_price(amount):

    if amount is None:
        return "XXX"

    converted = convert_currency(amount)

    return (
        f"{currency_symbol()} "
        f"{converted:,.2f}"
    )


# ============================================================
# HEADER
# ============================================================

st.title("🖥️ MDC Rack BOM Generator")

st.caption(
    "BOM configuration, quantities and costs are extracted "
    "from the supplied Excel workbook."
)


# ============================================================
# STEP 1 — SOLUTION
# ============================================================

st.header("1️⃣ MDC Solution")


solution_names = list(
    solutions.keys()
)

selected_solution = st.selectbox(
    "Select Solution",
    solution_names
)


selected_solution_items = solutions[
    selected_solution
]


# ============================================================
# BASE SOLUTION BOM
# ============================================================

st.subheader(
    "📦 Base Solution BOM"
)


base_df = pd.DataFrame(
    selected_solution_items
)


display_base_df = base_df.copy()

display_base_df[
    "Unit Price"
] = display_base_df[
    "Unit Price"
].apply(format_price)

display_base_df[
    "Amount"
] = display_base_df[
    "Amount"
].apply(format_price)


st.dataframe(
    display_base_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# BASE COST
# ============================================================

base_cost = sum(
    numeric_price(item["Amount"]) or 0
    for item in selected_solution_items
)


# ============================================================
# STEP 2 — OPTIONAL COMPONENTS
# ============================================================

st.header("2️⃣ Optional Components")

st.info(
    "Select an optional component and enter its required quantity. "
    "Amount = Quantity × Unit Price."
)


selected_optional_items = []


if optional_items:

    # Header row
    h1, h2, h3, h4, h5 = st.columns(
        [0.5, 3.8, 1.2, 1.2, 1.5]
    )

    with h1:
        st.write("**✓**")

    with h2:
        st.write("**Component**")

    with h3:
        st.write("**Unit Cost**")

    with h4:
        st.write("**Quantity**")

    with h5:
        st.write("**Amount**")


    for index, item in enumerate(
        optional_items
    ):

        col1, col2, col3, col4, col5 = st.columns(
            [0.5, 3.8, 1.2, 1.2, 1.5]
        )

        with col1:

            selected = st.checkbox(
                "",
                key=f"optional_select_{index}"
            )

        with col2:

            st.write(
                f"**{item['Description']}**"
            )

            if item["Part Number"]:

                st.caption(
                    f"Part Number: {item['Part Number']}"
                )

        with col3:

            st.write(
                format_price(
                    item["Unit Price"]
                )
            )

        with col4:

            if selected:

                quantity = st.number_input(
                    "Qty",
                    min_value=1,
                    max_value=10000,
                    value=1,
                    step=1,
                    key=f"optional_qty_{index}"
                )

            else:

                quantity = 0

        with col5:

            if selected:

                amount = calculate_amount(
                    item["Unit Price"],
                    quantity
                )

                st.write(
                    f"**{format_price(amount)}**"
                )

            else:

                amount = 0

        if selected:

            selected_optional_items.append(
                {
                    "Part Number":
                        item["Part Number"],

                    "Description":
                        item["Description"],

                    "Quantity":
                        quantity,

                    "UOM":
                        item["UOM"]
                        if item["UOM"]
                        else "EA",

                    "Unit Price":
                        item["Unit Price"],

                    "Amount":
                        amount
                }
            )

else:

    st.warning(
        "No optional components found."
    )


# ============================================================
# OPTIONAL SUMMARY
# ============================================================

st.subheader(
    "📋 Selected Optional Components"
)


if selected_optional_items:

    optional_summary_df = pd.DataFrame(
        selected_optional_items
    )

    display_optional_df = optional_summary_df.copy()

    display_optional_df[
        "Unit Price"
    ] = display_optional_df[
        "Unit Price"
    ].apply(format_price)

    display_optional_df[
        "Amount"
    ] = display_optional_df[
        "Amount"
    ].apply(format_price)

    st.dataframe(
        display_optional_df,
        use_container_width=True,
        hide_index=True
    )

else:

    st.info(
        "No optional components selected."
    )


optional_total = sum(
    item["Amount"]
    for item in selected_optional_items
)


# ============================================================
# STEP 3 — PDU SELECTION
# ============================================================

st.header("3️⃣ PDU Selection")


selected_pdu = None
pdu_quantity = 0
pdu_amount = 0


if pdu_items:

    pdu_labels = []

    for item in pdu_items:

        label = (
            f"{item['Part Number']} | "
            f"{item['Description']} | "
            f"{item['Type']} | "
            f"C13: {item['C13']} | "
            f"C19: {item['C19']} | "
            f"{format_price(item['Unit Price'])}"
        )

        pdu_labels.append(label)


    selected_pdu_index = st.selectbox(
        "Select PDU",
        range(len(pdu_items)),
        format_func=lambda i:
            pdu_labels[i]
    )


    selected_pdu = pdu_items[
        selected_pdu_index
    ]


    # --------------------------------------------------------
    # PDU details
    # --------------------------------------------------------

    pdu_info_1, pdu_info_2, pdu_info_3, pdu_info_4 = st.columns(4)

    with pdu_info_1:

        st.metric(
            "Part Number",
            selected_pdu["Part Number"]
        )

    with pdu_info_2:

        st.metric(
            "PDU Type",
            selected_pdu["Type"]
            if selected_pdu["Type"]
            else "—"
        )

    with pdu_info_3:

        st.metric(
            "C13 / C19",
            f"{selected_pdu['C13']} / {selected_pdu['C19']}"
        )

    with pdu_info_4:

        st.metric(
            "Unit Price",
            format_price(
                selected_pdu["Unit Price"]
            )
        )


    # --------------------------------------------------------
    # PDU QUANTITY
    # --------------------------------------------------------

    st.subheader(
        "🔢 PDU Quantity"
    )

    pdu_quantity = st.number_input(
        "Enter PDU Quantity",
        min_value=1,
        max_value=10000,
        value=1,
        step=1,
        key="pdu_quantity"
    )


    # --------------------------------------------------------
    # PDU AMOUNT
    # --------------------------------------------------------

    pdu_amount = calculate_amount(
        selected_pdu["Unit Price"],
        pdu_quantity
    )


    st.success(
        f"PDU Amount = "
        f"{pdu_quantity} × "
        f"{format_price(selected_pdu['Unit Price'])} "
        f"= **{format_price(pdu_amount)}**"
    )


else:

    st.warning(
        "No PDU items found in Excel."
    )


# ============================================================
# STEP 4 — TOTAL COST
# ============================================================

st.header("4️⃣ Cost Summary")


total_cost = (
    base_cost
    + optional_total
    + (pdu_amount or 0)
)


cost_col1, cost_col2, cost_col3, cost_col4 = st.columns(4)


with cost_col1:

    st.metric(
        "Base Solution Cost",
        format_price(base_cost)
    )


with cost_col2:

    st.metric(
        "Optional Cost",
        format_price(optional_total)
    )


with cost_col3:

    st.metric(
        "PDU Cost",
        format_price(pdu_amount)
    )


with cost_col4:

    st.metric(
        "TOTAL COST",
        format_price(total_cost)
    )


# ============================================================
# COST TABLE
# ============================================================

cost_rows = [

    {
        "Cost Item":
            "Base Solution",

        "Quantity":
            1,

        "Unit Cost":
            format_price(base_cost),

        "Total Cost":
            format_price(base_cost)
    }
]


for item in selected_optional_items:

    cost_rows.append(
        {
            "Cost Item":
                item["Description"],

            "Quantity":
                item["Quantity"],

            "Unit Cost":
                format_price(
                    item["Unit Price"]
                ),

            "Total Cost":
                format_price(
                    item["Amount"]
                )
        }
    )


if selected_pdu:

    cost_rows.append(
        {
            "Cost Item":
                selected_pdu["Description"],

            "Quantity":
                pdu_quantity,

            "Unit Cost":
                format_price(
                    selected_pdu["Unit Price"]
                ),

            "Total Cost":
                format_price(
                    pdu_amount
                )
        }
    )


cost_df = pd.DataFrame(
    cost_rows
)


st.subheader(
    "💵 Cost Summary — Before Pricing Factors"
)


st.dataframe(
    cost_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# STEP 5 — COST → PRICE
# ============================================================

st.header(
    "5️⃣ Cost → Price Build-up"
)


st.info(
    """
    Each pricing layer is applied sequentially to the
    running cumulative amount.

    Example:

    ₹1,000 + 15% = ₹1,150

    ₹1,150 + 20% = ₹1,380

    The final cumulative amount becomes the Selling Price.
    """
)


# ============================================================
# DEFAULT PRICING FACTORS
# ============================================================

DEFAULT_PRICING_FACTORS = [

    (
        "Factory Cost (COGS)",
        0.0
    ),

    (
        "Admin & R&D Overhead",
        15.0
    ),

    (
        "Marketing & Sales",
        20.0
    ),

    (
        "Manufacturer Profit",
        15.0
    ),

    (
        "Distribution & Retail",
        45.0
    )
]


pricing_factor_rows = []


# ============================================================
# PRICING FACTOR UI
# ============================================================

for factor_index, (
    default_name,
    default_percentage
) in enumerate(
    DEFAULT_PRICING_FACTORS,
    start=1
):

    factor_col1, factor_col2 = st.columns(
        [3, 1]
    )


    with factor_col1:

        factor_name = st.text_input(
            f"Layer {factor_index} — Name *",
            value=default_name,
            key=f"pricing_factor_name_{factor_index}"
        )


    with factor_col2:

        factor_percentage = st.number_input(
            f"Percentage *",
            min_value=0.0,
            max_value=1000.0,
            value=float(default_percentage),
            step=0.5,
            format="%.2f",
            key=f"pricing_factor_percentage_{factor_index}"
        )


    pricing_factor_rows.append(
        {
            "Layer":
                factor_index,

            "Name":
                factor_name.strip(),

            "Percentage":
                float(factor_percentage)
        }
    )


# ============================================================
# VALIDATE PRICING
# ============================================================

pricing_inputs_complete = all(
    row["Name"]
    and row["Percentage"] >= 0
    for row in pricing_factor_rows
)


if not pricing_inputs_complete:

    st.error(
        "Every pricing layer must have "
        "a name and percentage."
    )

    st.stop()


# ============================================================
# CALCULATE COST → PRICE
# ============================================================

pricing_build_up_rows = []


running_amount = total_cost


for row in pricing_factor_rows:

    layer = row["Layer"]

    name = row["Name"]

    percentage = row["Percentage"]


    # --------------------------------------------------------
    # Layer 1 = baseline
    # --------------------------------------------------------

    if layer == 1:

        previous_amount_display = (
            format_price(total_cost)
        )

        added_amount_display = "—"

        cumulative_display = (
            format_price(total_cost)
        )

        running_amount = total_cost


    # --------------------------------------------------------
    # Remaining layers
    # --------------------------------------------------------

    else:

        previous_amount = running_amount


        added_amount = (
            running_amount
            * (
                percentage
                / 100.0
            )
        )


        running_amount = (
            running_amount
            + added_amount
        )


        previous_amount_display = (
            format_price(
                previous_amount
            )
        )


        added_amount_display = (
            f"+{percentage:.2f}% = "
            f"{format_price(added_amount)}"
        )


        cumulative_display = (
            format_price(
                running_amount
            )
        )


    pricing_build_up_rows.append(
        {
            "Layer":
                layer,

            "Name":
                name,

            "Percentage Added":
                (
                    "Baseline"
                    if layer == 1
                    else f"{percentage:.2f}%"
                ),

            "Previous Amount":
                previous_amount_display,

            "Added Amount":
                added_amount_display,

            "Cumulative Price":
                cumulative_display
        }
    )


pricing_build_up_df = pd.DataFrame(
    pricing_build_up_rows
)


# ============================================================
# DISPLAY PRICE BUILD-UP
# ============================================================

st.subheader(
    "📈 Price Build-up"
)


st.dataframe(
    pricing_build_up_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# FINAL SELLING PRICE
# ============================================================

selling_price = running_amount


st.subheader(
    "🏷️ Final Selling Price"
)


st.success(
    f"# {format_price(selling_price)}"
)


# ============================================================
# PRICE CALCULATION DETAILS
# ============================================================

st.subheader(
    "🧮 Pricing Calculation"
)


pricing_detail_rows = []


for row in pricing_factor_rows:

    layer = row["Layer"]

    percentage = row["Percentage"]

    name = row["Name"]


    if layer == 1:

        previous = total_cost

        added = 0

        cumulative = total_cost

    else:

        # Recalculate for clean numerical table
        pass


running_numeric = total_cost


for row in pricing_factor_rows:

    layer = row["Layer"]

    name = row["Name"]

    percentage = row["Percentage"]


    if layer == 1:

        previous = total_cost

        added = 0

        cumulative = total_cost

    else:

        previous = running_numeric

        added = (
            previous
            * percentage
            / 100
        )

        cumulative = (
            previous
            + added
        )

        running_numeric = cumulative


    pricing_detail_rows.append(
        {
            "Layer":
                layer,

            "Pricing Factor":
                name,

            "Percentage":
                (
                    "Baseline"
                    if layer == 1
                    else f"{percentage:.2f}%"
                ),

            "Previous Amount":
                format_price(previous),

            "Added Amount":
                (
                    "—"
                    if layer == 1
                    else format_price(added)
                ),

            "Cumulative Price":
                format_price(cumulative)
        }
    )


pricing_detail_df = pd.DataFrame(
    pricing_detail_rows
)


st.dataframe(
    pricing_detail_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# STEP 6 — FINAL BOM
# ============================================================

st.header(
    "6️⃣ Final BOM"
)


final_bom_rows = []


# ============================================================
# BASE BOM
# ============================================================

for item in selected_solution_items:

    final_bom_rows.append(
        {
            "Category":
                "Base Solution",

            "Part Number":
                item["Part Number"],

            "Description":
                item["Description"],

            "Quantity":
                item["Qty"],

            "UOM":
                item["UOM"],

            "Unit Price":
                item["Unit Price"],

            "Amount":
                item["Amount"]
        }
    )


# ============================================================
# OPTIONAL BOM
# ============================================================

for item in selected_optional_items:

    final_bom_rows.append(
        {
            "Category":
                "Optional",

            "Part Number":
                item["Part Number"],

            "Description":
                item["Description"],

            "Quantity":
                item["Quantity"],

            "UOM":
                item["UOM"],

            "Unit Price":
                item["Unit Price"],

            "Amount":
                item["Amount"]
        }
    )


# ============================================================
# PDU BOM
# ============================================================

if selected_pdu:

    final_bom_rows.append(
        {
            "Category":
                "PDU",

            "Part Number":
                selected_pdu["Part Number"],

            "Description":
                selected_pdu["Description"],

            "Quantity":
                pdu_quantity,

            "UOM":
                "EA",

            "Unit Price":
                selected_pdu["Unit Price"],

            "Amount":
                pdu_amount
        }
    )


bom_df = pd.DataFrame(
    final_bom_rows
)


# ============================================================
# DISPLAY BOM
# ============================================================

display_bom_df = bom_df.copy()


display_bom_df[
    "Unit Price"
] = display_bom_df[
    "Unit Price"
].apply(format_price)


display_bom_df[
    "Amount"
] = display_bom_df[
    "Amount"
].apply(format_price)


st.dataframe(
    display_bom_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# STEP 7 — COMMERCIAL SUMMARY
# ============================================================

st.header(
    "7️⃣ Final BOM & Commercial Summary"
)


commercial_rows = [

    {
        "Part Number":
            "—",

        "Description":
            selected_solution,

        "Quantity":
            1,

        "Unit Cost":
            format_price(base_cost),

        "Amount":
            format_price(base_cost)
    }
]


for item in selected_optional_items:

    commercial_rows.append(
        {
            "Part Number":
                item["Part Number"],

            "Description":
                item["Description"],

            "Quantity":
                item["Quantity"],

            "Unit Cost":
                format_price(
                    item["Unit Price"]
                ),

            "Amount":
                format_price(
                    item["Amount"]
                )
        }
    )


if selected_pdu:

    commercial_rows.append(
        {
            "Part Number":
                selected_pdu["Part Number"],

            "Description":
                selected_pdu["Description"],

            "Quantity":
                pdu_quantity,

            "Unit Cost":
                format_price(
                    selected_pdu["Unit Price"]
                ),

            "Amount":
                format_price(
                    pdu_amount
                )
        }
    )


commercial_df = pd.DataFrame(
    commercial_rows
)


st.dataframe(
    commercial_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# STEP 8 — SUMMARY
# ============================================================

st.header(
    "8️⃣ Project Summary"
)


summary_rows_ui = [

    {
        "Metric":
            "Customer Name",

        "Value":
            customer_name or "—"
    },

    {
        "Metric":
            "Customer Place",

        "Value":
            customer_place or "—"
    },

    {
        "Metric":
            "Problem / Requirement",

        "Value":
            problem_statement or "—"
    },

    {
        "Metric":
            "Proposed Solution",

        "Value":
            proposed_solution or "—"
    },

    {
        "Metric":
            "Selected Configuration",

        "Value":
            selected_solution
    },

    {
        "Metric":
            "Base Solution Cost",

        "Value":
            format_price(base_cost)
    },

    {
        "Metric":
            "Optional Components Cost",

        "Value":
            format_price(optional_total)
    },

    {
        "Metric":
            "PDU Cost",

        "Value":
            format_price(pdu_amount)
    },

    {
        "Metric":
            "TOTAL COST",

        "Value":
            format_price(total_cost)
    },

    {
        "Metric":
            "FINAL SELLING PRICE",

        "Value":
            format_price(selling_price)
    }
]


summary_df = pd.DataFrame(
    summary_rows_ui
)


st.dataframe(
    summary_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# STEP 9 — EXCEL EXPORT
# ============================================================

st.header(
    "9️⃣ Excel Export"
)


# ============================================================
# CUSTOMER DETAILS
# ============================================================

customer_details_df = pd.DataFrame(
    [
        {
            "Field":
                "Customer Name",

            "Value":
                customer_name
        },

        {
            "Field":
                "Customer Place",

            "Value":
                customer_place
        },

        {
            "Field":
                "MDC Configuration",

            "Value":
                selected_solution
        },

        {
            "Field":
                "Problem / Requirement",

            "Value":
                problem_statement
        },

        {
            "Field":
                "Proposed Solution",

            "Value":
                proposed_solution
        }
    ]
)


# ============================================================
# BOM WITHOUT PRICE
# ============================================================

bom_without_price_df = bom_df[
    [
        "Category",
        "Part Number",
        "Description",
        "Quantity",
        "UOM"
    ]
].copy()


# ============================================================
# BOM WITH PRICE
# ============================================================

bom_with_price_df = bom_df.copy()


# ============================================================
# OPTIONAL EXPORT
# ============================================================

if selected_optional_items:

    export_optional_df = pd.DataFrame(
        [
            {
                "Part Number":
                    item["Part Number"],

                "Component":
                    item["Description"],

                "Quantity":
                    item["Quantity"],

                "Unit Cost":
                    format_price(
                        item["Unit Price"]
                    ),

                "Amount":
                    format_price(
                        item["Amount"]
                    )
            }

            for item in selected_optional_items
        ]
    )

else:

    export_optional_df = pd.DataFrame(
        columns=[
            "Part Number",
            "Component",
            "Quantity",
            "Unit Cost",
            "Amount"
        ]
    )


# ============================================================
# PDU EXPORT
# ============================================================

if selected_pdu:

    pdu_export_df = pd.DataFrame(
        [
            {
                "Part Number":
                    selected_pdu["Part Number"],

                "Description":
                    selected_pdu["Description"],

                "PDU Type":
                    selected_pdu["Type"],

                "C13":
                    selected_pdu["C13"],

                "C19":
                    selected_pdu["C19"],

                "Quantity":
                    pdu_quantity,

                "Unit Cost":
                    format_price(
                        selected_pdu["Unit Price"]
                    ),

                "Amount":
                    format_price(
                        pdu_amount
                    )
            }
        ]
    )

else:

    pdu_export_df = pd.DataFrame(
        columns=[
            "Part Number",
            "Description",
            "PDU Type",
            "C13",
            "C19",
            "Quantity",
            "Unit Cost",
            "Amount"
        ]
    )


# ============================================================
# PRICE SUMMARY EXPORT
# ============================================================

price_summary_rows = [

    {
        "Section":
            "Customer",

        "Item":
            "Customer Name",

        "Value":
            customer_name
    },

    {
        "Section":
            "Customer",

        "Item":
            "Customer Place",

        "Value":
            customer_place
    },

    {
        "Section":
            "Project",

        "Item":
            "Problem / Requirement",

        "Value":
            problem_statement
    },

    {
        "Section":
            "Project",

        "Item":
            "Proposed Solution",

        "Value":
            proposed_solution
    },

    {
        "Section":
            "Configuration",

        "Item":
            "Selected Solution",

        "Value":
            selected_solution
    },

    {
        "Section":
            "Cost",

        "Item":
            "Base Solution Cost",

        "Value":
            format_price(base_cost)
    },

    {
        "Section":
            "Cost",

        "Item":
            "Optional Components Cost",

        "Value":
            format_price(optional_total)
    },

    {
        "Section":
            "Cost",

        "Item":
            "PDU Cost",

        "Value":
            format_price(pdu_amount)
    },

    {
        "Section":
            "Cost",

        "Item":
            "TOTAL COST",

        "Value":
            format_price(total_cost)
    }
]


# ============================================================
# ADD PRICE BUILD-UP
# ============================================================

for row in pricing_build_up_rows:

    price_summary_rows.append(
        {
            "Section":
                "Cost → Price",

            "Item":
                f"Layer {row['Layer']} — "
                f"{row['Name']}",

            "Value":
                (
                    f"{row['Percentage Added']} | "
                    f"Previous: {row['Previous Amount']} | "
                    f"Added: {row['Added Amount']} | "
                    f"Cumulative: {row['Cumulative Price']}"
                )
        }
    )


price_summary_rows.append(
    {
        "Section":
            "Price",

        "Item":
            "FINAL SELLING PRICE",

        "Value":
            format_price(
                selling_price
            )
    }
)


price_summary_df = pd.DataFrame(
    price_summary_rows
)


# ============================================================
# PRICING FACTORS EXPORT
# ============================================================

pricing_factors_export_df = pd.DataFrame(
    pricing_factor_rows
)


pricing_factors_export_df[
    "Percentage Added"
] = pricing_factors_export_df[
    "Percentage"
].map(
    lambda x:
        "Baseline"
        if x == 0
        else f"{x:.2f}%"
)


pricing_factors_export_df = (
    pricing_factors_export_df[
        [
            "Layer",
            "Name",
            "Percentage Added"
        ]
    ]
)


# ============================================================
# FINAL COST SUMMARY EXPORT
# ============================================================

cost_export_df = cost_df.copy()


# ============================================================
# PRICE BUILD-UP NUMERICAL EXPORT
# ============================================================

price_build_up_export_df = pd.DataFrame(
    pricing_detail_rows
)


# ============================================================
# EXCEL FILE — WITHOUT PRICE
# ============================================================

excel_without_price = create_excel_file(
    {
        "Customer Details":
            customer_details_df,

        "BOM":
            bom_without_price_df
    }
)


# ============================================================
# EXCEL FILE — WITH PRICE
# ============================================================

excel_with_price = create_excel_file(
    {
        "Customer Details":
            customer_details_df,

        "BOM With Price":
            bom_with_price_df,

        "Optional Components":
            export_optional_df,

        "PDU":
            pdu_export_df,

        "Cost Summary":
            cost_export_df,

        "Price Build-up":
            price_build_up_export_df,

        "Pricing Factors":
            pricing_factors_export_df,

        "Price Summary":
            price_summary_df
    }
)


# ============================================================
# DOWNLOAD WITHOUT PRICE
# ============================================================

st.subheader(
    "📥 Download BOM"
)


st.download_button(
    label="📥 Download BOM — Without Price",

    data=excel_without_price,

    file_name=(
        "MDC_BOM_Without_Price.xlsx"
    ),

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    key="download_without_price"
)


# ============================================================
# DOWNLOAD WITH PRICE
# ============================================================

st.download_button(
    label="💰 Download BOM — With Price",

    data=excel_with_price,

    file_name=(
        "MDC_BOM_With_Price.xlsx"
    ),

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    key="download_with_price"
)


# ============================================================
# DATA SOURCE
# ============================================================

with st.expander(
    "📊 Excel Data Source"
):

    st.write(
        f"**Workbook:** {EXCEL_FILE}"
    )

    st.write(
        f"**Sheet:** {SHEET_NAME}"
    )

    st.write(
        f"**Solutions detected:** "
        f"{len(solutions)}"
    )

    st.write(
        f"**Optional items detected:** "
        f"{len(optional_items)}"
    )

    st.write(
        f"**PDU items detected:** "
        f"{len(pdu_items)}"
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "MDC Configuration & BOM Generator | "
    "Excel-driven BOM + Quantity-based Optional/PDU Cost + "
    "Editable Cost-to-Price Build-up"
)
