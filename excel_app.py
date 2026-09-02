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
# CURRENCY FUNCTIONS
# ============================================================

@st.cache_data(ttl=3600)
def get_live_exchange_rate(from_currency, to_currency):
    """
    Get live exchange rate.

    Cached for 1 hour so the application does not repeatedly
    call the API on every Streamlit rerun.
    """

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

    except Exception as e:

        st.warning(
            f"Unable to fetch live exchange rate: {e}"
        )

        return None


def convert_currency(amount, rate):
    """
    Convert amount using live exchange rate.
    """

    if amount is None:
        return None

    return float(amount) * float(rate)


def currency_symbol(currency):
    """
    Return currency symbol.
    """

    if currency == "INR":
        return "₹"

    if currency == "USD":
        return "$"

    return ""


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):

    if pd.isna(value):
        return ""

    return (
        str(value)
        .replace("\n", " ")
        .replace("\xa0", " ")
        .strip()
    )


def is_numeric_price(value):

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

        except (
            ValueError,
            TypeError
        ):

            return False

    return False


def numeric_price(value):

    if not is_numeric_price(value):
        return None

    return float(
        str(value)
        .strip()
        .replace(",", "")
    )


def to_number(value):

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

    price = numeric_price(unit_price)

    if price is None:
        return None

    try:

        qty = float(quantity)

    except Exception:

        qty = 0

    return price * qty


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
# SOLUTION DETECTION
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

    processed_headers = []

    for header in headers:

        duplicate = False

        for existing in processed_headers:

            if (
                existing["row"] == header["row"]
                and abs(
                    existing["col"]
                    - header["col"]
                ) <= 4
            ):

                duplicate = True
                break

        if not duplicate:

            processed_headers.append(header)

    for header in processed_headers:

        start_row = header["row"]
        start_col = header["col"]

        part_col = start_col
        desc_col = start_col + 1
        qty_col = start_col + 2
        uom_col = start_col + 3
        price_col = start_col + 4

        if price_col >= len(df.columns):
            continue

        end_row = len(df)

        for r in range(
            start_row + 1,
            len(df)
        ):

            row_text = " ".join(
                clean_text(
                    df.iloc[r, c]
                )
                for c in range(
                    len(df.columns)
                )
            ).upper()

            if (
                "OTHER OPTIONAL ITEMS"
                in row_text
                or
                "SINGLE PHASE PDU"
                in row_text
            ):

                end_row = r
                break

            solution_found = False

            for c in range(
                len(df.columns)
            ):

                cell = clean_text(
                    df.iloc[r, c]
                ).upper()

                if re.search(
                    r"SOLUTION\s*\d+",
                    cell
                ):

                    solution_found = True
                    break

            if solution_found:

                end_row = r
                break

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

            if (
                not part_number
                and not description
            ):
                continue

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

            if (
                "OPTIONAL ITEMS"
                in description.upper()
                or
                "SINGLE PHASE PDU"
                in description.upper()
            ):
                continue

            if not description:
                continue

            items.append(
                {
                    "Part Number":
                        part_number,

                    "Description":
                        description,

                    "Qty":
                        qty,

                    "UOM":
                        uom,

                    "Unit Price":
                        price,

                    "Amount":
                        calculate_amount(
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
# OPTIONAL ITEMS
# ============================================================

def extract_optional_items(df):

    optional_items = []

    optional_row = None

    for r in range(len(df)):

        row_text = " ".join(
            clean_text(
                df.iloc[r, c]
            )
            for c in range(
                len(df.columns)
            )
        ).upper()

        if "OTHER OPTIONAL ITEMS" in row_text:

            optional_row = r
            break

    if optional_row is None:
        return optional_items

    part_col = 0
    desc_col = 1
    qty_col = 2
    uom_col = 3
    price_col = 4

    for r in range(
        optional_row + 1,
        len(df)
    ):

        row_text = " ".join(
            clean_text(
                df.iloc[r, c]
            )
            for c in range(
                len(df.columns)
            )
        ).upper()

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

        if description.upper() in [
            "DESCRIPTION",
            "ITEM DESCRIPTION"
        ]:
            continue

        optional_items.append(
            {
                "Part Number":
                    part_number,

                "Description":
                    description,

                "Excel Qty":
                    excel_qty,

                "UOM":
                    uom,

                "Unit Price":
                    price
            }
        )

    return optional_items


# ============================================================
# PDU ITEMS
# ============================================================

def extract_pdu_items(df):

    pdu_items = []

    pdu_row = None

    for r in range(len(df)):

        row_text = " ".join(
            clean_text(
                df.iloc[r, c]
            )
            for c in range(
                len(df.columns)
            )
        ).upper()

        if "SINGLE PHASE PDU" in row_text:

            pdu_row = r
            break

    if pdu_row is None:
        return pdu_items

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

        if part_number.upper() == "PART NUMBER":
            continue

        if pdu_type:

            previous_type = pdu_type

        else:

            pdu_type = previous_type

        pdu_items.append(
            {
                "Part Number":
                    part_number,

                "Description":
                    description,

                "C13":
                    int(c13)
                    if float(c13).is_integer()
                    else c13,

                "C19":
                    int(c19)
                    if float(c19).is_integer()
                    else c19,

                "Type":
                    pdu_type,

                "Unit Price":
                    price
            }
        )

    return pdu_items


# ============================================================
# EXCEL EXPORT
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


if not solutions:

    st.error(
        "❌ No solution configurations were detected."
    )

    st.stop()


# ============================================================
# HEADER
# ============================================================

st.title(
    "🏭 MDC Configuration & BOM Generator"
)

st.markdown(
    """
    Select the MDC configuration, review the BOM,
    select optional components and PDU quantities,
    and generate the commercial BOM.
    """
)


# ============================================================
# CURRENCY SETTINGS
# ============================================================

st.header("💱 Currency Settings")


currency_col1, currency_col2 = st.columns(
    [2, 3]
)


with currency_col1:

    selected_currency = st.radio(
        "Display Currency",
        [
            "INR",
            "USD"
        ],
        horizontal=True,
        key="display_currency"
    )


# ============================================================
# LIVE RATE
# ============================================================

if selected_currency == "INR":

    exchange_rate = 1.0

    base_currency = "INR"
    target_currency = "INR"


else:

    base_currency = "INR"
    target_currency = "USD"

    exchange_rate = get_live_exchange_rate(
        base_currency,
        target_currency
    )


# ============================================================
# RATE ERROR
# ============================================================

if exchange_rate is None:

    st.error(
        """
        ❌ Unable to fetch the live INR → USD exchange rate.

        Please check your internet connection and refresh
        the application.
        """
    )

    st.stop()


# ============================================================
# LIVE RATE DISPLAY
# ============================================================

with currency_col2:

    if selected_currency == "USD":

        st.success(
            f"🔴 LIVE Exchange Rate: "
            f"₹1 = ${exchange_rate:.6f}"
        )

        st.caption(
            "Rate fetched automatically from the live "
            "exchange-rate service."
        )

    else:

        st.info(
            "🇮🇳 Displaying all values in Indian Rupees (INR)"
        )


st.divider()


# ============================================================
# FORMATTING FUNCTIONS
# ============================================================

def format_price(price):

    if not is_numeric_price(price):

        return "XXX"

    converted_price = convert_currency(
        float(price),
        exchange_rate
    )

    symbol = currency_symbol(
        selected_currency
    )

    return (
        f"{symbol} "
        f"{converted_price:,.2f}"
    )


# ============================================================
# CUSTOMER DETAILS
# ============================================================

st.header(
    "1️⃣ Customer & Project Details"
)


customer_col1, customer_col2 = st.columns(2)


with customer_col1:

    customer_name = st.text_input(
        "Customer Name *",
        placeholder="Enter customer / company name"
    )


with customer_col2:

    customer_place = st.text_input(
        "Customer Place *",
        placeholder="Enter city / location"
    )


problem_col, solution_col = st.columns(2)


with problem_col:

    problem_statement = st.text_area(
        "Problem / Requirement",
        placeholder="Describe the customer's requirement..."
    )


with solution_col:

    proposed_solution = st.text_area(
        "Proposed Solution",
        placeholder="Describe the proposed MDC solution..."
    )


# ============================================================
# SOLUTION SELECTION
# ============================================================

st.header(
    "2️⃣ Select MDC Solution"
)


solution_names = list(
    solutions.keys()
)


selected_solution = st.selectbox(
    "Configuration",
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
].apply(
    format_price
)


display_base_df[
    "Amount"
] = display_base_df[
    "Amount"
].apply(
    format_price
)


st.dataframe(
    display_base_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# BASE COST
# ============================================================

base_cost = sum(
    numeric_price(
        item["Amount"]
    ) or 0
    for item in selected_solution_items
)


# ============================================================
# OPTIONAL COMPONENTS
# ============================================================

st.header(
    "3️⃣ Optional Components"
)


st.info(
    "Select an optional component and enter its quantity. "
    "Amount = Quantity × Excel Unit Price."
)


selected_optional_items = []


if optional_items:

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
                    f"Part Number: "
                    f"{item['Part Number']}"
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
# OPTIONAL TOTAL
# ============================================================

optional_total = sum(
    item["Amount"]
    for item in selected_optional_items
)


# ============================================================
# PDU
# ============================================================

st.header(
    "4️⃣ PDU Selection"
)


selected_pdu = None

pdu_quantity = 0

pdu_amount = 0


if pdu_items:

    pdu_labels = []


    for item in pdu_items:

        pdu_labels.append(
            f"{item['Part Number']} | "
            f"{item['Description']} | "
            f"{item['Type']} | "
            f"C13: {item['C13']} | "
            f"C19: {item['C19']} | "
            f"{format_price(item['Unit Price'])}"
        )


    selected_pdu_index = st.selectbox(
        "Select PDU",
        range(len(pdu_items)),
        format_func=lambda i:
            pdu_labels[i]
    )


    selected_pdu = pdu_items[
        selected_pdu_index
    ]


    pdu_info1, pdu_info2, pdu_info3, pdu_info4 = st.columns(4)


    with pdu_info1:

        st.metric(
            "Part Number",
            selected_pdu["Part Number"]
        )


    with pdu_info2:

        st.metric(
            "PDU Type",
            selected_pdu["Type"]
            if selected_pdu["Type"]
            else "—"
        )


    with pdu_info3:

        st.metric(
            "C13 / C19",
            f"{selected_pdu['C13']} / "
            f"{selected_pdu['C19']}"
        )


    with pdu_info4:

        st.metric(
            "Unit Price",
            format_price(
                selected_pdu["Unit Price"]
            )
        )


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
# TOTAL COST
# ============================================================

st.header(
    "5️⃣ Cost Summary"
)


total_cost = (
    base_cost
    + optional_total
    + (pdu_amount or 0)
)


c1, c2, c3, c4 = st.columns(4)


with c1:

    st.metric(
        "Base Solution",
        format_price(base_cost)
    )


with c2:

    st.metric(
        "Optional Components",
        format_price(optional_total)
    )


with c3:

    st.metric(
        "PDU",
        format_price(pdu_amount)
    )


with c4:

    st.metric(
        "TOTAL COST",
        format_price(total_cost)
    )


# ============================================================
# COST → PRICE
# ============================================================

st.header(
    "6️⃣ Cost → Price Build-up"
)


st.info(
    """
    Each pricing layer is applied sequentially to the
    running cumulative amount.

    Example:

    ₹1,000 + 15% = ₹1,150

    ₹1,150 + 20% = ₹1,380
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
            "Percentage *",
            min_value=0.0,
            max_value=1000.0,
            value=float(
                default_percentage
            ),
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
                float(
                    factor_percentage
                )
        }
    )


# ============================================================
# PRICE CALCULATION
# ============================================================

pricing_build_up_rows = []

running_amount = total_cost


for row in pricing_factor_rows:

    layer = row["Layer"]

    name = row["Name"]

    percentage = row["Percentage"]


    if layer == 1:

        previous_amount = total_cost

        added_amount = 0

        cumulative_amount = total_cost


    else:

        previous_amount = running_amount

        added_amount = (
            running_amount
            * percentage
            / 100
        )

        cumulative_amount = (
            running_amount
            + added_amount
        )

        running_amount = cumulative_amount


    pricing_build_up_rows.append(
        {
            "Layer":
                layer,

            "Name":
                name,

            "Percentage":
                (
                    "Baseline"
                    if layer == 1
                    else f"{percentage:.2f}%"
                ),

            "Previous Amount":
                format_price(
                    previous_amount
                ),

            "Added Amount":
                (
                    "—"
                    if layer == 1
                    else format_price(
                        added_amount
                    )
                ),

            "Cumulative Price":
                format_price(
                    cumulative_amount
                )
        }
    )


pricing_build_up_df = pd.DataFrame(
    pricing_build_up_rows
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
# FINAL BOM
# ============================================================

st.header(
    "7️⃣ Final BOM"
)


final_bom_rows = []


# ------------------------------------------------------------
# BASE
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# OPTIONAL
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# PDU
# ------------------------------------------------------------

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


display_bom_df = bom_df.copy()


display_bom_df[
    "Unit Price"
] = display_bom_df[
    "Unit Price"
].apply(
    format_price
)


display_bom_df[
    "Amount"
] = display_bom_df[
    "Amount"
].apply(
    format_price
)


st.dataframe(
    display_bom_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# PROJECT SUMMARY
# ============================================================

st.header(
    "8️⃣ Project Summary"
)


summary_df = pd.DataFrame(
    [
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
                "Selected Solution",

            "Value":
                selected_solution
        },

        {
            "Metric":
                "Currency",

            "Value":
                selected_currency
        },

        {
            "Metric":
                "Live Exchange Rate",

            "Value":
                (
                    f"1 INR = "
                    f"{exchange_rate:.6f} USD"
                    if selected_currency == "USD"
                    else "1 INR = 1 INR"
                )
        },

        {
            "Metric":
                "Base Cost",

            "Value":
                format_price(base_cost)
        },

        {
            "Metric":
                "Optional Cost",

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
                "Total Cost",

            "Value":
                format_price(total_cost)
        },

        {
            "Metric":
                "Final Selling Price",

            "Value":
                format_price(selling_price)
        }
    ]
)


st.dataframe(
    summary_df,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# EXCEL EXPORT
# ============================================================

st.header(
    "9️⃣ Excel Export"
)


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
                "Selected Solution",

            "Value":
                selected_solution
        },

        {
            "Field":
                "Currency",

            "Value":
                selected_currency
        },

        {
            "Field":
                "Exchange Rate",

            "Value":
                exchange_rate
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


# ------------------------------------------------------------
# Pricing export
# ------------------------------------------------------------

pricing_export_df = pd.DataFrame(
    pricing_factor_rows
)


pricing_detail_export_df = pd.DataFrame(
    pricing_build_up_rows
)


# ------------------------------------------------------------
# Optional export
# ------------------------------------------------------------

optional_export_df = pd.DataFrame(
    selected_optional_items
)


# ------------------------------------------------------------
# PDU export
# ------------------------------------------------------------

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

                "Unit Price":
                    selected_pdu["Unit Price"],

                "Amount":
                    pdu_amount
            }
        ]
    )

else:

    pdu_export_df = pd.DataFrame()


# ------------------------------------------------------------
# Commercial summary
# ------------------------------------------------------------

commercial_summary_df = pd.DataFrame(
    [
        {
            "Item":
                "Base Solution",

            "Amount":
                base_cost
        },

        {
            "Item":
                "Optional Components",

            "Amount":
                optional_total
        },

        {
            "Item":
                "PDU",

            "Amount":
                pdu_amount
        },

        {
            "Item":
                "Total Cost",

            "Amount":
                total_cost
        },

        {
            "Item":
                "Final Selling Price",

            "Amount":
                selling_price
        }
    ]
)


# ============================================================
# CREATE EXCEL
# ============================================================

excel_file = create_excel_file(
    {
        "Customer Details":
            customer_details_df,

        "BOM":
            bom_df,

        "Optional Components":
            optional_export_df,

        "PDU":
            pdu_export_df,

        "Cost Summary":
            commercial_summary_df,

        "Price Build-up":
            pricing_detail_export_df,

        "Pricing Factors":
            pricing_export_df
    }
)


# ============================================================
# DOWNLOAD
# ============================================================

st.download_button(
    label="📥 Download Complete BOM Excel",

    data=excel_file,

    file_name="MDC_BOM_With_Live_Currency.xlsx",

    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )
)


# ============================================================
# DATA SOURCE
# ============================================================

with st.expander(
    "📊 Excel Data Source"
):

    st.write(
        f"**Workbook:** `{EXCEL_FILE}`"
    )

    st.write(
        f"**Sheet:** `{SHEET_NAME}`"
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

    if selected_currency == "USD":

        st.write(
            f"**Live INR → USD rate:** "
            f"{exchange_rate:.6f}"
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "MDC Configuration & BOM Generator | "
    "Excel-driven BOM | Quantity-based costing | "
    "Live INR/USD currency conversion | "
    "Editable Cost → Price build-up"
)
