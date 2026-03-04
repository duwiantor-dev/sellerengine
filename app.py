import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

st.set_page_config(page_title="Marketplace Tools", layout="wide")

st.title("Marketplace Bulk Tools")

# ==============================
# HELPER FUNCTIONS
# ==============================

def load_pricelist(file):

    df = pd.read_excel(file)

    price_map = {}

    for _, row in df.iterrows():

        sku = str(row[0]).strip()
        price = float(row[1])

        if price < 1000000:
            price *= 1000

        price_map[sku] = price

    return price_map


def load_addon(file):

    df = pd.read_excel(file)

    addon_map = {}

    for _, row in df.iterrows():

        addon = str(row[0]).strip()
        price = float(row[1])

        if price < 1000000:
            price *= 1000

        addon_map[addon] = price

    return addon_map


def process_price_file(template, price_map, addon_map):

    df = pd.read_excel(template)

    changed_rows = []

    for i, row in df.iterrows():

        try:

            sku = str(row["SKU"]).strip()

            base_price = price_map.get(sku)

            if base_price is None:
                continue

            addon = str(row.get("Addon", "")).strip()

            addon_price = addon_map.get(addon, 0)

            new_price = base_price + addon_price

            old_price = row["Price"]

            if new_price != old_price:

                row["Price"] = new_price

                changed_rows.append(row)

        except:
            continue

    if not changed_rows:
        return None

    result_df = pd.DataFrame(changed_rows)

    output = BytesIO()

    result_df.to_excel(output, index=False)

    return output.getvalue()


def process_stock_file(template):

    df = pd.read_excel(template)

    changed_rows = []

    for i, row in df.iterrows():

        try:

            new_stock = row["NewStock"]
            old_stock = row["Stock"]

            if new_stock != old_stock:

                row["Stock"] = new_stock

                changed_rows.append(row)

        except:
            continue

    if not changed_rows:
        return None

    result_df = pd.DataFrame(changed_rows)

    output = BytesIO()

    result_df.to_excel(output, index=False)

    return output.getvalue()


def make_zip(results):

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as z:

        for name, data in results:

            z.writestr(name, data)

    return zip_buffer.getvalue()


# ==============================
# SIDEBAR MENU
# ==============================

menu = st.sidebar.radio(
    "Menu",
    [
        "Harga Normal",
        "Harga Coret",
        "Update Stok"
    ]
)

# ==============================
# HARGA NORMAL
# ==============================

if menu == "Harga Normal":

    st.header("Harga Normal")

    templates = st.file_uploader(
        "Upload Template",
        type=["xlsx"],
        accept_multiple_files=True
    )

    pricelist = st.file_uploader(
        "Upload Pricelist",
        type=["xlsx"]
    )

    addon = st.file_uploader(
        "Upload Addon",
        type=["xlsx"]
    )

    if st.button("PROCESS"):

        if not templates or not pricelist or not addon:
            st.error("Upload semua file dulu")
            st.stop()

        price_map = load_pricelist(pricelist)
        addon_map = load_addon(addon)

        results = []

        progress = st.progress(0)

        for i, template in enumerate(templates):

            result = process_price_file(template, price_map, addon_map)

            if result:

                filename = template.name.replace(".xlsx", "_result.xlsx")

                results.append((filename, result))

            progress.progress((i + 1) / len(templates))

        if not results:

            st.warning("Tidak ada perubahan")

        elif len(results) == 1:

            st.download_button(
                "Download Result",
                results[0][1],
                file_name=results[0][0]
            )

        else:

            zip_file = make_zip(results)

            st.download_button(
                "Download ZIP",
                zip_file,
                file_name="results.zip"
            )


# ==============================
# HARGA CORET
# ==============================

elif menu == "Harga Coret":

    st.header("Harga Coret")

    templates = st.file_uploader(
        "Upload Template",
        type=["xlsx"],
        accept_multiple_files=True
    )

    pricelist = st.file_uploader(
        "Upload Pricelist",
        type=["xlsx"]
    )

    addon = st.file_uploader(
        "Upload Addon",
        type=["xlsx"]
    )

    if st.button("PROCESS"):

        price_map = load_pricelist(pricelist)
        addon_map = load_addon(addon)

        results = []

        progress = st.progress(0)

        for i, template in enumerate(templates):

            df = pd.read_excel(template)

            changed_rows = []

            for _, row in df.iterrows():

                try:

                    sku = str(row["SKU"]).strip()

                    base_price = price_map.get(sku)

                    if base_price is None:
                        continue

                    addon = str(row.get("Addon", "")).strip()

                    addon_price = addon_map.get(addon, 0)

                    new_price = base_price + addon_price

                    if new_price != row["Price"]:

                        row["Price"] = new_price

                        changed_rows.append(row)

                except:
                    continue

            if not changed_rows:
                continue

            df_out = pd.DataFrame(changed_rows)

            chunks = [df_out[i:i+1000] for i in range(0, len(df_out), 1000)]

            for idx, chunk in enumerate(chunks):

                output = BytesIO()

                chunk.to_excel(output, index=False)

                filename = f"{template.name}_part{idx+1}.xlsx"

                results.append((filename, output.getvalue()))

            progress.progress((i + 1) / len(templates))

        if len(results) == 1:

            st.download_button(
                "Download",
                results[0][1],
                file_name=results[0][0]
            )

        else:

            zip_file = make_zip(results)

            st.download_button(
                "Download ZIP",
                zip_file,
                file_name="results.zip"
            )


# ==============================
# UPDATE STOK
# ==============================

elif menu == "Update Stok":

    st.header("Update Stok")

    templates = st.file_uploader(
        "Upload Template",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if st.button("PROCESS"):

        results = []

        progress = st.progress(0)

        for i, template in enumerate(templates):

            result = process_stock_file(template)

            if result:

                filename = template.name.replace(".xlsx", "_stok.xlsx")

                results.append((filename, result))

            progress.progress((i + 1) / len(templates))

        if len(results) == 1:

            st.download_button(
                "Download Result",
                results[0][1],
                file_name=results[0][0]
            )

        else:

            zip_file = make_zip(results)

            st.download_button(
                "Download ZIP",
                zip_file,
                file_name="stok_results.zip"
            )
