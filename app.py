# app.py
import streamlit as st
import pandas as pd

from specs import SPECS
from engine import (
    build_pricelist_map,
    build_addon_map,
    process_price_inplace,
    process_discount_template,
    process_stock_inplace,
    build_stock_dataframe_from_range,
    build_stock_map_from_df,
    make_zip,
)

st.set_page_config(page_title="sellerengine", page_icon="⚙️", layout="wide")

# ✅ Compact UI CSS
st.markdown("""
<style>
/* kecilkan kotak upload */
[data-testid="stFileUploaderDropzone"] {
    min-height: 64px;
    padding: 8px 10px;
}
[data-testid="stFileUploaderDropzone"] * {
    font-size: 12px;
    line-height: 1.2;
}
/* kecilkan number input */
div[data-testid="stNumberInput"] input {
    padding-top: 6px;
    padding-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)

st.title("sellerengine")

def ui_row_uploads_and_discount(prefix: str, label_template: str):
    """1 baris: [template multi] [pricelist] [addon] [diskon]"""
    c1, c2, c3, c4 = st.columns([3.2, 2.2, 2.2, 1.4])

    with c1:
        templates = st.file_uploader(
            label_template,
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"{prefix}_tpl",
        )

    with c2:
        pricelist = st.file_uploader(
            "Upload Pricelist",
            type=["xlsx"],
            key=f"{prefix}_pl",
        )

    with c3:
        addon = st.file_uploader(
            "Upload Addon",
            type=["xlsx"],
            key=f"{prefix}_ad",
        )

    with c4:
        discount_rp = st.number_input(
            "Diskon Manual",
            min_value=0,
            value=0,
            step=1000,
            key=f"{prefix}_disc",
        )

    return templates, pricelist, addon, int(discount_rp)

def download_outputs(out_files, zip_name: str):
    if not out_files:
        st.warning("Tidak ada baris yang berubah / semua baris skip.")
        return

    if len(out_files) == 1:
        name, data = out_files[0]
        st.download_button(
            "Download XLSX",
            data=data,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        z = make_zip(out_files)
        st.download_button("Download ZIP", data=z, file_name=zip_name, mime="application/zip")

def build_maps_price(spec: dict, pricelist_uploader, addon_uploader):
    pl_map = build_pricelist_map(
        pricelist_bytes=pricelist_uploader.getvalue(),
        sheet_name=spec["pricelist"]["sheet_name"],
        header_row=spec["pricelist"]["header_row"],
        sku_header_candidates=spec["pricelist"]["sku_header_candidates"],
        price_col_letter=spec["pricelist"]["price_col_letter"],
    )
    ad_map = build_addon_map(
        addon_bytes=addon_uploader.getvalue(),
        code_candidates=spec["addon"]["code_candidates"],
        price_candidates=spec["addon"]["price_candidates"],
    )
    return pl_map, ad_map

# =========================
# Tabs utama
# =========================
tab_hn, tab_hc, tab_st = st.tabs(["Harga Normal", "Harga Coret", "Update Stok"])

# =========================================================
# HARGA NORMAL
# =========================================================
with tab_hn:
    t1, t2, t3, t4 = st.tabs(["TikTok", "Shopee", "PowerMerchant", "BigSeller"])

    def harga_normal_ui(platform: str, prefix: str, label: str):
        spec = SPECS[("harga_normal", platform)]
        templates, pricelist, addon, discount_rp = ui_row_uploads_and_discount(prefix, label)
        only_changed = st.checkbox("Hanya baris yang berubah", value=True, key=f"{prefix}_only")

        if st.button("PROCESS", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Wajib upload: Template + Pricelist + Addon.")
                st.stop()

            pl_map, ad_map = build_maps_price(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _changes = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=ad_map,
                    discount_rp=int(discount_rp),     # ✅ diskon manual harga normal
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, f"harga_normal_{platform}.zip")

    with t1:
        harga_normal_ui("tiktok", "hn_tt", "Upload File Mass Update (TikTok)")

    with t2:
        harga_normal_ui("shopee", "hn_sp", "Upload File Mass Update (Shopee)")

    with t3:
        harga_normal_ui("powermerchant", "hn_pm", "Upload File Mass Update (PowerMerchant)")

    with t4:
        harga_normal_ui("bigseller", "hn_bs", "Upload File Mass Update (BigSeller)")

# =========================================================
# HARGA CORET
# - TikTok & PowerMerchant: Discount Template (split 1000)
# - Shopee: in-place harga diskon
# =========================================================
with tab_hc:
    c1, c2, c3 = st.tabs(["TikTok (Discount Template)", "PowerMerchant (Discount Template)", "Shopee (Harga Diskon)"])

    # TikTok discount template
    with c1:
        platform = "tiktok"
        prefix = "hc_tt"
        spec = SPECS[("discount_template", platform)]

        templates, pricelist, addon, discount_rp = ui_row_uploads_and_discount(prefix, "Upload File Mass Update (TikTok Discount)")
        only_changed = st.checkbox("Hanya baris yang berubah", value=True, key=f"{prefix}_only")

        if st.button("PROCESS", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Wajib upload: Template + Pricelist + Addon.")
                st.stop()

            pl_map, ad_map = build_maps_price(spec, pricelist, addon)

            out_files = []
            previews = []
            prog = st.progress(0)

            for i, f in enumerate(templates):
                files_out, df_prev = process_discount_template(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=ad_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                out_files.extend(files_out)
                if df_prev is not None and len(df_prev) > 0:
                    previews.append(df_prev)
                prog.progress((i + 1) / len(templates))

            if previews:
                st.dataframe(pd.concat(previews, ignore_index=True).head(300), use_container_width=True)

            download_outputs(out_files, "tiktok_discount_output.zip")

    # PowerMerchant discount template (M4)
    with c2:
        platform = "powermerchant"
        prefix = "hc_pm"
        spec = SPECS[("discount_template", platform)]

        templates, pricelist, addon, discount_rp = ui_row_uploads_and_discount(prefix, "Upload File Mass Update (PowerMerchant Discount)")
        only_changed = st.checkbox("Hanya baris yang berubah", value=True, key=f"{prefix}_only")

        if st.button("PROCESS", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Wajib upload: Template + Pricelist + Addon.")
                st.stop()

            pl_map, ad_map = build_maps_price(spec, pricelist, addon)

            out_files = []
            previews = []
            prog = st.progress(0)

            for i, f in enumerate(templates):
                files_out, df_prev = process_discount_template(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=ad_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                out_files.extend(files_out)
                if df_prev is not None and len(df_prev) > 0:
                    previews.append(df_prev)
                prog.progress((i + 1) / len(templates))

            if previews:
                st.dataframe(pd.concat(previews, ignore_index=True).head(300), use_container_width=True)

            download_outputs(out_files, "powermerchant_discount_output.zip")

    # Shopee in-place harga diskon
    with c3:
        platform = "shopee"
        prefix = "hc_sp"
        spec = SPECS[("harga_coret", platform)]

        templates, pricelist, addon, discount_rp = ui_row_uploads_and_discount(prefix, "Upload File Mass Update (Shopee Harga Diskon)")
        only_changed = st.checkbox("Hanya baris yang berubah", value=True, key=f"{prefix}_only")

        if st.button("PROCESS", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Wajib upload: Template + Pricelist + Addon.")
                st.stop()

            pl_map, ad_map = build_maps_price(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)

            for i, f in enumerate(templates):
                out_bytes, _changes = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=ad_map,
                    discount_rp=int(discount_rp),  # ✅ diskon manual juga berlaku
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, "harga_coret_shopee.zip")

# =========================================================
# UPDATE STOK
# (Diskon manual di row ada tapi diabaikan sesuai rule "kecuali stok")
# Pricelist uploader dipakai sebagai "File stok sumber" (LAPTOP..SER OTH CON)
# =========================================================
with tab_st:
    s1, s2 = st.tabs(["TikTok", "Shopee"])

    def stok_ui(platform: str, prefix: str, label: str):
        spec = SPECS[("update_stok", platform)]
        sheets_from = spec["stock_source"]["sheets_from"]
        sheets_to = spec["stock_source"]["sheets_to"]

        # Pakai row yang sama agar konsisten:
        # - Template = file mass update marketplace
        # - Pricelist = file stok sumber (LAPTOP..SER OTH CON)
        # - Addon = tidak dipakai (biarkan kosong)
        # - Diskon = tidak dipakai (diabaikan)
        templates, stock_source_file, _addon_unused, _disc_unused = ui_row_uploads_and_discount(prefix, label)

        mode = st.radio("Mode Stok", ["Nasional", "Area", "Toko"], horizontal=True, key=f"{prefix}_mode")

        if st.button("PROCESS", type="primary", key=f"{prefix}_go"):
            if not templates or stock_source_file is None:
                st.error("Wajib upload: Template + File stok (pakai slot 'Upload Pricelist').")
                st.stop()

            df_stock = build_stock_dataframe_from_range(
                stock_file_bytes=stock_source_file.getvalue(),
                sheets_from=sheets_from,
                sheets_to=sheets_to,
            )

            # tentukan default kolom nasional
            candidates_nasional = ["TOT", "TOTAL", "NASIONAL"]
            default_nasional = None
            for c in candidates_nasional:
                if c in df_stock.columns:
                    default_nasional = c
                    break

            qty_cols = [c for c in df_stock.columns if c not in ["KODEBARANG", "KODE BARANG", "SKU"]]
            qty_cols_sorted = sorted(qty_cols)

            if mode == "Nasional":
                if default_nasional is None:
                    st.error("Mode Nasional butuh kolom TOT/TOTAL/NASIONAL di file stok.")
                    st.stop()
                qty_col = default_nasional
            else:
                qty_col = st.selectbox("Pilih Kolom Stok (Area/Toko)", options=qty_cols_sorted, key=f"{prefix}_qtycol")

            stock_map = build_stock_map_from_df(df_stock, qty_col=qty_col)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _changes = process_stock_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    stock_value_map=stock_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_stok_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, f"stok_{platform}.zip")

    with s1:
        stok_ui("tiktok", "st_tt", "Upload File Mass Update (TikTok Stok)")

    with s2:
        stok_ui("shopee", "st_sp", "Upload File Mass Update (Shopee Stok)")
