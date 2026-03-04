# app.py
import streamlit as st
import pandas as pd

from specs import SPECS
from engine import (
    build_pricelist_map,
    build_addon_map,
    process_price_inplace,
    process_stock_inplace,
    process_discount_template,
    make_zip,
    build_stock_map_from_pricelist_sheets,
    build_stock_map_from_column,
)

st.set_page_config(page_title="sellerengine", page_icon="⚙️", layout="wide")
st.title("sellerengine")

# =========================
# Helpers (KEY UNIK)
# =========================
def uploader_pricelist_addon(prefix: str):
    c1, c2 = st.columns(2)
    with c1:
        pricelist = st.file_uploader("Upload Pricelist (XLSX)", type=["xlsx"], key=f"{prefix}_pl")
    with c2:
        addon = st.file_uploader("Upload Addon (XLSX)", type=["xlsx"], key=f"{prefix}_ad")
    return pricelist, addon

def build_maps(spec: dict, pricelist_uploader, addon_uploader):
    pl_map = build_pricelist_map(
        pricelist_bytes=pricelist_uploader.getvalue(),
        sheet_name=spec["pricelist"]["sheet_name"],
        header_row=spec["pricelist"]["header_row"],
        sku_header_candidates=spec["pricelist"]["sku_header_candidates"],
        price_col_letter=spec["pricelist"]["price_col_letter"],
    )
    addon_map = build_addon_map(
        addon_bytes=addon_uploader.getvalue(),
        code_candidates=spec["addon"]["code_candidates"],
        price_candidates=spec["addon"]["price_candidates"],
    )
    return pl_map, addon_map

def download_outputs(out_files, zip_name: str, single_label="Download XLSX"):
    if not out_files:
        st.warning("Tidak ada baris yang berubah / semua row skip.")
        return
    if len(out_files) == 1:
        name, data = out_files[0]
        st.download_button(
            single_label,
            data=data,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        z = make_zip(out_files)
        st.download_button("Download ZIP", data=z, file_name=zip_name, mime="application/zip")

# =========================
# Tabs utama
# =========================
tab_produk, tab_promosi, tab_gudang = st.tabs(["Harga Normal", "Harga Coret", "Update Stok"])

# =========================================================
# TAB: HARGA NORMAL (✅ diskon manual ada)
# =========================================================
with tab_produk:
    st.subheader("Harga Normal")

    discount_rp = st.number_input("Diskon Manual (Rp) — berlaku untuk Harga Normal", min_value=0, value=0, step=1000, key="hn_disc")
    only_changed = st.checkbox("Hanya baris yang berubah", value=True, key="hn_only")

    t1, t2, t3, t4 = st.tabs(["TikTok", "Shopee", "PowerMerchant", "BigSeller"])

    # TikTok
    with t1:
        prefix = "hn_tt"
        spec = SPECS[("harga_normal", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS TikTok", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files, changes_all = [], []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, changes = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                changes_all.extend(changes)
                prog.progress((i + 1) / len(templates))

            if changes_all:
                st.dataframe(pd.DataFrame([c.__dict__ for c in changes_all]).head(300), use_container_width=True)
            download_outputs(out_files, "harga_normal_tiktok.zip")

    # Shopee
    with t2:
        prefix = "hn_sp"
        spec = SPECS[("harga_normal", "shopee")]
        templates = st.file_uploader("Upload Template Shopee (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS Shopee", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))
            download_outputs(out_files, "harga_normal_shopee.zip")

    # PowerMerchant
    with t3:
        prefix = "hn_pm"
        spec = SPECS[("harga_normal", "powermerchant")]
        templates = st.file_uploader("Upload Template PowerMerchant (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS PowerMerchant", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))
            download_outputs(out_files, "harga_normal_powermerchant.zip")

    # BigSeller
    with t4:
        prefix = "hn_bs"
        spec = SPECS[("harga_normal", "bigseller")]
        templates = st.file_uploader("Upload Template BigSeller (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS BigSeller", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))
            download_outputs(out_files, "harga_normal_bigseller.zip")

# =========================================================
# TAB: HARGA CORET (✅ diskon manual berlaku untuk semua)
# - TikTok & PowerMerchant pakai DISCOUNT TEMPLATE + split 1000
# - Shopee tetap in-place (Harga diskon)
# =========================================================
with tab_promosi:
    st.subheader("Harga Coret")

    discount_rp = st.number_input("Diskon Manual (Rp) — berlaku untuk Harga Coret", min_value=0, value=0, step=1000, key="hc_disc")
    only_changed = st.checkbox("Hanya baris yang berubah", value=True, key="hc_only")

    p1, p2, p3 = st.tabs(["TikTok (Discount Template)", "PowerMerchant (Discount Template)", "Shopee (In-place)"])

    # TikTok discount template (M3)
    with p1:
        prefix = "hc_tt"
        spec = SPECS[("harga_coret_discount_template", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok Discount (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS TikTok", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files, previews = [], []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                files_out, df_prev = process_discount_template(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
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

    # PowerMerchant discount template (M4) ✅
    with p2:
        prefix = "hc_pm"
        spec = SPECS[("harga_coret_discount_template", "powermerchant")]
        templates = st.file_uploader("Upload Template PowerMerchant Discount (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS PowerMerchant", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files, previews = [], []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                files_out, df_prev = process_discount_template(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
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

    # Shopee in-place
    with p3:
        prefix = "hc_sp"
        spec = SPECS[("harga_coret", "shopee")]
        templates = st.file_uploader("Upload Template Shopee Coret (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS Shopee", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=int(discount_rp),
                    only_changed=only_changed,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))
            download_outputs(out_files, "harga_coret_shopee.zip")

# =========================================================
# TAB: UPDATE STOK
# ✅ pilih Nasional / Area / Toko kembali
# ✅ baca sheet LAPTOP..SER OTH CON
# =========================================================
with tab_gudang:
    st.subheader("Update Stok")

    g1, g2 = st.tabs(["TikTok", "Shopee"])

    def stok_ui(platform_key: str, prefix: str):
        spec = SPECS[("update_stok", platform_key)]
        templates = st.file_uploader(f"Upload Template {platform_key.upper()} (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        stock_file = st.file_uploader("Upload File Stok (XLSX)", type=["xlsx"], key=f"{prefix}_stock")

        mode = st.radio("Mode Stok", ["Nasional", "Area", "Toko"], horizontal=True, key=f"{prefix}_mode")

        if st.button(f"PROCESS {platform_key.upper()}", type="primary", key=f"{prefix}_go"):
            if not templates or stock_file is None:
                st.error("Upload template + file stok dulu.")
                st.stop()

            sheets_from = spec["stock_source"]["sheets_from"]
            sheets_to = spec["stock_source"]["sheets_to"]

            # baca stok default (nasional) + ambil daftar kolom stok yang tersedia
            stock_default, qty_cols = build_stock_map_from_pricelist_sheets(
                stock_file_bytes=stock_file.getvalue(),
                sheets_from=sheets_from,
                sheets_to=sheets_to,
            )

            # normalize qty_cols sudah uppercase karena df.columns di engine normalize
            qty_cols_sorted = sorted(qty_cols)

            if mode == "Nasional":
                # pakai map default (TOT jika ada)
                if not stock_default:
                    st.error("Mode Nasional butuh kolom TOT/TOTAL/NASIONAL di file stok. Tidak ditemukan.")
                    st.stop()
                stock_map = stock_default
            else:
                st.info("Pilih kolom stok yang mau dipakai (Area/Toko) dari file stok kamu.")
                chosen_col = st.selectbox("Pilih Kolom Stok", options=qty_cols_sorted, key=f"{prefix}_col")
                stock_map = build_stock_map_from_column(
                    stock_file_bytes=stock_file.getvalue(),
                    sheets_from=sheets_from,
                    sheets_to=sheets_to,
                    qty_col_name_norm=chosen_col,
                )

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _ = process_stock_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    stock_value_map=stock_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_stok_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, f"stok_{platform_key}.zip")

    with g1:
        stok_ui("tiktok", "st_tt")

    with g2:
        stok_ui("shopee", "st_sp")
