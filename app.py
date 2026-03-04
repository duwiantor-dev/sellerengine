# app.py
import streamlit as st
import pandas as pd

from specs import SPECS
from engine import (
    build_pricelist_map,
    build_addon_map,
    process_price_inplace,
    process_stock_inplace,
    process_tiktok_discount,
    make_zip,
)

st.set_page_config(page_title="sellerengine", page_icon="⚙️", layout="wide")

# ===== Top bar ungu =====
st.markdown(
    """
    <style>
      .se-topbar {
        background: #5B2BBF;
        color: white;
        padding: 14px 18px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        gap: 14px;
        margin-bottom: 12px;
      }
      .se-logo {
        width: 34px; height: 34px;
        border-radius: 8px;
        background: rgba(255,255,255,0.18);
        display:flex; align-items:center; justify-content:center;
        font-weight: 800;
        letter-spacing: 0.5px;
      }
      .se-title { font-size: 18px; font-weight: 800; margin: 0; line-height: 1; }
      .se-subtitle { font-size: 12px; opacity: 0.9; margin-top: 3px; }
      div.block-container { padding-top: 1.0rem; }
      button[data-baseweb="tab"] { font-weight: 700; }
    </style>
    <div class="se-topbar">
      <div class="se-logo">SE</div>
      <div>
        <div class="se-title">sellerengine</div>
        <div class="se-subtitle">Marketplace Bulk Tools</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# Helpers (KEY UNIK!)
# =========================
def uploader_pricelist_addon(prefix: str):
    col1, col2 = st.columns(2)
    with col1:
        pricelist = st.file_uploader("Upload Pricelist (XLSX)", type=["xlsx"], key=f"{prefix}_pl")
    with col2:
        addon = st.file_uploader("Upload Addon (XLSX)", type=["xlsx"], key=f"{prefix}_ad")
    return pricelist, addon


def build_maps(spec: dict, pricelist_uploader, addon_uploader):
    pl_map = build_pricelist_map(
        pricelist_uploader.getvalue(),
        header_row=spec["pricelist"]["header_row"],
        sku_header_candidates=spec["pricelist"]["sku_header_candidates"],
        price_col_letter=spec["pricelist"]["price_col_letter"],
    )
    addon_map = build_addon_map(
        addon_uploader.getvalue(),
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


def build_stock_map_from_file(stock_file_uploader):
    from openpyxl import load_workbook
    import io, re

    def norm_sku(v) -> str:
        s = ("" if v is None else str(v)).strip().upper()
        if re.fullmatch(r"\d+\.0", s):
            s = s[:-2]
        s = re.sub(r"\s+", "", s)
        return s

    wb = load_workbook(io.BytesIO(stock_file_uploader.getvalue()), data_only=True)
    ws = wb.active

    header_row = None
    sku_col = None
    tot_col = None

    for r in range(1, min(12, ws.max_row) + 1):
        row_map = {}
        for c in range(1, ws.max_column + 1):
            key = str(ws.cell(r, c).value or "").strip().upper()
            row_map[key] = c

        for k in ["KODEBARANG", "KODE BARANG", "SKU"]:
            if k in row_map:
                sku_col = row_map[k]
                break
        if "TOT" in row_map:
            tot_col = row_map["TOT"]

        if sku_col and tot_col:
            header_row = r
            break

    if not (header_row and sku_col and tot_col):
        raise ValueError("File stok: tidak menemukan kolom KODEBARANG/KODE BARANG dan TOT.")

    stock_map = {}
    for r in range(header_row + 1, ws.max_row + 1):
        sku = norm_sku(ws.cell(r, sku_col).value)
        if not sku:
            continue
        qty = ws.cell(r, tot_col).value
        try:
            qty = int(float(qty))
        except Exception:
            continue
        stock_map[sku] = qty

    return stock_map


# =========================
# Main tabs (atas)
# =========================
tab_produk, tab_promosi, tab_gudang = st.tabs(["Produk", "Promosi", "Pergudangan"])

# =========================================================
# PRODUK: Harga Normal
# =========================================================
with tab_produk:
    st.subheader("Harga Normal")
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
                out_bytes, changes, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=0,
                    only_changed=True,
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
        st.caption("Jika template Shopee normal beda, ubah header_row/data_start_row di specs.py.")

        if st.button("PROCESS Shopee", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()
            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=0,
                    only_changed=True,
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
                out_bytes, _, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                    discount_rp=0,
                    only_changed=True,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, "harga_normal_powermerchant.zip")

    # BigSeller placeholder
    with t4:
        st.info(
            "BigSeller bisa dimasukin ke engine ini juga.\n\n"
            "Tapi aku butuh 2 info untuk mapping spec BigSeller:\n"
            "1) Nama header kolom SKU di template BigSeller\n"
            "2) Nama header kolom HARGA yang harus diupdate\n\n"
            "Kalau kamu upload 1 template BigSeller (dummy), aku bisa masukin spec-nya full."
        )

# =========================================================
# PROMOSI: Harga Coret (diskon manual berlaku untuk semua)
# =========================================================
with tab_promosi:
    st.subheader("Harga Coret / Promosi")

    discount_rp = st.number_input("Diskon (Rp)", min_value=0, value=0, step=1000, key="promo_disc")
    only_changed = st.checkbox("Hanya yang berubah harga", value=True, key="promo_only")

    pt1, pt2, pt3 = st.tabs(["TikTok (Discount Template)", "Shopee", "PowerMerchant"])

    # TikTok discount template
    with pt1:
        prefix = "hc_tt"
        spec = SPECS[("harga_coret_tiktok_discount", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok Discount (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS TikTok Discount", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            previews = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                files_out, df_prev = process_tiktok_discount(
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

    # Shopee coret (diskon manual berlaku)
    with pt2:
        prefix = "hc_sp"
        spec = SPECS[("harga_coret", "shopee")]
        templates = st.file_uploader("Upload Template Shopee (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS Shopee Coret", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
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

    # PowerMerchant coret (diskon manual berlaku)
    with pt3:
        prefix = "hc_pm"
        spec = SPECS[("harga_coret", "powermerchant")]
        templates = st.file_uploader("Upload Template PowerMerchant (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        pricelist, addon = uploader_pricelist_addon(prefix)

        if st.button("PROCESS PowerMerchant Coret", type="primary", key=f"{prefix}_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
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

            download_outputs(out_files, "harga_coret_powermerchant.zip")

# =========================================================
# PERGUDANGAN: Update Stok (tanpa diskon)
# =========================================================
with tab_gudang:
    st.subheader("Update Stok")

    gt1, gt2 = st.tabs(["TikTok", "Shopee"])

    with gt1:
        prefix = "st_tt"
        spec = SPECS[("update_stok", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        stock_file = st.file_uploader("Upload File Stok (XLSX) - kolom KODEBARANG + TOT", type=["xlsx"], key=f"{prefix}_stock")

        if st.button("PROCESS Stok TikTok", type="primary", key=f"{prefix}_go"):
            if not templates or stock_file is None:
                st.error("Upload template + file stok dulu.")
                st.stop()

            stock_map = build_stock_map_from_file(stock_file)

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

            download_outputs(out_files, "stok_tiktok.zip")

    with gt2:
        prefix = "st_sp"
        spec = SPECS[("update_stok", "shopee")]
        templates = st.file_uploader("Upload Template Shopee (boleh multi)", type=["xlsx"], accept_multiple_files=True, key=f"{prefix}_tpl")
        stock_file = st.file_uploader("Upload File Stok (XLSX) - kolom KODEBARANG + TOT", type=["xlsx"], key=f"{prefix}_stock")

        if st.button("PROCESS Stok Shopee", type="primary", key=f"{prefix}_go"):
            if not templates or stock_file is None:
                st.error("Upload template + file stok dulu.")
                st.stop()

            stock_map = build_stock_map_from_file(stock_file)

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

            download_outputs(out_files, "stok_shopee.zip")
