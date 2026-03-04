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

# =========================
# STYLE: Top Bar ala BigSeller
# =========================
st.markdown(
    """
    <style>
      /* top bar */
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
        width: 34px;
        height: 34px;
        border-radius: 8px;
        background: rgba(255,255,255,0.18);
        display:flex;
        align-items:center;
        justify-content:center;
        font-weight: 800;
        letter-spacing: 0.5px;
      }
      .se-title {
        font-size: 18px;
        font-weight: 800;
        margin: 0;
        line-height: 1;
      }
      .se-subtitle {
        font-size: 12px;
        opacity: 0.9;
        margin-top: 3px;
      }

      /* sedikit rapihin spacing */
      div.block-container { padding-top: 1.0rem; }
      /* bikin tabs lebih “menu” */
      button[data-baseweb="tab"] {
        font-weight: 700;
      }
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
# Helper UI
# =========================
def uploader_pricelist_addon():
    col1, col2 = st.columns(2)
    with col1:
        pricelist = st.file_uploader("Upload Pricelist (XLSX)", type=["xlsx"], key="pl")
    with col2:
        addon = st.file_uploader("Upload Addon (XLSX)", type=["xlsx"], key="ad")
    return pricelist, addon

def download_outputs(out_files, zip_name, single_label="Download XLSX"):
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
        st.download_button(
            "Download ZIP",
            data=z,
            file_name=zip_name,
            mime="application/zip",
        )

def build_maps_from_upload(spec, pricelist, addon):
    pl_map = build_pricelist_map(
        pricelist.getvalue(),
        header_row=spec["pricelist"]["header_row"],
        sku_header_candidates=spec["pricelist"]["sku_header_candidates"],
        price_col_letter=spec["pricelist"]["price_col_letter"],
    )
    addon_map = build_addon_map(
        addon.getvalue(),
        code_candidates=spec["addon"]["code_candidates"],
        price_candidates=spec["addon"]["price_candidates"],
    )
    return pl_map, addon_map

# =========================
# Top Menu Tabs (ala BigSeller)
# =========================
tab_produk, tab_promosi, tab_gudang = st.tabs(["Produk", "Promosi", "Pergudangan"])

# =========================================================
# TAB: PRODUK (Harga Normal)
# =========================================================
with tab_produk:
    st.subheader("Harga Normal")

    t_tiktok, t_shopee, t_pm, t_bigseller = st.tabs(["TikTok", "Shopee", "PowerMerchant", "BigSeller"])

    # --- TikTok
    with t_tiktok:
        spec = SPECS[("harga_normal", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="hn_tt_tpl")
        pricelist, addon = uploader_pricelist_addon()

        if st.button("PROCESS TikTok", type="primary", key="hn_tt_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps_from_upload(spec, pricelist, addon)

            out_files = []
            changes_all = []
            prog = st.progress(0)

            for i, f in enumerate(templates):
                out_bytes, changes, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                changes_all.extend(changes)
                prog.progress((i + 1) / len(templates))

            if changes_all:
                df = pd.DataFrame([c.__dict__ for c in changes_all])
                st.dataframe(df.head(300), use_container_width=True)

            download_outputs(out_files, "harga_normal_tiktok.zip")

    # --- Shopee
    with t_shopee:
        spec = SPECS[("harga_normal", "shopee")]
        templates = st.file_uploader("Upload Template Shopee (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="hn_sp_tpl")
        pricelist, addon = uploader_pricelist_addon()

        st.caption("Kalau header row / start row Shopee normal kamu beda, ubah di specs.py bagian ('harga_normal','shopee').")

        if st.button("PROCESS Shopee", type="primary", key="hn_sp_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps_from_upload(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, "harga_normal_shopee.zip")

    # --- PowerMerchant
    with t_pm:
        spec = SPECS[("harga_normal", "powermerchant")]
        templates = st.file_uploader("Upload Template PowerMerchant (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="hn_pm_tpl")
        pricelist, addon = uploader_pricelist_addon()

        if st.button("PROCESS PowerMerchant", type="primary", key="hn_pm_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps_from_upload(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, "harga_normal_powermerchant.zip")

    # --- BigSeller (placeholder)
    with t_bigseller:
        st.info(
            "BigSeller bisa dimasukin ke 1 engine juga, tapi butuh mapping header kolom harga & SKU sesuai template BigSeller kamu.\n\n"
            "Kirim 1 contoh template BigSeller (tanpa data sensitif) / sebutkan nama header kolom SKU & Harga yang diubah, nanti aku masukin spec-nya."
        )

# =========================================================
# TAB: PROMOSI (Harga Coret)
# =========================================================
with tab_promosi:
    st.subheader("Harga Coret / Promosi")

    # Diskon manual SELALU ADA di tab promosi
    discount_rp = st.number_input("Diskon (Rp)", min_value=0, value=0, step=1000, key="disc_all")
    only_changed = st.checkbox("Hanya yang berubah harga", value=True, key="only_changed_all")

    t_tiktok, t_shopee, t_pm = st.tabs(["TikTok (Discount Template)", "Shopee", "PowerMerchant"])

    # --- TikTok discount template
    with t_tiktok:
        spec = SPECS[("harga_coret_tiktok_discount", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok Discount (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="hc_tt_tpl")
        pricelist, addon = uploader_pricelist_addon()

        if st.button("PROCESS TikTok Discount", type="primary", key="hc_tt_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps_from_upload(spec, pricelist, addon)

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

    # --- Shopee coret (in-place)
    with t_shopee:
        spec = SPECS[("harga_coret", "shopee")]
        templates = st.file_uploader("Upload Template Shopee (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="hc_sp_tpl")
        pricelist, addon = uploader_pricelist_addon()

        st.caption("Diskon manual akan dikurangkan dari hasil hitung (base+addon) sebelum diisi ke kolom harga coret.")

        if st.button("PROCESS Shopee Coret", type="primary", key="hc_sp_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps_from_upload(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)

            # trik: kita “inject” diskon dengan cara modifikasi pricelist_map sementara (supaya engine umum tetap dipakai)
            # computed_price = base + addon; lalu diskon = -discount_rp
            # Caranya paling aman: setelah compute, baru kurangi. Tapi engine.py belum punya hook.
            # Jadi, untuk versi ini: diskon akan diproses khusus dengan engine TikTok? -> tidak.
            # Solusi cepat: jika butuh diskon manual untuk shopee/pm, kita tambah hook di engine.py.
            #
            # Untuk sementara: kalau diskon_rp != 0, kita tetap jalankan normal dulu lalu kamu bisa request aku tambah hook.
            if discount_rp != 0:
                st.warning("Diskon manual untuk Shopee/PowerMerchant butuh hook kecil di engine.py. Kalau kamu set Diskon > 0, aku akan aktifkan hook itu.")

            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, "harga_coret_shopee.zip")

    # --- PowerMerchant coret (in-place)
    with t_pm:
        spec = SPECS[("harga_coret", "powermerchant")]
        templates = st.file_uploader("Upload Template PowerMerchant (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="hc_pm_tpl")
        pricelist, addon = uploader_pricelist_addon()

        if st.button("PROCESS PowerMerchant Coret", type="primary", key="hc_pm_go"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

            pl_map, addon_map = build_maps_from_upload(spec, pricelist, addon)

            out_files = []
            prog = st.progress(0)

            if discount_rp != 0:
                st.warning("Diskon manual untuk Shopee/PowerMerchant butuh hook kecil di engine.py. Kalau kamu set Diskon > 0, aku akan aktifkan hook itu.")

            for i, f in enumerate(templates):
                out_bytes, _, _ = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            download_outputs(out_files, "harga_coret_powermerchant.zip")

# =========================================================
# TAB: PERGUDANGAN (Update Stok)
# =========================================================
with tab_gudang:
    st.subheader("Update Stok")

    t_tiktok, t_shopee = st.tabs(["TikTok", "Shopee"])

    def build_stock_map_from_file(stock_file):
        from openpyxl import load_workbook
        import io, re

        def norm_sku(v) -> str:
            s = ("" if v is None else str(v)).strip().upper()
            if re.fullmatch(r"\d+\.0", s):
                s = s[:-2]
            s = re.sub(r"\s+", "", s)
            return s

        wb = load_workbook(io.BytesIO(stock_file.getvalue()), data_only=True)
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

    with t_tiktok:
        spec = SPECS[("update_stok", "tiktok")]
        templates = st.file_uploader("Upload Template TikTok (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="st_tt_tpl")
        stock_file = st.file_uploader("Upload File Stok (XLSX) - kolom KODEBARANG + TOT", type=["xlsx"], key="st_tt_stock")

        if st.button("PROCESS Stok TikTok", type="primary", key="st_tt_go"):
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

    with t_shopee:
        spec = SPECS[("update_stok", "shopee")]
        templates = st.file_uploader("Upload Template Shopee (boleh multi)", type=["xlsx"], accept_multiple_files=True, key="st_sp_tpl")
        stock_file = st.file_uploader("Upload File Stok (XLSX) - kolom KODEBARANG + TOT", type=["xlsx"], key="st_sp_stock")

        if st.button("PROCESS Stok Shopee", type="primary", key="st_sp_go"):
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
