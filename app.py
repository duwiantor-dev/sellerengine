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
st.title("sellerengine")

menu = st.sidebar.radio("Menu", ["Harga Normal", "Harga Coret", "Update Stok"])

if menu == "Harga Normal":
    marketplace = st.selectbox("Pilih Marketplace", ["tiktok", "shopee", "powermerchant"])
    job_key = ("harga_normal", marketplace)

    st.subheader(f"Harga Normal — {marketplace.upper()}")

    templates = st.file_uploader("Upload Template (boleh multi file)", type=["xlsx"], accept_multiple_files=True)
    pricelist = st.file_uploader("Upload Pricelist", type=["xlsx"])
    addon = st.file_uploader("Upload Addon", type=["xlsx"])

    if st.button("PROCESS", type="primary"):
        spec = SPECS[job_key]

        if not templates or pricelist is None or addon is None:
            st.error("Upload template + pricelist + addon dulu.")
            st.stop()

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

        out_files = []
        all_changes = []

        prog = st.progress(0)
        for i, f in enumerate(templates):
            out_bytes, changes, _issues = process_price_inplace(
                template_bytes=f.getvalue(),
                template_name=f.name,
                spec=spec,
                pricelist_map=pl_map,
                addon_map=addon_map,
            )
            if out_bytes:
                out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
            all_changes.extend(changes)
            prog.progress((i + 1) / len(templates))

        if all_changes:
            df = pd.DataFrame([c.__dict__ for c in all_changes])
            st.dataframe(df.head(300), use_container_width=True)

        if not out_files:
            st.warning("Tidak ada baris yang berubah.")
        elif len(out_files) == 1:
            name, data = out_files[0]
            st.download_button("Download XLSX", data=data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            z = make_zip(out_files)
            st.download_button("Download ZIP", data=z, file_name="hasil_harga_normal.zip", mime="application/zip")


elif menu == "Harga Coret":
    marketplace = st.selectbox("Pilih Marketplace", ["tiktok (discount template)", "shopee", "powermerchant"])

    if marketplace.startswith("tiktok"):
        job_key = ("harga_coret_tiktok_discount", "tiktok")
        spec = SPECS[job_key]

        st.subheader("Harga Coret — TikTok (Product Discount Template, split 1000)")

        templates = st.file_uploader("Upload Template TikTok Discount (boleh multi file)", type=["xlsx"], accept_multiple_files=True)
        pricelist = st.file_uploader("Upload Pricelist", type=["xlsx"])
        addon = st.file_uploader("Upload Addon", type=["xlsx"])
        discount_rp = st.number_input("Diskon (Rp)", min_value=0, value=0, step=1000)
        only_changed = st.checkbox("Hanya yang berubah harga", value=True)

        if st.button("PROCESS", type="primary"):
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

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

            if not out_files:
                st.warning("Tidak ada output (mungkin tidak ada yang berubah / semua skip).")
            elif len(out_files) == 1:
                name, data = out_files[0]
                st.download_button("Download XLSX", data=data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                z = make_zip(out_files)
                st.download_button("Download ZIP", data=z, file_name="tiktok_discount_output.zip", mime="application/zip")

    else:
        marketplace_key = "shopee" if marketplace == "shopee" else "powermerchant"
        job_key = ("harga_coret", marketplace_key)

        st.subheader(f"Harga Coret — {marketplace_key.upper()}")

        templates = st.file_uploader("Upload Template (boleh multi file)", type=["xlsx"], accept_multiple_files=True)
        pricelist = st.file_uploader("Upload Pricelist", type=["xlsx"])
        addon = st.file_uploader("Upload Addon", type=["xlsx"])

        if st.button("PROCESS", type="primary"):
            spec = SPECS[job_key]
            if not templates or pricelist is None or addon is None:
                st.error("Upload template + pricelist + addon dulu.")
                st.stop()

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

            out_files = []
            prog = st.progress(0)
            for i, f in enumerate(templates):
                out_bytes, _changes, _issues = process_price_inplace(
                    template_bytes=f.getvalue(),
                    template_name=f.name,
                    spec=spec,
                    pricelist_map=pl_map,
                    addon_map=addon_map,
                )
                if out_bytes:
                    out_files.append((f.name.replace(".xlsx", "_changed.xlsx"), out_bytes))
                prog.progress((i + 1) / len(templates))

            if not out_files:
                st.warning("Tidak ada baris yang berubah.")
            elif len(out_files) == 1:
                name, data = out_files[0]
                st.download_button("Download XLSX", data=data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                z = make_zip(out_files)
                st.download_button("Download ZIP", data=z, file_name="hasil_harga_coret.zip", mime="application/zip")


else:  # Update Stok
    marketplace = st.selectbox("Pilih Marketplace", ["tiktok", "shopee"])
    job_key = ("update_stok", marketplace)
    spec = SPECS[job_key]

    st.subheader(f"Update Stok — {marketplace.upper()}")

    templates = st.file_uploader("Upload Template (boleh multi file)", type=["xlsx"], accept_multiple_files=True)

    st.caption("Upload file stok (XLSX): minimal ada kolom KODEBARANG/KODE BARANG dan kolom TOT (stok nasional).")
    stock_file = st.file_uploader("Upload File Stok", type=["xlsx"])

    if st.button("PROCESS", type="primary"):
        if not templates or stock_file is None:
            st.error("Upload template + file stok dulu.")
            st.stop()

        # Build stock map dari file stok: ambil SKU dari KODEBARANG/KODE BARANG dan qty dari TOT
        from openpyxl import load_workbook
        import io
        import re

        def norm_sku(v) -> str:
            s = ("" if v is None else str(v)).strip().upper()
            if re.fullmatch(r"\d+\.0", s):
                s = s[:-2]
            s = re.sub(r"\s+", "", s)
            return s

        wb = load_workbook(io.BytesIO(stock_file.getvalue()), data_only=True)
        ws = wb.active

        # cari header row (scan 1..12), cari kolom KODEBARANG dan TOT
        header_row = None
        sku_col = None
        tot_col = None

        for r in range(1, min(12, ws.max_row) + 1):
            row_map = {}
            for c in range(1, ws.max_column + 1):
                row_map[(ws.cell(r, c).value or "").__str__().strip().upper()] = c

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
            st.error("File stok: tidak menemukan kolom KODEBARANG/KODE BARANG dan TOT.")
            st.stop()

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

        if not out_files:
            st.warning("Tidak ada baris yang berubah.")
        elif len(out_files) == 1:
            name, data = out_files[0]
            st.download_button("Download XLSX", data=data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            z = make_zip(out_files)
            st.download_button("Download ZIP", data=z, file_name="hasil_update_stok.zip", mime="application/zip")
