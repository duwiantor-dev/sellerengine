# specs.py

SPECS = {
    # =========================
    # HARGA NORMAL
    # =========================
    ("harga_normal", "tiktok"): {
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "price_headers": ["Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {
            "header_row": 2,
            "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"],
            "price_col_letter": "M3",
        },
        "addon": {
            "code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
            "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"],
        },
    },

    ("harga_normal", "powermerchant"): {
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "price_headers": ["Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {
            "header_row": 2,
            "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"],
            "price_col_letter": "M4",
        },
        "addon": {
            "code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
            "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"],
        },
    },

    ("harga_normal", "shopee"): {
        # Kalau template Shopee normal kamu beda header/data row,
        # ubah 2 angka ini saja.
        "template": {
            "header_row": 3,
            "data_start_row": 7,
            "sku_headers": ["SKU", "SKU Ref. No.(Optional)", "SKU\u00a0Ref.\u00a0No.(Optional)"],
            "price_headers": ["Harga", "Price", "Harga Normal", "Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {
            "header_row": 2,
            "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"],
            "price_col_letter": "M4",
        },
        "addon": {
            "code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
            "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"],
        },
    },

    # =========================
    # HARGA CORET (PROMOSI) - in-place
    # Diskon manual akan dikurangkan di engine.py
    # =========================
    ("harga_coret", "shopee"): {
        "template": {
            "header_row": 1,
            "data_start_row": 2,
            "sku_headers": ["SKU Ref. No.(Optional)", "SKU\u00a0Ref.\u00a0No.(Optional)", "SKU"],
            "price_headers": ["Harga diskon", "Discount Price", "Harga Diskon"],
        },
        "pricelist": {
            "header_row": 2,
            "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"],
            "price_col_letter": "M4",
        },
        "addon": {
            "code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
            "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"],
        },
    },

    ("harga_coret", "powermerchant"): {
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "price_headers": ["Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {
            "header_row": 2,
            "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"],
            "price_col_letter": "M4",
        },
        "addon": {
            "code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
            "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"],
        },
    },

    # =========================
    # UPDATE STOK
    # =========================
    ("update_stok", "tiktok"): {
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "stock_headers": ["Kuantitas", "Quantity", "Qty"],
        },
    },

    ("update_stok", "shopee"): {
        "template": {
            "header_row": 3,
            "data_start_row": 7,
            "sku_headers": ["SKU"],
            "stock_headers": ["Stok", "Stock"],
        },
    },

    # =========================
    # TikTok discount template (Product Discount) + split 1000
    # =========================
    ("harga_coret_tiktok_discount", "tiktok"): {
        "input": {
            "data_start_row": 6,
            "col_product_id": "A",
            "col_sku_id": "D",
            "col_price": "F",
            "col_stock": "G",
            "col_seller_sku": "H",  # fallback ke E jika kosong
        },
        "output": {
            "max_rows_per_file": 1000,
            "headers": [
                "Product_id (wajib)",
                "SKU_id (wajib)",
                "Harga Penawaran (wajib)",
                "Total Stok Promosi (opsional)\n1. Total Stok Promosi≤ Stok \n2. Jika tidak diisi artinya tidak terbatas",
                "Batas Pembelian (opsional)\n1. 1 ≤ Batas pembelian≤ 99\n2. Jika tidak diisi artinya tidak terbatas",
            ],
        },
        "pricelist": {
            "header_row": 2,
            "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"],
            "price_col_letter": "M3",
        },
        "addon": {
            "code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
            "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"],
        },
    },
}
