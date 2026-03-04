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
        "pricelist": {"header_row": 2, "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"], "price_col_letter": "M3"},
        "addon": {"code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
                  "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"]},
        "pricing": {"mode": "base_plus_addon"},
    },

    ("harga_normal", "powermerchant"): {
        # dari file kamu: layout template sama seperti tiktok normal, tapi pricelist pakai M4 (lihat script power merchant)
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "price_headers": ["Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {"header_row": 2, "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"], "price_col_letter": "M4"},
        "addon": {"code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
                  "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"]},
        "pricing": {"mode": "base_plus_addon"},
    },

    ("harga_normal", "shopee"): {
        # Shopee normal di file kamu pakai pendekatan berbeda (pandas),
        # tapi engine ini tetap pakai openpyxl agar template identik.
        # Header/data row untuk Shopee normal tidak kamu sebut di awal chat;
        # kalau template Shopee normal kamu sebenarnya sama model header row 3/data 7 seperti stok,
        # silakan sesuaikan angka ini bila perlu.
        "template": {
            "header_row": 3,
            "data_start_row": 7,
            "sku_headers": ["SKU", "SKU Ref. No.(Optional)", "SKU\u00a0Ref.\u00a0No.(Optional)"],
            "price_headers": ["Harga", "Price", "Harga Normal", "Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {"header_row": 2, "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"], "price_col_letter": "M4"},
        "addon": {"code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
                  "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"]},
        "pricing": {"mode": "base_plus_addon"},
    },

    # =========================
    # HARGA CORET (Shopee/PM pakai template yang sama model "update in-place")
    # =========================
    ("harga_coret", "shopee"): {
        "template": {
            "header_row": 1,
            "data_start_row": 2,
            "sku_headers": ["SKU Ref. No.(Optional)", "SKU\u00a0Ref.\u00a0No.(Optional)", "SKU"],
            "price_headers": ["Harga diskon"],
        },
        "pricelist": {"header_row": 2, "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"], "price_col_letter": "M4"},
        "addon": {"code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
                  "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"]},
        "pricing": {"mode": "base_plus_addon"},
    },

    ("harga_coret", "powermerchant"): {
        # di file kamu: Pricelist pakai M4
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "price_headers": ["Harga Ritel (Mata Uang Lokal)"],
        },
        "pricelist": {"header_row": 2, "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"], "price_col_letter": "M4"},
        "addon": {"code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
                  "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"]},
        "pricing": {"mode": "base_plus_addon"},
    },

    # =========================
    # UPDATE STOK
    # =========================
    ("update_stok", "tiktok"): {
        "template": {
            "header_row": 3,
            "data_start_row": 6,
            "sku_headers": ["SKU Penjual", "Seller SKU"],
            "stock_headers": ["Kuantitas"],
        },
    },

    ("update_stok", "shopee"): {
        "template": {
            "header_row": 3,
            "data_start_row": 7,
            "sku_headers": ["SKU"],
            "stock_headers": ["Stok"],
        },
    },

    # =========================
    # TIKTOK DISKON / HARGA CORET (template output berbeda + max 1000)
    # =========================
    ("harga_coret_tiktok_discount", "tiktok"): {
        "input": {
            "header_row": 3,
            "data_start_row": 6,
            # default kolom (sesuai file kamu)
            "col_product_id": "A",
            "col_sku_id": "D",
            "col_price": "F",
            "col_stock": "G",
            "col_seller_sku": "H",  # fallback E jika kosong
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
        "pricelist": {"header_row": 2, "sku_header_candidates": ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO"], "price_col_letter": "M3"},
        "addon": {"code_candidates": ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"],
                  "price_candidates": ["harga", "HARGA", "Price", "PRICE", "Harga"]},
    },
}
