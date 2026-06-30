# -*- coding: utf-8 -*-
"""
Modul: system_config.py
Tanggung Jawab: Menyimpan kamus aturan pemusnahan (Blacklist), perlindungan (Whitelist),
               dan konstanta pengaman warna tab (Red Tab Safeguard).
Arsitektur: Pasif, sebagai single source of truth untuk regulasi dokumen Excel.
Developed by: An_
"""

# =============================================================================
# 🚨 1. CONSTANTS PROTOKOL PENGAMAN (SAFETY CONSTANTS)
# =============================================================================
# Kode warna Hex RGB untuk Tab Merah di Excel (Protokol Red Tab Safeguard).
# Sheet dengan warna tab ini akan di-bypass total dari modifikasi/penghapusan.
RED_TAB_COLOR = "FF0000"

# =============================================================================
# 📋 2. DAFTAR PERLINDUNGAN UTAMAA (WHITELIST SHEET)
# =============================================================================
# Daftar nama-nama sheet (dalam huruf kecil) yang sifatnya statis/sakral 
# dan wajib dilindungi dari pembersihan massal atau modifikasi struktural.
WHITELIST_GLOBAL_SHEETS = [
    "cover",
    "atp cw cover",
    "cover cw opm",
    "ba lapangan",
    "ba rectifikasi opm",
    "ba rectifikasi cw atp",
    "cw punchpoint defect list",
    "opm punchpoint defect list",
    "support table",
    "fdt"
]

# =============================================================================
# ❌ 3. DAFTAR PEMUSNAHAN SHEET MODE CLUSTER (DISTRIBUSI HILIR)
# =============================================================================
# Kata kunci (huruf kecil) untuk sheet yang HARUS dimusnahkan secara otomatis 
# oleh mesin ketika operator memilih mode "Cluster Jaringan".
BLACKLIST_CLUSTER = [
    "opm subfeeder",
    "otdr summary 1550",
    "otdr summary 1310",
    "bast key"
]

# =============================================================================
# ❌ 4. DAFTAR PEMUSNAHAN SHEET MODE SUBFEEDER (BACKBONE HULU)
# =============================================================================
# Kata kunci (huruf kecil) untuk sheet yang HARUS dimusnahkan secara otomatis
# oleh mesin ketika operator memilih mode "Subfeeder Jaringan".
# Seluruh jejak rute distribusi hilir, splitter FAT, dan OPM distribusi dibersihkan.
BLACKLIST_SUBFEEDER = [
    "otdr summary (fdt-fat)",
    "otdr sumary (fdt-fat)",
    "opm distribution",
    "opm feeder",
    "ba splitter fat",
    "ba spitter fat"
]

# =============================================================================
# 🔑 5. MATRIKS METADATA MAPPING (FOR BACKWARD COMPATIBILITY)
# =============================================================================
# Kamus internal untuk memastikan sinkronisasi antara kunci variabel JSON 
# dan penulisan token placeholder asli di dalam file Template.xlsx.
METADATA_TOKEN_MAP = {
    "NAMA_PROYEK": "[NAMA_PROYEK]",
    "REGION": "[REGION]",
    "NAMA_LOKASI": "[NAMA_LOKASI]",
    "ID_LOKASI": "[ID_LOKASI]",
    "ALAMAT": "[ALAMAT]",
    "NAMA_OLT": "[NAMA_OLT]",
    "ID_FDT_FROM": "[ID_FDT_FROM]",
    "ID_FAT_TO": "[ID_FAT_TO]",
    "NAMA_PT_VENDOR": "[NAMA_PT_VENDOR]",
    "REP_VENDOR": "[REP_VENDOR]",
    "JABATAN_VENDOR": "[JABATAN_VENDOR]",
    "NAMA_PT_CUSTOMER": "[NAMA_PT_CUSTOMER]",
    "REP_CUSTOMER": "[REP_CUSTOMER]",
    "JABATAN_CUSTOMER": "[JABATAN_CUSTOMER]",
    "TANGGAL_TEST": "[TANGGAL_TEST]",
    "NO_PO": "[NO_PO]"
}
