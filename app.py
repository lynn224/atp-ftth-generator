# -*- coding: utf-8 -*-
"""
Modul: app.py
Tanggung Jawab: Pusat konduktor (Orchestrator) utama aplikasi Web UI.
               Mengelola Tab navigasi, session state memori, dan pipa estafet RAM.
Arsitektur: Web Front-End (Streamlit) & Integrator Pipeline.
Developed by: An_
"""

import streamlit as st
import datetime
import random
import os
import io
import json
import pandas as pd
import traceback

# Integrasi Modul-Modul Modular Sistem
from system_config import METADATA_TOKEN_MAP
from login_engine import verify_login, scan_all_project_logs, register_user, delete_user_account, get_all_registered_users
from parser_engine import parse_all_fat, parse_all_poles
from ui_technical_forms import render_modul_2a_splitter, render_modul_2b_opm_distribution, render_modul_2c_2d_otdr_summary
from excel_injector_fase1 import inject_excel_fase1
from excel_injector_fase2 import inject_excel_fase2

TEMPLATE_PATH = "Template.xlsx"
DB_DIR = "history_database"
GREETINGS_PATH = "greetings.txt"
os.makedirs(DB_DIR, exist_ok=True)
st.set_page_config(page_title="Universal ATP Generator", layout="wide")
# =============================================================================
# ”9ä6 [BLOK RUANG STANDBY SISTEM LOGIN & HAK AKSES ADMIN]
# =============================================================================
# =============================================================================
# SISTEM AUTENTIKASI (PENGUNCI GERBANG)
# =============================================================================
if "is_logged_in" not in st.session_state:
    st.session_state.is_logged_in = False
    st.session_state.user_role = None
    st.session_state.user_full_name = None

if not st.session_state.is_logged_in:
    st.title("”9ä8 Log In - Universal ATP Generator")
    user = st.text_input("Username")
    pwd = st.text_input("Password", type="password")
    if st.button("Masuk"):
        user_info = verify_login(user, pwd)
        if user_info:
            st.session_state.is_logged_in = True
            st.session_state.user_role = user_info["role"]
            st.session_state.user_full_name = user_info["nama_lengkap"]
            st.rerun()
        else:
            st.error("Kredensial salah!")
    st.stop()

# =============================================================================
# PANEL KENDALI ADMINISTRATOR
# =============================================================================
if st.session_state.user_role == "administrator":
    st.title(f"7±5„1‚5 Panel Administrator - {st.session_state.user_full_name}")
    
    tab_admin1, tab_admin2 = st.tabs(["”9Ó5 Manajemen Operator DC", "”9Ý6 Pengawasan Proyek"])
    
    with tab_admin1:
        st.subheader("Registrasi Operator Baru")
        with st.form("reg_user"):
            new_u = st.text_input("Username")
            new_p = st.text_input("Password")
            new_n = st.text_input("Nama Lengkap")
            if st.form_submit_button("Daftar"):
                if register_user(new_u, new_p, new_n, "document_control"): st.success("User dibuat!")
        
        st.subheader("Daftar Operator")
        users = get_all_registered_users()
        st.table(pd.DataFrame.from_dict(users, orient='index'))
        
    with tab_admin2:
        st.subheader("Monitoring Proyek DC")
        logs = scan_all_project_logs()
        if logs: st.dataframe(pd.DataFrame(logs))
        else: st.info("Belum ada proyek yang dikerjakan.")
        
    if st.button("Log Out"): st.session_state.is_logged_in = False; st.rerun()
    st.stop()

# =============================================================================
# PANEL OPERASIONAL DOCUMENT CONTROL
# =============================================================================
st.sidebar.success(f"DC: {st.session_state.user_full_name}")
if st.sidebar.button("Log Out"): st.session_state.is_logged_in = False; st.rerun()

# if "authenticated" not in st.session_state:
#     st.session_state.authenticated = False
# =============================================================================

DEFAULT_METADATA = {
    "NAMA_PROYEK": "EMR FTTH PROJECT", "REGION": "JAWA TIMUR",
    "NAMA_LOKASI": "DUSUN BOGO RW 08 FDT-2", "ID_LOKASI": "NJK000095",
    "ALAMAT": "Nglawak Kecamatan Kertosono", "NAMA_OLT": "KERTOSONO",
    "ID_FDT_FROM": "NJK.100.021.DSBG08-FDT2.019.110", "ID_FAT_TO": "DSBG08FDT2.019",
    "NAMA_PT_VENDOR": "PT Buana Menara Indonesia", "REP_VENDOR": "ERFIN FIRMANSYAH",
    "JABATAN_VENDOR": "BMI FIELD SUPERVISOR", "NAMA_PT_CUSTOMER": "PT Ekamas Mora Republik Tbk",
    "REP_CUSTOMER": "M. NUGROHO", "JABATAN_CUSTOMER": "EMR FIELD SUPERVISOR",
    "TANGGAL_TEST": "2026-06-27", "NO_PO": "PO-EMR-2026-001"
}

st.set_page_config(page_title="Universal ATP Generator", page_icon="7²3", layout="wide")

# Inisialisasi State Memori Sementara Web
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.current_theme = "light"
    st.session_state.metadata = {key: "" for key in DEFAULT_METADATA.keys()}
    st.session_state.fat_commands = ["A12"]
    st.session_state.pole_commands = ["pole 73=3", "ext 74=2"]
    st.session_state.parsed_fat = []
    st.session_state.parsed_poles = []
    st.session_state.fase1_extracted = False

def load_greetings_from_file():
    db = {"MINGGU": [], "PAGI": [], "SIANG": [], "SORE": [], "MALAM": []}
    if not os.path.exists(GREETINGS_PATH):
        return db
    current_section = None
    with open(GREETINGS_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("[") and line.endswith("]"):
                current_section = line[1:-1].upper()
            elif line and current_section in db:
                db[current_section].append(line)
    return db

def generate_dc_greeting():
    now = datetime.datetime.now()
    hour = now.hour
    hari = now.strftime("%A")
    panggilan = "An_"
    db = load_greetings_from_file()
    fallback = f"Halo {panggilan}! Selamat bekerja mengawal dokumen ATP."
    
    if hari == "Sunday" and db.get("MINGGU"):
        terpilih = random.choice(db["MINGGU"])
    elif 0 <= hour < 11 and db.get("PAGI"):
        terpilih = random.choice(db["PAGI"])
    elif 11 <= hour < 15 and db.get("SIANG"):
        terpilih = random.choice(db["SIANG"])
    elif 15 <= hour < 19 and db.get("SORE"):
        terpilih = random.choice(db["SORE"])
    elif db.get("MALAM"):
        terpilih = random.choice(db["MALAM"])
    else:
        terpilih = fallback
    return terpilih.format(nama_lengkap=panggilan)

# Pengaturan Tema Spasial Web
if st.session_state.current_theme == "dark":
    st.markdown("""<style>.stApp { background-color: #0E1117; color: #FFFFFF; }</style>""", unsafe_allow_html=True)
else:
    st.markdown("""<style>.stApp { background-color: #F9FAFB; color: #1F2937; }</style>""", unsafe_allow_html=True)

with st.sidebar:
    st.title("7²3 Pusat Kendali")
    st.session_state.current_theme = "light" if st.toggle("”9Ü1 Mode Terang Web", value=(st.session_state.current_theme == "light")) else "dark"
    st.divider()
    st.markdown("”9ä8 **Panel Akun Administrator (Standby)**\n*Menu khusus pengelolaan data kredensial user baru.*")

st.info(generate_dc_greeting())

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"7²2„1‚5 BERKAS UTAMA HILANG: File `{TEMPLATE_PATH}` wajib diletakkan di root folder aplikasi!")
    st.stop()

# PENENTUAN MODE PROJECT SECARA GLOBAL KAKU
project_mode = st.radio(
    "”9ß9 Tentukan Mode Karakteristik Dokumen ATP Jaringan:",
    ["Cluster Jaringan (Distribusi Hilir)", "Subfeeder Jaringan (Backbone Hulu)"],
    horizontal=True
)
mode_slug = "cluster" if "Cluster" in project_mode else "subfeeder"

# Pembentukan Struktur Tiga Tab Kerja
tab1, tab2, tab3 = st.tabs(["”9Ý7 FASE 1: Struktur Kerangka", "”9Ý6 FASE 2: Angka Hasil Ukur", "”9ö6„1‚5 PUSAT ARSIP DIGITAL"])

with tab1:
    st.header("1. Metadata Identitas Administratif")
    with st.form("form_fase1_metadata"):
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.metadata["NAMA_PROYEK"] = st.text_input("Nama Proyek / Site:", value=st.session_state.metadata["NAMA_PROYEK"] or DEFAULT_METADATA["NAMA_PROYEK"])
            st.session_state.metadata["NO_PO"] = st.text_input("Nomor PO (Purchase Order):", value=st.session_state.metadata["NO_PO"] or DEFAULT_METADATA["NO_PO"])
            st.session_state.metadata["TANGGAL_TEST"] = st.text_input("Tanggal Pengujian:", value=st.session_state.metadata["TANGGAL_TEST"] or DEFAULT_METADATA["TANGGAL_TEST"])
            st.session_state.metadata["REGION"] = st.text_input("Wilayah / Region:", value=st.session_state.metadata["REGION"] or DEFAULT_METADATA["REGION"])
        with col2:
            st.session_state.metadata["NAMA_LOKASI"] = st.text_input("Nama Cluster Jaringan:", value=st.session_state.metadata["NAMA_LOKASI"] or DEFAULT_METADATA["NAMA_LOKASI"])
            st.session_state.metadata["ID_LOKASI"] = st.text_input("ID Cluster / FDT:", value=st.session_state.metadata["ID_LOKASI"] or DEFAULT_METADATA["ID_LOKASI"])
            st.session_state.metadata["ALAMAT"] = st.text_input("Alamat Lokasi Riil Jaringan:", value=st.session_state.metadata["ALAMAT"] or DEFAULT_METADATA["ALAMAT"])
            st.session_state.metadata["NAMA_OLT"] = st.text_input("Nama Sentral OLT Target:", value=st.session_state.metadata["NAMA_OLT"] or DEFAULT_METADATA["NAMA_OLT"])
            
        with st.expander("”9ä3 Detail Perusahaan & Penanggung Jawab Lapangan (Opsional)"):
            c1, c2 = st.columns(2)
            with c1:
                st.session_state.metadata["ID_FDT_FROM"] = st.text_input("ID Link From (FDT Hulu):", value=st.session_state.metadata["ID_FDT_FROM"] or DEFAULT_METADATA["ID_FDT_FROM"])
                st.session_state.metadata["ID_FAT_TO"] = st.text_input("ID Link To (FAT Hilir):", value=st.session_state.metadata["ID_FAT_TO"] or DEFAULT_METADATA["ID_FAT_TO"])
                st.session_state.metadata["NAMA_PT_VENDOR"] = st.text_input("Nama PT Vendor Rekanan:", value=st.session_state.metadata["NAMA_PT_VENDOR"] or DEFAULT_METADATA["NAMA_PT_VENDOR"])
                st.session_state.metadata["REP_VENDOR"] = st.text_input("Nama Pengawas Lapangan Vendor:", value=st.session_state.metadata["REP_VENDOR"] or DEFAULT_METADATA["REP_VENDOR"])
            with c2:
                st.session_state.metadata["JABATAN_VENDOR"] = st.text_input("Jabatan Pengawas Vendor:", value=st.session_state.metadata["JABATAN_VENDOR"] or DEFAULT_METADATA["JABATAN_VENDOR"])
                st.session_state.metadata["NAMA_PT_CUSTOMER"] = st.text_input("Nama PT Pemilik Jaringan:", value=st.session_state.metadata["NAMA_PT_CUSTOMER"] or DEFAULT_METADATA["NAMA_PT_CUSTOMER"])
                st.session_state.metadata["REP_CUSTOMER"] = st.text_input("Nama Pengawas Pihak Customer:", value=st.session_state.metadata["REP_CUSTOMER"] or DEFAULT_METADATA["REP_CUSTOMER"])
                st.session_state.metadata["JABATAN_CUSTOMER"] = st.text_input("Jabatan Pengawas Customer:", value=st.session_state.metadata["JABATAN_CUSTOMER"] or DEFAULT_METADATA["JABATAN_CUSTOMER"])
        st.form_submit_button("”9ä8 Kunci Parameter Administrasi")

    st.header("2. Komando Ringkas Pembentuk Topologi Jaringan")
    col_cmd1, col_cmd2 = st.columns(2)
    with col_cmd1:
        st.subheader("A. Pembangun Rute Perangkat FAT")
        for i in range(len(st.session_state.fat_commands)):
            st.session_state.fat_commands[i] = st.text_input(f"Instruksi Baris FAT-{i+1}:", value=st.session_state.fat_commands[i], key=f"fat_cmd_key_{i}", placeholder="Contoh: A12 atau B08")
        if st.button("7Ê7 Tambah Baris Instruksi FAT"):
            st.session_state.fat_commands.append("")
            st.rerun()
    with col_cmd2:
        st.subheader("B. Pembangun Distribusi Tiang")
        for i in range(len(st.session_state.pole_commands)):
            st.session_state.pole_commands[i] = st.text_input(f"Instruksi Baris Tiang-{i+1}:", value=st.session_state.pole_commands[i], key=f"pole_cmd_key_{i}", placeholder="Contoh: pole 73 = 14 atau ext 72.5 = 5")
        if st.button("7Ê7 Tambah Baris Instruksi Tiang"):
            st.session_state.pole_commands.append("")
            st.rerun()

    if st.button("”9ã4 EKSTRAK STRUKTUR UTAMA FASE 1", use_container_width=True):
        st.session_state.parsed_fat = parse_all_fat([c for c in st.session_state.fat_commands if c.strip()])
        st.session_state.parsed_poles = parse_all_poles([c for c in st.session_state.pole_commands if c.strip()])
        st.session_state.fase1_extracted = True
        st.success(f"Topologi Sah! Berhasil mengunci {len(st.session_state.parsed_fat)} FAT riil dan {len(st.session_state.parsed_poles)} Kelompok Tiang Jaringan.")

    if st.session_state.get("fase1_extracted"):
        st.divider()
        st.subheader("”9à3 Pencetakan Berkas Draf Kerangka (Fase 1)")
        st.caption("Pencetakan draf ini mengosongkan seluruh sel teknis (Fase 2 tetap dibiarkan utuh berupa token placeholder di Excel).")
        if st.button("”9Ü4 KOMPILASI BERKAS DRAF FASE 1", use_container_width=True):
            try:
                with open(TEMPLATE_PATH, "rb") as f:
                    template_bytes = f.read()
                with st.spinner("Mesin Pass-1 merakit draf struktural..."):
                    draf_ram = inject_excel_fase1(io.BytesIO(template_bytes), st.session_state.metadata, st.session_state.parsed_fat, st.session_state.parsed_poles, mode=mode_slug)
                    st.session_state.stream_f1_output = draf_ram.getvalue()
                st.session_state.f1_download_ready = True
                st.success("Berkas draf kerangka Fase 1 sukses dipacking!")
            except Exception as e:
                st.error(traceback.format_exc())
                
        if st.session_state.get("f1_download_ready"):
            st.download_button("•0‹4 UNDUH FILE DRAF FASE 1 (.XLSX)", data=st.session_state.stream_f1_output, file_name=f"DRAF_KERANGKA_F1_{mode_slug.upper()}.xlsx", use_container_width=True)

with tab2:
    st.header("Formulir Input Nilai Pengukuran Sektor Optik (Fase 2)")
    if st.session_state.get("fase1_extracted"):
        # PANGGIL KOMPONEN VISUAL TABEL DINAMIS SECARA TERISOLASI
        render_modul_2a_splitter()
        st.divider()
        render_modul_2b_opm_distribution(st.session_state.parsed_fat)
        st.divider()
        render_modul_2c_2d_otdr_summary(st.session_state.parsed_fat, mode_slug)
        
        st.divider()
        st.subheader("”9à3 Eksekusi Akhir Pipeline Dokumen Gabungan")
        st.caption("Proses ini merakit dokumen final murni di memori RAM secara estafet: Template -> Mesin Fase 1 -> Mesin Fase 2 -> File Excel Bersih.")
        
        if st.button("”9Ü4 COMPILING & GENERATE FINAL ATP EXCEL", use_container_width=True):
            try:
                with open(TEMPLATE_PATH, "rb") as f:
                    template_bytes = f.read()
                with st.spinner("Menjalankan estafet kompilasi data murni di RAM..."):
                    # PIPELINE ESTAFET (In-Memory Processing)
                    # Langkah 1: Kirim bahan baku ke Injektor Fase 1
                    stream_draf_hidup = inject_excel_fase1(io.BytesIO(template_bytes), st.session_state.metadata, st.session_state.parsed_fat, st.session_state.parsed_poles, mode=mode_slug)
                    # Langkah 2: Operasikan stream draf hidup tersebut langsung ke Injektor Fase 2
                    stream_final_matang = inject_excel_fase2(stream_draf_hidup, mode=mode_slug)
                    st.session_state.stream_final_output = stream_final_matang.getvalue()
                    
                st.session_state.final_download_ready = True
                st.success("Dokumen Serah Terima (ATP Final) Selesai Dirakit Sempurna!")
            except Exception as e:
                st.error(traceback.format_exc())
                
        if st.session_state.get("final_download_ready"):
            st.download_button("•0‹4 DOWNLOAD BERKAS FINAL ATP EXCEL (.XLSX)", data=st.session_state.stream_final_output, file_name=f"FINAL_ATP_BERSIH_{mode_slug.upper()}.xlsx", use_container_width=True)
    else:
        st.warning("7²2„1‚5 Perhatian: Ekstrak dan kunci data topologi rute di Tab Fase 1 terlebih dahulu agar form input angka ini dapat terbuka.")

with tab3:
    st.header("”9ö6„1‚5 Brankas Digital Manajemen database Proyek")
    st.caption("Pusat pengamanan terpusat untuk menyimpan progress kerja harian (Administrasi F1 + Angka Isian F2) murni ke format biner JSON lokal.")
    
    if st.button("”9Ü4 KUNCI PROGRESS KERJA KE DATABASE JSON", use_container_width=True):
        loc_raw = st.session_state.metadata.get("NAMA_LOKASI", "PROYEK_BARU").strip()
        filename_clean = loc_raw.replace(" ", "_").replace("-", "_").upper()
        
        # Konversi dataframe editor aktif ke bentuk dictionary agar bisa diserialisasi JSON
        json_contract = {
            "saved_timestamp": str(datetime.datetime.now()),
            "project_mode_slug": mode_slug,
            "metadata": st.session_state.metadata,
            "fat_commands": st.session_state.fat_commands,
            "pole_commands": st.session_state.pole_commands,
            "grid_splitter_backup": st.session_state.grid_splitter_data.to_dict(orient="records") if "grid_splitter_data" in st.session_state else None,
            "grid_otdr_cluster_backup": st.session_state.grid_otdr_cluster_data.to_dict(orient="records") if "grid_otdr_cluster_data" in st.session_state else None,
            "grid_otdr_subfeeder_backup": st.session_state.grid_otdr_subfeeder_data.to_dict(orient="records") if "grid_otdr_subfeeder_data" in st.session_state else None
        }
        
        # Simpan rekam isian data OPM distribusi per line secara utuh
        for key in list(st.session_state.keys()):
            if key.startswith("grid_opm_dist_line_"):
                json_contract[key] = st.session_state[key].to_dict(orient="records")
                
        with open(f"{DB_DIR}/{filename_clean}.json", "w", encoding="utf-8") as j_file:
            json.dump(json_contract, j_file, indent=4)
        st.success(f"7¼3 Progres Berhasil Diamankan! Berkas database `{filename_clean}.json` terkunci aman.")

    st.divider()
    available_json_files = [f for f in os.listdir(DB_DIR) if f.endswith(".json")]
    if available_json_files:
        query_search = st.text_input("”9ä3 Cari Berkas Arsip Berdasarkan Nama Lokasi:")
        for file in available_json_files:
            if query_search.lower() in file.lower():
                col_name, col_act = st.columns([3, 1])
                with col_name:
                    st.markdown(f"”9Ü7 Arsip Terdaftar: `{file}`")
                with col_act:
                    if st.button("”9à3 Muat Kembali", key=f"btn_load_{file}", use_container_width=True):
                        with open(f"{DB_DIR}/{file}", "r", encoding="utf-8") as j_file:
                            loaded = json.load(j_file)
                        
                        # Restorasi Sesi Memori Utama Aplikasi Web
                        st.session_state.metadata = loaded["metadata"]
                        st.session_state.fat_commands = loaded["fat_commands"]
                        st.session_state.pole_commands = loaded["pole_commands"]
                        st.session_state.fase1_extracted = True
                        
                        # Restorasi Isian Grid Data Angka Lapangan
                        if loaded.get("grid_splitter_backup") is not None:
                            st.session_state.grid_splitter_data = pd.DataFrame(loaded["grid_splitter_backup"])
                        if loaded.get("grid_otdr_cluster_backup") is not None:
                            st.session_state.grid_otdr_cluster_data = pd.DataFrame(loaded["grid_otdr_cluster_backup"])
                        if loaded.get("grid_otdr_subfeeder_backup") is not None:
                            st.session_state.grid_otdr_subfeeder_data = pd.DataFrame(loaded["grid_otdr_subfeeder_backup"])
                            
                        # Restorasi Isian Dinamis OPM Distribusi Jaringan
                        for k, v in loaded.items():
                            if k.startswith("grid_opm_dist_line_"):
                                st.session_state[k] = pd.DataFrame(v)
                                
                        st.success("Progress berhasil dipulihkan total ke kondisi terakhir!")
                        st.rerun()

st.markdown('<div style="text-align: center; color: #888888; padding: 20px; font-size: 13px;">Universal ATP Generator Application 6¦1 Developed by An_</div>', unsafe_allow_html=True)
