# -*- coding: utf-8 -*-
"""
Modul: ui_technical_forms.py
Tanggung Jawab: Menyediakan fungsi visual untuk merender form grid data teknis (Fase 2)
               secara dinamis dan otomatis mengunci list FAT hasil Fase 1.
Arsitektur: Modular UI Component, terisolasi dari core engine injektor.
Developed by: An_
"""

import streamlit as st
import pandas as pd

def render_modul_2a_splitter():
    """
    Grid Form 2A: Mengakomodasi OPM Feeder, OPM Subfeeder, dan Splitter FDT.
    Desain antarmuka dibuat berurutan menurun agar operator mudah menginput data.
    Mesin injektor yang akan memetakan secara kaku: Slot 1 (Kiri Atas), Slot 2 (Kiri Bawah),
    Slot 3 (Kanan Atas), Slot 4 (Kanan Bawah). Jika lebih, meluap ke sheet cadangan_2.
    """
    st.markdown("### 🎛️ Modul 2A: OPM Feeder, Subfeeder & Splitter FDT")
    st.caption("Masukkan data Serial Number, nilai OPM Before, dan 8 Nilai Port OPM After secara berurutan.")

    # Kolom kaku 8 port sesuai spesifikasi visual template
    kolom_splitter = [
        "Splitter ID", "Splitter SN", "OPM Before (dBm)",
        "Port 1", "Port 2", "Port 3", "Port 4",
        "Port 5", "Port 6", "Port 7", "Port 8"
    ]

    # Inisialisasi state jika belum ada data agar input tidak hilang saat pindah tab
    if "grid_splitter_data" not in st.session_state:
        # Menyediakan 4 baris default (mewakili 1 sheet penuh template kaku: 2 Kiri, 2 Kanan)
        initial_data = [
            {"Splitter ID": f"Splitter {i+1}", "Splitter SN": "NO SN", "OPM Before (dBm)": "-12.00",
             "Port 1": "-18.10", "Port 2": "-18.20", "Port 3": "", "Port 4": "",
             "Port 5": "", "Port 6": "", "Port 7": "", "Port 8": ""}
            for i in range(4)
        ]
        st.session_state.grid_splitter_data = pd.DataFrame(initial_data, columns=kolom_splitter)

    # Render data editor dengan fitur baris dinamis (bisa tambah baris jika splitter > 4)
    st.session_state.grid_splitter_data = st.data_editor(
        st.session_state.grid_splitter_data,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_modul_2a",
        column_config={
            "Splitter ID": st.column_config.TextColumn("Splitter ID", help="ID urutan fisik splitter", required=True),
            "Splitter SN": st.column_config.TextColumn("Serial Number (SN)", help="Masukkan SN perangkat"),
            "OPM Before (dBm)": st.column_config.TextColumn("OPM Before", help="Nilai input redaman masuk FDT")
        }
    )

def render_modul_2b_opm_distribution(parsed_fat_list: list):
    """
    Grid Form 2B: OPM Distribusi (Arah rumah pelanggan).
    Mengunci kolom pertama dengan nama FAT hasil Fase 1 (Tanpa Batas Baris).
    Setiap FAT wajib menyediakan 8 kolom port dropcore yang datanya berbeda-beda.
    """
    st.markdown("### ⚡ Modul 2B: OPM Distribusi (Arah FAT Jaringan)")
    st.caption("Baris tabel otomatis terkunci berdasarkan rute FAT hasil ekstraksi Fase 1. Isi nilai redaman pada Port 1 s.d Port 8.")

    if not parsed_fat_list:
        st.warning("⚠️ Perhatian: Silakan lakukan validasi dan ekstraksi rute FAT di Tab Fase 1 terlebih dahulu agar tabel ini dapat terbuat.")
        return

    # Siapkan format penamaan baris FAT kaku (Cth: FAT A01, FAT A02)
    fats_formatted = [f"FAT {fat}" for fat in parsed_fat_list]
    
    # Check sinkronisasi jumlah baris state dengan hasil ekstraksi terbaru
    key_state = "grid_opm_dist_data"
    if key_state not in st.session_state or len(st.session_state[key_state]) != len(fats_formatted):
        # Bangun struktur data baru jika ada perubahan rute di Fase 1
        st.session_state[key_state] = pd.DataFrame({
            "FAT Name (Locked)": fats_formatted,
            "Port 1": ["-19.10"] * len(fats_formatted),
            "Port 2": ["-19.15"] * len(fats_formatted),
            "Port 3": [""] * len(fats_formatted),
            "Port 4": [""] * len(fats_formatted),
            "Port 5": [""] * len(fats_formatted),
            "Port 6": [""] * len(fats_formatted),
            "Port 7": [""] * len(fats_formatted),
            "Port 8": [""] * len(fats_formatted)
        })

    # Render data editor dengan jumlah baris yang dikunci (num_rows="fixed") agar nama FAT tidak bisa diubah/diacak-acak user
    st.session_state[key_state] = st.data_editor(
        st.session_state[key_state],
        num_rows="fixed",
        use_container_width=True,
        key="editor_modul_2b",
        disabled=["FAT Name (Locked)"] # Mengunci kolom nama FAT agar kebal typo
    )

def render_modul_2c_2d_otdr_summary(parsed_fat_list: list, mode_slug: str):
    """
    Grid Form 2C & 2D: Penanganan Hasil Ukur OTDR.
    Memisahkan logika input dan tampilan visual secara tegas antara Cluster dan Subfeeder.
    """
    st.markdown("### 📊 Modul 2C/2D: Hasil Pengukuran Alat Ukur OTDR")

    if mode_slug == "cluster":
        st.caption("Mode Cluster: Masukkan jarak (Km) beserta nilai Loss gelombang 1310nm dan 1550nm per FAT (Disuntik Bersandingan).")
        if not parsed_fat_list:
            st.warning("⚠️ Perhatian: Silakan lakukan validasi dan ekstraksi rute FAT di Tab Fase 1 terlebih dahulu.")
            return

        fats_formatted = [f"FAT {fat}" for fat in parsed_fat_list]
        key_state = "grid_otdr_cluster_data"
        
        if key_state not in st.session_state or len(st.session_state[key_state]) != len(fats_formatted):
            st.session_state[key_state] = pd.DataFrame({
                "FAT Name (Locked)": fats_formatted,
                "Distance (Km)": ["1.50"] * len(fats_formatted),
                "Loss 1310 nm (dB)": ["0.12"] * len(fats_formatted),
                "Loss 1550 nm (dB)": ["0.11"] * len(fats_formatted)
            })

        st.session_state[key_state] = st.data_editor(
            st.session_state[key_state],
            num_rows="fixed",
            use_container_width=True,
            key="editor_modul_2c",
            disabled=["FAT Name (Locked)"]
        )

    else:
        st.caption("Mode Subfeeder: Masukkan parameter jarak dan loss backbone utama. Data otomatis dibelah ke sheet 1310 & 1550 secara terpisah.")
        
        # Penanganan Sektor Subfeeder: Baris diisi per Nomor Core utama secara sekuensial menurun
        if "grid_otdr_subfeeder_data" not in st.session_state:
            initial_cores = [
                {"Core No": f"Core {i+1}", "Distance (Km)": "0.52", "Loss 1310 nm (dB)": "0.04", "Loss 1550 nm (dB)": "0.05"}
                for i in range(4)
            ]
            st.session_state.grid_otdr_subfeeder_data = pd.DataFrame(
                initial_cores, 
                columns=["Core No", "Distance (Km)", "Loss 1310 nm (dB)", "Loss 1550 nm (dB)"]
            )

        st.session_state.grid_otdr_subfeeder_data = st.data_editor(
            st.session_state.grid_otdr_subfeeder_data,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_modul_2d"
        )
