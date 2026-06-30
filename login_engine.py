# -*- coding: utf-8 -*-
"""
Modul: login_engine.py
Tanggung Jawab: Mengelola database kredensial user, otentikasi login sesi,
               dan penarikan data log arsip proyek untuk fungsi pengawasan Admin.
Arsitektur: Modular Security Engine & Data Registry.
Developed by: An_
"""

import os
import json
import datetime

USER_DB_PATH = "history_database/user_registry.json"
DB_DIR = "history_database"

def initialize_user_registry():
    """
    Membuat file registrasi user default jika belum tersedia di folder database.
    Menyediakan 1 akun Administrator utama dan 1 akun Document Control bawaan.
    """
    if not os.path.exists(USER_DB_PATH):
        default_users = {
            "admin": {
                "password_hash": "admin123",
                "role": "administrator",
                "nama_lengkap": "Anjas Prasetyo"
            },
            "dc_operator": {
                "password_hash": "dc2026",
                "role": "document_control",
                "nama_lengkap": "Operator DC Lapangan"
            }
        }
        with open(USER_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(default_users, f, indent=4)

def verify_login(username_input: str, password_input: str) -> dict:
    """
    Memverifikasi kecocokan username dan password terhadap user_registry.json.
    Mengembalikan data user jika sukses, mengembalikan None jika gagal.
    """
    initialize_user_registry()
    username_clean = username_input.strip().lower()
    
    try:
        with open(USER_DB_PATH, "r", encoding="utf-8") as f:
            users = json.load(f)
            
        if username_clean in users:
            user_data = users[username_clean]
            if user_data["password_hash"] == password_input.strip():
                return {
                    "username": username_clean,
                    "role": user_data["role"],
                    "nama_lengkap": user_data["nama_lengkap"]
                }
    except Exception:
        pass
    return None

def register_user(username: str, password_txt: str, name_txt: str, role_type: str) -> bool:
    """
    FUNGSI EKSKLUSIF ADMIN: Mendaftarkan akun Document Control baru ke dalam sistem.
    """
    initialize_user_registry()
    username_clean = username.strip().lower()
    if not username_clean or not password_txt.strip():
        return False
        
    try:
        with open(USER_DB_PATH, "r", encoding="utf-8") as f:
            users = json.load(f)
            
        # Cegah penumpukan jika username sudah dipakai
        if username_clean in users:
            return False
            
        users[username_clean] = {
            "password_hash": password_txt.strip(),
            "role": role_type,
            "nama_lengkap": name_txt.strip()
        }
        
        with open(USER_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=4)
        return True
    except Exception:
        return False

def get_all_registered_users() -> dict:
    """
    FUNGSI EKSKLUSIF ADMIN: Menarik daftar seluruh akun yang terdaftar untuk ditampilkan di Web UI.
    """
    initialize_user_registry()
    try:
        with open(USER_DB_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def delete_user_account(username: str) -> bool:
    """
    FUNGSI EKSKLUSIF ADMIN: Menghapus hak akses akun operator DC tertentu.
    """
    try:
        with open(USER_DB_PATH, "r", encoding="utf-8") as f:
            users = json.load(f)
        
        username_clean = username.strip().lower()
        if username_clean in users and username_clean != "admin": # Melarang penghapusan admin utama
            del users[username_clean]
            with open(USER_DB_PATH, "w", encoding="utf-8") as f:
                json.dump(users, f, indent=4)
            return True
    except Exception:
        pass
    return False

def scan_all_project_logs() -> list:
    """
    FUNGSI STRUKTUR PENGAWASAN ADMIN:
    Memindai seluruh berkas JSON proyek yang sedang dikerjakan oleh para DC.
    Mengekstrak metadata penanggung jawab, waktu simpan, dan status kelengkapan data.
    """
    project_logs = []
    if not os.path.exists(DB_DIR):
        return project_logs
        
    for file_name in os.listdir(DB_DIR):
        # Scan semua berkas arsip proyek kecuali file registrasi user
        if file_name.endswith(".json") and file_name != "user_registry.json":
            file_path = os.path.exists(os.path.join(DB_DIR, file_name))
            try:
                with open(os.path.join(DB_DIR, file_name), "r", encoding="utf-8") as f:
                    proj_data = json.load(f)
                    
                meta = proj_data.get("metadata", {})
                # Cek status kelengkapan data angka Fase 2
                has_splitter = "Yes" if proj_data.get("grid_splitter_backup") else "No"
                has_fat_list = len(proj_data.get("fat_commands", [])) > 0
                
                status_fase = "Selesai Fase 2 (Angka Matang)" if has_splitter == "Yes" else "Fase 1 (Draf Administrasi)"
                
                project_logs.append({
                    "File Name": file_name,
                    "Nama Lokasi": meta.get("NAMA_LOKASI", "Tidak Diketahui"),
                    "Mode Jaringan": str(proj_data.get("project_mode_slug", "cluster")).upper(),
                    "Operator DC": str(proj_data.get("locked_by_operator", "Tidak Tercatat")),
                    "Penyimpanan Terakhir": str(proj_data.get("saved_timestamp", "N/A"))[:19],
                    "Status Dokumen": status_fase
                })
            except Exception:
                pass
    return project_logs
