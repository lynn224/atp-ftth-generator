# -*- coding: utf-8 -*-
"""
Modul: parser_engine.py
Tanggung Jawab: Menerjemahkan komando teks ringkas (FAT & Tiang) menjadi 
               struktur data terformat (List & Dictionary) di memori RAM.
Arsitektur: Modular, independen, dan fail-safe.
Developed by: An_
"""

import re

def generate_fat_sequence(command: str) -> list:
    """
    Mengubah perintah ringkas FAT tunggal menjadi rangkaian urutan array sekuensial.
    Contoh: 'A12' -> ['A01', 'A02', 'A03', ..., 'A12']
    Jika input tidak sesuai format ringkas, akan dianggap sebagai nama FAT literal.
    """
    cmd_clean = command.strip()
    if not cmd_clean:
        return []
        
    # Regex untuk mendeteksi pola Huruf Jalur diikuti oleh Angka Jumlah (Cth: A12, B08, JAF15)
    match = re.match(r"^([A-Za-z]+)(\d+)$", cmd_clean)
    if match:
        line_letter = match.group(1).upper()
        try:
            count = int(match.group(2))
            # Menghasilkan list berurutan dengan padding 2 digit angka (zfill)
            return [f"{line_letter}{str(i).zfill(2)}" for i in range(1, count + 1)]
        except ValueError:
            return [cmd_clean.upper()]
    else:
        # Jika bukan format ringkas, return sebagai nama literal kapital
        return [cmd_clean.upper()]

def parse_all_fat(fat_commands_list: list) -> list:
    """
    Memproses seluruh daftar baris perintah FAT yang dikirim dari antarmuka Web UI.
    Menggabungkan hasil ekstraksi menjadi satu array tunggal yang flat.
    """
    final_fat_list = []
    for cmd in fat_commands_list:
        if cmd and str(cmd).strip():
            final_fat_list.extend(generate_fat_sequence(str(cmd)))
    return final_fat_list

def convert_pole_size_to_ejaan(size_str: str) -> str:
    """
    Mengonversi angka dimensi mentah tiang menjadi kalimat ejaan resmi METER dan INCH.
    Aturan Konversi:
    - 73 atau 7.3 -> 7 METER 3 INCH
    - 74 -> 7 METER 4 INCH
    - 72.5 atau 7.25 -> 7 METER 2.5 INCH
    """
    # Bersihkan spasi dan karakter non-numerik kecuali titik desimal
    clean_num = size_str.strip().replace(" ", "")
    match = re.search(r"(\d+(\.\d+)?)", clean_num)
    if not match:
        return size_str.upper()
        
    num_val = match.group(1)
    
    if "." in num_val:
        parts = num_val.split(".")
        # Penanganan Kasus Format Desimal Titik di Tengah (Cth: 7.25)
        if len(parts[0]) == 1:
            meter = parts[0]
            inch = parts[1]
            # Normalisasi otomatis jika tertulis 25 menjadi 2.5 inch sesuai standar lapangan
            if inch == "25":
                inch = "2.5"
        # Penanganan Kasus Format Desimal Beruntun (Cth: 72.5)
        else:
            meter = parts[0][0]
            inch = parts[0][1:] + "." + parts[1]
        return f"{meter} METER {inch} INCH"
    else:
        # Penanganan Kasus Format Angka Bulat Gabungan (Cth: 73, 74)
        if len(num_val) >= 2:
            meter = num_val[0]
            inch = num_val[1:]
            return f"{meter} METER {inch} INCH"
        else:
            return f"{num_val} METER"

def parse_single_pole_command(command: str) -> dict:
    """
    Memecah satu baris perintah tiang menjadi dictionary data terstruktur kaku.
    Mendukung penentuan jenis tipe otomatis (pole -> NEW POLE, ext -> EXT POLE).
    Contoh: 'pole 73 = 14' -> {'title': 'Pole Erection 73', 'type': 'NEW POLE', 'size_clean': '7 METER 3 INCH', 'qty': 14}
    """
    cmd_clean = command.strip()
    if not cmd_clean or "=" not in cmd_clean:
        return None
        
    # Belah instruksi antara spesifikasi (kiri) dan kuantitas jumlah (kanan)
    left_side, right_side = cmd_clean.split("=", 1)
    left_side = left_side.strip()
    
    try:
        qty = int(right_side.strip())
    except ValueError:
        return None  # Abaikan baris jika kuantitas bukan angka bulat valid
        
    # Deteksi Tipe Tiang secara otomatis berdasarkan awalan kata kunci
    if left_side.lower().startswith("ext "):
        pole_type = "EXT POLE"
        raw_size = left_side[4:].strip()
        # Normalisasi penomoran sub-ukuran untuk penamaan sheet (Cth: 7.25 atau 72.5 menjadi 7.25)
        sheet_size = raw_size if "." in raw_size else f"{raw_size[0]}.{raw_size[1:]}" if len(raw_size) > 1 else raw_size
        title = f"Pole Erection EXT {sheet_size}"
    else:
        pole_type = "NEW POLE"
        raw_size = left_side[5:].strip() if left_side.lower().startswith("pole ") else left_side
        title = f"Pole Erection {raw_size}"
        
    # Terjemahkan dimensi menjadi kalimat ejaan resmi
    size_clean = convert_pole_size_to_ejaan(raw_size)
    
    return {
        "title": title,
        "type": pole_type,
        "size_clean": size_clean,
        "qty": qty
    }

def parse_all_poles(pole_commands_list: list) -> list:
    """
    Mengeksekusi seluruh daftar baris perintah tiang dari Web UI.
    Mengembalikan array objek dictionary terstruktur yang siap dikonsumsi oleh Injektor Excel.
    """
    final_pole_list = []
    for cmd in pole_commands_list:
        if cmd and str(cmd).strip():
            parsed_obj = parse_single_pole_command(str(cmd))
            if parsed_obj:
                final_pole_list.append(parsed_obj)
    return final_pole_list
