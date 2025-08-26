import streamlit as st
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from io import BytesIO

st.set_page_config(page_title="Mapa przedmiotów", layout="wide")
st.title("Mapa przedmiotów")

uploaded_file = st.file_uploader("Prześlij plik Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Wczytanie Excela z pamięcią
    uploaded_file.seek(0)
    wb = load_workbook(uploaded_file, data_only=True)

    if "Baza" not in wb.sheetnames or "STARA mapa" not in wb.sheetnames:
        st.error("Plik musi zawierać arkusze 'Baza' i 'STARA mapa'")
    else:
        ws_baza = wb["Baza"]
        ws_stara = wb["STARA mapa"]

        # Funkcja czyszczenia nazwiska
        def clean_nazwisko(nazwisko):
            if not nazwisko:
                return ""
            return re.sub(r"\(.*?\)", "", str(nazwisko)).strip().lower()

        # Tworzymy mapę maili
        mail_map = {}
        for row in ws_stara.iter_rows(min_row=2, values_only=True):
            imie = str(row[2]).strip() if row[2] else ""
            nazwisko = str(row[3]).strip() if row[3] else ""
            mail = str(row[1]).strip() if row[1] else ""
            if imie and nazwisko and mail:
                key = f"{imie.lower()} {clean_nazwisko(nazwisko.lower())}"
                mail_map[key] = mail

        def find_mail_safe(imie, nazwisko):
            key_full = f"{imie.lower()} {clean_nazwisko(nazwisko.lower())}"
            if key_full in mail_map:
                return mail_map[key_full]
            # alternatywnie szukaj po czystym nazwisku
            candidates = []
            for k, mail in mail_map.items():
                k_imie, k_nazwisko = k.split(" ", 1)
                k_nazwisko_clean = clean_nazwisko(k_nazwisko)
                if imie.lower() == k_imie and nazwisko.lower() == k_nazwisko_clean:
                    candidates.append(mail)
            return candidates[0] if len(candidates)==1 else "BRAK MAILA"

        # Style Excel
        color_map = {"LO": "ADD8E6", "1-3 SP": "FFFF99", "4-8 SP": "CCFFCC"}
        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        red_fill = PatternFill(start_color="FF7F7F", end_color="FF7F7F", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

        # Nowy skoroszyt
        new_wb = Workbook()
        ws_mapa = new_wb.active
        ws_mapa.title = "MAPA PRZEDMIOTÓW"
        ws_brak_id = new_wb.create_sheet("BRAKUJĄCE ID")
        ws_brak_mail = new_wb.create_sheet("BRAKUJĄCE EMAILE")

        # Nagłówki
        headers = ["ID", "Email", "Imię", "Nazwisko", "Przedmiot", "Poziom edukacyjny", "Rozszerzenie", "Aktywny"]
        for col_num, header in enumerate(headers, 1):
            cell = ws_mapa.cell(row=1, column=col_num, value=header)
            cell.font = Font(bold=True)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Arkusze BRAK ID i BRAK MAIL
        ws_brak_id.append(["Imię", "Nazwisko"])
        ws_brak_mail.append(["Imię", "Nazwisko"])

        row_out = 2
        brak_id_set = set()
        brak_mail_set = set()

        def add_row(_id, imie, nazwisko, mail, subj, typ, rozszerzenie, aktywny):
            nonlocal row_out
            values = [_id, mail, imie, nazwisko, subj, typ, rozszerzenie, aktywny]
            for col_num, val in enumerate(values, 1):
                cell = ws_mapa.cell(row=row_out, column=col_num, value=val)
                if col_num == 6 and typ in color_map:
                    cell.fill = PatternFill(start_color=color_map[typ], end_color=color_map[typ], fill_type="solid")
                if col_num == 7:
                    cell.fill = red_fill if rozszerzenie=="NIE" else green_fill
                if col_num == 8:
                    cell.fill = green_fill
                cell.border = thin_border
            row_out +=1

            if not _id and (imie,nazwisko) not in brak_id_set:
                ws_brak_id.append([imie,nazwisko])
                brak_id_set.add((imie,nazwisko))
            if mail=="BRAK MAILA" and (imie,nazwisko) not in brak_mail_set:
                ws_brak_mail.append([imie,nazwisko])
                brak_mail_set.add((imie,nazwisko))

        # Przetwarzanie danych z Baza
        for row in ws_baza.iter_rows(min_row=2, values_only=True):
            _id = row[0]
            imie = str(row[1]).strip() if row[1] else ""
            nazwisko = str(row[2]).strip() if row[2] else ""
            sp_subjects = row[4] if row[4] else ""
            lo_subjects = row[5] if row[5] else ""

            mail = find_mail_safe(imie,nazwisko)

            if sp_subjects:
                for subj in str(sp_subjects).split(","):
                    subj=subj.strip()
                    if subj:
                        typ="1-3 SP" if subj.lower()=="edukacja wczesnoszkolna" else "4-8 SP"
                        add_row(_id, imie, nazwisko, mail, subj, typ, "NIE", "TAK")
            if lo_subjects:
                for subj in str(lo_subjects).split(","):
                    subj=subj.strip()
                    if subj:
                        add_row(_id, imie, nazwisko, mail, subj, "LO", "TAK", "TAK")

        # Zapis do BytesIO
        output = BytesIO()
        new_wb.save(output)
        output.seek(0)

        st.download_button(
            label="Pobierz przetworzony plik",
            data=output,
            file_name=f"{uploaded_file.name.split('.')[0]}_wynik.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
