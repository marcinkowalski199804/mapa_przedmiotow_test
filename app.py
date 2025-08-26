import streamlit as st
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from io import BytesIO

st.title("Mapa przedmiotów")

uploaded_file = st.file_uploader("Prześlij plik Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Funkcja przetwarzająca plik (Twoja logika z Tkintera)
    def process_file(file_stream, filename):
        wb = load_workbook(file_stream)
        ws_baza = wb["Baza"]
        ws_stara = wb["STARA mapa"]

        def clean_nazwisko(nazwisko):
            if not nazwisko:
                return ""
            return re.sub(r"\(.*?\)", "", str(nazwisko)).strip().lower()

        mail_map = {}
        for row in ws_stara.iter_rows(min_row=2, values_only=True):
            imie = str(row[2]).strip() if row[2] else ""
            nazwisko = str(row[3]).strip() if row[3] else ""
            mail = str(row[1]).strip() if row[1] else ""
            if imie and nazwisko and mail:
                key = f"{imie.lower()} {nazwisko.lower()}"
                mail_map[key] = mail

        def find_mail_safe(imie, nazwisko, mail_map):
            key_full = f"{imie.lower()} {nazwisko.lower()}"
            if key_full in mail_map:
                return mail_map[key_full]
            candidates = []
            for k, mail in mail_map.items():
                k_imie, k_nazwisko = k.split(" ", 1)
                k_nazwisko_clean = clean_nazwisko(k_nazwisko)
                if imie.lower() == k_imie and nazwisko.lower() == k_nazwisko_clean:
                    candidates.append(mail)
            if len(candidates) == 1:
                return candidates[0]
            else:
                return "BRAK MAILA"

        color_map = {"LO": "ADD8E6", "1-3 SP": "FFFF99", "4-8 SP": "CCFFCC"}
        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        red_fill = PatternFill(start_color="FF7F7F", end_color="FF7F7F", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

        new_wb = Workbook()
        ws_mapa = new_wb.active
        ws_mapa.title = "MAPA PRZEDMIOTÓW"
        ws_brak_id = new_wb.create_sheet("BRAKUJĄCE ID")
        ws_brak_mail = new_wb.create_sheet("BRAKUJĄCE EMAILE")

        # Tu możesz wkleić resztę Twojej logiki kopiując z Tkintera (nagłówki, kolory, pętle itp.)

        # Na końcu zapis do BytesIO
        output = BytesIO()
        new_wb.save(output)
        output.seek(0)
        return output, f"{filename.split('.')[0]}_wynik.xlsx"

    output_file, out_name = process_file(uploaded_file, uploaded_file.name)
    st.download_button(
        label="Pobierz przetworzony plik",
        data=output_file,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
