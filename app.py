import re
import os
from flask import Flask, request, send_file, render_template_string
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from io import BytesIO

app = Flask(__name__)

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

    headers = ["ID", "Email", "Imię", "Nazwisko", "Przedmiot", "Poziom edukacyjny", "Rozszerzenie", "Aktywny"]
    for col_num, header in enumerate(headers, 1):
        cell = ws_mapa.cell(row=1, column=col_num, value=header)
        cell.border = thin_border
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # BRAKUJĄCE ID
    ws_brak_id.merge_cells('A1:B1')
    cell = ws_brak_id['A1']
    cell.value = "Osoby, które nie mają ID"
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_brak_id.append(["Imię", "Nazwisko"])
    for c in ws_brak_id[2]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    # BRAKUJĄCE EMAILE
    ws_brak_mail.merge_cells('A1:B1')
    cell = ws_brak_mail['A1']
    cell.value = "Osoby, które nie mają Email"
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_brak_mail.append(["Imię", "Nazwisko"])
    for c in ws_brak_mail[2]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    brak_id_set = set()
    brak_mail_set = set()
    row_out = 2

    for row in ws_baza.iter_rows(min_row=2, values_only=True):
        _id = row[0]
        imie = str(row[1]).strip() if row[1] else ""
        nazwisko = str(row[2]).strip() if row[2] else ""
        sp_subjects = row[4]
        lo_subjects = row[5]

        mail = find_mail_safe(imie, nazwisko, mail_map)

        # SP
        if sp_subjects:
            for subj in [s.strip() for s in str(sp_subjects).split(",")]:
                if subj:
                    typ = "1-3 SP" if subj.lower() == "edukacja wczesnoszkolna" else "4-8 SP"
                    values = [_id, mail, imie, nazwisko, subj, typ, "NIE", "TAK"]

                    if not _id:
                        key = (imie, nazwisko)
                        if key not in brak_id_set:
                            ws_brak_id.append([imie, nazwisko])
                            for c in ws_brak_id[ws_brak_id.max_row]:
                                c.border = thin_border
                            brak_id_set.add(key)

                    if mail == "BRAK MAILA":
                        key = (imie, nazwisko)
                        if key not in brak_mail_set:
                            ws_brak_mail.append([imie, nazwisko])
                            for c in ws_brak_mail[ws_brak_mail.max_row]:
                                c.border = thin_border
                            brak_mail_set.add(key)

                    for col_num, val in enumerate(values, 1):
                        cell = ws_mapa.cell(row=row_out, column=col_num, value=val)
                        if col_num == 6 and typ in color_map:
                            cell.fill = PatternFill(start_color=color_map[typ],
                                                    end_color=color_map[typ], fill_type="solid")
                        if col_num == 7:
                            cell.fill = red_fill
                        if col_num == 8:
                            cell.fill = green_fill
                        cell.border = thin_border
                    row_out += 1

        # LO
        if lo_subjects:
            for subj in [s.strip() for s in str(lo_subjects).split(",")]:
                if subj:
                    typ = "LO"
                    values = [_id, mail, imie, nazwisko, subj, typ, "TAK", "TAK"]

                    if not _id:
                        key = (imie, nazwisko)
                        if key not in brak_id_set:
                            ws_brak_id.append([imie, nazwisko])
                            for c in ws_brak_id[ws_brak_id.max_row]:
                                c.border = thin_border
                            brak_id_set.add(key)

                    if mail == "BRAK MAILA":
                        key = (imie, nazwisko)
                        if key not in brak_mail_set:
                            ws_brak_mail.append([imie, nazwisko])
                            for c in ws_brak_mail[ws_brak_mail.max_row]:
                                c.border = thin_border
                            brak_mail_set.add(key)

                    for col_num, val in enumerate(values, 1):
                        cell = ws_mapa.cell(row=row_out, column=col_num, value=val)
                        if col_num == 6 and typ in color_map:
                            cell.fill = PatternFill(start_color=color_map[typ],
                                                    end_color=color_map[typ], fill_type="solid")
                        if col_num == 7:
                            cell.fill = green_fill
                        if col_num == 8:
                            cell.fill = green_fill
                        cell.border = thin_border
                    row_out += 1

    # Zapis do BytesIO
    output = BytesIO()
    new_wb.save(output)
    output.seek(0)
    return output, f"{os.path.splitext(filename)[0]}_wynik.xlsx"

# Strona HTML z uploadem
HTML = """
<!doctype html>
<title>Mapa przedmiotów</title>
<h2>Prześlij plik Excel (.xlsx)</h2>
<form method=post enctype=multipart/form-data>
  <input type=file name=file>
  <input type=submit value=Prześlij>
</form>
"""

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]
        if not file.filename.lower().endswith(".xlsx"):
            return "To nie jest plik Excel (.xlsx)"
        output, out_name = process_file(file, file.filename)
        return send_file(output, download_name=out_name, as_attachment=True)
    return render_template_string(HTML)

if __name__ == "__main__":
    app.run(debug=True)
