import re
import os
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from tkinterdnd2 import TkinterDnD, DND_FILES

def process_file(filepath):
    if not filepath.lower().endswith(".xlsx"):
        messagebox.showerror("Błąd", "To nie jest plik Excel (.xlsx)")
        return

    folder, filename = os.path.split(filepath)
    name, ext = os.path.splitext(filename)
    output_path = os.path.join(folder, f"{name}_wynik{ext}")

    wb = load_workbook(filepath)
    ws_baza = wb["Baza"]
    ws_stara = wb["STARA mapa"]

    def clean_nazwisko(nazwisko):
        if not nazwisko:
            return ""
        return re.sub(r"\(.*?\)", "", str(nazwisko)).strip().lower()

    # Mapa maili
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

    # Kolory i obramowania
    color_map = {"LO": "ADD8E6", "1-3 SP": "FFFF99", "4-8 SP": "CCFFCC"}
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    red_fill = PatternFill(start_color="FF7F7F", end_color="FF7F7F", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # Nowy skoroszyt i arkusze
    new_wb = Workbook()
    ws_mapa = new_wb.active
    ws_mapa.title = "MAPA PRZEDMIOTÓW"
    ws_brak_id = new_wb.create_sheet("BRAKUJĄCE ID")
    ws_brak_mail = new_wb.create_sheet("BRAKUJĄCE EMAILE")

    # Nagłówki MAPA PRZEDMIOTÓW
    headers = ["ID", "Email", "Imię", "Nazwisko", "Przedmiot", "Poziom edukacyjny", "Rozszerzenie", "Aktywny"]
    for col_num, header in enumerate(headers, 1):
        cell = ws_mapa.cell(row=1, column=col_num, value=header)
        cell.border = thin_border
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Arkusz BRAKUJĄCE ID
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

    # Arkusz BRAKUJĄCE EMAILE
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

    # Zbiory dla unikalnych osób
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

    new_wb.save(output_path)
    messagebox.showinfo("Gotowe!", f"Plik zapisany jako:\n{output_path}")


# Okno GUI z drag & drop
root = TkinterDnD.Tk()
root.title("Mapa przedmotów")
root.geometry("500x200")

label = tk.Label(root, text="Przeciągnij plik Excel tutaj", font=("Arial", 14))
label.pack(expand=True, padx=20, pady=20)

def drop(event):
    filepath = event.data.strip("{}")  # usuń ewentualne nawiasy
    process_file(filepath)

label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', drop)

# Przyciski alternatywne
button = tk.Button(root, text="Wybierz plik ręcznie", command=lambda: process_file(askopenfilename(filetypes=[("Excel files", "*.xlsx")])))
button.pack(pady=10)

root.mainloop()
