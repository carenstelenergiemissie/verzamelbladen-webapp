# # Verzamelblad PNH — automatische verwerking Kenter / Liander / Vattenfall
# import pandas as pd
# import datetime as dt
# import os
# from openpyxl import load_workbook
# from openpyxl.styles import PatternFill
# import win32com.client as win32
# import random
# import shutil 


# # ---------------------------------------------------
# # Instellingen
# # ---------------------------------------------------
# bronbestand = r"C:\Users\CarenStel\Downloads\VerzamelLijstFacturen-2025-MM-Boekingen-07122025_1025.xlsx"
# # Basisoutputmap
# BASE_OUTPUT_DIR = r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - Provincie Noord Holland\4. Uitvoering\DB eFactuur"

# # Datum voor mapnaam
# map_datum = dt.date.today().strftime("%d-%m-%y")

# # Eindmap, bv. "Betaalbestand 07-12-25"
# OUTPUT_FOLDER = os.path.join(BASE_OUTPUT_DIR, f"Betaalbestand {map_datum}")

# # Map aanmaken indien niet bestaat
# os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# suppliers = {
#     "kenter": {
#         "tabnaam": "NL59INGB0006779355_D",
#         "template": r"C:\Users\CarenStel\OneDrive - Energiemissie\Documenten\Verzamelbladen templates PNH\Verzamel Kenter PNH  Kenter_3011025.xlsx",
#     },
#     "liander": {
#         "tabnaam": "NL95INGB0000005585_D",
#         "template": r"C:\Users\CarenStel\OneDrive - Energiemissie\Documenten\Verzamelbladen templates PNH\Verzamel Liander PNH  3008162625_131125.xlsx",
#     },
#     "vattenfall": {
#         "tabnaam": "NL42INGB0000827935_D",
#         "template": r"C:\Users\CarenStel\OneDrive - Energiemissie\Documenten\Verzamelbladen templates PNH\Verzamel Vattenfall PNH  Vattenfall_13112025.xlsx",
#     },
# }


# # ===================================================
# #              FUNCTIE — VERWERKING PER LEVERANCIER
# # ===================================================
# def verwerk_supplier(keuze: str, config: dict) -> None:

#     tabnaam = config["tabnaam"]
#     template_file = config["template"]

#     # ---------------------------------------------------
#     # Datumformaten
#     # ---------------------------------------------------
#     datum_excel = dt.date.today().strftime("%d%m%y")
#     datum_tekst = dt.date.today().strftime("%d-%m-%Y")

#     # ---------------------------------------------------
#     # NIEUWE BESTANDSNAAM op basis van template
#     # ---------------------------------------------------
#     base_name = os.path.basename(template_file)
#     name_no_ext, ext = os.path.splitext(base_name)

#     prefix, old_date = name_no_ext.rsplit("_", 1)
#     new_name = f"{prefix}_{datum_excel}{ext}"

#     doelbestand = os.path.join(OUTPUT_FOLDER, new_name)

#     print(f"\n➡ Verwerking gestart voor: {keuze}")

#     # ---------------------------------------------------
#     # BRONBESTAND → controleren of tabblad bestaat
#     # ---------------------------------------------------
#     available_sheets = pd.ExcelFile(bronbestand).sheet_names

#     if tabnaam not in available_sheets:
#         print(f"⚠ SKIPPED: Tabblad '{tabnaam}' bestaat niet in bronbestand. {keuze} wordt overgeslagen.\n")
#         return  # skip deze leverancier en ga door met de volgende

#     # ---------------------------------------------------
#     # BRONBESTAND inlezen
#     # ---------------------------------------------------
#     df_raw = pd.read_excel(bronbestand, sheet_name=tabnaam, header=None).dropna(how="all")
#     df_headers = df_raw.iloc[0]
#     df_data = df_raw[1:].copy()
#     df_data.columns = df_headers

#     # Kolom C = factuurnummers
#     factnr_bron = df_data.iloc[:, 2].astype(str).tolist()

#     # Kies één willekeurig factuurnummer voor de vraag aan de gebruiker
#     import random
#     random_factuur = random.choice(factnr_bron)


#     # ---------------------------------------------------
#     # TEMPLATE inlezen
#     # ---------------------------------------------------
#     wb = load_workbook(template_file)
#     ws_spec = wb["Specificatie"]

#     # ---------------------------------------------------
#     # SPECIFICATIE leegmaken & nieuwe data schrijven
#     # ---------------------------------------------------
#     last_used_row = 1
#     while ws_spec.cell(row=last_used_row + 1, column=1).value not in (None, ""):
#         last_used_row += 1

#     if last_used_row > 1:
#         ws_spec.delete_rows(2, last_used_row - 1)

#     for r_idx, row_values in enumerate(df_data.values, start=2):
#         for c_idx, value in enumerate(row_values, start=1):
#             ws_spec.cell(row=r_idx, column=c_idx, value=value)

#     # ---------------------------------------------------
#     # CREDIT-regels verwijderen (kolom N)
#     # ---------------------------------------------------
#     credit_rows = []
#     max_row = ws_spec.max_row
#     for r in range(2, max_row + 1):
#         val = ws_spec.cell(row=r, column=14).value
#         if isinstance(val, str) and val.strip().lower() == "credit":
#             credit_rows.append(r)

#     for r in reversed(credit_rows):
#         ws_spec.delete_rows(r)

#     # ---------------------------------------------------
#     # Verzamelblad datum aanpassen
#     # ---------------------------------------------------
#     # Verzamelblad datum aanpassen
#     ws_verz = wb["Verzamelblad"]
#     ws_verz["B4"].value = datum_tekst
#     ws_verz["C24"].value = f"{keuze.capitalize()}_{datum_excel}"

#     ws_verz["B4"].fill = PatternFill(fill_type=None)
#     ws_verz["C24"].fill = PatternFill(fill_type=None)

# # Handmatige invoer voor periode met random factuurnummer
# # ---------------------------------------------------
#     print(f"\nVoor {keuze.capitalize()} kun je deze factuur opzoeken ter controle:")
#     print(f"➡ Factuurnummer: {random_factuur}")

#     invoer_periode = input(
#         f"Voer de periode in voor {keuze.capitalize()} (bijv. 01-12-2025 t/m 31-12-2025): "
#     )

#     ws_verz["C26"].value = invoer_periode



#     # ---------------------------------------------------
#     # NIEUWE BESTAND OPSLAAN
#     # ---------------------------------------------------
#     wb.save(doelbestand)
#     print(f"✔ Excel opgeslagen → {doelbestand}")

#     # ---------------------------------------------------
#     # Nieuwe Specificatie opnieuw inlezen voor controles
#     # ---------------------------------------------------
#     wb_new = load_workbook(doelbestand)
#     ws_spec_new = wb_new["Specificatie"]

#     rows_new = []
#     row = 2

#     while ws_spec_new.cell(row=row, column=1).value not in (None, ""):
#         rows_new.append([
#             ws_spec_new.cell(row=row, column=c).value
#             for c in range(1, df_data.shape[1] + 1)
#         ])
#         row += 1

#     df_spec_new = pd.DataFrame(rows_new, columns=df_data.columns)

#     # ---------------------------------------------------
#     # FACTUURNUMMERS controleren
#     # ---------------------------------------------------
#     print("\n=== FACTUURNUMMERS ===")

#     factnr_spec_new = df_spec_new.iloc[:, 2].astype(str)
#     ontbrekend = set(factnr_bron) - set(factnr_spec_new)
#     extra = set(factnr_spec_new) - set(factnr_bron)

#     if not ontbrekend and not extra:
#         print("✔ Factuurnummers kloppen.")
#     else:
#         if ontbrekend:
#             print("⚠ Ontbrekend in nieuwe specificatie:")
#             for f in sorted(ontbrekend):
#                 print(" -", f)
#         if extra:
#             print("⚠ Staat wel in nieuwe maar niet in bron:")
#             for f in sorted(extra):
#                 print(" -", f)

#     # ---------------------------------------------------
#     # TOTALEN controleren
#     # ---------------------------------------------------
#     print("\n=== CONTROLE EXCL / BTW / INCL ===")

#     kol_excl = df_data.columns.get_loc("Excl. BTW")
#     kol_btw = df_data.columns.get_loc("BTW")
#     kol_incl = df_data.columns.get_loc("Incl. BTW")

#     sum_excl_bron = df_data.iloc[:, kol_excl].astype(float).sum()
#     sum_btw_bron  = df_data.iloc[:, kol_btw].astype(float).sum()
#     sum_incl_bron = df_data.iloc[:, kol_incl].astype(float).sum()

#     sum_excl_new = df_spec_new.iloc[:, kol_excl].astype(float).sum()
#     sum_btw_new  = df_spec_new.iloc[:, kol_btw].astype(float).sum()
#     sum_incl_new = df_spec_new.iloc[:, kol_incl].astype(float).sum()

#     print(f"Excl: bron {sum_excl_bron:.2f} | nieuw {sum_excl_new:.2f}")
#     print(f"BTW : bron {sum_btw_bron:.2f}  | nieuw {sum_btw_new:.2f}")
#     print(f"Incl: bron {sum_incl_bron:.2f} | nieuw {sum_incl_new:.2f}")

#     if (
#         sum_excl_bron == sum_excl_new and
#         sum_btw_bron == sum_btw_new and
#         sum_incl_bron == sum_incl_new
#     ):
#         print("✔ Totalen kloppen.")
#     else:
#         print("⚠ Totalen komen NIET overeen.")

#     # ---------------------------------------------------
#     # PDF EXPORT (Excel blijft onzichtbaar)
#     # ---------------------------------------------------
#     try:
#         excel = win32.Dispatch("Excel.Application")
#         excel.Visible = False
#         excel.DisplayAlerts = False

#         wb_pdf = excel.Workbooks.Open(doelbestand, ReadOnly=True)
#         sheet = wb_pdf.Worksheets("Verzamelblad")

#         pdf_pad = doelbestand.replace(".xlsx", ".pdf")
#         sheet.ExportAsFixedFormat(0, pdf_pad)

#         wb_pdf.Close(False)
#         excel.Quit()

#         print(f"✔ PDF opgeslagen → {pdf_pad}")

#     except Exception as e:
#         print(f"⚠ Fout bij PDF-export: {e}")


# # ===================================================
# #                MAIN LOOP
# # ===================================================
# if __name__ == "__main__":
#     for keuze, config in suppliers.items():
#                 # ---------------------------------------------------
#         # Kopie van bronbestand opslaan in OUTPUT_FOLDER
#         # ---------------------------------------------------
#         print("\nKopieer bronbestand naar Betaalbestand-map...")

#         bron_bestandsnaam = os.path.basename(bronbestand)
#         bron_kopie_pad = os.path.join(OUTPUT_FOLDER, bron_bestandsnaam)

#         shutil.copy2(bronbestand, bron_kopie_pad)

#         print(f"✔ Bronbestand gekopieerd naar: {bron_kopie_pad}")
#         print("\n======================================")
#         print(f"Start verwerking voor: {keuze}")
#         print("======================================")
#         verwerk_supplier(keuze, config)

# # Verzamelblad PNH & GGZ — automatische verwerking met GUI
# import os
# import shutil
# import random
# import datetime as dt
# import tkinter as tk
# from tkinter import filedialog, ttk, messagebox
# import re

# import pandas as pd
# from openpyxl import load_workbook
# from openpyxl.styles import PatternFill
# import win32com.client as win32


# # =====================================================================
# #                         LEVERANCIERCONFIG
# # =====================================================================
# suppliers = {
#     "kenter": {
#         "tabnaam": "NL59INGB0006779355_D",
#         "template": r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - Provincie Noord Holland\4. Uitvoering\Verzamelbladen templates PNH\Verzamel Kenter PNH  Kenter_3011025.xlsx",
#     },
#     "liander": {
#         "tabnaam": "NL95INGB0000005585_D",
#         "template": r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - Provincie Noord Holland\4. Uitvoering\Verzamelbladen templates PNH\Verzamel Liander PNH  3008162625_131125.xlsx",
#     },
#     "vattenfall": {
#         "tabnaam": "NL42INGB0000827935_D",
#         "template": r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - Provincie Noord Holland\4. Uitvoering\Verzamelbladen templates PNH\Verzamel Vattenfall PNH  Vattenfall_13112025.xlsx",
#     },
#     # GGZ Centraal leveranciers
#     "eneco": {
#         "tabnaam": "NL13ABNA0640000797_D",
#         "template": r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - GGZ Centraal\4. Uitvoering\Betaalbestand\Template\Verzamel Eneco GGZ 04122025_0923.xlsx",
#     },
#     "vitens": {
#         "tabnaam": "NL94INGB0000869000_D",
#         "template": r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - GGZ Centraal\4. Uitvoering\Betaalbestand\Template\Verzamel Vitens GGZ 26112025_1120.xlsx",
#     },
# }

# supplier_settings = {}
# supplier_frames = {}


# # =====================================================================
# #                        GUI FUNCTIE – POPUP
# # =====================================================================
# def open_settings_popup(suppliers_config):
#     root = tk.Tk()
#     root.title("Instellingen Verzamelbestand")
#     root.geometry("780x780")

#     settings = {
#         "bronbestand": None,
#         "klant": None,
#         "per_supplier": {}
#     }

#     # ----------------------------- Bronbestand -----------------------------
#     def kies_bronbestand():
#         pad = filedialog.askopenfilename(
#             title="Kies bronbestand (DB Energie export)",
#             filetypes=[("Excel bestanden", "*.xlsx *.xls")]
#         )
#         if not pad:
#             return

#         bron_entry.delete(0, tk.END)
#         bron_entry.insert(0, pad)

#         # Probeer tabbladen uit te lezen en leveranciers automatisch aan te vinken
#         try:
#             xls = pd.ExcelFile(pad)
#             sheets = set(xls.sheet_names)
#         except Exception as e:
#             messagebox.showerror("Fout", f"Kon tabbladen niet lezen:\n{e}")
#             return

#         # Auto-select leveranciers op basis van tabbladen
#         for naam, cfg in suppliers_config.items():
#             tabnaam = cfg["tabnaam"]
#             frame_info = supplier_frames[naam]
#             if tabnaam in sheets:
#                 frame_info["selected_var"].set(True)
#                 frame_info["checkbutton"].config(state="normal")
#             else:
#                 frame_info["selected_var"].set(False)
#                 frame_info["checkbutton"].config(state="disabled")

#     tk.Label(root, text="Bronbestand (download via DB Energie):").pack(anchor="w", padx=10, pady=5)
#     bron_frame = tk.Frame(root)
#     bron_frame.pack(anchor="w", padx=10)

#     bron_entry = tk.Entry(bron_frame, width=80)
#     bron_entry.pack(side="left")
#     tk.Button(bron_frame, text="Bladeren", command=kies_bronbestand).pack(side="left", padx=5)

#     # ----------------------------- Klant kiezen ----------------------------
#     tk.Label(root, text="Selecteer klant (exact 1):").pack(anchor="w", padx=10, pady=10)

#     klant_var = tk.StringVar()
#     klantenlijst = [
#         "Provincie Noord-Holland",
#         "GGZ Centraal",
#     ]
#     klant_dropdown = ttk.Combobox(root, textvariable=klant_var, values=klantenlijst, width=50, state="readonly")
#     klant_dropdown.pack(anchor="w", padx=10)

#     # ---------------------- Instellingen per leverancier ---------------------
#     tk.Label(root, text="Per leverancier instellen:").pack(anchor="w", padx=10, pady=10)

#     global supplier_frames
#     supplier_frames = {}

#     for naam in suppliers_config.keys():
#         frame = tk.LabelFrame(root, text=naam.capitalize(), padx=10, pady=10)
#         frame.pack(fill="x", padx=10, pady=5)

#         selected_var = tk.BooleanVar(value=False)
#         chk_sel = tk.Checkbutton(frame, text="Deze leverancier verwerken", variable=selected_var)
#         chk_sel.pack(anchor="w")

#         credit_var = tk.BooleanVar(value=True)
#         tk.Checkbutton(frame, text="Creditfacturen verwerken", variable=credit_var).pack(anchor="w")

#         tk.Label(frame, text="Periode (bijv. 01-12-2025 t/m 31-12-2025):").pack(anchor="w", pady=(5, 0))
#         periode_var = tk.StringVar()
#         tk.Entry(frame, textvariable=periode_var, width=40).pack(anchor="w")

#         supplier_frames[naam] = {
#             "selected_var": selected_var,
#             "checkbutton": chk_sel,
#             "credit_var": credit_var,
#             "periode_var": periode_var,
#         }

#     # ----------------------------- VERWERK-knop -----------------------------
#     def bevestigen():
#         if not bron_entry.get():
#             messagebox.showerror("Fout", "Geen bronbestand gekozen.")
#             return

#         if not klant_var.get():
#             messagebox.showerror("Fout", "Geen klant gekozen.")
#             return

#         if not any(info["selected_var"].get() for info in supplier_frames.values()):
#             messagebox.showerror("Fout", "Geen leverancier geselecteerd.")
#             return

#         settings["bronbestand"] = bron_entry.get()
#         settings["klant"] = klant_var.get()

#         per_supplier = {}
#         for naam, info in supplier_frames.items():
#             per_supplier[naam] = {
#                 "selected": info["selected_var"].get(),
#                 "credit": info["credit_var"].get(),
#                 "periode": info["periode_var"].get().strip(),
#             }

#         settings["per_supplier"] = per_supplier
#         root.destroy()

#     tk.Button(
#         root,
#         text="VERWERK",
#         command=bevestigen,
#         bg="#008F39",
#         fg="white",
#         font=("Arial", 20, "bold"),
#         width=22,
#         height=2
#     ).pack(pady=30)

#     root.mainloop()
#     return settings


# # =====================================================================
# #                   FUNCTIE – VERWERKING PER LEVERANCIER
# # =====================================================================
# def verwerk_supplier(keuze: str, config: dict, bronbestand: str, output_folder: str):
#     global supplier_settings

#     tabnaam = config["tabnaam"]
#     template_file = config["template"]

#     # datum_excel in formaat ddmmjjjj, bv 09122025
#     datum_excel = dt.date.today().strftime("%d%m%Y")
#     datum_tekst = dt.date.today().strftime("%d-%m-%Y")

#     # Nieuwe bestandsnaam
#     base_name = os.path.basename(template_file)
#     name_no_ext, ext = os.path.splitext(base_name)

#     if keuze in ("kenter", "liander", "vattenfall"):
#         # alles na laatste '_' vervangen door ddmmjjjj
#         try:
#             prefix, _ = name_no_ext.rsplit("_", 1)
#             new_name_no_ext = f"{prefix}_{datum_excel}"
#         except ValueError:
#             new_name_no_ext = f"{name_no_ext}_{datum_excel}"
#     else:
#         # GGZ: eerste 8-cijferige datum vervangen
#         new_name_no_ext = re.sub(r"\d{8}", datum_excel, name_no_ext, count=1)
#         if new_name_no_ext == name_no_ext:
#             # fallback
#             new_name_no_ext = f"{name_no_ext}_{datum_excel}"

#     new_name = f"{new_name_no_ext}{ext}"
#     doelbestand = os.path.join(output_folder, new_name)

#     print(f"\n➡ Verwerking gestart voor: {keuze}")
#     print(f"  Template: {template_file}")
#     print(f"  Nieuw bestand: {doelbestand}")

#     # Controleer tabblad
#     available_sheets = pd.ExcelFile(bronbestand).sheet_names
#     if tabnaam not in available_sheets:
#         print(f"⚠ SKIPPED: Tabblad '{tabnaam}' ontbreekt.")
#         return

#     # Inlezen bronbestand
#     df_raw = pd.read_excel(bronbestand, sheet_name=tabnaam, header=None).dropna(how="all")
#     df_headers = df_raw.iloc[0]
#     df_data = df_raw[1:].copy()
#     df_data.columns = df_headers

#     # kolom C = factuurnummers
#     factnr_bron = df_data.iloc[:, 2].astype(str).tolist()
#     random_factuur = random.choice(factnr_bron)

#     wb = load_workbook(template_file)
#     ws_spec = wb["Specificatie"]

#     # SPECIFICATIE leegmaken
#     last_used_row = 1
#     while ws_spec.cell(row=last_used_row + 1, column=1).value not in (None, ""):
#         last_used_row += 1
#     if last_used_row > 1:
#         ws_spec.delete_rows(2, last_used_row - 1)

#     # nieuwe data schrijven
#     for r_idx, row_values in enumerate(df_data.values, start=2):
#         for c_idx, value in enumerate(row_values, start=1):
#             ws_spec.cell(row=r_idx, column=c_idx, value=value)

#     # CREDIT filter
#     if not supplier_settings[keuze]["credit"]:
#         credit_rows = []
#         for r in range(2, ws_spec.max_row + 1):
#             val = ws_spec.cell(row=r, column=14).value
#             if isinstance(val, str) and val.strip().lower() == "credit":
#                 credit_rows.append(r)
#         for r in reversed(credit_rows):
#             ws_spec.delete_rows(r)

#     # Verzamelblad invullen
#     ws_verz = wb["Verzamelblad"]
#     ws_verz["B4"].value = datum_tekst

#     # C24: PNH-leveranciers hard, GGZ via regex
#     if keuze in ("kenter", "liander", "vattenfall"):
#         ws_verz["C24"].value = f"{keuze.capitalize()}_{datum_excel}"
#     else:
#         old_c24 = ws_verz["C24"].value
#         if isinstance(old_c24, str):
#             new_c24 = re.sub(r"\d{8}", datum_excel, old_c24, count=1)
#             if new_c24 != old_c24:
#                 ws_verz["C24"].value = new_c24
#             else:
#                 ws_verz["C24"].value = f"{keuze.upper()}_{datum_excel}"
#         else:
#             ws_verz["C24"].value = f"{keuze.upper()}_{datum_excel}"

#     # Periode (uit popup)
#     invoer_periode = supplier_settings[keuze]["periode"]
#     ws_verz["C26"].value = invoer_periode

#     print(f"• Periode: {invoer_periode}")
#     print(f"• Controle factuur: {random_factuur}")

#     # Opslaan Excel
#     wb.save(doelbestand)
#     print(f"✔ Excel opgeslagen → {doelbestand}")

#     # PDF EXPORT (zelfde naam, .pdf) – altijd eerste werkblad
#     try:
#         excel = win32.Dispatch("Excel.Application")
#         excel.Visible = False
#         excel.DisplayAlerts = False

#         wb_pdf = excel.Workbooks.Open(doelbestand, ReadOnly=True)
#         sheet = wb_pdf.Worksheets(1)

#         pdf_pad = doelbestand.replace(".xlsx", ".pdf")
#         sheet.ExportAsFixedFormat(0, pdf_pad)

#         wb_pdf.Close(False)
#         excel.Quit()

#         print(f"✔ PDF opgeslagen → {pdf_pad}")

#     except Exception as e:
#         print(f"⚠ Fout bij PDF-export: {e}")


# # =====================================================================
# #                             MAIN SCRIPT
# # =====================================================================
# if __name__ == "__main__":
#     # Popup starten
#     user_settings = open_settings_popup(suppliers)

#     bronbestand = user_settings["bronbestand"]
#     gekozen_klant = user_settings["klant"]
#     supplier_settings = user_settings["per_supplier"]

#     # Koppeling klant → outputmap
#     klant_output_paden = {
#         "Provincie Noord-Holland":
#             r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - Provincie Noord Holland\4. Uitvoering\DB eFactuur",

#         "GGZ Centraal":
#             r"C:\Users\CarenStel\Energiemissie\Customer Service - Documenten\1 - Directe klanten-Trenton\Trenton - GGZ Centraal\4. Uitvoering\Betaalbestand\2025",

#     }

#     if gekozen_klant in klant_output_paden:
#         BASE_OUTPUT_DIR = klant_output_paden[gekozen_klant]
#     else:
#         BASE_OUTPUT_DIR = r"C:\Users\CarenStel\Energiemissie"

#     # Map met datum aanmaken
#     map_datum = dt.date.today().strftime("%d-%m-%Y")
#     OUTPUT_FOLDER = os.path.join(BASE_OUTPUT_DIR, f"Betaalbestand {map_datum}")
#     os.makedirs(OUTPUT_FOLDER, exist_ok=True)

#     # Kopie bronbestand
#     bron_bestandsnaam = os.path.basename(bronbestand)
#     bron_kopie_pad = os.path.join(OUTPUT_FOLDER, bron_bestandsnaam)
#     shutil.copy2(bronbestand, bron_kopie_pad)
#     print(f"✔ Bronbestand gekopieerd naar: {bron_kopie_pad}")

#     # Verwerking leveranciers
#     for keuze, config in suppliers.items():
#         if not supplier_settings.get(keuze, {}).get("selected", False):
#             print(f"⏭ Leverancier '{keuze}' overgeslagen.")
#             continue

#         print("\n====================================================")
#         print(f"Start verwerking voor: {keuze}")
#         print("====================================================")

#         verwerk_supplier(keuze, config, bronbestand, OUTPUT_FOLDER)

#     print("\n✔ Verwerking compleet.")

# import os
# import sys
# import shutil
# import random
# import datetime as dt
# import tkinter as tk
# from tkinter import filedialog, ttk, messagebox
# import re
# import subprocess
# import pandas as pd
# from openpyxl import load_workbook


# # WIN32COM alleen op Windows
# try:
#     import win32com.client as win32
# except:
#     win32 = None


# # =====================================================================
# # LEVERANCIERSCONFIG - OneDrive paden (platformonafhankelijk)
# # =====================================================================
# import platform

# def get_onedrive_path():
#     """Detecteer OneDrive pad op basis van platform"""
#     system = platform.system()
    
#     if system == "Windows":
#         # Windows OneDrive pad
#         onedrive = os.path.expandvars("%USERPROFILE%\OneDrive")
#         if os.path.exists(onedrive):
#             return onedrive
#         # Fallback naar standaard locatie
#         return os.path.join(os.path.expanduser("~"), "OneDrive")
#     elif system == "Darwin":  # macOS
#         # macOS OneDrive pad
#         onedrive = os.path.expanduser("~/Library/CloudStorage/OneDrive")
#         if os.path.exists(onedrive):
#             return onedrive
#         # Alternatieve macOS locatie
#         onedrive_alt = os.path.expanduser("~/OneDrive")
#         if os.path.exists(onedrive_alt):
#             return onedrive_alt
#         return os.path.expanduser("~/OneDrive")  # fallback
#     else:
#         # Linux/overige - gebruik home directory
#         return os.path.expanduser("~/OneDrive")

# # Basis OneDrive pad
# ONEDRIVE_BASE = get_onedrive_path()

# # Leverancier templates - relatief aan OneDrive
# suppliers = {
#     "kenter": {
#         "tabnaam": "NL59INGB0006779355_D",
#         "template": os.path.join(ONEDRIVE_BASE, "Energiemissie", "Documenten", "Verzamelbladen templates PNH", "Verzamel Kenter PNH  Kenter_3011025.xlsx"),
#     },
#     "liander": {
#         "tabnaam": "NL95INGB0000005585_D",
#         "template": os.path.join(ONEDRIVE_BASE, "Energiemissie", "Documenten", "Verzamelbladen templates PNH", "Verzamel Liander PNH  3008162625_131125.xlsx"),
#     },
#     "vattenfall": {
#         "tabnaam": "NL42INGB0000827935_D",
#         "template": os.path.join(ONEDRIVE_BASE, "Energiemissie", "Documenten", "Verzamelbladen templates PNH", "Verzamel Vattenfall PNH  Vattenfall_13112025.xlsx"),
#     },
#     "eneco": {
#         "tabnaam": "NL13ABNA0640000797_D",
#         "template": os.path.join(ONEDRIVE_BASE, "Energiemissie", "Customer Service - Documenten", "1 - Directe klanten-Trenton", "Trenton - GGZ Centraal", "4. Uitvoering", "Betaalbestand", "Template", "Verzamel Eneco GGZ 04122025_0923.xlsx"),
#     },
#     "vitens": {
#         "tabnaam": "NL94INGB0000869000_D",
#         "template": os.path.join(ONEDRIVE_BASE, "Energiemissie", "Customer Service - Documenten", "1 - Directe klanten-Trenton", "Trenton - GGZ Centraal", "4. Uitvoering", "Betaalbestand", "Template", "Verzamel Vitens GGZ 26112025_1120.xlsx"),
#     },
# }

# supplier_settings = {}
# supplier_frames = {}


# # =====================================================================
# # POPUP VOOR INSTELLINGEN
# # =====================================================================
# def open_settings_popup(suppliers_config):

#     root = tk.Tk()
#     root.title("Instellingen Verzamelbestand")
#     root.geometry("780x780")

#     settings = {
#         "bronbestand": None,
#         "klant": None,
#         "per_supplier": {}
#     }

#     # ----------------------------- Bronbestand -----------------------------
#     def kies_bronbestand():
#         pad = filedialog.askopenfilename(
#             title="Kies bronbestand (DB Energie export)",
#             filetypes=[("Excel bestanden", "*.xlsx *.xls")])

#         if not pad:
#             return

#         bron_entry.delete(0, tk.END)
#         bron_entry.insert(0, pad)

#         try:
#             xls = pd.ExcelFile(pad)
#             sheets = set(xls.sheet_names)
#         except:
#             messagebox.showerror("Fout", "Tabbladen konden niet gelezen worden.")
#             return

#         # Auto-select leveranciers waar tabblad bestaat
#         for naam, cfg in suppliers_config.items():
#             tabnaam = cfg["tabnaam"]
#             f = supplier_frames[naam]
#             if tabnaam in sheets:
#                 f["selected_var"].set(True)
#                 f["checkbutton"].config(state="normal")
#             else:
#                 f["selected_var"].set(False)
#                 f["checkbutton"].config(state="disabled")

#     tk.Label(root, text="Bronbestand (DB Energie):").pack(anchor="w", padx=10)
#     bron_frame = tk.Frame(root)
#     bron_frame.pack(anchor="w", padx=10)
#     bron_entry = tk.Entry(bron_frame, width=80)
#     bron_entry.pack(side="left")
#     tk.Button(bron_frame, text="Bladeren", command=kies_bronbestand).pack(side="left", padx=5)

#     # ----------------------------- Klant kiezen ----------------------------
#     tk.Label(root, text="Selecteer klant:").pack(anchor="w", padx=10, pady=10)
#     klant_var = tk.StringVar()
#     klantenlijst = [
#         "Provincie Noord-Holland",
#         "GGZ Centraal",
#         "Gemeente Amsterdam",
#         "Trenton",
#     ]
#     ttk.Combobox(root, textvariable=klant_var, values=klantenlijst,
#                  width=50, state="readonly").pack(anchor="w", padx=10)

#     # ----------------------------- Leveranciers -----------------------------
#     tk.Label(root, text="Selecteer leverancier(s):").pack(anchor="w", padx=10, pady=10)
#     global supplier_frames
#     supplier_frames = {}

#     for naam in suppliers_config.keys():
#         frame = tk.LabelFrame(root, text=naam.capitalize(), padx=10, pady=10)
#         frame.pack(fill="x", padx=10, pady=5)

#         sel = tk.BooleanVar(value=False)
#         chk = tk.Checkbutton(frame, text="Verwerken", variable=sel)
#         chk.pack(anchor="w")

#         credit = tk.BooleanVar(value=True)
#         tk.Checkbutton(frame, text="Creditfacturen verwerken", variable=credit).pack(anchor="w")

#         tk.Label(frame, text="Periode:").pack(anchor="w")
#         periode_var = tk.StringVar()
#         tk.Entry(frame, textvariable=periode_var, width=30).pack(anchor="w")

#         supplier_frames[naam] = {
#             "selected_var": sel,
#             "checkbutton": chk,
#             "credit_var": credit,
#             "periode_var": periode_var,
#         }

#     # ----------------------------- VERWERK knop -----------------------------
#     def bevestigen():
#         if not bron_entry.get():
#             messagebox.showerror("Fout", "Geen bronbestand gekozen.")
#             return

#         if not klant_var.get():
#             messagebox.showerror("Fout", "Geen klant geselecteerd.")
#             return

#         if not any(f["selected_var"].get() for f in supplier_frames.values()):
#             messagebox.showerror("Fout", "Geen leverancier geselecteerd.")
#             return

#         settings["bronbestand"] = bron_entry.get()
#         settings["klant"] = klant_var.get()

#         per_supplier = {}
#         for naam, f in supplier_frames.items():
#             per_supplier[naam] = {
#                 "selected": f["selected_var"].get(),
#                 "credit": f["credit_var"].get(),
#                 "periode": f["periode_var"].get().strip()
#             }
#         settings["per_supplier"] = per_supplier

#         root.destroy()

#     tk.Button(root,
#               text="VERWERK",
#               command=bevestigen,
#               bg="#008F39",
#               fg="white",
#               font=("Arial", 20, "bold"),
#               width=22,
#               height=2).pack(pady=30)

#     root.mainloop()
#     return settings


# # =====================================================================
# # PDF EXPORT – Windows/macOS
# # =====================================================================
# def export_pdf(doelbestand: str):
#     pdf_pad = doelbestand.replace(".xlsx", ".pdf")

#     # ---------- Windows ----------
#     if sys.platform.startswith("win") and win32 is not None:
#         try:
#             excel = win32.Dispatch("Excel.Application")
#             excel.Visible = False
#             excel.DisplayAlerts = False

#             wb_pdf = excel.Workbooks.Open(doelbestand, ReadOnly=True)
#             sheet = wb_pdf.Worksheets(1)
#             sheet.ExportAsFixedFormat(0, pdf_pad)

#             wb_pdf.Close(False)
#             excel.Quit()
#             print(f"✔ PDF opgeslagen (Windows) → {pdf_pad}")

#         except Exception as e:
#             print(f"⚠ PDF fout (Windows): {e}")

#     # ---------- macOS ----------
#     elif sys.platform == "darwin":
#         xlsx_posix = os.path.abspath(doelbestand)
#         pdf_posix = os.path.abspath(pdf_pad)

#         applescript = f'''
#         tell application "Microsoft Excel"
#             activate
#             set wb to open POSIX file "{xlsx_posix}"
#             tell wb
#                 save workbook as wb filename (POSIX file "{pdf_posix}") file format PDF file format
#                 close saving no
#             end tell
#         end tell
#         '''

#         try:
#             subprocess.run(["osascript", "-e", applescript], check=True)
#             print(f"✔ PDF opgeslagen (macOS) → {pdf_pad}")
#         except Exception as e:
#             print(f"⚠ PDF fout (macOS): {e}")

#     return pdf_pad


# # =====================================================================
# # VERWERKING PER LEVERANCIER
# # =====================================================================
# def verwerk_supplier(keuze: str, cfg: dict, bronbestand: str, output_folder: str):

#     tabnaam = cfg["tabnaam"]
#     template_file = cfg["template"]

#     datum_excel = dt.date.today().strftime("%d%m%Y")
#     datum_tekst = dt.date.today().strftime("%d-%m-%Y")

#     base_name = os.path.basename(template_file)
#     name_no_ext, ext = os.path.splitext(base_name)

#     # ------------------------- nieuwe bestandsnaam -------------------------
#     if keuze in ("kenter", "liander", "vattenfall"):
#         try:
#             prefix, _ = name_no_ext.rsplit("_", 1)
#             new_name_no_ext = f"{prefix}_{datum_excel}"
#         except:
#             new_name_no_ext = f"{name_no_ext}_{datum_excel}"
#     else:
#         new_name_no_ext = re.sub(r"\d{8}", datum_excel, name_no_ext, 1)
#         if new_name_no_ext == name_no_ext:
#             new_name_no_ext = f"{name_no_ext}_{datum_excel}"

#     new_name = f"{new_name_no_ext}{ext}"
#     doelbestand = os.path.join(output_folder, new_name)

#     print(f"\n➡ Verwerking gestart voor: {keuze}")

#     # ------------------------- bronbestand check -------------------------
#     sheets = pd.ExcelFile(bronbestand).sheet_names
#     if tabnaam not in sheets:
#         print(f"⚠ Tabblad {tabnaam} ontbreekt → leverancier overgeslagen.")
#         return

#     # ------------------------- data inlezen -------------------------
#     df_raw = pd.read_excel(bronbestand, sheet_name=tabnaam, header=None).dropna(how="all")
#     df_headers = df_raw.iloc[0]
#     df_data = df_raw[1:].copy()
#     df_data.columns = df_headers

#     factnr_bron = df_data.iloc[:, 2].astype(str).tolist()

#     wb = load_workbook(template_file)
#     ws_spec = wb["Specificatie"]

#     # ------------------------- Specificatie leegmaken -------------------------
#     last = 1
#     while ws_spec.cell(last + 1, 1).value not in (None, ""):
#         last += 1
#     if last > 1:
#         ws_spec.delete_rows(2, last - 1)

#     # ------------------------- nieuwe data -------------------------
#     for r, row in enumerate(df_data.values, start=2):
#         for c, v in enumerate(row, start=1):
#             ws_spec.cell(r, c, v)

#     # ------------------------- creditfilter -------------------------
#     if not supplier_settings[keuze]["credit"]:
#         deletelist = []
#         for r in range(2, ws_spec.max_row + 1):
#             val = ws_spec.cell(r, 14).value
#             if isinstance(val, str) and val.lower() == "credit":
#                 deletelist.append(r)
#         for r in reversed(deletelist):
#             ws_spec.delete_rows(r)

#     # ------------------------- Verzamelblad vullen -------------------------
#     ws_verz = wb["Verzamelblad"]
#     ws_verz["B4"].value = datum_tekst

#     if keuze in ("kenter", "liander", "vattenfall"):
#         ws_verz["C24"].value = f"{keuze.capitalize()}_{datum_excel}"
#     else:
#         old = ws_verz["C24"].value
#         if isinstance(old, str):
#             new = re.sub(r"\d{8}", datum_excel, old, 1)
#             ws_verz["C24"].value = new
#         else:
#             ws_verz["C24"].value = f"{keuze.upper()}_{datum_excel}"

#     ws_verz["C26"].value = supplier_settings[keuze]["periode"]

#     wb.save(doelbestand)
#     print(f"✔ Excel opgeslagen → {doelbestand}")

#     # =====================================================================
#     # SOMCONTROLE
#     # =====================================================================
#     wb_new = load_workbook(doelbestand)
#     ws_new = wb_new["Specificatie"]

#     rows_new = []
#     r = 2
#     while ws_new.cell(r, 1).value not in (None, ""):
#         row = [ws_new.cell(r, c).value for c in range(1, df_data.shape[1] + 1)]
#         rows_new.append(row)
#         r += 1

#     df_spec_new = pd.DataFrame(rows_new, columns=df_data.columns)

#     # totalen
#     kol_excl = df_data.columns.get_loc("Excl. BTW")
#     kol_btw  = df_data.columns.get_loc("BTW")
#     kol_incl = df_data.columns.get_loc("Incl. BTW")

#     sum_excl_bron = df_data.iloc[:, kol_excl].astype(float).sum()
#     sum_btw_bron  = df_data.iloc[:, kol_btw].astype(float).sum()
#     sum_incl_bron = df_data.iloc[:, kol_incl].astype(float).sum()

#     sum_excl_new = df_spec_new.iloc[:, kol_excl].astype(float).sum()
#     sum_btw_new  = df_spec_new.iloc[:, kol_btw].astype(float).sum()
#     sum_incl_new = df_spec_new.iloc[:, kol_incl].astype(float).sum()

#     bedragen_kloppen = (
#         abs(sum_excl_bron - sum_excl_new) < 0.001 and
#         abs(sum_btw_bron  - sum_btw_new ) < 0.001 and
#         abs(sum_incl_bron - sum_incl_new) < 0.001
#     )

#     if not bedragen_kloppen:
#         print("⚠ Totalen verschillen → leverancier wordt overgeslagen.")

#         # POPUP
#         try:
#             tk.Tk().withdraw()
#             messagebox.showerror(
#                 "Fout in bedragen",
#                 f"De totalen voor '{keuze}' komen NIET overeen.\n\n"
#                 f"Excl: bron {sum_excl_bron:.2f} | nieuw {sum_excl_new:.2f}\n"
#                 f"BTW : bron {sum_btw_bron:.2f}  | nieuw {sum_btw_new:.2f}\n"
#                 f"Incl: bron {sum_incl_bron:.2f} | nieuw {sum_incl_new:.2f}\n\n"
#                 "PDF wordt NIET gemaakt.\n\n"
#                 "De verwerking gaat WEL door met de volgende leverancier."
#             )
#         except:
#             pass

#         return  # GA DOOR MET VOLGENDE LEVERANCIER

#     print("✔ Bedragen kloppen.")

#     # ------------------------- PDF EXPORT -------------------------
#     export_pdf(doelbestand)


# # =====================================================================
# # MAIN
# # =====================================================================
# if __name__ == "__main__":
#     settings = open_settings_popup(suppliers)

#     bronbestand = settings["bronbestand"]
#     gekozen_klant = settings["klant"]
#     supplier_settings = settings["per_supplier"]

#     # Basis outputpaden - platformonafhankelijk
#     klant_paden = {
#         "Provincie Noord-Holland": os.path.join(
#             ONEDRIVE_BASE, 
#             "Energiemissie", 
#             "Customer Service - Documenten", 
#             "1 - Directe klanten-Trenton", 
#             "Trenton - Provincie Noord Holland", 
#             "4. Uitvoering", 
#             "DB eFactuur"
#         ),

#         "GGZ Centraal": os.path.join(
#             ONEDRIVE_BASE,
#             "Energiemissie",
#             "Customer Service - Documenten",
#             "1 - Directe klanten-Trenton",
#             "Trenton - GGZ Centraal",
#             "4. Uitvoering",
#             "Betaalbestand",
#             "2025"
#         ),
        
#         "Gemeente Amsterdam": os.path.join(
#             ONEDRIVE_BASE,
#             "Energiemissie",
#             "Customer Service - Documenten",
#             "1 - Directe klanten-Trenton",
#             "Trenton - Gemeente Amsterdam",
#             "4. Uitvoering",
#             "Betaalbestand"
#         ),
        
#         "Trenton": os.path.join(
#             ONEDRIVE_BASE,
#             "Energiemissie",
#             "Customer Service - Documenten",
#             "1 - Directe klanten-Trenton",
#             "Trenton - Algemeen",
#             "4. Uitvoering",
#             "Betaalbestand"
#         ),
#     }

#     base = klant_paden.get(gekozen_klant, os.path.join(ONEDRIVE_BASE, "Energiemissie"))

#     datum_map = dt.date.today().strftime("%d-%m-%Y")
#     OUTPUT_FOLDER = os.path.join(base, f"Betaalbestand {datum_map}")
#     os.makedirs(OUTPUT_FOLDER, exist_ok=True)

#     # kopie bronbestand
#     shutil.copy2(bronbestand, os.path.join(OUTPUT_FOLDER, os.path.basename(bronbestand)))

#     # leveranciers verwerken
#     for naam, cfg in suppliers.items():
#         if supplier_settings[naam]["selected"]:
#             verwerk_supplier(naam, cfg, bronbestand, OUTPUT_FOLDER)

#     print("\n✔ Verwerking compleet.")


import streamlit as st
import pandas as pd
import os
import tempfile
import datetime as dt
from openpyxl import load_workbook
import re

# PDF
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table
from reportlab.lib import colors


# ===============================================
# PDF: exporteer eerste tabblad
# ===============================================
def excel_first_sheet_to_pdf(excel_path: str, pdf_path: str):
    df = pd.read_excel(excel_path, sheet_name=0)
    pdf = SimpleDocTemplate(pdf_path, pagesize=landscape(A4), leftMargin=20, rightMargin=20)
    data = [df.columns.tolist()] + df.values.tolist()
    table = Table(data)
    pdf.build([table])


# ===============================================
# Leveranciers
# ===============================================
suppliers = {
    "kenter": {"tabnaam": "NL59INGB0006779355_D"},
    "liander": {"tabnaam": "NL95INGB0000005585_D"},
    "vattenfall": {"tabnaam": "NL42INGB0000827935_D"},
    "eneco": {"tabnaam": "NL13ABNA0640000797_D"},
    "vitens": {"tabnaam": "NL94INGB0000869000_D"},
}


# ===============================================
# Pagina & thema
# ===============================================
st.set_page_config(page_title="Energiemissie Verzamelbladen", layout="wide", page_icon="⚡")

st.markdown("""
<style>

body {
    background-color: #f5f7fa;
}

h1, h2, h3 {
    color: #003d66 !important;
    font-weight: 800 !important;
}

.section-card {
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #d9e3ec;
    padding: 25px;
    margin-bottom: 28px;
    box-shadow: 0 3px 10px rgba(0,0,0,0.06);
}

.energy-accent {
    color: #009879 !important;
    font-weight: 700;
}

</style>
""", unsafe_allow_html=True)


# ===============================================
# Header
# ===============================================
st.markdown("<h1>⚡ Verzamelbestand Generator – Energiebeheer</h1>", unsafe_allow_html=True)
st.write("Professionele webapp voor het verwerken van energiefacturatie.")


# ===============================================
# 1. Bronbestand
# ===============================================
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.subheader("1️⃣ Upload bronbestand (DB Energie)")

uploaded_source = st.file_uploader("Upload .xlsx bronbestand", type=["xlsx"])
st.markdown("</div>", unsafe_allow_html=True)

if not uploaded_source:
    st.stop()

df_sheets = pd.ExcelFile(uploaded_source).sheet_names


# ===============================================
# 2. Klantselectie
# ===============================================
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.subheader("2️⃣ Kies klant")

klant = st.selectbox("Klant", [
    "Provincie Noord-Holland",
    "GGZ Centraal",
    "Gemeente Amsterdam",
    "Trenton"
])
st.markdown("</div>", unsafe_allow_html=True)


# ===============================================
# 3. Leveranciers (automatische herkenning)
# ===============================================
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.subheader("3️⃣ Automatisch herkende leveranciers")

supplier_inputs = {}
col1, col2 = st.columns(2)

for i, (naam, cfg) in enumerate(suppliers.items()):
    tabnaam = cfg["tabnaam"]
    auto_selected = tabnaam in df_sheets

    container = col1 if i % 2 == 0 else col2

    with container:
        st.markdown(f"### 🔌 {naam.capitalize()}")

        geselecteerd = st.checkbox(
            f"Verwerken ({naam}) – { '✔️ gevonden' if auto_selected else '❌ niet gevonden' }",
            value=auto_selected
        )

        credit = st.checkbox(f"Creditfacturen verwerken ({naam})", value=False)

        periode = st.text_input(f"Periode ({naam})", placeholder="Bijv. november 2025")

        template = st.file_uploader(f"Template voor {naam}", type=["xlsx"])

    supplier_inputs[naam] = {
        "selected": geselecteerd,
        "credit": credit,
        "periode": periode,
        "template": template
    }

st.markdown("</div>", unsafe_allow_html=True)


# ===============================================
# 4. Verwerking starten
# ===============================================
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.subheader("4️⃣ Verwerking starten")

start = st.button("🚀 Start verwerking")
st.markdown("</div>", unsafe_allow_html=True)

if not start:
    st.stop()


# ===============================================
# VERWERKING
# ===============================================
st.header("📥 Downloads & Controle")


for naam, cfg in suppliers.items():
    settings = supplier_inputs[naam]

    if not settings["selected"]:
        continue

    tabnaam = cfg["tabnaam"]

    if tabnaam not in df_sheets:
        st.warning(f"⛔ {naam.capitalize()}: Tabblad '{tabnaam}' ontbreekt.")
        continue

    if not settings["template"]:
        st.error(f"Template ontbreekt voor {naam}")
        continue

    st.markdown(f"## 🔧 {naam.capitalize()} – verwerking")

    with tempfile.TemporaryDirectory() as tmp:
        src_path = os.path.join(tmp, "bron.xlsx")
        tpl_path = os.path.join(tmp, f"template_{naam}.xlsx")

        with open(src_path, "wb") as f:
            f.write(uploaded_source.getbuffer())

        with open(tpl_path, "wb") as f:
            f.write(settings["template"].getbuffer())

        df_raw = pd.read_excel(src_path, sheet_name=tabnaam, header=None).dropna(how="all")
        df_headers = df_raw.iloc[0]
        df_data = df_raw[1:].copy()
        df_data.columns = df_headers

        wb = load_workbook(tpl_path)
        ws_spec = wb["Specificatie"]
        ws_verz = wb["Verzamelblad"]

        # Spec leegmaken
        last = 1
        while ws_spec.cell(last + 1, 1).value not in ("", None):
            last += 1
        if last > 1:
            ws_spec.delete_rows(2, last-1)

        # Nieuwe data
        for r, row in enumerate(df_data.values, start=2):
            for c, v in enumerate(row, start=1):
                ws_spec.cell(r, c, v)

        # Creditfilter
        if not settings["credit"]:
            remove = []
            for r in range(2, ws_spec.max_row+1):
                if str(ws_spec.cell(r, 14).value).lower() == "credit":
                    remove.append(r)
            for r in reversed(remove):
                ws_spec.delete_rows(r)

        # Verzamelblad
        vandaag_excel = dt.date.today().strftime("%d%m%Y")
        ws_verz["B4"].value = dt.date.today().strftime("%d-%m-%Y")
        ws_verz["C24"].value = f"{naam}_{vandaag_excel}"
        ws_verz["C26"].value = settings["periode"]

        # Opslaan
        out_xlsx = os.path.join(tmp, f"{naam}_{vandaag_excel}.xlsx")
        wb.save(out_xlsx)

        # ===============================================
        # SOMCONTROLE ✔️/❌
        # ===============================================
        kol_excl = df_data.columns.get_loc("Excl. BTW")
        kol_btw = df_data.columns.get_loc("BTW")
        kol_incl = df_data.columns.get_loc("Incl. BTW")

        sum_src = (
            df_data.iloc[:, kol_excl].sum(),
            df_data.iloc[:, kol_btw].sum(),
            df_data.iloc[:, kol_incl].sum(),
        )

        df_new = pd.read_excel(out_xlsx, sheet_name="Specificatie")

        sum_new = (
            df_new.iloc[:, kol_excl].sum(),
            df_new.iloc[:, kol_btw].sum(),
            df_new.iloc[:, kol_incl].sum(),
        )

        # Bericht
        if all(abs(a - b) < 0.01 for a, b in zip(sum_src, sum_new)):
            st.success(f"✔ Bedragen voor {naam} kloppen volledig.")
            ok_to_pdf = True
        else:
            st.error(f"❌ Bedragen komen niet overeen voor {naam}. PDF wordt niet gemaakt.")
            ok_to_pdf = False

        # ===============================================
        # DOWNLOADS
        # ===============================================
        with open(out_xlsx, "rb") as f:
            st.download_button(
                label=f"📊 Download Excel – {naam.capitalize()}",
                data=f,
                file_name=os.path.basename(out_xlsx)
            )

        # PDF
        if ok_to_pdf:
            out_pdf = out_xlsx.replace(".xlsx", ".pdf")
            excel_first_sheet_to_pdf(out_xlsx, out_pdf)

            with open(out_pdf, "rb") as f:
                st.download_button(
                    label=f"📄 Download PDF – {naam.capitalize()}",
                    data=f,
                    file_name=os.path.basename(out_pdf)
                )

st.success("🎉 Verwerking voltooid!")
