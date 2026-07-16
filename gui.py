import os
import sys
import threading
from tkinter import filedialog, messagebox
import customtkinter as ctk

# Biztosítjuk, hogy az excel_diff importálható legyen, akárhonnan is futtatjuk a kódot
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Biztonságos importálás: ha az excel_diff.py-ban nincs load_excel, helyileg pótoljuk pandas-szal
try:
    from excel_diff import compare_dataframes, write_report, load_excel
except ImportError:
    from excel_diff import compare_dataframes, write_report
    import pandas as pd
    def load_excel(file_path):
        return pd.read_excel(file_path)

# Globális téma beállítások (Sötét/Világos követése és Modern kék stílus)
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelDiffApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 1. Ablak alapbeállításai
        self.title("Excel Diff Tool")
        self.geometry("650x450")
        
        # AKTIVÁLJUK az átméretezhetőséget (most már tudod húzni a széleit!)
        self.resizable(True, True)
        self.minsize(580, 400)  # Minimális ablakméret a szétesés ellen

        # 2. Reszponzív rács (Grid) konfiguráció
        self.grid_columnconfigure(1, weight=1)  # A középső (beviteli mező) oszlop nyúlik dinamikusan
        
        # Biztosítjuk a sorok egyenletes függőleges eloszlását
        for i in range(6):
            self.grid_rowconfigure(i, weight=1)

        # Útvonalakat tároló StringVar változók
        self.old_path = ctk.StringVar()
        self.new_path = ctk.StringVar()
        self.output_path = ctk.StringVar(value="diff_eredmeny.xlsx")

        # --- Felhasználói felület elemei (Widgets) ---

        # Régi fájl sor
        self.label_old = ctk.CTkLabel(self, text="Régi fájl:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_old.grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        self.entry_old = ctk.CTkEntry(self, textvariable=self.old_path, placeholder_text="Válaszd ki a régi Excel fájlt...")
        self.entry_old.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        self.btn_old = ctk.CTkButton(self, text="Tallózás", width=100, command=self.pick_old)
        self.btn_old.grid(row=0, column=2, padx=20, pady=10, sticky="e")

        # Új fájl sor
        self.label_new = ctk.CTkLabel(self, text="Új fájl:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_new.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        
        self.entry_new = ctk.CTkEntry(self, textvariable=self.new_path, placeholder_text="Válaszd ki az új Excel fájlt...")
        self.entry_new.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        
        self.btn_new = ctk.CTkButton(self, text="Tallózás", width=100, command=self.pick_new)
        self.btn_new.grid(row=1, column=2, padx=20, pady=10, sticky="e")

        # Kulcsoszlop(ok) sor
        self.label_keys = ctk.CTkLabel(self, text="Kulcsoszlop(ok):", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_keys.grid(row=2, column=0, padx=20, pady=10, sticky="w")
        
        self.key_cols_entry = ctk.CTkEntry(self, placeholder_text="pl. id vagy id, nev (vesszővel elválasztva)")
        self.key_cols_entry.grid(row=2, column=1, columnspan=2, padx=10, pady=10, sticky="ew")

        # Kimeneti fájl sor
        self.label_out = ctk.CTkLabel(self, text="Kimeneti fájl:", font=ctk.CTkFont(size=13, weight="bold"))
        self.label_out.grid(row=3, column=0, padx=20, pady=10, sticky="w")
        
        self.entry_out = ctk.CTkEntry(self, textvariable=self.output_path)
        self.entry_out.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        
        self.btn_out = ctk.CTkButton(self, text="Tallózás", width=100, command=self.pick_output)
        self.btn_out.grid(row=3, column=2, padx=20, pady=10, sticky="e")

        # Futtatás gomb
        self.btn_run = ctk.CTkButton(
            self, 
            text="Összehasonlítás futtatása", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            fg_color="#1f538d",
            hover_color="#14375e",
            command=self.run_compare
        )
        self.btn_run.grid(row=4, column=0, columnspan=3, padx=20, pady=15, sticky="ew")

        # Állapotsor (Status Label)
        self.status_label = ctk.CTkLabel(self, text="Készen áll az összehasonlításra.", font=ctk.CTkFont(size=12))
        self.status_label.grid(row=5, column=0, columnspan=3, padx=20, pady=5, sticky="ew")

    # Tallózási folyamatok
    def pick_old(self):
        path = filedialog.askopenfilename(
            title="Régi Excel fájl kiválasztása",
            filetypes=[("Excel táblázatok", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            self.old_path.set(path)

    def pick_new(self):
        path = filedialog.askopenfilename(
            title="Új Excel fájl kiválasztása",
            filetypes=[("Excel táblázatok", "*.xlsx *.xls *.xlsm")]
        )
        if path:
            self.new_path.set(path)

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Kimeneti fájl mentése",
            defaultextension=".xlsx",
            filetypes=[("Excel táblázat", "*.xlsx")]
        )
        if path:
            self.output_path.set(path)

    # Összehasonlítás indítása
    def run_compare(self):
        old_f = self.old_path.get().strip()
        new_f = self.new_path.get().strip()
        keys_raw = self.key_cols_entry.get().strip()
        out_f = self.output_path.get().strip()

        # Érvényesség ellenőrzése
        if not old_f or not new_f or not keys_raw:
            messagebox.showerror("Hiba", "Töltsd ki mindkét fájl elérési útját és a kulcsoszlopot!")
            return

        # Kulcsoszlopok tisztítása és feldolgozása
        key_cols = [k.strip() for k in keys_raw.replace(";", ",").split(",") if k.strip()]

        # GUI felület frissítése
        self.status_label.configure(text="Összehasonlítás futása...", text_color="yellow")
        self.update_idletasks()

        # Futtatás külön szálon, hogy ne fagyjon ki az ablak
        thread = threading.Thread(target=self._do_compare, args=(old_f, new_f, key_cols, out_f))
        thread.start()

    # Háttérben futó logika
    def _do_compare(self, old_f, new_f, key_cols, out_f):
        try:
            df_old = load_excel(old_f)
            df_new = load_excel(new_f)

            # Eltérések keresése
            only_old, only_new, changed = compare_dataframes(df_old, df_new, key_cols)

            # Riport generálása
            write_report(only_old, only_new, changed, out_f)

            # Sikeres visszajelzés
            self.status_label.configure(text="Sikeres futás! Riport elmentve.", text_color="green")
            messagebox.showinfo("Siker", f"Az összehasonlítás sikeresen lefutott!\nA riport elkészült:\n{out_f}")

        except Exception as e:
            self.status_label.configure(text="Hiba történt a futtatás során!", text_color="red")
            messagebox.showerror("Hiba", f"Hiba történt a fájlok feldolgozása közben:\n{str(e)}")

if __name__ == "__main__":
    app = ExcelDiffApp()
    app.mainloop()