"""
Wysyłka SMS do aktywnych kierowców.
Czyta plik Excel (.xls/.xlsx) i wysyła SMS-y przez telefon z Androidem
(aplikacja "SMS Gateway for Android" w trybie Local Server).

Ustawienia są w pliku config.json obok programu (lub w oknie "Ustawienia").

Instalacja:  pip install pandas openpyxl xlrd requests sv-ttk
Budowa .exe: pyinstaller --onefile --windowed --collect-data sv_ttk --name WysylkaSMS wysylka_sms.py
"""

import hashlib
import json
import re
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import requests

try:
    import sv_ttk
except ImportError:
    sv_ttk = None

# Ostre czcionki na ekranach o wysokiej rozdzielczości
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ----------------------------- konfiguracja -----------------------------
BAZA = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
CONFIG_PLIK = BAZA / "config.json"
LOG_PLIK = BAZA / "wyslane_log.json"

DOMYSLNE = {
    "kolumna_nazwisko": "nazwisko i imię",
    "kolumna_telefon": "numer telefonu",
    "telefon_adres": "http://192.168.1.50:8080",
    "telefon_endpoint": "/message",
    "telefon_login": "sms",
    "telefon_haslo": "haslo",
    "przerwa_sekundy": 1.5,
    "timeout_sekundy": 15,
}
CFG = dict(DOMYSLNE)


def wczytaj_config():
    if CONFIG_PLIK.exists():
        try:
            CFG.update(json.loads(CONFIG_PLIK.read_text(encoding="utf-8")))
        except Exception:
            messagebox.showwarning("Ustawienia", "Nie udało się odczytać config.json, użyto domyślnych.")
    else:
        zapisz_config()


def zapisz_config():
    CONFIG_PLIK.write_text(json.dumps(CFG, ensure_ascii=False, indent=2), encoding="utf-8")


# ------------------------------ logika ------------------------------
def normalizuj_numer(surowy):
    """Zwraca numer w formacie +48XXXXXXXXX albo None, jeśli nieprawidłowy."""
    if surowy is None:
        return None
    s = str(surowy).strip()
    if s.lower() in ("", "nan", "none"):
        return None
    if re.fullmatch(r"\d+\.0", s):          # Excel czasem zapisuje 600123456.0
        s = s[:-2]
    cyfry = re.sub(r"\D", "", s)
    if cyfry.startswith("00"):
        cyfry = cyfry[2:]
    if len(cyfry) == 9:
        cyfry = "48" + cyfry
    numer = "+" + cyfry
    return numer if re.fullmatch(r"\+48\d{9}", numer) else None


def wczytaj_log():
    try:
        return json.loads(LOG_PLIK.read_text(encoding="utf-8"))
    except Exception:
        return {}


def zapisz_log(dane):
    LOG_PLIK.write_text(json.dumps(dane, ensure_ascii=False, indent=1), encoding="utf-8")


def odcisk(tresc, numery):
    surowy = datetime.now().strftime("%Y-%m-%d") + "|" + tresc + "|" + ",".join(sorted(numery))
    return hashlib.sha256(surowy.encode("utf-8")).hexdigest()


def wyslij_sms(numer, tresc):
    """Wysyła jeden SMS przez telefon. Zwraca (ok, komunikat)."""
    try:
        r = requests.post(
            CFG["telefon_adres"].rstrip("/") + CFG["telefon_endpoint"],
            json={"message": tresc, "phoneNumbers": [numer]},
            auth=(CFG["telefon_login"], CFG["telefon_haslo"]),
            timeout=float(CFG["timeout_sekundy"]),
        )
        if r.status_code in (200, 201, 202):
            return True, "OK"
        return False, f"HTTP {r.status_code}"
    except requests.exceptions.ConnectTimeout:
        return False, "brak połączenia z telefonem (timeout)"
    except requests.exceptions.ConnectionError:
        return False, "telefon niedostępny w sieci"
    except Exception as e:
        return False, str(e)[:80]


# ------------------------------ okno ustawień ------------------------------
class OknoUstawien(tk.Toplevel):
    POLA = [
        ("kolumna_nazwisko", "Kolumna z nazwiskiem (Excel)"),
        ("kolumna_telefon", "Kolumna z numerem telefonu (Excel)"),
        ("telefon_adres", "Adres telefonu (np. http://192.168.1.50:8080)"),
        ("telefon_endpoint", "Ścieżka API (zwykle /message)"),
        ("telefon_login", "Login bramki"),
        ("telefon_haslo", "Hasło bramki"),
        ("przerwa_sekundy", "Przerwa między SMS-ami (sekundy)"),
    ]

    def __init__(self, rodzic):
        super().__init__(rodzic)
        self.title("Ustawienia")
        self.resizable(False, False)
        self.transient(rodzic)
        self.grab_set()
        ramka = ttk.Frame(self, padding=20)
        ramka.pack(fill="both", expand=True)
        self.zmienne = {}
        for i, (klucz, opis) in enumerate(self.POLA):
            ttk.Label(ramka, text=opis).grid(row=i * 2, column=0, sticky="w", pady=(8, 0))
            v = tk.StringVar(value=str(CFG[klucz]))
            self.zmienne[klucz] = v
            ttk.Entry(ramka, textvariable=v, width=44,
                      show="•" if klucz == "telefon_haslo" else "").grid(
                row=i * 2 + 1, column=0, sticky="we")
        przyciski = ttk.Frame(ramka)
        przyciski.grid(row=len(self.POLA) * 2, column=0, pady=(18, 0), sticky="e")
        ttk.Button(przyciski, text="Anuluj", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(przyciski, text="Zapisz", style="Accent.TButton",
                   command=self.zapisz).pack(side="left")

    def zapisz(self):
        try:
            przerwa = float(self.zmienne["przerwa_sekundy"].get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Błąd", "Przerwa musi być liczbą.", parent=self)
            return
        for klucz, v in self.zmienne.items():
            CFG[klucz] = przerwa if klucz == "przerwa_sekundy" else v.get().strip()
        zapisz_config()
        self.destroy()


# ------------------------------ aplikacja ------------------------------
class Aplikacja(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Wysyłka SMS do kierowców")
        self.geometry("860x760")
        self.minsize(760, 680)
        if sv_ttk:
            sv_ttk.set_theme("light")
        self.odbiorcy, self.nieudane = [], []
        self.trwa = False
        self._buduj()
        self._licznik()

    def _buduj(self):
        glowna = ttk.Frame(self, padding=(24, 18))
        glowna.pack(fill="both", expand=True)

        # Nagłówek
        gora = ttk.Frame(glowna)
        gora.pack(fill="x")
        ttk.Label(gora, text="Wysyłka SMS do kierowców",
                  font=("Segoe UI Semibold", 18)).pack(side="left")
        ttk.Button(gora, text="⚙  Ustawienia",
                   command=lambda: OknoUstawien(self)).pack(side="right")
        ttk.Button(gora, text="Test połączenia", command=self.test_polaczenia).pack(side="right", padx=8)

        # Krok 1
        ttk.Label(glowna, text="1  Lista odbiorców", font=("Segoe UI Semibold", 11)).pack(
            anchor="w", pady=(18, 6))
        k1 = ttk.Frame(glowna)
        k1.pack(fill="x")
        ttk.Button(k1, text="Wybierz plik Excel…", style="Accent.TButton",
                   command=self.wybierz_plik).pack(side="left")
        self.lbl_plik = ttk.Label(k1, text="Nie wybrano pliku", foreground="#777")
        self.lbl_plik.pack(side="left", padx=12)
        self.lbl_info = ttk.Label(glowna, text="", foreground="#444")
        self.lbl_info.pack(anchor="w", pady=(6, 6))

        tabela_ramka = ttk.Frame(glowna)
        tabela_ramka.pack(fill="both", expand=True)
        self.tabela = ttk.Treeview(tabela_ramka, columns=("nazwisko", "numer", "status"),
                                   show="headings", height=9)
        for kol, nagl, szer in (("nazwisko", "Nazwisko i imię", 320),
                                ("numer", "Numer", 150), ("status", "Status", 260)):
            self.tabela.heading(kol, text=nagl, anchor="w")
            self.tabela.column(kol, width=szer, anchor="w")
        self.tabela.tag_configure("ok", foreground="#1a7f37")
        self.tabela.tag_configure("blad", foreground="#c62828")
        self.tabela.tag_configure("czeka", foreground="#888")
        sb = ttk.Scrollbar(tabela_ramka, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=sb.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Krok 2
        ttk.Label(glowna, text="2  Treść wiadomości", font=("Segoe UI Semibold", 11)).pack(
            anchor="w", pady=(16, 6))
        self.txt = tk.Text(glowna, height=5, wrap="word", font=("Segoe UI", 11),
                           relief="flat", highlightthickness=1,
                           highlightbackground="#ccc", highlightcolor="#0067c0",
                           padx=10, pady=8)
        self.txt.pack(fill="x")
        self.txt.bind("<KeyRelease>", lambda e: self._licznik())
        self.lbl_licznik = ttk.Label(glowna, text="", foreground="#777")
        self.lbl_licznik.pack(anchor="w", pady=(4, 0))

        # Krok 3
        ttk.Label(glowna, text="3  Wysyłka", font=("Segoe UI Semibold", 11)).pack(
            anchor="w", pady=(14, 6))
        dol = ttk.Frame(glowna)
        dol.pack(fill="x")
        self.btn_wyslij = ttk.Button(dol, text="Wyślij SMS", style="Accent.TButton",
                                     command=self.wyslij)
        self.btn_wyslij.pack(side="left")
        self.btn_ponow = ttk.Button(dol, text="Ponów nieudane", command=self.ponow,
                                    state="disabled")
        self.btn_ponow.pack(side="left", padx=10)
        self.lbl_postep = ttk.Label(dol, text="")
        self.lbl_postep.pack(side="right")
        self.pasek = ttk.Progressbar(dol, length=260)
        self.pasek.pack(side="right", padx=12)

    def _licznik(self):
        t = self.txt.get("1.0", "end").strip()
        n = len(t)
        polskie = bool(re.search(r"[ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]", t))
        limit = 70 if polskie else 160
        sms = max(1, -(-n // limit)) if n else 0
        self.lbl_licznik.config(
            text=f"{n} znaków  •  około {sms} SMS na odbiorcę"
                 + ("  •  polskie znaki skracają limit do 70 znaków" if polskie else ""))

    # ---------- test połączenia ----------
    def test_polaczenia(self):
        try:
            requests.get(CFG["telefon_adres"], timeout=5,
                         auth=(CFG["telefon_login"], CFG["telefon_haslo"]))
            messagebox.showinfo("Test połączenia", "Telefon odpowiada. Połączenie działa.")
        except Exception:
            messagebox.showerror(
                "Test połączenia",
                "Telefon nie odpowiada.\n\nSprawdź, czy telefon i komputer są w tej samej "
                "sieci Wi-Fi, czy aplikacja bramki jest włączona i czy adres w "
                "Ustawieniach jest poprawny.")

    # ---------- wczytywanie pliku ----------
    def wybierz_plik(self):
        sciezka = filedialog.askopenfilename(
            filetypes=[("Pliki Excel", "*.xls *.xlsx"), ("Wszystkie", "*.*")])
        if not sciezka:
            return
        try:
            df = pd.read_excel(sciezka, dtype=str)
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się wczytać pliku:\n{e}")
            return
        df.columns = [str(c).strip().lower() for c in df.columns]
        k_n = CFG["kolumna_nazwisko"].strip().lower()
        k_t = CFG["kolumna_telefon"].strip().lower()
        if k_n not in df.columns or k_t not in df.columns:
            messagebox.showerror(
                "Brak kolumn",
                f"Nie znaleziono kolumn '{CFG['kolumna_nazwisko']}' i '{CFG['kolumna_telefon']}'.\n\n"
                f"Kolumny w pliku:\n{', '.join(df.columns)}\n\n"
                "Popraw nazwy w oknie Ustawienia.")
            return

        widziane, self.odbiorcy = set(), []
        bledne = duplikaty = 0
        for _, w in df.iterrows():
            numer = normalizuj_numer(w[k_t])
            nazwisko = str(w[k_n]).strip() if pd.notna(w[k_n]) else ""
            if numer is None:
                if nazwisko or pd.notna(w[k_t]):
                    bledne += 1
                continue
            if numer in widziane:
                duplikaty += 1
                continue
            widziane.add(numer)
            self.odbiorcy.append((nazwisko, numer))

        self.nieudane = []
        self.btn_ponow.config(state="disabled")
        self.lbl_plik.config(text=Path(sciezka).name, foreground="#222")
        self.lbl_info.config(
            text=f"Odbiorców: {len(self.odbiorcy)}   •   błędne/puste numery pominięte: {bledne}"
                 f"   •   duplikaty: {duplikaty}")
        self.tabela.delete(*self.tabela.get_children())
        for n, t in self.odbiorcy:
            self.tabela.insert("", "end", iid=t, values=(n, t, ""))
        self.pasek["value"] = 0
        self.lbl_postep.config(text="")

    # ---------- wysyłka ----------
    def wyslij(self):
        tresc = self.txt.get("1.0", "end").strip()
        if self.trwa:
            return
        if not self.odbiorcy:
            messagebox.showwarning("Brak odbiorców", "Najpierw wczytaj plik Excel.")
            return
        if not tresc:
            messagebox.showwarning("Brak treści", "Wpisz treść wiadomości.")
            return
        log = wczytaj_log()
        h = odcisk(tresc, [t for _, t in self.odbiorcy])
        if h in log and not messagebox.askyesno(
                "Uwaga: możliwy duplikat",
                f"Ta sama wiadomość do tej samej listy została już wysłana ({log[h]}).\n"
                "Wysłać ponownie?"):
            return
        if not messagebox.askyesno(
                "Potwierdzenie", f"Wysłać SMS do {len(self.odbiorcy)} osób?\n\n{tresc}"):
            return
        log[h] = datetime.now().strftime("%Y-%m-%d %H:%M")
        zapisz_log(log)
        self._start(list(self.odbiorcy), tresc)

    def ponow(self):
        tresc = self.txt.get("1.0", "end").strip()
        if self.nieudane and tresc and not self.trwa:
            self._start(list(self.nieudane), tresc)

    def _start(self, lista, tresc):
        self.trwa = True
        self.btn_wyslij.config(state="disabled")
        self.btn_ponow.config(state="disabled")
        self.pasek.config(maximum=len(lista), value=0)
        for _, t in lista:
            self.tabela.set(t, "status", "oczekuje")
            self.tabela.item(t, tags=("czeka",))
        threading.Thread(target=self._petla, args=(lista, tresc), daemon=True).start()

    def _petla(self, lista, tresc):
        nieudane, ok_licznik = [], 0
        for i, (nazwisko, numer) in enumerate(lista, 1):
            ok, kom = wyslij_sms(numer, tresc)
            if ok:
                ok_licznik += 1
            else:
                nieudane.append((nazwisko, numer))
            self.after(0, self._aktualizuj, numer, ok, kom, i, len(lista))
            time.sleep(float(CFG["przerwa_sekundy"]))
        self.after(0, self._koniec, nieudane, ok_licznik, len(lista))

    def _aktualizuj(self, numer, ok, kom, i, razem):
        self.tabela.set(numer, "status", "✔ wysłano" if ok else "✖ " + kom)
        self.tabela.item(numer, tags=("ok" if ok else "blad",))
        self.pasek["value"] = i
        self.lbl_postep.config(text=f"{i} / {razem}")

    def _koniec(self, nieudane, ok_licznik, razem):
        self.trwa = False
        self.nieudane = nieudane
        self.btn_wyslij.config(state="normal")
        if nieudane:
            self.btn_ponow.config(state="normal")
        messagebox.showinfo(
            "Zakończono",
            f"Wysłano: {ok_licznik} z {razem}\nBłędy: {len(nieudane)}"
            + ("\n\nUżyj przycisku „Ponów nieudane”." if nieudane else ""))


if __name__ == "__main__":
    wczytaj_config()
    Aplikacja().mainloop()
