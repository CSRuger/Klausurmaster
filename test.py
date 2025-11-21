import tkinter as tk
from tkinter import messagebox, simpledialog, ttk, filedialog
import json, random, os

CONFIG_FILE = "config.json"

class CardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Karteikarten App")
        self.root.geometry("1280x720")
        self.root.minsize(800, 600)

        # Icon für die App setzen
        self.root.iconbitmap("app_icon.ico")

        # Tabellenstruktur
        self.base_columns = ["Anfang", "2.0"]
        self.columns = []
        self.rows = []
        self.cards = {}
        self.row_colors = {}  # Speichert die Hintergrundfarben pro Fach

        # Ausgewählte Karte für Bewegung
        self.selected_card = None
        self.selected_card_position = None

        self.delete_mode = False  # Löschmodus deaktiviert starten

        # Speicherverzeichnis laden oder abfragen
        self.save_file = self.load_or_ask_save_directory()

        # Anzahl der Progressionsstufen festlegen
        self.set_progression_stages()

        # Haupt-Container
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Tabellenbereich
        self.table = ttk.Frame(self.main_frame)
        self.table.grid(row=1, column=0, sticky=tk.NSEW)

        # Steuerungsbereich
        self.controls = ttk.Frame(self.main_frame, padding=10)
        self.controls.grid(row=2, column=0, sticky=tk.EW)

        # Optionen-Button
        options_menu = tk.Menu(self.root)
        self.root.config(menu=options_menu)
        options_menu.add_command(label="Optionen", command=self.open_options_menu)

        # Steuerungsbuttons
        ttk.Button(self.controls, text="Zeile hinzufügen", command=self.add_row).pack(side=tk.LEFT, padx=5)
        self.add_card_button = ttk.Button(self.controls, text="Karte hinzufügen", command=self.add_card)
        self.add_card_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.controls, text="Speichern", command=self.save_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.controls, text="Laden", command=self.load_data).pack(side=tk.LEFT, padx=5)

        # Grid-Konfiguration für dynamische Größenanpassung
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        self.create_table()
        self.load_data()

        # ESC-Abbruch-Event binden
        self.root.bind("<Escape>", self.cancel_operation)

    def load_or_ask_save_directory(self):
        """Lädt das Speicherverzeichnis aus der Konfigurationsdatei oder fragt danach."""
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
                return config.get("save_file")
        else:
            dir_path = filedialog.askdirectory(title="Wähle ein Verzeichnis für den Arbeitsstand")
            if not dir_path:
                messagebox.showerror("Fehler", "Kein Verzeichnis ausgewählt. Das Programm wird beendet.")
                self.root.destroy()
            save_file = os.path.join(dir_path, "card_app_data.json")
            self.save_config(save_file)
            return save_file

    def save_config(self, save_file):
        """Speichert das Speicherverzeichnis in der Konfigurationsdatei."""
        with open(CONFIG_FILE, "w") as f:
            json.dump({"save_file": save_file}, f)

    def set_progression_stages(self):
        """Fragt nach der Anzahl der Progressionsstufen und erstellt die Spalten."""
        while True:
            try:
                num_stages = simpledialog.askinteger("Progressionsstufen", "Wie viele Progressionsstufen möchtest du haben?")
                if num_stages is None:
                    messagebox.showerror("Fehler", "Keine Anzahl eingegeben. Das Programm wird beendet.")
                    self.root.destroy()
                    return
                if num_stages > 0:
                    self.columns = ["Anfang"] + [f"{i}. Stufe" for i in range(1, num_stages + 1)] + ["2.0"]
                    break
                else:
                    messagebox.showerror("Ungültige Eingabe", "Bitte gib eine Zahl größer als 0 ein.")
            except ValueError:
                messagebox.showerror("Ungültige Eingabe", "Bitte gib eine gültige Zahl ein.")

    def open_options_menu(self):
        """Öffnet das Optionen-Menü."""
        options_window = tk.Toplevel(self.root)
        options_window.title("Optionen")
        options_window.geometry("300x150")

        ttk.Button(options_window, text="Speicherverzeichnis ändern", command=self.change_save_directory).pack(pady=20)

    def change_save_directory(self):
        """Ändert das Speicherverzeichnis."""
        dir_path = filedialog.askdirectory(title="Neues Speicherverzeichnis auswählen")
        if dir_path:
            self.save_file = os.path.join(dir_path, "card_app_data.json")
            self.save_config(self.save_file)
            messagebox.showinfo("Speicherverzeichnis geändert", f"Neues Speicherverzeichnis: {self.save_file}")

    def create_table(self):
        """Erstellt die Tabellenstruktur."""
        for col_idx, col_name in enumerate(self.columns):
            label = ttk.Label(self.table, text=col_name, anchor="center", padding=5, style="TableHeader.TLabel")
            label.grid(row=0, column=col_idx, sticky=tk.NSEW, padx=2, pady=2)

        for col in range(len(self.columns)):
            self.table.columnconfigure(col, weight=1)

    def cancel_operation(self, event):
        """Bricht die aktuelle Operation ab."""
        self.root.focus_set()
        messagebox.showinfo("Abbruch", "Operation wurde abgebrochen.")

def random_pastel_color():
    """Erzeugt eine zufällige Pastellfarbe."""
    r = random.randint(200, 255)
    g = random.randint(200, 255)
    b = random.randint(200, 255)
    return f"#{r:02x}{g:02x}{b:02x}"

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style()
    style.configure("TableHeader.TLabel", font=("Arial", 12, "bold"), background="#4A90E2", foreground="white")
    app = CardApp(root)
    root.mainloop()
