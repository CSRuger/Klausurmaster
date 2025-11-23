"""Tkinter UI for managing cards, tables, and aggregated statistics."""

from __future__ import annotations

import copy
import json
import os
import sys
from contextlib import contextmanager
from datetime import datetime
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
import tkinter as tk

from cards import create_card, find_card, normalize_cards_tree
from formula import calculate_ratio, calculate_expected_grade
from table import generate_columns, interpolate_color, random_pastel_color
from main.runtime_paths import (
    APP_NAME,
    APP_VERSION,
    DEFAULT_LANGUAGE,
    DEFAULT_SAVE_FILENAME,
    get_user_config_path,
    load_save_file_path,
    load_user_language,
    persist_save_file_path,
    persist_user_language,
    resource_path,
)

SAVE_FILE = ""
BIN_HISTORY_FILENAME = "tabellenspeicher_bin.json"

THEMES = {
    "dark": {
        "APP_BG": "#070b16",
        "PANEL_BG": "#0f172a",
        "CONTENT_BG": "#0b1120",
        "CARD_BG": "#16243a",
        "CARD_BORDER": "#243b53",
        "CARD_DELETE_BG": "#3b1d1d",
        "PRIMARY_ACCENT": "#38bdf8",
        "TEXT_PRIMARY": "#f1f5f9",
        "TEXT_MUTED": "#94a3b8",
    },
    "beige": {
        "APP_BG": "#fdfdfc",
        "PANEL_BG": "#f1f5fb",
        "CONTENT_BG": "#ffffff",
        "CARD_BG": "#fefefe",
        "CARD_BORDER": "#5490d9",
        "CARD_DELETE_BG": "#ffe5e5",
        "PRIMARY_ACCENT": "#d64045",
        "TEXT_PRIMARY": "#1e2a3b",
        "TEXT_MUTED": "#55657a",
    },
}

THEME_KEYS = [
    "APP_BG",
    "PANEL_BG",
    "CONTENT_BG",
    "CARD_BG",
    "CARD_BORDER",
    "CARD_DELETE_BG",
    "PRIMARY_ACCENT",
    "TEXT_PRIMARY",
    "TEXT_MUTED",
]

APP_BG = THEMES["dark"]["APP_BG"]
PANEL_BG = THEMES["dark"]["PANEL_BG"]
CONTENT_BG = THEMES["dark"]["CONTENT_BG"]
CARD_BG = THEMES["dark"]["CARD_BG"]
CARD_BORDER = THEMES["dark"]["CARD_BORDER"]
CARD_DELETE_BG = THEMES["dark"]["CARD_DELETE_BG"]
PRIMARY_ACCENT = THEMES["dark"]["PRIMARY_ACCENT"]
TEXT_PRIMARY = THEMES["dark"]["TEXT_PRIMARY"]
TEXT_MUTED = THEMES["dark"]["TEXT_MUTED"]

LANGUAGE_OPTIONS = {
    "de": "Deutsch",
    "en": "English",
    "es": "Español",
    "fr": "Français",
    "sq": "Shqip",
}

TRANSLATIONS: dict[str, dict[str, str]] = {
    "app.subtitle": {
        "de": "Wähle links Space und Tabelle, um Karten zu fokussieren.",
        "en": "Choose a space and table on the left to focus your cards.",
        "es": "Elige a la izquierda un espacio y una tabla para enfocarte en las tarjetas.",
        "fr": "Choisissez à gauche un espace et une table pour vous concentrer sur les cartes.",
        "sq": "Zgjidh në të majtë një hapësirë dhe tabelë për t'u fokusuar te kartat.",
    },
    "nav.collections": {
        "de": "Sammlungen",
        "en": "Collections",
        "es": "Colecciones",
        "fr": "Collections",
        "sq": "Koleksione",
    },
    "menu.file": {
        "de": "Datei",
        "en": "File",
        "es": "Archivo",
        "fr": "Fichier",
        "sq": "Skedar",
    },
    "menu.edit": {
        "de": "Bearbeiten",
        "en": "Edit",
        "es": "Editar",
        "fr": "Édition",
        "sq": "Redakto",
    },
    "menu.space": {
        "de": "Space",
        "en": "Space",
        "es": "Espacio",
        "fr": "Espace",
        "sq": "Hapësirë",
    },
    "menu.help": {
        "de": "Hilfe",
        "en": "Help",
        "es": "Ayuda",
        "fr": "Aide",
        "sq": "Ndihmë",
    },
    "menu.language": {
        "de": "Sprache",
        "en": "Language",
        "es": "Idioma",
        "fr": "Langue",
        "sq": "Gjuha",
    },
    "menu.actions": {
        "de": "Aktionen",
        "en": "Actions",
        "es": "Acciones",
        "fr": "Actions",
        "sq": "Veprime",
    },
    "menu.file.save": {
        "de": "Speichern",
        "en": "Save",
        "es": "Guardar",
        "fr": "Enregistrer",
        "sq": "Ruaj",
    },
    "menu.file.reload": {
        "de": "Neu laden",
        "en": "Reload",
        "es": "Recargar",
        "fr": "Recharger",
        "sq": "Ringarko",
    },
    "menu.file.view_state": {
        "de": "Status ansehen aus…",
        "en": "View state from…",
        "es": "Ver estado desde…",
        "fr": "Afficher l'état depuis…",
        "sq": "Shiko gjendjen nga…",
    },
    "menu.file.load_state": {
        "de": "Status laden aus…",
        "en": "Load state from…",
        "es": "Cargar estado desde…",
        "fr": "Charger l'état depuis…",
        "sq": "Ngarko gjendjen nga…",
    },
    "menu.file.add_table": {
        "de": "Tabelle hinzufügen (Strg+N)",
        "en": "Add table (Ctrl+N)",
        "es": "Añadir tabla (Ctrl+N)",
        "fr": "Ajouter une table (Ctrl+N)",
        "sq": "Shto tabelë (Ctrl+N)",
    },
    "menu.file.delete_table": {
        "de": "Tabelle löschen",
        "en": "Delete table",
        "es": "Eliminar tabla",
        "fr": "Supprimer la table",
        "sq": "Fshi tabelën",
    },
    "menu.file.add_card": {
        "de": "Karte hinzufügen (1–9)",
        "en": "Add card (1–9)",
        "es": "Añadir tarjeta (1–9)",
        "fr": "Ajouter une carte (1–9)",
        "sq": "Shto kartë (1–9)",
    },
    "menu.file.toggle_delete": {
        "de": "Löschmodus umschalten (Strg+W)",
        "en": "Toggle delete mode (Ctrl+W)",
        "es": "Cambiar modo borrar (Ctrl+W)",
        "fr": "Basculer le mode suppression (Ctrl+W)",
        "sq": "Ndrysho modalitetin e fshirjes (Ctrl+W)",
    },
    "menu.file.cancel_action": {
        "de": "Aktion abbrechen (Esc)",
        "en": "Cancel action (Esc)",
        "es": "Cancelar acción (Esc)",
        "fr": "Annuler l'action (Esc)",
        "sq": "Anulo veprimin (Esc)",
    },
    "menu.file.choose_folder": {
        "de": "Speicherordner wählen…",
        "en": "Choose save folder…",
        "es": "Elegir carpeta de guardado…",
        "fr": "Choisir le dossier de sauvegarde…",
        "sq": "Zgjidh dosjen e ruajtjes…",
    },
    "menu.file.change_file": {
        "de": "Speicherdatei ändern…",
        "en": "Change save file…",
        "es": "Cambiar archivo de guardado…",
        "fr": "Changer le fichier de sauvegarde…",
        "sq": "Ndrysho skedarin e ruajtjes…",
    },
    "menu.file.exit": {
        "de": "Beenden",
        "en": "Exit",
        "es": "Salir",
        "fr": "Quitter",
        "sq": "Dalje",
    },
    "menu.edit.undo": {
        "de": "Rückgängig (Strg+Z)",
        "en": "Undo (Ctrl+Z)",
        "es": "Deshacer (Ctrl+Z)",
        "fr": "Annuler (Ctrl+Z)",
        "sq": "Zhbëj (Ctrl+Z)",
    },
    "menu.edit.redo": {
        "de": "Wiederholen (Strg+Shift+Z)",
        "en": "Redo (Ctrl+Shift+Z)",
        "es": "Rehacer (Ctrl+Shift+Z)",
        "fr": "Rétablir (Ctrl+Shift+Z)",
        "sq": "Ribëj (Ctrl+Shift+Z)",
    },
    "menu.space.new": {
        "de": "Neuer Space (Strg+S)",
        "en": "New space (Ctrl+S)",
        "es": "Nuevo espacio (Ctrl+S)",
        "fr": "Nouvel espace (Ctrl+S)",
        "sq": "Hapësirë e re (Ctrl+S)",
    },
    "menu.space.load": {
        "de": "Space laden…",
        "en": "Load space…",
        "es": "Cargar espacio…",
        "fr": "Charger un espace…",
        "sq": "Ngarko hapësirë…",
    },
    "menu.space.import_space": {
        "de": "Space importieren…",
        "en": "Import space…",
        "es": "Importar espacio…",
        "fr": "Importer un espace…",
        "sq": "Importo hapësirë…",
    },
    "menu.space.export_space": {
        "de": "Space exportieren…",
        "en": "Export space…",
        "es": "Exportar espacio…",
        "fr": "Exporter un espace…",
        "sq": "Eksporto hapësirë…",
    },
    "menu.space.import_cards": {
        "de": "Karten importieren…",
        "en": "Import cards…",
        "es": "Importar tarjetas…",
        "fr": "Importer des cartes…",
        "sq": "Importo karta…",
    },
    "menu.space.import_short": {
        "de": "Import",
        "en": "Import",
        "es": "Importar",
        "fr": "Importer",
        "sq": "Importo",
    },
    "menu.space.duplicate": {
        "de": "Space duplizieren…",
        "en": "Duplicate space…",
        "es": "Duplicar espacio…",
        "fr": "Dupliquer l'espace…",
        "sq": "Dupliko hapësirën…",
    },
    "menu.space.transfer": {
        "de": "Tabelle verschieben/kopieren…",
        "en": "Move/copy table…",
        "es": "Mover/copiar tabla…",
        "fr": "Déplacer/copier la table…",
        "sq": "Zhvendos/kopjo tabelën…",
    },
    "menu.space.delete": {
        "de": "Space löschen…",
        "en": "Delete space…",
        "es": "Eliminar espacio…",
        "fr": "Supprimer l'espace…",
        "sq": "Fshi hapësirën…",
    },
    "menu.space.theme": {
        "de": "Theme",
        "en": "Theme",
        "es": "Tema",
        "fr": "Thème",
        "sq": "Tema",
    },
    "menu.space.theme.dark": {
        "de": "Dunkel",
        "en": "Dark",
        "es": "Oscuro",
        "fr": "Sombre",
        "sq": "I errët",
    },
    "menu.space.theme.beige": {
        "de": "Beige",
        "en": "Beige",
        "es": "Beige",
        "fr": "Beige",
        "sq": "Bezhë",
    },
    "menu.space.theme.custom": {
        "de": "Eigenes Theme…",
        "en": "Custom theme…",
        "es": "Tema personalizado…",
        "fr": "Thème personnalisé…",
        "sq": "Temë e personalizuar…",
    },
    "menu.help.about": {
        "de": "Über",
        "en": "About",
        "es": "Acerca de",
        "fr": "À propos",
        "sq": "Rreth",
    },
    "action.toggle_mark_mode": {
        "de": "Markiermodus umschalten",
        "en": "Toggle mark mode",
        "es": "Cambiar modo marcado",
        "fr": "Basculer le mode de marquage",
        "sq": "Ndrysho modalitetin e shënimit",
    },
    "save_prompt.title": {
        "de": "Speicherort wählen",
        "en": "Choose storage location",
        "es": "Elegir ubicación de guardado",
        "fr": "Choisir l'emplacement de stockage",
        "sq": "Zgjidh vendndodhjen e ruajtjes",
    },
    "save_prompt.body": {
        "de": "Klausurmaster benötigt einen Speicherordner für Ihre Tabellen.\n\nMöchten Sie jetzt einen Ordner wählen? Andernfalls wird der Standardordner im Benutzerverzeichnis verwendet.",
        "en": "Klausurmaster needs a folder to store your tables.\n\nDo you want to choose one now? Otherwise the default directory in your user profile will be used.",
        "es": "Klausurmaster necesita una carpeta para guardar tus tablas.\n\n¿Quieres elegir una ahora? De lo contrario se usará la carpeta predeterminada de tu perfil.",
        "fr": "Klausurmaster a besoin d'un dossier pour enregistrer vos tableaux.\n\nSouhaitez-vous en choisir un maintenant ? Sinon, le dossier par défaut de votre profil sera utilisé.",
        "sq": "Klausurmaster ka nevojë për një dosje për të ruajtur tabelat tuaja.\n\nDëshiron të zgjedhësh tani një dosje? Përndryshe përdoret dosja standarde në profilin tënd.",
    },
    "save_prompt.default_title": {
        "de": "Standard verwendet",
        "en": "Default used",
        "es": "Se usa el predeterminado",
        "fr": "Emplacement par défaut",
        "sq": "Është përdorur parazgjedhja",
    },
    "save_prompt.default_body": {
        "de": "Es wird der Standardordner verwendet.",
        "en": "The default folder will be used.",
        "es": "Se usará la carpeta predeterminada.",
        "fr": "Le dossier par défaut sera utilisé.",
        "sq": "Do të përdoret dosja e parazgjedhur.",
    },
    "status.current_path": {
        "de": "Aktueller Speicherpfad: {path}",
        "en": "Current save path: {path}",
        "es": "Ruta de guardado actual: {path}",
        "fr": "Chemin d'enregistrement actuel : {path}",
        "sq": "Shtegu aktual i ruajtjes: {path}",
    },
    "choose_directory.success_title": {
        "de": "Erfolg",
        "en": "Success",
        "es": "Éxito",
        "fr": "Succès",
        "sq": "Sukses",
    },
    "choose_directory.success_body": {
        "de": "Speicherpfad gesetzt auf:\n{path}",
        "en": "Save path set to:\n{path}",
        "es": "Ruta de guardado establecida en:\n{path}",
        "fr": "Chemin d'enregistrement défini sur :\n{path}",
        "sq": "Shtegu i ruajtjes u vendos në:\n{path}",
    },
    "change_save.option_title": {
        "de": "Speicheroption",
        "en": "Storage option",
        "es": "Opción de guardado",
        "fr": "Option de stockage",
        "sq": "Opsioni i ruajtjes",
    },
    "change_save.option_question": {
        "de": "Möchten Sie eine neue JSON-Datei erstellen oder eine bestehende Datei laden?",
        "en": "Do you want to create a new JSON file or load an existing one?",
        "es": "¿Quieres crear un archivo JSON nuevo o cargar uno existente?",
        "fr": "Souhaitez-vous créer un nouveau fichier JSON ou en charger un existant ?",
        "sq": "Dëshiron të krijosh një skedar të ri JSON apo të ngarkosh një ekzistues?",
    },
    "change_save.new_dialog_title": {
        "de": "Neuen Speicherort für die JSON-Datei wählen",
        "en": "Choose a new location for the JSON file",
        "es": "Elige una nueva ubicación para el archivo JSON",
        "fr": "Choisir un nouvel emplacement pour le fichier JSON",
        "sq": "Zgjidh një vendndodhje të re për skedarin JSON",
    },
    "change_save.existing_dialog_title": {
        "de": "Vorhandene JSON-Datei auswählen",
        "en": "Select an existing JSON file",
        "es": "Selecciona un archivo JSON existente",
        "fr": "Sélectionner un fichier JSON existant",
        "sq": "Zgjidh një skedar ekzistues JSON",
    },
    "change_save.no_change": {
        "de": "Keine Änderung vorgenommen.",
        "en": "No changes were made.",
        "es": "No se realizaron cambios.",
        "fr": "Aucun changement effectué.",
        "sq": "Asnjë ndryshim nuk u bë.",
    },
}


def translate_text(key: str, language: str, **kwargs) -> str:
    values = TRANSLATIONS.get(key)
    if not values:
        return key
    text = values.get(language) or values.get(DEFAULT_LANGUAGE) or next(iter(values.values()))
    if kwargs:
        try:
            return text.format(**kwargs)
        except (KeyError, ValueError):
            return text
    return text


def ensure_save_path_initialized(root: tk.Misc | None = None) -> None:
    """Ensure SAVE_FILE is set, prompting the user for a directory on first run."""
    global SAVE_FILE
    if SAVE_FILE:
        return

    config_exists = get_user_config_path().exists()
    env_override = os.environ.get("KLAUSURMASTER_SAVE_FILE")
    language_code = load_user_language()

    # Offer a directory picker the very first time the config is created (installer-like flow).
    if not config_exists and env_override is None and root is not None:
        root.withdraw()
        root.update_idletasks()
        wants_custom = messagebox.askyesno(
            translate_text("save_prompt.title", language_code),
            translate_text("save_prompt.body", language_code),
            parent=root,
        )
        if wants_custom:
            directory = filedialog.askdirectory(parent=root, title="Speicherordner wählen")
            if directory:
                filename = DEFAULT_SAVE_FILENAME
                selected_path = os.path.join(directory, filename)
                SAVE_FILE = persist_save_file_path(selected_path)
            else:
                messagebox.showinfo(
                    translate_text("save_prompt.default_title", language_code),
                    translate_text("save_prompt.default_body", language_code),
                    parent=root,
                )
        root.deiconify()

    if not SAVE_FILE:
        SAVE_FILE = str(load_save_file_path())


class CardApp:
    def __init__(self, root: tk.Tk):
        # Initializes widgets, state, and loads persisted data.
        self.root = root
        self.root.title(f"{APP_NAME} {APP_VERSION}")
        self.root.geometry("1400x820")
        self.root.minsize(960, 640)
        self.language = load_user_language()
        self.language_var = tk.StringVar(master=self.root, value=self.language)
        self._icon_images: list[tk.PhotoImage] = []
        self._set_window_icon()

        self.theme_name = "dark"
        self.data = {"tables": {}, "current_table": None}
        self.delete_mode = False
        self.moving_card: tuple[str, str, str] | None = None
        self.mark_mode = False
        self.selected_row_name: str | None = None
        self.tree_nodes: dict[tuple[str, ...], str] = {}
        self.tree_row_lookup: dict[str, tuple[str, ...]] = {}
        self.column_frames: list[tk.Frame] = []
        self.card_columns: dict[tuple[str, str], tk.Frame] = {}
        self.drag_data: dict | None = None
        self.drag_preview: tk.Toplevel | None = None
        self.current_highlighted_column: tk.Frame | None = None
        self.row_header_widgets: dict[str, tk.Widget] = {}
        self.card_weight_step = 10.0
        self.card_weight_min = 10.0
        self.card_weight_max = 200.0
        self.history: list[dict] = []
        self.future: list[dict] = []
        self.max_history = 7
        self.bin_history: list[dict] = []
        self.active_dialogs: list[tk.Toplevel] = []
        self.custom_theme = copy.deepcopy(THEMES.get("beige", THEMES["dark"]))
        self._ui_transition_depth = 0

        self.style = ttk.Style()
        self.primary_accent = PRIMARY_ACCENT
        self.configure_styles()
        self.build_main_layout()
        self.build_menu()
        self.update_file_path_label()

        self.bin_history = self.load_binary_history()
        self.load_data_or_create_new_table()
        self.bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self.handle_exit_request)

    def configure_styles(self):
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.apply_theme(self.theme_name, initial=True)

    def apply_theme(self, theme_name: str, initial: bool = False):
        with self.smooth_state_transition():
            theme = THEMES.get(theme_name, THEMES["dark"])
            self.theme_name = theme_name

            global APP_BG, PANEL_BG, CONTENT_BG, CARD_BG, CARD_BORDER, CARD_DELETE_BG, PRIMARY_ACCENT, TEXT_PRIMARY, TEXT_MUTED
            APP_BG = theme["APP_BG"]
            PANEL_BG = theme["PANEL_BG"]
            CONTENT_BG = theme["CONTENT_BG"]
            CARD_BG = theme["CARD_BG"]
            CARD_BORDER = theme["CARD_BORDER"]
            CARD_DELETE_BG = theme["CARD_DELETE_BG"]
            PRIMARY_ACCENT = theme["PRIMARY_ACCENT"]
            TEXT_PRIMARY = theme["TEXT_PRIMARY"]
            TEXT_MUTED = theme["TEXT_MUTED"]

            self.primary_accent = PRIMARY_ACCENT
            self.root.configure(bg=APP_BG)

            self.style.configure("App.TFrame", background=APP_BG)
            self.style.configure("Toolbar.TFrame", background=APP_BG)
            self.style.configure("Nav.TFrame", background=APP_BG)
            self.style.configure("Content.TFrame", background=APP_BG)
            self.style.configure("Controls.TFrame", background=APP_BG)
            self.style.configure("Board.TFrame", background=CONTENT_BG)
            self.style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), foreground=TEXT_PRIMARY, background=APP_BG)
            self.style.configure("Subtitle.TLabel", font=("Segoe UI", 11), foreground=TEXT_MUTED, background=APP_BG)
            self.style.configure("Placeholder.TFrame", background=CONTENT_BG)
            self.style.configure("PlaceholderTitle.TLabel", font=("Segoe UI", 16, "bold"), foreground=TEXT_PRIMARY, background=CONTENT_BG)
            self.style.configure("PlaceholderBody.TLabel", font=("Segoe UI", 11), foreground=TEXT_MUTED, background=CONTENT_BG)
            self.style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), foreground="#04111f", padding=(12, 6))
            self.style.map(
                "Primary.TButton",
                background=[("pressed", PRIMARY_ACCENT), ("active", PRIMARY_ACCENT), ("!disabled", PRIMARY_ACCENT)],
            )
            self.style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(10, 6), foreground=TEXT_PRIMARY)
            self.style.map(
                "Secondary.TButton",
                background=[("pressed", PANEL_BG), ("active", PANEL_BG), ("!disabled", PANEL_BG)],
                foreground=[("disabled", TEXT_MUTED), ("!disabled", TEXT_PRIMARY)],
            )

            self.style.configure(
                "Navigation.Treeview",
                rowheight=28,
                background=PANEL_BG,
                fieldbackground=PANEL_BG,
                foreground=TEXT_PRIMARY,
                borderwidth=0,
            )
            self.style.map(
                "Navigation.Treeview",
                background=[("selected", PRIMARY_ACCENT)],
                foreground=[("selected", "#ffffff")],
            )

            self.style.configure("TScrollbar", troughcolor=PANEL_BG, background=CARD_BORDER)
            self._style_file_path_label()

            if not initial:
                self.update_file_path_label()
                self.build_navigation_tree()
                self.update_table()

    def tr(self, key: str, **kwargs) -> str:
        return translate_text(key, self.language, **kwargs)

    def change_language(self, language_code: str) -> None:
        if language_code not in LANGUAGE_OPTIONS:
            return
        self.language = persist_user_language(language_code)
        if self.language_var.get() != language_code:
            self.language_var.set(language_code)
        self.build_menu()
        self.refresh_language_labels()
        self.update_file_path_label()

    def refresh_language_labels(self) -> None:
        if hasattr(self, "toolbar_subtitle_label"):
            self.toolbar_subtitle_label.config(text=self.tr("app.subtitle"))
        if hasattr(self, "nav_title_label"):
            self.nav_title_label.config(text=self.tr("nav.collections"))

    def _normalize_hex_color(self, value: str) -> str:
        candidate = (value or "").strip()
        if not candidate:
            raise ValueError("Farbwert darf nicht leer sein")
        if not candidate.startswith("#"):
            candidate = f"#{candidate}"
        if len(candidate) != 7:
            raise ValueError("Bitte hexadezimale Farben im Format #RRGGBB eingeben")
        try:
            int(candidate[1:], 16)
        except ValueError as exc:
            raise ValueError("Ungültiger Hexwert") from exc
        return candidate.lower()

    def show_custom_theme_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Eigenes Theme")
        dialog.resizable(False, False)
        self.register_dialog(dialog)

        current_values = {key: self.custom_theme.get(key, THEMES[self.theme_name].get(key, "#ffffff")) for key in THEME_KEYS}
        vars_by_key: dict[str, tk.StringVar] = {}

        ttk.Label(dialog, text="Wähle individuelle Farben für die Oberfläche.", style="Subtitle.TLabel").pack(padx=16, pady=(16, 8), anchor=tk.W)

        for key in THEME_KEYS:
            row = ttk.Frame(dialog)
            row.pack(fill=tk.X, padx=16, pady=4)
            label_text = key.replace("_", " ")
            ttk.Label(row, text=label_text).pack(side=tk.LEFT)
            var = tk.StringVar(value=current_values.get(key, "#ffffff"))
            vars_by_key[key] = var
            entry = ttk.Entry(row, textvariable=var, width=12)
            entry.pack(side=tk.LEFT, padx=(8, 8))

            def make_picker(target_key: str):
                def _pick_color():
                    initial = vars_by_key[target_key].get()
                    rgb, hex_value = colorchooser.askcolor(color=initial or "#ffffff", parent=dialog)
                    if hex_value:
                        vars_by_key[target_key].set(hex_value)
                return _pick_color

            ttk.Button(row, text="Farbe wählen", command=make_picker(key)).pack(side=tk.LEFT)

        status_var = tk.StringVar()
        ttk.Label(dialog, textvariable=status_var, foreground="red").pack(padx=16, pady=(8, 0), anchor=tk.W)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=16, pady=(16, 16))

        def apply_custom_theme(close_after: bool = False):
            new_theme = {}
            try:
                for key, var in vars_by_key.items():
                    new_theme[key] = self._normalize_hex_color(var.get())
            except ValueError as exc:
                status_var.set(str(exc))
                return

            self.custom_theme = new_theme
            THEMES["custom"] = new_theme
            status_var.set("Benutzerdefiniertes Theme aktiv.")
            self.apply_theme("custom")
            if close_after:
                self.close_dialog(dialog)

        ttk.Button(button_frame, text="Abbrechen", command=lambda: self.close_dialog(dialog)).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(button_frame, text="Anwenden", command=lambda: apply_custom_theme(False)).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(button_frame, text="Speichern & schließen", command=lambda: apply_custom_theme(True)).pack(side=tk.RIGHT)

    def _style_file_path_label(self):
        if hasattr(self, "file_path_label"):
            self.file_path_label.config(bg=PANEL_BG, fg=TEXT_MUTED)

    def build_main_layout(self):
        self.main_frame = ttk.Frame(self.root, padding=0, style="App.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.columnconfigure(1, weight=1)

        self._build_toolbar()
        self._build_navigation()
        self._build_content_area()

        self.file_path_label = tk.Label(
            self.root,
            text="",
            anchor="w",
            padx=12,
            pady=4,
        )
        self.file_path_label.pack(side=tk.BOTTOM, fill=tk.X)
        self._style_file_path_label()

    def build_menu(self):
        self.menubar = tk.Menu(self.root)

        file_menu = tk.Menu(self.menubar, tearoff=0)
        file_menu.add_command(label=self.tr("menu.file.save"), command=self.save_data)
        file_menu.add_command(label=self.tr("menu.file.reload"), command=self.load_data_or_create_new_table)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("menu.file.view_state"), command=self.view_history_state)
        file_menu.add_command(label=self.tr("menu.file.load_state"), command=self.load_history_state)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("menu.file.add_table"), command=self.add_row)
        file_menu.add_command(label=self.tr("menu.file.delete_table"), command=self.delete_row)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("menu.file.add_card"), command=self.add_card_via_button)
        file_menu.add_command(label=self.tr("menu.file.toggle_delete"), command=self.toggle_delete_mode)
        file_menu.add_command(label=self.tr("menu.file.cancel_action"), command=self.cancel_operations)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("menu.file.choose_folder"), command=self.choose_save_directory)
        file_menu.add_command(label=self.tr("menu.file.change_file"), command=self.change_save_location)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("menu.file.exit"), command=self.handle_exit_request)

        edit_menu = tk.Menu(self.menubar, tearoff=0)
        edit_menu.add_command(label=self.tr("menu.edit.undo"), command=self.undo_action)
        edit_menu.add_command(label=self.tr("menu.edit.redo"), command=self.redo_action)

        space_menu = tk.Menu(self.menubar, tearoff=0)
        space_menu.add_command(label=self.tr("menu.space.new"), command=self.new_table)
        space_menu.add_command(label=self.tr("menu.space.load"), command=self.load_table_via_menu)
        space_menu.add_separator()
        space_menu.add_command(label=self.tr("menu.space.import_space"), command=self.import_table_from_file)
        space_menu.add_command(label=self.tr("menu.space.export_space"), command=self.export_current_table)
        space_menu.add_command(label=self.tr("menu.space.import_cards"), command=self.import_cards_via_text)
        space_menu.add_command(label=self.tr("menu.space.import_short"), command=self.import_cards_via_text)
        space_menu.add_separator()
        space_menu.add_command(label=self.tr("menu.space.duplicate"), command=self.duplicate_current_table)
        space_menu.add_command(label=self.tr("menu.space.transfer"), command=self.transfer_table_between_spaces)
        space_menu.add_command(label=self.tr("menu.space.delete"), command=self.delete_table_via_menu)
        space_menu.add_separator()
        theme_menu = tk.Menu(space_menu, tearoff=0)
        theme_menu.add_command(label=self.tr("menu.space.theme.dark"), command=lambda: self.apply_theme("dark"))
        theme_menu.add_command(label=self.tr("menu.space.theme.beige"), command=lambda: self.apply_theme("beige"))
        theme_menu.add_separator()
        theme_menu.add_command(label=self.tr("menu.space.theme.custom"), command=self.show_custom_theme_dialog)
        space_menu.add_cascade(label=self.tr("menu.space.theme"), menu=theme_menu)

        help_menu = tk.Menu(self.menubar, tearoff=0)
        help_menu.add_command(label=self.tr("menu.help.about"), command=self.show_about_dialog)

        language_menu = tk.Menu(self.menubar, tearoff=0)
        for code, label in LANGUAGE_OPTIONS.items():
            language_menu.add_radiobutton(
                label=label,
                value=code,
                variable=self.language_var,
                command=lambda c=code: self.change_language(c),
            )

        actions_menu = tk.Menu(self.menubar, tearoff=0)
        for action in self._get_actions_menu_items():
            accelerator = action.get("accelerator") or ""
            actions_menu.add_command(
                label=self.tr(action["label_key"]),
                accelerator=accelerator,
                command=action["command"],
            )

        self.menubar.add_cascade(label=self.tr("menu.file"), menu=file_menu)
        self.menubar.add_cascade(label=self.tr("menu.edit"), menu=edit_menu)
        self.menubar.add_cascade(label=self.tr("menu.space"), menu=space_menu)
        self.menubar.add_cascade(label=self.tr("menu.actions"), menu=actions_menu)
        self.menubar.add_cascade(label=self.tr("menu.help"), menu=help_menu)
        self.menubar.add_cascade(label=self.tr("menu.language"), menu=language_menu)
        self.root.config(menu=self.menubar)

    def _get_actions_menu_items(self) -> list[dict[str, object]]:
        return [
            {"label_key": "menu.file.save", "command": self.save_data},
            {"label_key": "menu.file.reload", "command": self.load_data_or_create_new_table},
            {"label_key": "menu.file.view_state", "command": self.view_history_state},
            {"label_key": "menu.file.load_state", "command": self.load_history_state},
            {"label_key": "menu.file.add_table", "command": self.add_row, "accelerator": "Ctrl+N"},
            {"label_key": "menu.file.delete_table", "command": self.delete_row},
            {"label_key": "menu.file.add_card", "command": self.add_card_via_button, "accelerator": "1–9"},
            {"label_key": "menu.file.toggle_delete", "command": self.toggle_delete_mode, "accelerator": "Ctrl+W"},
            {"label_key": "action.toggle_mark_mode", "command": self.toggle_mark_mode, "accelerator": "Ctrl+M"},
            {"label_key": "menu.file.cancel_action", "command": self.cancel_operations, "accelerator": "Esc"},
            {"label_key": "menu.file.choose_folder", "command": self.choose_save_directory},
            {"label_key": "menu.file.change_file", "command": self.change_save_location},
            {"label_key": "menu.edit.undo", "command": self.undo_action, "accelerator": "Ctrl+Z"},
            {"label_key": "menu.edit.redo", "command": self.redo_action, "accelerator": "Ctrl+Shift+Z / Ctrl+Y"},
            {"label_key": "menu.space.new", "command": self.new_table, "accelerator": "Ctrl+S"},
            {"label_key": "menu.space.load", "command": self.load_table_via_menu},
            {"label_key": "menu.space.import_space", "command": self.import_table_from_file},
            {"label_key": "menu.space.export_space", "command": self.export_current_table},
            {"label_key": "menu.space.import_cards", "command": self.import_cards_via_text},
            {"label_key": "menu.space.duplicate", "command": self.duplicate_current_table},
            {"label_key": "menu.space.transfer", "command": self.transfer_table_between_spaces},
            {"label_key": "menu.space.delete", "command": self.delete_table_via_menu},
            {"label_key": "menu.help.about", "command": self.show_about_dialog},
        ]

    def _build_toolbar(self):
        self.toolbar = ttk.Frame(self.main_frame, padding=(24, 16), style="Toolbar.TFrame")
        self.toolbar.grid(row=0, column=0, columnspan=2, sticky=tk.EW)
        self.toolbar.columnconfigure(0, weight=1)

        title_wrapper = ttk.Frame(self.toolbar, style="Toolbar.TFrame")
        title_wrapper.grid(row=0, column=0, sticky=tk.W)
        self.toolbar_title_label = ttk.Label(title_wrapper, text=f"{APP_NAME} {APP_VERSION}", style="Title.TLabel")
        self.toolbar_title_label.pack(anchor=tk.W)
        self.toolbar_subtitle_label = ttk.Label(title_wrapper, text=self.tr("app.subtitle"), style="Subtitle.TLabel")
        self.toolbar_subtitle_label.pack(anchor=tk.W)

    def _build_navigation(self):
        self.nav_frame = ttk.Frame(self.main_frame, padding=(24, 16, 12, 24), style="Nav.TFrame")
        self.nav_frame.grid(row=1, column=0, sticky=tk.NS)
        self.nav_frame.rowconfigure(1, weight=1)

        self.nav_title_label = ttk.Label(self.nav_frame, text=self.tr("nav.collections"), style="Subtitle.TLabel")
        self.nav_title_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 6))

        self.navigation_tree = ttk.Treeview(self.nav_frame, style="Navigation.Treeview", show="tree")
        self.navigation_tree.grid(row=1, column=0, sticky=tk.NS)
        nav_scroll = ttk.Scrollbar(self.nav_frame, orient=tk.VERTICAL, command=self.navigation_tree.yview)
        nav_scroll.grid(row=1, column=1, sticky=tk.NS, padx=(8, 0))
        self.navigation_tree.configure(yscrollcommand=nav_scroll.set)
        self.navigation_tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    def _build_content_area(self):
        self.content_frame = ttk.Frame(self.main_frame, padding=(0, 16, 24, 24), style="Content.TFrame")
        self.content_frame.grid(row=1, column=1, sticky=tk.NSEW)
        self.content_frame.rowconfigure(0, weight=1)
        self.content_frame.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.content_frame, bg=CONTENT_BG, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky=tk.NSEW)
        self.v_scrollbar = ttk.Scrollbar(self.content_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.h_scrollbar = ttk.Scrollbar(self.content_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky=tk.EW, pady=(8, 0))
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)

        self.table = ttk.Frame(self.canvas, style="Content.TFrame")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.table, anchor="nw")
        self.table.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

    def on_frame_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)

    def update_file_path_label(self):
        global SAVE_FILE
        self.file_path_label.config(text=self.tr("status.current_path", path=SAVE_FILE))

    def _set_window_icon(self):
        icon_candidates = ["favicon.ico", "favicon.png"]
        for candidate in icon_candidates:
            icon_path = resource_path(candidate)
            if not icon_path.exists():
                continue
            try:
                if candidate.lower().endswith(".ico") and sys.platform.startswith("win"):
                    self.root.iconbitmap(str(icon_path))
                    return
                icon_image = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, icon_image)
                self._icon_images.append(icon_image)
                return
            except tk.TclError:
                continue

    def load_data_or_create_new_table(self):
        self.load_data()
        self.convert_old_cards_format()
        self._reset_history()

        if not self.data.get("tables"):
            self.data["tables"] = {}
        if "current_table" not in self.data:
            self.data["current_table"] = None

        if not self.data["tables"]:
            self.new_table(skip_history=True)

        if self.data["current_table"] is None and self.data["tables"]:
            self.data["current_table"] = list(self.data["tables"].keys())[0]

        self.update_file_path_label()
        self.ensure_row_selection_valid()
        self.build_navigation_tree()
        self.update_table()

    def convert_old_cards_format(self):
        tables = self.data.get("tables", {})
        for table_data in tables.values():
            cards = table_data.get("cards", {})
            normalize_cards_tree(cards)

    def _reset_history(self):
        self.history.clear()
        self.future.clear()

    def _bring_dialog_to_front(self):
        try:
            windows = [self.root] + self.active_dialogs
            for win in windows:
                win.lift()
                win.attributes("-topmost", True)
            def release_topmost(targets):
                for win in targets:
                    try:
                        win.attributes("-topmost", False)
                    except tk.TclError:
                        continue
            self.root.after(200, lambda: release_topmost(list(windows)))
        except tk.TclError:
            pass

    def register_dialog(self, dialog: tk.Toplevel):
        if dialog not in self.active_dialogs:
            self.active_dialogs.append(dialog)
        dialog.protocol("WM_DELETE_WINDOW", lambda d=dialog: self.close_dialog(d))
        dialog.transient(self.root)
        dialog.grab_set()
        self._bring_dialog_to_front()

    def close_dialog(self, dialog: tk.Toplevel):
        try:
            dialog.grab_release()
        except tk.TclError:
            pass
        if dialog in self.active_dialogs:
            self.active_dialogs.remove(dialog)
        dialog.destroy()

    @contextmanager
    def smooth_state_transition(self):
        self._ui_transition_depth += 1
        if self._ui_transition_depth == 1:
            try:
                self.root.configure(cursor="watch")
            except tk.TclError:
                pass
            self.root.update_idletasks()
        try:
            yield
        finally:
            self._ui_transition_depth = max(0, self._ui_transition_depth - 1)
            if self._ui_transition_depth == 0:
                try:
                    self.root.configure(cursor="")
                except tk.TclError:
                    pass
                self.root.after_idle(self.root.update_idletasks)

    def _capture_history_state(self) -> dict:
        return {
            "data": copy.deepcopy(self.data),
            "selected_row": self.selected_row_name,
        }

    def _record_history(self):
        snapshot = self._capture_history_state()
        self.history.append(snapshot)
        if len(self.history) > self.max_history:
            self.history.pop(0)
        self.future.clear()

    def _apply_history_state(self, state: dict):
        self.data = copy.deepcopy(state.get("data", {}))
        self.selected_row_name = state.get("selected_row")
        self.ensure_row_selection_valid()
        self.build_navigation_tree()
        self.update_table()
        self.save_data()

    def undo_action(self, event=None):
        if not self.history:
            return "break"
        self.future.append(self._capture_history_state())
        if len(self.future) > self.max_history:
            self.future.pop(0)
        state = self.history.pop()
        self._apply_history_state(state)
        return "break"

    def redo_action(self, event=None):
        if not self.future:
            return "break"
        self.history.append(self._capture_history_state())
        if len(self.history) > self.max_history:
            self.history.pop(0)
        state = self.future.pop()
        self._apply_history_state(state)
        return "break"

    def shortcut_new_space(self, event=None):
        self.new_table()
        return "break"

    def get_bin_file_path(self) -> str:
        directory = os.path.dirname(SAVE_FILE)
        if not directory:
            directory = "."
        return os.path.join(directory, BIN_HISTORY_FILENAME)

    def load_binary_history(self) -> list[dict]:
        path = self.get_bin_file_path()
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except FileNotFoundError:
            return []
        except json.JSONDecodeError:
            return []
        if isinstance(payload, dict):
            history = payload.get("history", [])
            if isinstance(history, list):
                return history[-100:]
        return []

    def save_binary_history(self):
        path = self.get_bin_file_path()
        directory = os.path.dirname(path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        payload = {"history": self.bin_history[-100:]}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)

    def add_snapshot_to_bin(self):
        timestamp = datetime.now()
        entry = {
            "timestamp": timestamp.isoformat(),
            "label": timestamp.strftime("%d.%m.%Y %H:%M"),
            "data": copy.deepcopy(self.data),
        }
        self.bin_history.append(entry)
        if len(self.bin_history) > 100:
            self.bin_history = self.bin_history[-100:]

    def refresh_bin_history(self):
        self.bin_history = self.load_binary_history()

    def handle_exit_request(self):
        try:
            self.save_data()
            self.add_snapshot_to_bin()
            self.save_binary_history()
        except Exception as exc:
            messagebox.showerror("Fehler", f"Konnte Versionierung nicht speichern: {exc}")
        finally:
            self.root.destroy()

    def _pick_history_entry(self, action: str) -> dict | None:
        if not self.bin_history:
            messagebox.showinfo("Keine Versionen", "Es sind noch keine gespeicherten Versionen vorhanden.")
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Version {action}")
        dialog.transient(self.root)
        dialog.grab_set()
        ttk.Label(dialog, text="Wähle einen gespeicherten Stand:").pack(padx=16, pady=(16, 8))

        listbox = tk.Listbox(dialog, width=40, height=10)
        listbox.pack(padx=16, pady=8, fill=tk.BOTH, expand=True)
        entries = list(reversed(self.bin_history))
        for entry in entries:
            listbox.insert(tk.END, entry.get("label", "Unbekannt"))

        selection = {"value": None}

        def confirm():
            idxs = listbox.curselection()
            if not idxs:
                messagebox.showerror("Auswahl erforderlich", "Bitte einen Eintrag auswählen.")
                return
            selection["value"] = entries[idxs[0]]
            dialog.destroy()

        ttk.Button(dialog, text="Auswählen", command=confirm).pack(pady=(0, 12))
        dialog.bind("<Return>", lambda _event: confirm())
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        listbox.focus_set()
        self.root.wait_window(dialog)
        return selection["value"]

    def view_history_state(self):
        entry = self._pick_history_entry("anzeigen")
        if not entry:
            return
        viewer = tk.Toplevel(self.root)
        viewer.title(f"Stand vom {entry.get('label','')}")
        viewer.transient(self.root)
        viewer.grab_set()
        text = tk.Text(viewer, wrap="word")
        text.pack(fill=tk.BOTH, expand=True)
        text.insert("1.0", json.dumps(entry.get("data", {}), indent=2, ensure_ascii=False))
        text.config(state="disabled")
        ttk.Button(viewer, text="Schließen", command=viewer.destroy).pack(pady=8)

    def load_history_state(self):
        entry = self._pick_history_entry("laden")
        if not entry:
            return
        confirm = messagebox.askyesno(
            "Version laden",
            f"Soll der Stand vom {entry.get('label','')} geladen werden? Aktuelle Daten werden überschrieben.",
        )
        if not confirm:
            return
        self.data = copy.deepcopy(entry.get("data", {})) or {"tables": {}, "current_table": None}
        self._reset_history()
        self.save_data()
        self.ensure_row_selection_valid()
        self.build_navigation_tree()
        self.update_table()

    def build_navigation_tree(self):
        with self.smooth_state_transition():
            self._build_navigation_tree_impl()

    def _build_navigation_tree_impl(self):
        if not hasattr(self, "navigation_tree"):
            return

        self.navigation_tree.delete(*self.navigation_tree.get_children())
        self.tree_nodes.clear()
        self.tree_row_lookup.clear()

        current_table = self.data.get("current_table")
        selected_row = self.selected_row_name

        space_averages: dict[str, str] = {}
        for table_name, table_data in self.data.get("tables", {}).items():
            rows = table_data.get("rows", [])
            if not rows:
                space_averages[table_name] = "—"
                continue
            columns = table_data.get("columns", [])
            cards = table_data.get("cards", {})
            values = []
            for row_name in rows:
                ratio = calculate_ratio(row_name, cards, columns)
                expected_grade = calculate_expected_grade(row_name, cards, columns)
                if ratio is None or expected_grade is None:
                    continue
                values.append(expected_grade)
            space_averages[table_name] = "—" if not values else f"{sum(values) / len(values):.2f}"

        for table_index, (table_name, table_data) in enumerate(self.data.get("tables", {}).items(), start=1):
            avg = space_averages.get(table_name, "—")
            table_label = f"{table_index}. {table_name} · xG {avg}"
            table_id = self.navigation_tree.insert("", "end", text=table_label, open=(table_name == current_table))
            self.tree_nodes[("table", table_name)] = table_id
            self.tree_row_lookup[table_id] = ("table", table_name)
            for row_index, row_name in enumerate(table_data.get("rows", []), start=1):
                row_label = f"{row_name} ({row_index})"
                row_id = self.navigation_tree.insert(table_id, "end", text=row_label)
                self.tree_nodes[("row", table_name, row_name)] = row_id
                self.tree_row_lookup[row_id] = ("row", table_name, row_name)

        target_id = None
        if current_table and selected_row:
            target_id = self.tree_nodes.get(("row", current_table, selected_row))
        if not target_id and current_table:
            target_id = self.tree_nodes.get(("table", current_table))
        if target_id:
            self.navigation_tree.selection_set(target_id)
            self.navigation_tree.see(target_id)

    def ensure_row_selection_valid(self):
        table_data = self.get_current_table_data()
        if table_data is None:
            self.selected_row_name = None
            return
        if not table_data["rows"]:
            self.selected_row_name = None
            return
        if self.selected_row_name not in table_data["rows"]:
            self.selected_row_name = None

    def get_current_table_data(self):
        if self.data["current_table"] is None:
            return None
        self.root.title(f"Klausurmaster2D - {self.data['current_table']}")
        return self.data["tables"].get(self.data["current_table"], None)

    def get_current_columns(self):
        table_data = self.get_current_table_data()
        if table_data is None:
            return []
        return table_data["columns"]

    def render_placeholder(self, title: str, subtitle: str):
        self.card_columns = {}
        self.row_header_widgets = {}
        for widget in self.table.winfo_children():
            widget.destroy()

        placeholder = ttk.Frame(self.table, style="Placeholder.TFrame", padding=60)
        placeholder.grid(row=0, column=0, sticky=tk.NSEW, padx=80, pady=80)
        self.table.columnconfigure(0, weight=1)
        self.table.rowconfigure(0, weight=1)
        ttk.Label(placeholder, text=title, style="PlaceholderTitle.TLabel").pack(anchor=tk.CENTER)
        ttk.Label(placeholder, text=subtitle, style="PlaceholderBody.TLabel", wraplength=520, justify="center").pack(anchor=tk.CENTER, pady=(12, 0))

    def _get_contrast_color(self, hex_color: str | None) -> str:
        if not hex_color:
            return TEXT_PRIMARY
        try:
            cleaned = hex_color.lstrip("#")
            r = int(cleaned[0:2], 16)
            g = int(cleaned[2:4], 16)
            b = int(cleaned[4:6], 16)
        except (ValueError, IndexError, TypeError):
            return TEXT_PRIMARY
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return "#000000" if luminance > 0.6 else "#ffffff"

    def render_row_header(self, row_name: str, ratio: float, expected_grade: float, row_color: str, total_cards: int):
        header_color = row_color or PANEL_BG
        header = tk.Frame(self.table, bg=header_color, padx=18, pady=18, highlightthickness=0)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 18))
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=0)

        title = tk.Label(header, text=row_name, font=("Segoe UI", 18, "bold"), bg=header_color, fg=TEXT_PRIMARY)
        title.grid(row=0, column=0, sticky=tk.W)
        expected = "—" if total_cards == 0 else f"{expected_grade:.2f}"
        subtitle = tk.Label(
            header,
            text=f"{total_cards} Karten · xG {expected}",
            font=("Segoe UI", 11),
            bg=header_color,
            fg=TEXT_PRIMARY,
        )
        subtitle.grid(row=1, column=0, sticky=tk.W, pady=(8, 0))

        button_bg = row_color or PANEL_BG
        contrast_fg = self._get_contrast_color(button_bg)
        color_button = tk.Button(
            header,
            text="Farbe",
            font=("Segoe UI", 9, "bold"),
            command=lambda rn=row_name: self.pick_row_color(rn),
            bg=button_bg,
            fg=contrast_fg,
            activebackground=button_bg,
            activeforeground=contrast_fg,
            relief="flat",
            bd=0,
            padx=12,
            pady=8,
        )
        color_button.grid(row=0, column=1, rowspan=2, sticky=tk.NE, padx=(12, 0))

        self.row_header_widgets = {
            "row_name": row_name,
            "frame": header,
            "title": title,
            "subtitle": subtitle,
            "color_button": color_button,
        }

    def render_board_for_row(self, row_name: str, columns: list[str], cards: dict, row_color: str):
        self._clear_column_highlight()
        board = ttk.Frame(self.table, style="Board.TFrame", padding=(0, 0, 0, 40))
        board.grid(row=1, column=0, sticky=tk.NSEW)
        board.rowconfigure(0, weight=1)
        self.column_frames = []
        self.card_columns = {}

        for idx, col_name in enumerate(columns):
            board.columnconfigure(idx, weight=1)
            column_frame = tk.Frame(
                board,
                bg=PANEL_BG,
                padx=16,
                pady=16,
                highlightbackground=row_color,
                highlightthickness=0,
            )
            column_frame.grid(row=0, column=idx, sticky=tk.NSEW, padx=10)
            column_frame._row_name = row_name
            column_frame._col_name = col_name
            column_frame._base_highlight_color = row_color
            column_frame.bind("<ButtonPress-1>", lambda event, rn=row_name, cn=col_name: self.on_column_click(rn, cn))
            self.column_frames.append(column_frame)

            header = tk.Label(column_frame, text=col_name, font=("Segoe UI", 12, "bold"), bg=PANEL_BG, fg=TEXT_PRIMARY)
            header.pack(fill=tk.X, pady=(0, 12))
            header.bind("<ButtonPress-1>", lambda event, rn=row_name, cn=col_name: self.on_column_click(rn, cn))

            cards_container = tk.Frame(column_frame, bg=PANEL_BG)
            cards_container.pack(fill=tk.BOTH, expand=True)
            self.card_columns[(row_name, col_name)] = cards_container
            self.render_cards_in_column(row_name, col_name, cards_container, cards)

    def _bind_card_widget(self, widget: tk.Widget, row_name: str, col_name: str, card_front: str):
        widget.bind(
            "<ButtonPress-1>",
            lambda event, rn=row_name, cn=col_name, cf=card_front: self.on_card_press(event, rn, cn, cf),
        )
        widget.bind("<B1-Motion>", self.on_card_motion)
        widget.bind(
            "<ButtonRelease-1>",
            lambda event, rn=row_name, cn=col_name, cf=card_front: self.on_card_release(event, rn, cn, cf),
        )
        widget.bind(
            "<Button-3>",
            lambda event, rn=row_name, cn=col_name, cf=card_front: self.on_card_right_click(rn, cn, cf),
        )

    def render_cards_in_column(self, row_name: str, col_name: str, container: tk.Frame, cards: dict):
        for widget in container.winfo_children():
            widget.destroy()

        if row_name not in cards or col_name not in cards[row_name]:
            return

        for card_dict in cards[row_name][col_name]:
            card_front = card_dict["front"]
            base_bg = CARD_DELETE_BG if self.delete_mode else CARD_BG
            highlight_color = self.primary_accent if card_dict["marked"] else CARD_BORDER
            thickness = 2 if (card_dict["marked"] or self.delete_mode) else 1
            card_frame = tk.Frame(
                container,
                bg=base_bg,
                padx=12,
                pady=10,
                bd=0,
                highlightbackground=highlight_color,
                highlightthickness=thickness,
            )
            card_frame.pack(fill=tk.X, pady=6)
            self._bind_card_widget(card_frame, row_name, col_name, card_front)
            card_frame._is_card_frame = True
            card_frame._card_payload = card_dict

            header_bar = tk.Frame(card_frame, bg=base_bg)
            header_bar.pack(fill=tk.X)
            self._bind_card_widget(header_bar, row_name, col_name, card_front)

            title = tk.Label(
                header_bar,
                text=card_front,
                font=("Segoe UI", 11, "bold"),
                bg=base_bg,
                fg=TEXT_PRIMARY,
                wraplength=220,
                justify=tk.LEFT,
            )
            title.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self._bind_card_widget(title, row_name, col_name, card_front)

            has_back_text = bool(card_dict.get("back", "").strip())
            note_bg = self.primary_accent if has_back_text else CARD_BORDER
            note_fg = "#ffffff" if has_back_text else TEXT_PRIMARY
            note_button = tk.Button(
                header_bar,
                text="📝",
                width=2,
                bg=note_bg,
                fg=note_fg,
                bd=0,
                relief="flat",
                activebackground=self.primary_accent,
                activeforeground="#ffffff",
                command=lambda rn=row_name, cn=col_name, cf=card_front: self._open_card_back_editor(rn, cn, cf),
            )
            note_button.pack(side=tk.RIGHT, padx=(8, 0))

            if card_dict.get("back"):
                display_text = card_dict["back"]
                if len(display_text) > 120:
                    display_text = display_text[:120] + "…"
                snippet = tk.Label(
                    card_frame,
                    text=display_text,
                    font=("Segoe UI", 9),
                    bg=base_bg,
                    fg=TEXT_MUTED,
                    justify=tk.LEFT,
                    wraplength=220,
                )
                snippet.pack(fill=tk.X, pady=(6, 0))
                self._bind_card_widget(snippet, row_name, col_name, card_front)

            weight = self._extract_card_weight(card_dict)
            weight_controls = tk.Frame(card_frame, bg=base_bg)
            weight_controls.pack(fill=tk.X, pady=(10, 0))
            control_bg = CARD_BORDER

            minus_btn = tk.Button(
                weight_controls,
                text="<",
                width=2,
                bg=control_bg,
                fg=TEXT_PRIMARY,
                activebackground=self.primary_accent,
                activeforeground=TEXT_PRIMARY,
                relief="flat",
                bd=0,
                command=lambda rn=row_name, cn=col_name, cf=card_front: self.adjust_card_weight(rn, cn, cf, -self.card_weight_step),
            )
            minus_btn.pack(side=tk.LEFT)

            weight_label = tk.Label(
                weight_controls,
                text=f"Gewicht: {weight:.0f}",
                bg=base_bg,
                fg=TEXT_MUTED,
                font=("Segoe UI", 9, "bold"),
            )
            weight_label.pack(side=tk.LEFT, expand=True, padx=6)

            plus_btn = tk.Button(
                weight_controls,
                text=">",
                width=2,
                bg=control_bg,
                fg=TEXT_PRIMARY,
                activebackground=self.primary_accent,
                activeforeground=TEXT_PRIMARY,
                relief="flat",
                bd=0,
                command=lambda rn=row_name, cn=col_name, cf=card_front: self.adjust_card_weight(rn, cn, cf, self.card_weight_step),
            )
            plus_btn.pack(side=tk.RIGHT)

            weight_bar = tk.Frame(card_frame, bg=base_bg)
            weight_bar.pack(fill=tk.X, pady=(4, 0))
            weight_canvas = tk.Canvas(weight_bar, height=6, bg=base_bg, highlightthickness=0)
            weight_canvas.pack(fill=tk.X)
            width = 150
            weight_canvas.configure(width=width)
            weight_canvas.create_rectangle(0, 0, width, 6, fill=CARD_BORDER, outline="")
            normalized = (weight - self.card_weight_min) / (self.card_weight_max - self.card_weight_min)
            normalized = max(0.0, min(1.0, normalized))
            weight_canvas.create_rectangle(0, 0, int(width * normalized), 6, fill=self.primary_accent, outline="")

    def refresh_card_column(self, row_name: str, col_name: str) -> bool:
        container = self.card_columns.get((row_name, col_name))
        table_data = self.get_current_table_data()
        if container is None or table_data is None:
            return False
        cards = table_data["cards"]
        self.render_cards_in_column(row_name, col_name, container, cards)
        return True

    def update_row_header_info(self, row_name: str) -> bool:
        header_info = getattr(self, "row_header_widgets", None)
        if not header_info or header_info.get("row_name") != row_name:
            return False
        table_data = self.get_current_table_data()
        if table_data is None:
            return False

        cards = table_data["cards"]
        columns = table_data["columns"]
        row_colors = table_data["row_colors"]
        if row_name not in cards:
            return False

        total_cards = 0
        for col in columns:
            if col in cards[row_name]:
                total_cards += len(cards[row_name][col])

        row_color = row_colors.get(row_name, self.primary_accent)
        ratio = calculate_ratio(row_name, cards, columns)
        expected_grade = calculate_expected_grade(row_name, cards, columns)
        header_color = row_color or PANEL_BG

        header = header_info.get("frame")
        title = header_info.get("title")
        subtitle = header_info.get("subtitle")
        color_button = header_info.get("color_button")
        if not header or not title or not subtitle:
            return False

        header.configure(bg=header_color)
        title.configure(bg=header_color)
        expected = "—" if total_cards == 0 else f"{expected_grade:.2f}"
        subtitle.configure(bg=header_color, text=f"{total_cards} Karten · xG {expected}")
        for child in header.winfo_children():
            try:
                child.configure(bg=header_color)
            except tk.TclError:
                continue
        if color_button:
            button_bg = row_color or PANEL_BG
            contrast_fg = self._get_contrast_color(button_bg)
            color_button.configure(
                bg=button_bg,
                fg=contrast_fg,
                activebackground=button_bg,
                activeforeground=contrast_fg,
            )
        return True

    def pick_row_color(self, row_name: str):
        table_data = self.get_current_table_data()
        if table_data is None:
            return
        row_colors = table_data.get("row_colors", {})
        if row_name not in row_colors:
            return

        current_color = row_colors.get(row_name, self.primary_accent)
        chosen = colorchooser.askcolor(color=current_color, title="Tabellenfarbe wählen")
        if not chosen or not chosen[1]:
            return
        new_hex = chosen[1]
        if not new_hex or new_hex == current_color:
            return

        self._record_history()
        row_colors[row_name] = new_hex
        self.save_data()
        if not self.update_row_header_info(row_name):
            self.update_table()

    def _extract_card_weight(self, card_dict: dict) -> float:
        try:
            return float(card_dict.get("weight", 100.0))
        except (TypeError, ValueError):
            return 100.0

    def adjust_card_weight(self, row_name: str, col_name: str, card_front: str, delta: float):
        table_data = self.get_current_table_data()
        if table_data is None:
            return
        cards = table_data["cards"]
        if row_name not in cards or col_name not in cards[row_name]:
            return
        card_dict = find_card(cards[row_name][col_name], card_front)
        if not card_dict:
            return
        self._record_history()
        current_weight = self._extract_card_weight(card_dict)
        new_weight = max(self.card_weight_min, min(self.card_weight_max, current_weight + delta))
        card_dict["weight"] = round(new_weight, 2)
        self.save_data()
        refreshed = self.refresh_card_column(row_name, col_name)
        header_updated = True
        if self.selected_row_name == row_name:
            header_updated = self.update_row_header_info(row_name)
        if not refreshed or (self.selected_row_name == row_name and not header_updated):
            self.update_table()

    def _create_row_entry(self, row_name: str, preferred_color: str | None = None) -> tuple[bool, str | None]:
        table_data = self.get_current_table_data()
        if table_data is None:
            return False, "Kein Space geladen!"

        row_name = row_name.strip()
        if not row_name:
            return False, "Name darf nicht leer sein"

        rows = table_data["rows"]
        if len(rows) >= 9:
            return False, "Maximal 9 Tabellen sind erlaubt!"
        if row_name in rows:
            return False, "Tabelle existiert bereits"

        cards = table_data["cards"]
        columns = table_data["columns"]
        row_colors = table_data["row_colors"]

        self._record_history()
        rows.append(row_name)
        cards[row_name] = {col: [] for col in columns}
        row_colors[row_name] = preferred_color or random_pastel_color()
        self.selected_row_name = row_name
        self.save_data()
        self.ensure_row_selection_valid()
        self.build_navigation_tree()
        self.update_table()
        return True, None

    def add_row(self):
        table_data = self.get_current_table_data()
        if table_data is None:
            messagebox.showerror("Fehler", "Kein Space geladen!")
            return

        rows = table_data["rows"]
        if len(rows) >= 9:
            messagebox.showerror("Fehler", "Maximal 9 Tabellen sind erlaubt!")
            return

        self._bring_dialog_to_front()
        dialog = tk.Toplevel(self.root)
        dialog.title("Neue Tabelle")
        dialog.resizable(False, False)
        self.register_dialog(dialog)

        name_var = tk.StringVar()
        preview_var = tk.StringVar()
        color_preview = tk.StringVar()
        color_value = tk.StringVar(value=random_pastel_color())

        color_chip = tk.Canvas(dialog, width=120, height=24, highlightthickness=0, bd=0)

        def update_preview(*_args):
            # Keep preview text and color swatch in sync with user input.
            row_name = name_var.get().strip() or "Neue Tabelle"
            preview_var.set(f"Name: {row_name}")
            color_preview.set(f"Vorschlagsfarbe: {color_value.get()}")
            color_chip.configure(bg=color_value.get())

        def shuffle_color():
            color_value.set(random_pastel_color())
            update_preview()

        ttk.Label(dialog, text="Tabellenname").pack(padx=16, pady=(16, 4), anchor=tk.W)
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(padx=16, fill=tk.X)
        ttk.Label(dialog, textvariable=preview_var).pack(padx=16, pady=(10, 0), anchor=tk.W)
        ttk.Label(dialog, textvariable=color_preview, foreground=TEXT_MUTED).pack(padx=16, pady=(0, 4), anchor=tk.W)
        color_chip.pack(padx=16, pady=(0, 8), anchor=tk.W, fill=tk.X)
        ttk.Button(dialog, text="Farbe neu würfeln", command=shuffle_color).pack(padx=16, pady=(0, 8), anchor=tk.W)
        error_var = tk.StringVar()
        ttk.Label(dialog, textvariable=error_var, foreground="red").pack(padx=16, pady=(0, 8), anchor=tk.W)

        def submit():
            success, error_message = self._create_row_entry(name_var.get(), color_value.get())
            if not success:
                error_var.set(error_message or "Unbekannter Fehler")
                return
            self.close_dialog(dialog)

        ttk.Button(dialog, text="Anlegen", command=submit).pack(padx=16, pady=(0, 16))
        entry.focus_set()
        dialog.bind("<Return>", lambda event: submit())
        name_var.trace_add("write", update_preview)
        update_preview()

    def delete_row(self):
        table_data = self.get_current_table_data()
        if table_data is None:
            messagebox.showerror("Fehler", "Kein Space geladen!")
            return

        rows = table_data["rows"]
        cards = table_data["cards"]
        row_colors = table_data["row_colors"]

        if not rows:
            messagebox.showerror("Fehler", "Keine Tabelle vorhanden, die gelöscht werden könnte.")
            return

        row_name = self._prompt_row_name("Tabelle löschen", "Welche Tabelle möchtest du löschen?", rows)
        if row_name and row_name in rows:
            self._record_history()
            rows.remove(row_name)
            cards.pop(row_name, None)
            row_colors.pop(row_name, None)
            if self.selected_row_name == row_name:
                self.selected_row_name = None
            self.save_data()
            self.ensure_row_selection_valid()
            self.build_navigation_tree()
            self.update_table()
        else:
            messagebox.showerror("Fehler", "Tabelle nicht gefunden.")

    def _prompt_row_name(self, title: str, prompt: str, options: list[str]) -> str | None:
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.resizable(False, False)
        self.register_dialog(dialog)

        ttk.Label(dialog, text=prompt).pack(padx=16, pady=(16, 8))
        name_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=name_var, width=32)
        entry.pack(padx=16, pady=(0, 12))

        matching_list = tk.Listbox(dialog, height=6)
        matching_list.pack(fill=tk.X, padx=16, pady=(0, 12))
        for option in options:
            matching_list.insert(tk.END, option)

        result = {"value": None}

        def accept_selection(value: str | None = None):
            candidate = value or name_var.get().strip()
            if candidate:
                result["value"] = candidate
            self.close_dialog(dialog)

        def on_click(_event):
            selection = matching_list.curselection()
            if selection:
                idx = selection[0]
                accept_selection(matching_list.get(idx))

        matching_list.bind("<Double-Button-1>", on_click)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 16), fill=tk.X)
        ttk.Button(button_frame, text="Abbrechen", command=lambda: self.close_dialog(dialog)).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(button_frame, text="Bestätigen", command=lambda: accept_selection()).pack(side=tk.RIGHT)

        entry.focus_set()
        dialog.bind("<Return>", lambda event: accept_selection())
        dialog.bind("<Escape>", lambda event: self.close_dialog(dialog))
        self.root.wait_window(dialog)
        return result["value"]

    def update_table(self):
        with self.smooth_state_transition():
            self._update_table_impl()

    def _update_table_impl(self):
        table_data = self.get_current_table_data()
        if table_data is None:
            self.render_placeholder("Kein Space vorhanden", "Erstelle über Optionen einen neuen Space oder importiere einen Speicherstand.")
            return

        rows = table_data["rows"]
        if not rows:
            self.render_placeholder("Keine Tabellen vorhanden", "Füge unten über die Aktion eine neue Tabelle hinzu.")
            return

        if not self.selected_row_name or self.selected_row_name not in rows:
            self.render_placeholder("Tabelle auswählen", "Nutze die Navigation links, um eine Tabelle zu wählen. Danach erscheinen hier die Karten.")
            return

        cards = table_data["cards"]
        row_colors = table_data["row_colors"]
        columns = table_data["columns"]
        row_name = self.selected_row_name

        for widget in self.table.winfo_children():
            widget.destroy()

        total_cards = sum(len(cards[row_name][col]) for col in columns)
        ratio = calculate_ratio(row_name, cards, columns)
        expected_grade = calculate_expected_grade(row_name, cards, columns)
        row_color = row_colors.get(row_name, self.primary_accent)

        self.render_row_header(row_name, ratio, expected_grade, row_color, total_cards)
        self.render_board_for_row(row_name, columns, cards, row_color)

    def add_card_to_row(self, row_index: int):
        table_data = self.get_current_table_data()
        if table_data is None:
            messagebox.showerror("Fehler", "Kein Space geladen!")
            return
        rows = table_data["rows"]

        idx = row_index - 1
        if 0 <= idx < len(rows):
            row_name = rows[idx]
            self.add_card_to_row_by_name(row_name)
        else:
            messagebox.showerror("Fehler", f"Es existiert keine Tabelle {row_index}.")

    def add_card_to_row_by_name(self, row_name: str):
        table_data = self.get_current_table_data()
        if table_data is None:
            messagebox.showerror("Fehler", "Kein Space geladen!")
            return
        cards = table_data["cards"]
        columns = table_data["columns"]

        prompt = self._build_card_prompt(row_name)
        if prompt is None:
            return
        card_name = prompt.get("value")
        if card_name:
            first_col = columns[0]
            self._record_history()
            cards[row_name][first_col].append(create_card(card_name))
            self.selected_row_name = row_name
            self.save_data()
            self.build_navigation_tree()
            self.update_table()

    def _build_card_prompt(self, row_name: str) -> dict | None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Karte hinzufügen")
        dialog.resizable(False, False)
        self.register_dialog(dialog)

        ttk.Label(dialog, text=f"Wie soll die Karte für '{row_name}' heißen?").pack(padx=16, pady=(16, 8))
        name_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=name_var, width=32)
        entry.pack(padx=16, pady=(0, 12))

        result = {"value": None}

        def submit():
            value = name_var.get().strip()
            if not value:
                return
            result["value"] = value
            self.close_dialog(dialog)

        def cancel():
            result["value"] = None
            self.close_dialog(dialog)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(padx=16, pady=(0, 16), fill=tk.X)
        ttk.Button(button_frame, text="Abbrechen", command=cancel).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(button_frame, text="Speichern", command=submit).pack(side=tk.RIGHT)

        entry.focus_set()
        dialog.bind("<Return>", lambda event: submit())
        dialog.bind("<Escape>", lambda event: cancel())
        self.root.wait_window(dialog)
        return result

    def add_card_via_button(self):
        table_data = self.get_current_table_data()
        if table_data is None:
            messagebox.showerror("Fehler", "Kein Space geladen!")
            return
        rows = table_data["rows"]

        if not rows:
            messagebox.showerror("Fehler", "Füge zuerst eine Tabelle hinzu!")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Tabelle auswählen")
        tk.Label(popup, text="Wähle eine Tabelle:").pack(padx=10, pady=10)

        button_frame = tk.Frame(popup)
        button_frame.pack(padx=10, pady=10)

        for _, row_name in enumerate(rows, start=1):
            btn = ttk.Button(button_frame, text=row_name, command=lambda rn=row_name: self.select_line_and_close_by_name(popup, rn))
            btn.pack(fill=tk.X, pady=2)

    def select_line_and_close_by_name(self, popup: tk.Toplevel, row_name: str):
        popup.destroy()
        self.selected_row_name = row_name
        self.add_card_to_row_by_name(row_name)

    def save_data(self):
        global SAVE_FILE
        with self.smooth_state_transition():
            try:
                directory = os.path.dirname(SAVE_FILE)
                if directory:
                    os.makedirs(directory, exist_ok=True)
                with open(SAVE_FILE, "w", encoding='utf-8') as f:
                    json.dump(self.data, f, indent=4, ensure_ascii=False)
            except Exception as exc:
                messagebox.showerror("Fehler", f"Fehler beim Speichern der Daten: {exc}")
            self.update_file_path_label()

    def load_data(self):
        global SAVE_FILE
        with self.smooth_state_transition():
            try:
                with open(SAVE_FILE, "r", encoding='utf-8') as f:
                    self.data = json.load(f)
            except FileNotFoundError:
                self.data = {"tables": {}, "current_table": None}
            except json.JSONDecodeError:
                messagebox.showerror("Fehler", "Die JSON-Datei ist beschädigt oder hat ein ungültiges Format.")
                self.data = {"tables": {}, "current_table": None}

            if "tables" not in self.data:
                self.data["tables"] = {}
            if "current_table" not in self.data:
                self.data["current_table"] = None

            self.update_file_path_label()

    def open_options_dialog(self):
        options_popup = tk.Toplevel(self.root)
        options_popup.title("Optionen")
        options_popup.resizable(False, False)
        button_kwargs = {"pady": 4, "fill": tk.X, "padx": 20}

        ttk.Button(options_popup, text="Neuer Space", command=lambda: self.new_table_wrapper(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Space laden", command=lambda: self.load_table_and_close(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Space importieren", command=lambda: self.import_table_and_close(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Space exportieren", command=lambda: self.export_table_and_close(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Space löschen", command=lambda: self.delete_table_from_buttons(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Speicherordner wählen", command=lambda: self.choose_dir_and_close(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Speicherdatei wechseln", command=lambda: self.change_save_location_wrapper(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Themenfarbe ändern", command=lambda: self.change_theme_and_close(options_popup)).pack(**button_kwargs)
        ttk.Button(options_popup, text="Speicherstand importieren", command=lambda: self.import_save_state(options_popup)).pack(**button_kwargs)

        ttk.Separator(options_popup, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=20, pady=12)

        inline_frame = ttk.Frame(options_popup, padding=(20, 0))
        inline_frame.pack(fill=tk.X, pady=(0, 16))
        ttk.Label(inline_frame, text="Tabelle hinzufügen", style="Subtitle.TLabel").pack(anchor=tk.W)

        inline_name_var = tk.StringVar()
        inline_preview_var = tk.StringVar()
        inline_color_label = tk.StringVar()
        inline_color_value = tk.StringVar(value=random_pastel_color())
        inline_error_var = tk.StringVar()
        inline_status_var = tk.StringVar()

        inline_color_chip = tk.Canvas(inline_frame, width=120, height=24, highlightthickness=0, bd=0)

        def inline_update_preview(*_args):
            row_name = inline_name_var.get().strip() or "Neue Tabelle"
            inline_preview_var.set(f"Name: {row_name}")
            inline_color_label.set(f"Farbe: {inline_color_value.get()}")
            inline_color_chip.configure(bg=inline_color_value.get())

        def inline_shuffle_color():
            inline_color_value.set(random_pastel_color())
            inline_update_preview()

        def submit_inline_row():
            desired_name = inline_name_var.get().strip() or "Neue Tabelle"
            success, message = self._create_row_entry(desired_name, inline_color_value.get())
            if success:
                inline_status_var.set(f"Tabelle '{desired_name}' erstellt.")
                inline_error_var.set("")
                inline_name_var.set("")
                inline_color_value.set(random_pastel_color())
                inline_update_preview()
            else:
                inline_error_var.set(message or "Unbekannter Fehler")
                inline_status_var.set("")

        ttk.Entry(inline_frame, textvariable=inline_name_var).pack(fill=tk.X, pady=(8, 4))
        ttk.Label(inline_frame, textvariable=inline_preview_var).pack(anchor=tk.W)
        ttk.Label(inline_frame, textvariable=inline_color_label, foreground=TEXT_MUTED).pack(anchor=tk.W, pady=(2, 0))
        inline_color_chip.pack(anchor=tk.W, pady=(2, 4))
        ttk.Button(inline_frame, text="Farbe neu würfeln", command=inline_shuffle_color).pack(anchor=tk.W)
        ttk.Button(inline_frame, text="Tabelle erstellen", command=submit_inline_row).pack(fill=tk.X, pady=(8, 4))
        ttk.Label(inline_frame, textvariable=inline_error_var, foreground="red").pack(anchor=tk.W)
        ttk.Label(inline_frame, textvariable=inline_status_var, foreground=self.primary_accent).pack(anchor=tk.W)

        inline_update_preview()

    def delete_table_from_buttons(self, popup: tk.Toplevel):
        if not self.data["tables"]:
            messagebox.showerror("Fehler", "Keine Spaces vorhanden!")
            return

        delete_popup = tk.Toplevel(self.root)
        delete_popup.title("Space löschen")
        tk.Label(delete_popup, text="Wähle einen Space zum Löschen:").pack(padx=10, pady=10)

        button_frame = tk.Frame(delete_popup)
        button_frame.pack(padx=10, pady=10)

        def delete_selected_table(table_name: str):
            confirm = messagebox.askokcancel("Bestätigung", f"Soll der Space '{table_name}' wirklich gelöscht werden?")
            if not confirm:
                return

            self._record_history()
            self.data["tables"].pop(table_name, None)
            if self.data["current_table"] == table_name:
                if self.data["tables"]:
                    self.data["current_table"] = list(self.data["tables"].keys())[0]
                else:
                    self.data["current_table"] = None
                    self.new_table(skip_history=True)
            self.selected_row_name = None

            self.save_data()
            self.ensure_row_selection_valid()
            self.build_navigation_tree()
            self.update_table()
            messagebox.showinfo("Erfolg", f"Space '{table_name}' wurde gelöscht.")
            delete_popup.destroy()
            popup.destroy()

        for table_name in self.data["tables"].keys():
            btn = ttk.Button(button_frame, text=table_name, command=lambda tn=table_name: delete_selected_table(tn))
            btn.pack(fill=tk.X, pady=2)

    def new_table_wrapper(self, popup: tk.Toplevel):
        self.new_table()
        popup.destroy()

    def load_table_and_close(self, popup: tk.Toplevel):
        self.load_table_via_menu()
        popup.destroy()

    def import_table_and_close(self, popup: tk.Toplevel):
        self.import_table_from_file()
        popup.destroy()

    def export_table_and_close(self, popup: tk.Toplevel):
        self.export_current_table()
        popup.destroy()

    def choose_dir_and_close(self, popup: tk.Toplevel):
        self.choose_save_directory()
        popup.destroy()

    def change_theme_and_close(self, popup: tk.Toplevel):
        popup.destroy()

    def new_table(self, skip_history: bool = False):
        dialog = tk.Toplevel(self.root)
        dialog.title("Neuer Space")
        dialog.resizable(False, False)
        self.register_dialog(dialog)

        name_var = tk.StringVar()
        cols_var = tk.StringVar(value="5")
        preview_var = tk.StringVar()
        pending_rows: list[str] = []
        pending_error_var = tk.StringVar()
        pending_status_var = tk.StringVar()
        pending_name_var = tk.StringVar()

        def update_preview(*_args):
            title = name_var.get().strip() or "Neuer Space"
            cols = cols_var.get()
            try:
                cols_int = int(cols)
            except ValueError:
                cols_int = 0
            if cols_int < 2:
                preview = f"{title}: Mindestens 2 Spalten notwendig"
            else:
                sample_cols = ", ".join(generate_columns(cols_int))
                preview = f"{title} · Spalten: {sample_cols}"
            preview_var.set(preview)

        ttk.Label(dialog, text="Space-Name").pack(padx=16, pady=(16, 4), anchor=tk.W)
        space_name_entry = ttk.Entry(dialog, textvariable=name_var, width=32)
        space_name_entry.pack(padx=16, fill=tk.X)
        ttk.Label(dialog, text="Spaltenanzahl").pack(padx=16, pady=(12, 4), anchor=tk.W)
        cols_entry = ttk.Entry(dialog, textvariable=cols_var, width=10)
        cols_entry.pack(padx=16, fill=tk.X)
        ttk.Label(dialog, textvariable=preview_var, wraplength=280, foreground=TEXT_MUTED).pack(padx=16, pady=12)

        error_var = tk.StringVar()
        error_label = ttk.Label(dialog, textvariable=error_var, foreground="red")
        error_label.pack(padx=16, pady=(0, 8))

        ttk.Separator(dialog, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=16, pady=(0, 12))
        ttk.Label(dialog, text="Tabellen vorbereiten", font=("Segoe UI", 10, "bold")).pack(padx=16, anchor=tk.W)
        ttk.Label(dialog, text="Tabellen verwenden die oben festgelegte Spaltenanzahl.", foreground=TEXT_MUTED).pack(padx=16, pady=(0, 8), anchor=tk.W)

        pending_frame = ttk.Frame(dialog)
        pending_frame.pack(fill=tk.X, padx=16)

        pending_name_entry = ttk.Entry(pending_frame, textvariable=pending_name_var)
        pending_name_entry.grid(row=0, column=0, sticky=tk.EW, pady=(0, 6))
        pending_frame.columnconfigure(0, weight=1)

        def refresh_pending_list():
            pending_listbox.delete(0, tk.END)
            for idx, row_name in enumerate(pending_rows, start=1):
                pending_listbox.insert(tk.END, f"{idx}. {row_name}")
            pending_status_var.set(f"{len(pending_rows)} Tabellen geplant")

        def add_pending_row():
            row_name = pending_name_var.get().strip()
            if not row_name:
                pending_error_var.set("Tabellenname darf nicht leer sein")
                return
            if len(pending_rows) >= 9:
                pending_error_var.set("Maximal 9 Tabellen pro Space")
                return
            if row_name in pending_rows:
                pending_error_var.set("Tabelle bereits geplant")
                return
            pending_rows.append(row_name)
            pending_name_var.set("")
            pending_error_var.set("")
            refresh_pending_list()

        def remove_selected_row():
            selection = pending_listbox.curselection()
            if not selection:
                return
            idx = selection[0]
            if 0 <= idx < len(pending_rows):
                pending_rows.pop(idx)
                pending_error_var.set("")
                refresh_pending_list()

        ttk.Button(pending_frame, text="Tabelle hinzufügen", command=add_pending_row).grid(row=0, column=1, padx=(8, 0), sticky=tk.E)
        pending_listbox = tk.Listbox(dialog, height=5, activestyle="dotbox")
        pending_listbox.pack(fill=tk.X, padx=16, pady=(0, 4))
        ttk.Button(dialog, text="Ausgewählte entfernen", command=remove_selected_row).pack(padx=16, anchor=tk.E)
        ttk.Label(dialog, textvariable=pending_error_var, foreground="red").pack(padx=16, pady=(4, 0), anchor=tk.W)
        ttk.Label(dialog, textvariable=pending_status_var, foreground=TEXT_MUTED).pack(padx=16, pady=(0, 8), anchor=tk.W)
        pending_status_var.set("0 Tabellen geplant")
        refresh_pending_list()

        def submit():
            table_name = name_var.get().strip()
            if not table_name:
                error_var.set("Name darf nicht leer sein")
                return
            if table_name in self.data["tables"]:
                error_var.set("Space existiert bereits")
                return
            cols = cols_var.get().strip()
            if not cols.isdigit():
                error_var.set("Spaltenzahl muss numerisch sein")
                return
            num_cols = int(cols)
            if num_cols < 2:
                error_var.set("Mindestens 2 Spalten erforderlich")
                return

            columns = generate_columns(num_cols)
            if not skip_history:
                self._record_history()
            self.data["tables"][table_name] = {"rows": [], "cards": {}, "row_colors": {}, "columns": columns}
            self.data["current_table"] = table_name
            rows = self.data["tables"][table_name]["rows"]
            cards = self.data["tables"][table_name]["cards"]
            row_colors = self.data["tables"][table_name]["row_colors"]
            for row_name in pending_rows:
                rows.append(row_name)
                cards[row_name] = {col: [] for col in columns}
                row_colors[row_name] = random_pastel_color()
            self.selected_row_name = pending_rows[0] if pending_rows else None
            self.root.title(f"Klausurmaster2D - {table_name}")
            self.save_data()
            self.build_navigation_tree()
            self.update_table()
            self.close_dialog(dialog)

        ttk.Button(dialog, text="Erstellen", command=submit).pack(padx=16, pady=(0, 16))
        name_var.trace_add("write", update_preview)
        cols_var.trace_add("write", update_preview)
        update_preview()
        space_name_entry.focus_set()
        dialog.bind("<Return>", lambda event: submit())

    def change_save_location_wrapper(self, popup: tk.Toplevel):
        self.change_save_location()
        popup.destroy()

    def load_from_json_buttons(self, popup: tk.Toplevel):
        if not self.data["tables"]:
            messagebox.showerror("Fehler", "Keine Spaces vorhanden!")
            return

        load_popup = tk.Toplevel(self.root)
        load_popup.title("Space laden")
        tk.Label(load_popup, text="Wähle einen Space:").pack(padx=10, pady=10)

        button_frame = tk.Frame(load_popup)
        button_frame.pack(padx=10, pady=10)

        def load_selected_table(table_name: str):
            self.data["current_table"] = table_name
            self.selected_row_name = None
            self.save_data()
            self.ensure_row_selection_valid()
            self.build_navigation_tree()
            self.update_table()
            messagebox.showinfo("Erfolg", f"Space '{table_name}' wurde geladen.")
            load_popup.destroy()
            popup.destroy()

        for table_name in self.data["tables"].keys():
            btn = ttk.Button(button_frame, text=table_name, command=lambda tn=table_name: load_selected_table(tn))
            btn.pack(fill=tk.X, pady=2)

    def import_save_state(self, parent_popup: tk.Toplevel):
        import_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")], title="Speicherstand importieren")
        if not import_path:
            return

        try:
            with open(import_path, "r", encoding='utf-8') as f:
                imported_data = json.load(f)
        except Exception as exc:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {exc}")
            return

        if not isinstance(imported_data, dict) or "tables" not in imported_data or "current_table" not in imported_data:
            messagebox.showerror("Fehler", "Die importierte Datei hat ein ungültiges Format.")
            return

        merge = messagebox.askyesno(
            "Importoption",
            "Möchten Sie die importierten Daten zu den bestehenden hinzufügen?\n\nJa: Hinzufügen\nNein: Ersetzen",
        )

        if merge:
            self._record_history()
            for table_name, table_data in imported_data["tables"].items():
                if table_name in self.data["tables"]:
                    overwrite = messagebox.askyesno(
                        "Doppelter Space", f"Der Space '{table_name}' existiert bereits. Möchten Sie ihn überschreiben?"
                    )
                    if overwrite:
                        self.data["tables"][table_name] = table_data
                else:
                    self.data["tables"][table_name] = table_data

            if imported_data["current_table"] in self.data["tables"]:
                self.data["current_table"] = imported_data["current_table"]
        else:
            self._record_history()
            self.data = imported_data

        self.selected_row_name = None
        self.save_data()
        self.ensure_row_selection_valid()
        self.build_navigation_tree()
        self.update_table()
        messagebox.showinfo("Erfolg", "Speicherstand erfolgreich importiert.")
        parent_popup.destroy()

    def change_save_location(self):
        global SAVE_FILE

        user_choice = messagebox.askquestion(
            self.tr("change_save.option_title"),
            self.tr("change_save.option_question"),
            icon='question',
            default='yes',
        )

        if user_choice == 'yes':
            new_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                title=self.tr("change_save.new_dialog_title"),
            )
            if new_path:
                SAVE_FILE = persist_save_file_path(new_path)
                print(f"Neuer Speicherort gesetzt: {SAVE_FILE}")
                self.save_data()

        elif user_choice == 'no':
            existing_path = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json")],
                title=self.tr("change_save.existing_dialog_title"),
            )
            if existing_path:
                SAVE_FILE = persist_save_file_path(existing_path)
                print(f"Bestehende Datei geladen: {SAVE_FILE}")
                self.load_data()
        else:
            print(self.tr("change_save.no_change"))

        self.update_file_path_label()
        self.refresh_bin_history()

    def choose_save_directory(self):
        global SAVE_FILE
        directory = filedialog.askdirectory(title="Speicherordner wählen")
        if not directory:
            return

        filename = os.path.basename(SAVE_FILE) if SAVE_FILE else "Tabellenspeicher_neu.json"
        new_path = os.path.join(directory, filename)
        SAVE_FILE = persist_save_file_path(new_path)
        self.save_data()
        messagebox.showinfo(
            self.tr("choose_directory.success_title"),
            self.tr("choose_directory.success_body", path=SAVE_FILE),
        )
        self.update_file_path_label()
        self.refresh_bin_history()

    def prompt_table_choice(self, title: str, prompt: str) -> str | None:
        tables = list(self.data.get("tables", {}).keys())
        if not tables:
            messagebox.showerror("Fehler", "Keine Spaces vorhanden!")
            return None

        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.resizable(False, False)
        ttk.Label(dialog, text=prompt).pack(padx=20, pady=(20, 10))

        selection = tk.StringVar(value=tables[0])
        combo = ttk.Combobox(dialog, values=tables, textvariable=selection, state="readonly")
        combo.pack(padx=20, pady=(0, 20))
        combo.focus_set()

        chosen = {"value": None}

        def confirm_choice():
            chosen["value"] = selection.get()
            dialog.destroy()

        ttk.Button(dialog, text="Auswählen", command=confirm_choice).pack(padx=20, pady=(0, 10))
        ttk.Button(dialog, text="Abbrechen", command=dialog.destroy).pack(padx=20, pady=(0, 20))

        self.root.wait_window(dialog)
        return chosen["value"]

    def load_table_via_menu(self):
        table_name = self.prompt_table_choice("Space laden", "Wähle einen Space zum Laden:")
        if not table_name:
            return
        self.set_active_table(table_name)
        self.build_navigation_tree()
        self.update_table()
        messagebox.showinfo("Erfolg", f"Space '{table_name}' wurde geladen.")

    def delete_table_via_menu(self):
        table_name = self.prompt_table_choice("Space löschen", "Wähle einen Space zum Löschen:")
        if not table_name:
            return
        confirm = messagebox.askokcancel("Bestätigung", f"Soll der Space '{table_name}' wirklich gelöscht werden?")
        if not confirm:
            return
        self._record_history()
        self.data["tables"].pop(table_name, None)
        if self.data["current_table"] == table_name:
            if self.data["tables"]:
                self.data["current_table"] = next(iter(self.data["tables"]), None)
            else:
                self.data["current_table"] = None
                self.new_table(skip_history=True)
        self.selected_row_name = None
        self.save_data()
        self.ensure_row_selection_valid()
        self.build_navigation_tree()
        self.update_table()
        messagebox.showinfo("Erfolg", f"Space '{table_name}' wurde gelöscht.")

    def transfer_table_between_spaces(self):
        spaces = list(self.data.get("tables", {}).keys())
        if len(spaces) < 2:
            messagebox.showerror("Fehler", "Es werden mindestens zwei Spaces benötigt.")
            return

        sources_with_rows = [name for name in spaces if self.data["tables"].get(name, {}).get("rows")]
        if not sources_with_rows:
            messagebox.showerror("Fehler", "Keine Tabellen vorhanden, die übertragen werden könnten.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Tabelle verschieben oder kopieren")
        dialog.resizable(False, False)
        self.register_dialog(dialog)

        action_var = tk.StringVar(value="move")
        src_var = tk.StringVar(value=sources_with_rows[0])
        dest_default = next((name for name in spaces if name != src_var.get()), spaces[0])
        dest_var = tk.StringVar(value=dest_default)
        table_var = tk.StringVar()
        name_var = tk.StringVar()
        info_var = tk.StringVar()
        error_var = tk.StringVar()

        ttk.Label(dialog, text="Aktion wählen").pack(padx=16, pady=(16, 4), anchor=tk.W)
        action_frame = ttk.Frame(dialog)
        action_frame.pack(padx=16, anchor=tk.W)
        ttk.Radiobutton(action_frame, text="Verschieben", variable=action_var, value="move").pack(side=tk.LEFT, padx=(0, 16))
        ttk.Radiobutton(action_frame, text="Kopieren", variable=action_var, value="copy").pack(side=tk.LEFT)

        ttk.Label(dialog, text="Quell-Space").pack(padx=16, pady=(12, 4), anchor=tk.W)
        src_combo = ttk.Combobox(dialog, values=spaces, textvariable=src_var, state="readonly")
        src_combo.pack(padx=16, fill=tk.X)

        ttk.Label(dialog, text="Tabelle auswählen").pack(padx=16, pady=(12, 4), anchor=tk.W)
        table_combo = ttk.Combobox(dialog, values=[], textvariable=table_var, state="readonly")
        table_combo.pack(padx=16, fill=tk.X)

        ttk.Label(dialog, text="Ziel-Space").pack(padx=16, pady=(12, 4), anchor=tk.W)
        dest_combo = ttk.Combobox(dialog, values=spaces, textvariable=dest_var, state="readonly")
        dest_combo.pack(padx=16, fill=tk.X)

        ttk.Label(dialog, text="Neuer Tabellenname (optional)").pack(padx=16, pady=(12, 4), anchor=tk.W)
        name_entry = ttk.Entry(dialog, textvariable=name_var)
        name_entry.pack(padx=16, fill=tk.X)

        ttk.Label(dialog, textvariable=info_var, foreground=TEXT_MUTED, wraplength=320, justify=tk.LEFT).pack(padx=16, pady=(8, 0), anchor=tk.W)
        ttk.Label(dialog, textvariable=error_var, foreground="red", wraplength=320, justify=tk.LEFT).pack(padx=16, pady=(4, 0), anchor=tk.W)

        def update_table_choices(*_args):
            source = src_var.get()
            rows = self.data["tables"].get(source, {}).get("rows", [])
            table_combo["values"] = rows
            if rows:
                table_var.set(rows[0])
                if not name_var.get():
                    name_var.set(rows[0])
            else:
                table_var.set("")
            update_info()

        def update_info(*_args):
            source = src_var.get()
            dest = dest_var.get()
            src_data = self.data["tables"].get(source)
            dest_data = self.data["tables"].get(dest)
            details = []
            if src_data and dest_data:
                if src_data["columns"] == dest_data["columns"]:
                    details.append(f"Spalten kompatibel ({len(src_data['columns'])})")
                else:
                    details.append("Warnung: Spalten unterscheiden sich")
                details.append(f"Zieltabelle: {len(dest_data['rows'])}/9 belegt")
            info_var.set(" · ".join(details))

        def sync_name_with_selection(*_args):
            if not name_var.get():
                name_var.set(table_var.get())

        def submit_transfer():
            action = action_var.get()
            source = src_var.get()
            dest = dest_var.get()
            table_name = table_var.get()
            desired_name = name_var.get().strip() or table_name

            error_var.set("")

            if action not in {"move", "copy"}:
                action = "move"
            if not source or not dest or not table_name:
                error_var.set("Bitte Quelle, Tabelle und Ziel auswählen.")
                return
            if source == dest:
                error_var.set("Quelle und Ziel müssen unterschiedlich sein.")
                return
            src_data = self.data["tables"].get(source)
            dest_data = self.data["tables"].get(dest)
            if not src_data or not dest_data:
                error_var.set("Ungültige Space-Auswahl.")
                return
            if table_name not in src_data["rows"]:
                error_var.set("Die ausgewählte Tabelle existiert nicht mehr.")
                return
            if len(dest_data["rows"]) >= 9:
                error_var.set("Der Ziel-Space ist bereits voll (9 Tabellen).")
                return
            if desired_name in dest_data["rows"]:
                error_var.set("Im Ziel existiert bereits eine Tabelle mit diesem Namen.")
                return
            if src_data["columns"] != dest_data["columns"]:
                error_var.set("Spaces haben unterschiedliche Spalten. Bitte zuerst angleichen.")
                return

            self._record_history()

            columns = dest_data["columns"]
            source_cards = src_data["cards"].get(table_name, {})
            cloned_cards = {col: copy.deepcopy(source_cards.get(col, [])) for col in columns}
            dest_data["rows"].append(desired_name)
            dest_data["cards"][desired_name] = cloned_cards
            dest_data["row_colors"][desired_name] = src_data["row_colors"].get(table_name, random_pastel_color())

            if action == "move":
                src_data["rows"].remove(table_name)
                src_data["cards"].pop(table_name, None)
                src_data["row_colors"].pop(table_name, None)
                if self.data.get("current_table") == source and self.selected_row_name == table_name:
                    self.selected_row_name = None

            if self.data.get("current_table") == dest:
                self.selected_row_name = desired_name

            self.save_data()
            self.ensure_row_selection_valid()
            self.build_navigation_tree()
            self.update_table()

            self.close_dialog(dialog)
            verb = "verschoben" if action == "move" else "kopiert"
            messagebox.showinfo("Erfolg", f"Tabelle '{table_name}' wurde nach '{dest}' {verb} (als '{desired_name}').")

        src_combo.bind("<<ComboboxSelected>>", update_table_choices)
        dest_combo.bind("<<ComboboxSelected>>", update_info)
        table_combo.bind("<<ComboboxSelected>>", sync_name_with_selection)

        update_table_choices()
        if not table_var.get():
            # If the current source has no rows, try switching to another
            for candidate in spaces:
                if self.data["tables"].get(candidate, {}).get("rows"):
                    src_var.set(candidate)
                    update_table_choices()
                    break
        update_info()
        sync_name_with_selection()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(padx=16, pady=(16, 16), fill=tk.X)
        ttk.Button(button_frame, text="Abbrechen", command=lambda: self.close_dialog(dialog)).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(button_frame, text="Ausführen", command=submit_transfer).pack(side=tk.RIGHT)

        name_entry.focus_set()

    def import_table_from_file(self):
        import_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")], title="Space importieren")
        if not import_path:
            return

        try:
            with open(import_path, "r", encoding='utf-8') as f:
                imported = json.load(f)
        except Exception as exc:
            messagebox.showerror("Fehler", f"Fehler beim Import: {exc}")
            return

        if isinstance(imported, dict) and "name" in imported and "table" in imported:
            suggested_name = imported["name"]
            table_payload = imported["table"]
        else:
            suggested_name = imported.get("name") if isinstance(imported, dict) else None
            table_payload = imported

        if not isinstance(table_payload, dict) or not {"rows", "cards", "row_colors", "columns"}.issubset(table_payload.keys()):
            messagebox.showerror("Fehler", "Die importierte Datei enthält keinen gültigen Space.")
            return

        default_name = suggested_name or "Neuer Space"
        self._bring_dialog_to_front()
        table_name = simpledialog.askstring("Space-Name", "Name des importierten Space:", initialvalue=default_name)
        if not table_name:
            return
        if table_name in self.data["tables"]:
            overwrite = messagebox.askyesno("Überschreiben?", f"'{table_name}' existiert bereits. Überschreiben?")
            if not overwrite:
                return

        self._record_history()
        self.data["tables"][table_name] = table_payload
        self.data["current_table"] = table_name
        self.selected_row_name = None
        self.save_data()
        self.build_navigation_tree()
        self.update_table()
        messagebox.showinfo("Erfolg", f"Space '{table_name}' importiert.")

    def _parse_card_import_block(self, raw_text: str) -> tuple[str, list[tuple[str, str]]]:
        text = raw_text.strip()
        if not text:
            raise ValueError("Kein Inhalt zum Importieren gefunden.")

        def strip_braces(payload: str) -> str:
            payload = payload.strip()
            if payload.startswith("{") and payload.endswith("}"):
                return payload[1:-1].strip()
            return payload

        text = strip_braces(text)
        if "::" not in text:
            raise ValueError("Format ungültig: '::' zwischen Space und Kartenblock fehlt.")

        space_part, remainder = text.split("::", 1)
        space_name = space_part.strip()
        if not space_name:
            raise ValueError("Space-Name konnte nicht gelesen werden.")

        remainder = strip_braces(remainder)
        if not remainder:
            raise ValueError("Der Kartenblock ist leer.")

        cards_payload: list[tuple[str, str]] = []
        for raw_line in remainder.split(";"):
            line = raw_line.strip()
            if not line:
                continue
            if "::" in line:
                front, back = line.split("::", 1)
            else:
                front, back = line, ""
            front = front.strip()
            back = back.strip()
            if not front:
                continue
            cards_payload.append((front, back))

        if not cards_payload:
            raise ValueError("Es konnten keine gültigen Kartenzeilen gelesen werden.")

        return space_name, cards_payload

    def _apply_card_import(self, space_name: str, cards_payload: list[tuple[str, str]]):
        table_data = self.data.get("tables", {}).get(space_name)
        if not table_data:
            messagebox.showerror("Fehler", f"Space '{space_name}' existiert nicht.")
            return

        rows = table_data.get("rows", [])
        if not rows:
            messagebox.showerror("Fehler", f"Space '{space_name}' enthält keine Tabellen zum Befüllen.")
            return

        if self.data.get("current_table") == space_name and self.selected_row_name in rows:
            row_name = self.selected_row_name
        else:
            row_name = self._prompt_row_name(
                "Zieltabelle wählen",
                f"In welche Tabelle innerhalb von '{space_name}' sollen die Karten importiert werden?",
                rows,
            )
            if not row_name or row_name not in rows:
                return

        columns = table_data.get("columns", [])
        if not columns:
            messagebox.showerror("Fehler", "Der Space verfügt über keine Spalten.")
            return

        cards = table_data.setdefault("cards", {})
        if row_name not in cards:
            cards[row_name] = {col: [] for col in columns}
        else:
            for col in columns:
                cards[row_name].setdefault(col, [])

        first_column = columns[0]
        existing_fronts = set()
        for col in columns:
            for card in cards[row_name].get(col, []):
                existing_fronts.add(str(card.get("front", "")))

        unique_payload: list[tuple[str, str]] = []
        skipped: list[str] = []
        for front, back in cards_payload:
            if front in existing_fronts:
                skipped.append(front)
                continue
            existing_fronts.add(front)
            unique_payload.append((front, back))

        if not unique_payload:
            messagebox.showinfo("Keine neuen Karten", "Alle übergebenen Karten existieren bereits in dieser Tabelle.")
            return

        self._record_history()
        for front, back in unique_payload:
            cards[row_name][first_column].append(create_card(front, back))

        self.data["current_table"] = space_name
        self.selected_row_name = row_name
        self.save_data()
        self.build_navigation_tree()
        self.update_table()

        summary = f"{len(unique_payload)} Karten importiert."
        if skipped:
            summary += f" {len(skipped)} doppelte Einträge übersprungen."
        messagebox.showinfo("Import abgeschlossen", summary)

    def import_cards_via_text(self):
        if not self.data.get("tables"):
            messagebox.showerror("Fehler", "Keine Spaces vorhanden, in die importiert werden könnte.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Karten importieren")
        dialog.geometry("560x400")
        dialog.minsize(460, 320)
        self.register_dialog(dialog)

        instructions = (
            "Format: {Space::{Kartenvorderseite :: optionale Rückseite; …}}\n"
            "Der angegebene Space muss bereits existieren."
        )
        ttk.Label(dialog, text=instructions, wraplength=520, justify=tk.LEFT).pack(padx=16, pady=(16, 8), anchor=tk.W)

        text_widget = tk.Text(dialog, wrap="word")
        text_widget.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        example = "{\nMeinSpace::{\nFragenname :: Hinweis;\nWeiterer Begriff :: ;\n}\n}"
        text_widget.insert("1.0", example)

        error_var = tk.StringVar()
        ttk.Label(dialog, textvariable=error_var, foreground="red").pack(padx=16, pady=(0, 8), anchor=tk.W)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=16, pady=(0, 16))

        def close_dialog_local():
            self.close_dialog(dialog)

        def confirm_import():
            raw_text = text_widget.get("1.0", "end-1c")
            try:
                space_name, payload = self._parse_card_import_block(raw_text)
            except ValueError as exc:
                error_var.set(str(exc))
                return

            if space_name not in self.data.get("tables", {}):
                error_var.set(f"Space '{space_name}' existiert nicht.")
                return

            self.close_dialog(dialog)
            self._apply_card_import(space_name, payload)

        ttk.Button(button_frame, text="Abbrechen", command=close_dialog_local).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(button_frame, text="Importieren", command=confirm_import).pack(side=tk.RIGHT)

        text_widget.focus_set()
        dialog.bind("<Control-Return>", lambda event: (confirm_import(), "break"))

    def export_current_table(self):
        current_table = self.data.get("current_table")
        if not current_table or current_table not in self.data.get("tables", {}):
            messagebox.showerror("Fehler", "Kein Space zum Exportieren ausgewählt.")
            return

        export_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Space exportieren",
            initialfile=f"{current_table}.json",
        )
        if not export_path:
            return

        payload = {"name": current_table, "table": self.data["tables"][current_table]}
        try:
            with open(export_path, "w", encoding='utf-8') as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
        except Exception as exc:
            messagebox.showerror("Fehler", f"Export fehlgeschlagen: {exc}")
            return
        messagebox.showinfo("Erfolg", f"Space wurde nach\n{export_path}\nexportiert.")

    def duplicate_current_table(self):
        current_table = self.data.get("current_table")
        if not current_table:
            messagebox.showerror("Fehler", "Kein Space ausgewählt.")
            return
        source = self.data["tables"].get(current_table)
        if not source:
            messagebox.showerror("Fehler", "Der aktuelle Space ist leer oder ungültig.")
            return

        suggestion = f"{current_table}_kopie"
        self._bring_dialog_to_front()
        new_name = simpledialog.askstring("Space duplizieren", "Name der Kopie:", initialvalue=suggestion)
        if not new_name:
            return
        if new_name in self.data["tables"]:
            messagebox.showerror("Fehler", "Ein Space mit diesem Namen existiert bereits.")
            return

        self._record_history()
        self.data["tables"][new_name] = copy.deepcopy(source)
        self.data["current_table"] = new_name
        self.selected_row_name = None
        self.save_data()
        self.build_navigation_tree()
        self.update_table()
        messagebox.showinfo("Erfolg", f"Space '{current_table}' wurde als '{new_name}' kopiert.")

    def show_about_dialog(self):
        messagebox.showinfo(
            "Über",
            "Karteisystem 2.0\nModerne Navigation & Optionenmenü\n© 2025",
        )

    def bind_shortcuts(self):
        self.root.bind("<Control-n>", lambda event: self.add_row())
        self.root.bind("<Control-w>", lambda event: self.toggle_delete_mode())
        self.root.bind("<Escape>", lambda event: self.cancel_operations())
        self.root.bind("<Control-s>", self.shortcut_new_space)
        self.root.bind("<Control-z>", self.undo_action)
        self.root.bind("<Control-Shift-Z>", self.redo_action)
        self.root.bind("<Control-y>", self.redo_action)

        for i in range(1, 10):
            self.root.bind(f"<Key-{i}>", lambda event, x=i: self.add_card_to_row(x))

        self.root.bind("<Control-m>", lambda event: self.toggle_mark_mode())

        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind_all("<Button-4>", self.on_mousewheel)
        self.canvas.bind_all("<Button-5>", self.on_mousewheel)

    def on_mousewheel(self, event):
        if getattr(event, "num", None) == 4:
            self.canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def cancel_operations(self):
        if self.delete_mode:
            self.delete_mode = False
        if self.moving_card:
            self.moving_card = None
        self.update_table()

    def toggle_delete_mode(self):
        self.delete_mode = not self.delete_mode
        if self.delete_mode and self.mark_mode:
            self.mark_mode = False
        self.update_table()

    def toggle_mark_mode(self):
        self.mark_mode = not self.mark_mode
        if self.mark_mode and self.delete_mode:
            self.delete_mode = False
        self.update_table()

    def on_column_click(self, row_name: str, col_name: str):
        if self.moving_card:
            table_data = self.get_current_table_data()
            if table_data is None:
                return
            cards = table_data["cards"]
            old_row, old_col, card_front = self.moving_card
            old_card = find_card(cards[old_row][old_col], card_front)
            if old_card:
                self._record_history()
                cards[old_row][old_col].remove(old_card)
                cards[row_name][col_name].append(old_card)
                self.moving_card = None
                self.save_data()
                self.update_table()

    def on_card_click(self, row_name: str, col_name: str, card_front: str):
        table_data = self.get_current_table_data()
        if table_data is None:
            return

        cards = table_data["cards"]
        card_dict = find_card(cards[row_name][col_name], card_front)
        if not card_dict:
            return

        if self.mark_mode:
            self._record_history()
            card_dict["marked"] = not card_dict["marked"]
            self.save_data()
            self.update_table()
            return

        if self.delete_mode:
            self._record_history()
            cards[row_name][col_name].remove(card_dict)
            self.delete_mode = False
            self.save_data()
            self.update_table()
        else:
            if self.moving_card is None:
                self.moving_card = (row_name, col_name, card_front)
                self.update_table()
            else:
                old_row, old_col, moving_front = self.moving_card
                old_card = find_card(cards[old_row][old_col], moving_front)
                if old_card:
                    self._record_history()
                    cards[old_row][old_col].remove(old_card)
                    cards[row_name][col_name].append(old_card)
                    self.moving_card = None
                    self.save_data()
                    self.update_table()

    def _open_card_back_editor(self, row_name: str, col_name: str, card_front: str):
        table_data = self.get_current_table_data()
        if table_data is None:
            return
        cards = table_data["cards"]

        card_dict = find_card(cards[row_name][col_name], card_front)
        if not card_dict:
            return

        editor = tk.Toplevel(self.root)
        editor.title(f"Kartenrückseite - {card_front}")
        editor.geometry("420x320")
        self.register_dialog(editor)

        text_widget = tk.Text(editor, wrap="word", font=("Segoe UI", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 0))
        text_widget.insert("1.0", card_dict.get("back", ""))

        status_var = tk.StringVar()
        status_label = ttk.Label(editor, textvariable=status_var, foreground=TEXT_MUTED)
        status_label.pack(anchor=tk.W, padx=12, pady=(6, 0))

        button_frame = ttk.Frame(editor)
        button_frame.pack(fill=tk.X, padx=12, pady=12)

        def persist_changes() -> bool:
            new_text = text_widget.get("1.0", "end-1c")
            if new_text == card_dict.get("back", ""):
                return False
            self._record_history()
            card_dict["back"] = new_text
            self.save_data()
            self.refresh_card_column(row_name, col_name)
            return True

        def save_and_close(_event=None):
            changed = persist_changes()
            if changed:
                status_var.set("Änderungen gespeichert.")
            self.close_dialog(editor)

        def save_only(_event=None):
            changed = persist_changes()
            status_var.set("Änderungen gespeichert." if changed else "Keine Änderungen.")

        ttk.Button(button_frame, text="Speichern", command=save_only).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Speichern & schließen", command=save_and_close).pack(side=tk.RIGHT)

        def _on_ctrl_s(event):
            save_only()
            return "break"

        editor.bind("<Control-s>", _on_ctrl_s)
        editor.bind("<Escape>", save_and_close)

    def on_card_right_click(self, row_name: str, col_name: str, card_front: str):
        self._open_card_back_editor(row_name, col_name, card_front)

    def on_card_press(self, event, row_name: str, col_name: str, card_front: str):
        if self.delete_mode or self.mark_mode:
            self.drag_data = None
            return
        card_frame = self._get_card_frame(event.widget)
        payload = getattr(card_frame, "_card_payload", None)
        self.drag_data = {
            "row": row_name,
            "col": col_name,
            "card": card_front,
            "start": (event.x_root, event.y_root),
            "widget": card_frame,
            "moved": False,
            "payload": payload,
        }
        self.moving_card = None

    def on_card_motion(self, event):
        if not self.drag_data:
            return
        start_x, start_y = self.drag_data["start"]
        dx = abs(event.x_root - start_x)
        dy = abs(event.y_root - start_y)
        if not self.drag_data["moved"] and max(dx, dy) > 8:
            self.drag_data["moved"] = True
            self.create_drag_preview(self.drag_data.get("payload"))
        if not self.drag_data["moved"]:
            return
        self.update_drag_preview_position(event.x_root, event.y_root)
        self.highlight_column_under_pointer(event.x_root, event.y_root)

    def on_card_release(self, event, row_name: str, col_name: str, card_front: str):
        if self.drag_data and self.drag_data.get("moved"):
            widget = self.root.winfo_containing(event.x_root, event.y_root)
            column_frame = self._find_column_frame(widget)
            if column_frame:
                target_row = column_frame._row_name
                target_col = column_frame._col_name
                self.move_card_between_columns(row_name, col_name, card_front, target_row, target_col)
            self.reset_drag_state()
            return

        self.reset_drag_state()
        self.on_card_click(row_name, col_name, card_front)

    def create_drag_preview(self, card_payload: dict | None):
        self.destroy_drag_preview()
        if not card_payload:
            return
        self.drag_preview = tk.Toplevel(self.root)
        self.drag_preview.overrideredirect(True)
        self.drag_preview.attributes("-topmost", True)
        self.drag_preview.attributes("-alpha", 0.95)

        preview_frame = tk.Frame(
            self.drag_preview,
            bg=CARD_BG,
            padx=12,
            pady=10,
            bd=0,
            highlightbackground=CARD_BORDER,
            highlightthickness=1,
        )
        preview_frame.pack()

        tk.Label(
            preview_frame,
            text=str(card_payload.get("front", "")),
            font=("Segoe UI", 11, "bold"),
            bg=CARD_BG,
            fg=TEXT_PRIMARY,
            wraplength=220,
            justify=tk.LEFT,
        ).pack(fill=tk.X)

        back_text = str(card_payload.get("back", "")).strip()
        if back_text:
            snippet = back_text if len(back_text) <= 120 else f"{back_text[:120]}…"
            tk.Label(
                preview_frame,
                text=snippet,
                font=("Segoe UI", 9),
                bg=CARD_BG,
                fg=TEXT_MUTED,
                wraplength=220,
                justify=tk.LEFT,
            ).pack(fill=tk.X, pady=(6, 0))

        weight = self._extract_card_weight(card_payload)
        tk.Label(
            preview_frame,
            text=f"Gewicht: {weight:.0f}",
            font=("Segoe UI", 8, "bold"),
            bg=CARD_BG,
            fg=TEXT_MUTED,
        ).pack(anchor=tk.W, pady=(8, 0))

    def update_drag_preview_position(self, x_root: int, y_root: int):
        if not self.drag_preview:
            return
        self.drag_preview.geometry(f"+{x_root + 12}+{y_root + 12}")

    def destroy_drag_preview(self):
        if self.drag_preview is not None:
            self.drag_preview.destroy()
            self.drag_preview = None

    def highlight_column_under_pointer(self, x_root: int, y_root: int):
        widget = self.root.winfo_containing(x_root, y_root)
        column_frame = self._find_column_frame(widget)
        if column_frame == self.current_highlighted_column:
            return
        self._clear_column_highlight()
        if column_frame:
            column_frame.configure(highlightthickness=2, highlightbackground=self.primary_accent)
            self.current_highlighted_column = column_frame

    def _clear_column_highlight(self):
        if self.current_highlighted_column is None:
            return
        base_color = getattr(self.current_highlighted_column, "_base_highlight_color", CARD_BORDER)
        try:
            self.current_highlighted_column.configure(highlightthickness=0, highlightbackground=base_color)
        except tk.TclError:
            pass
        self.current_highlighted_column = None

    def reset_drag_state(self):
        self.destroy_drag_preview()
        self._clear_column_highlight()
        self.drag_data = None

    def _get_card_frame(self, widget: tk.Widget) -> tk.Widget:
        current = widget
        while current is not None:
            if getattr(current, "_is_card_frame", False):
                return current
            current = getattr(current, "master", None)
        return widget

    def _find_column_frame(self, widget: tk.Widget | None) -> tk.Frame | None:
        current = widget
        while current is not None:
            if getattr(current, "_col_name", None) is not None:
                return current
            current = getattr(current, "master", None)
        return None

    def move_card_between_columns(
        self,
        source_row: str,
        source_col: str,
        card_front: str,
        target_row: str,
        target_col: str,
    ):
        table_data = self.get_current_table_data()
        if table_data is None:
            return
        if source_row == target_row and source_col == target_col:
            return

        cards = table_data["cards"]
        if target_row not in cards or target_col not in cards[target_row]:
            return
        old_card = find_card(cards[source_row][source_col], card_front)
        if not old_card:
            return

        self._record_history()
        cards[source_row][source_col].remove(old_card)
        cards[target_row][target_col].append(old_card)
        self.save_data()
        updated_source = self.refresh_card_column(source_row, source_col)
        updated_target = self.refresh_card_column(target_row, target_col)
        need_header_update = self.selected_row_name in {source_row, target_row}
        header_updated = True
        if need_header_update and self.selected_row_name:
            header_updated = self.update_row_header_info(self.selected_row_name)

        if not (updated_source or updated_target):
            self.update_table()
            return
        if need_header_update and not header_updated:
            self.update_table()

    def on_tree_select(self, _event):
        selection = self.navigation_tree.selection()
        if not selection:
            return
        item_id = selection[0]
        node_info = self.tree_row_lookup.get(item_id)
        if not node_info:
            return

        if node_info[0] == "row":
            _, table_name, row_name = node_info
            self.set_active_table(table_name, rebuild=False)
            self.selected_row_name = row_name
        else:
            _, table_name = node_info
            self.set_active_table(table_name, rebuild=False)
            self.selected_row_name = None

        self.update_table()

    def set_active_table(self, table_name: str, rebuild: bool = True):
        if table_name not in self.data.get("tables", {}):
            return
        if self.data["current_table"] != table_name:
            self.data["current_table"] = table_name
            self.selected_row_name = None
            self.save_data()
        self.ensure_row_selection_valid()
        if rebuild:
            self.build_navigation_tree()


def run_app():
    root = tk.Tk()
    ensure_save_path_initialized(root)
    app = CardApp(root)
    root.after(100, app.load_data_or_create_new_table)
    root.mainloop()
