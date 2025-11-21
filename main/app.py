"""Tkinter UI for managing cards, tables, and aggregated statistics."""

from __future__ import annotations

import copy
import json
import os
from datetime import datetime
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
import tkinter as tk

from cards import create_card, find_card, normalize_cards_tree
from formula import calculate_ratio, calculate_expected_grade
from table import generate_columns, interpolate_color, random_pastel_color

SAVE_FILE = r"P:\\Sync\\Karteisystem\\Tabellenspeicher_neu.json"
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
        "APP_BG": "#f1e7d3",
        "PANEL_BG": "#f5eddc",
        "CONTENT_BG": "#f9f3e6",
        "CARD_BG": "#fffaf2",
        "CARD_BORDER": "#d8c8ab",
        "CARD_DELETE_BG": "#f6d6d0",
        "PRIMARY_ACCENT": "#c89f65",
        "TEXT_PRIMARY": "#2c2114",
        "TEXT_MUTED": "#6b5c4b",
    },
}

APP_BG = THEMES["dark"]["APP_BG"]
PANEL_BG = THEMES["dark"]["PANEL_BG"]
CONTENT_BG = THEMES["dark"]["CONTENT_BG"]
CARD_BG = THEMES["dark"]["CARD_BG"]
CARD_BORDER = THEMES["dark"]["CARD_BORDER"]
CARD_DELETE_BG = THEMES["dark"]["CARD_DELETE_BG"]
PRIMARY_ACCENT = THEMES["dark"]["PRIMARY_ACCENT"]
TEXT_PRIMARY = THEMES["dark"]["TEXT_PRIMARY"]
TEXT_MUTED = THEMES["dark"]["TEXT_MUTED"]


class CardApp:
    def __init__(self, root: tk.Tk):
        # Initializes widgets, state, and loads persisted data.
        self.root = root
        self.root.title("Klausurmaster2D")
        self.root.geometry("1400x820")
        self.root.minsize(960, 640)
        self.root.iconbitmap("favicon.ico")

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

        self.style = ttk.Style()
        self.primary_accent = PRIMARY_ACCENT
        self.configure_styles()
        self.build_main_layout()
        self.build_menu()

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
        file_menu.add_command(label="Speichern", command=self.save_data)
        file_menu.add_command(label="Neu laden", command=self.load_data_or_create_new_table)
        file_menu.add_separator()
        file_menu.add_command(label="View state from…", command=self.view_history_state)
        file_menu.add_command(label="Load state from…", command=self.load_history_state)
        file_menu.add_separator()
        file_menu.add_command(label="Tabelle hinzufügen (Strg+N)", command=self.add_row)
        file_menu.add_command(label="Tabelle löschen", command=self.delete_row)
        file_menu.add_separator()
        file_menu.add_command(label="Karte hinzufügen (1–9)", command=self.add_card_via_button)
        file_menu.add_command(label="Löschmodus umschalten (Strg+W)", command=self.toggle_delete_mode)
        file_menu.add_command(label="Aktion abbrechen (Esc)", command=self.cancel_operations)
        file_menu.add_separator()
        file_menu.add_command(label="Speicherordner wählen…", command=self.choose_save_directory)
        file_menu.add_command(label="Speicherdatei ändern…", command=self.change_save_location)
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self.handle_exit_request)

        edit_menu = tk.Menu(self.menubar, tearoff=0)
        edit_menu.add_command(label="Rückgängig (Strg+Z)", command=self.undo_action)
        edit_menu.add_command(label="Wiederholen (Strg+Shift+Z)", command=self.redo_action)

        space_menu = tk.Menu(self.menubar, tearoff=0)
        space_menu.add_command(label="Neuer Space (Strg+S)", command=self.new_table)
        space_menu.add_command(label="Space laden…", command=self.load_table_via_menu)
        space_menu.add_separator()
        space_menu.add_command(label="Space importieren…", command=self.import_table_from_file)
        space_menu.add_command(label="Space exportieren…", command=self.export_current_table)
        space_menu.add_separator()
        space_menu.add_command(label="Space duplizieren…", command=self.duplicate_current_table)
        space_menu.add_command(label="Tabelle verschieben/kopieren…", command=self.transfer_table_between_spaces)
        space_menu.add_command(label="Space löschen…", command=self.delete_table_via_menu)
        space_menu.add_separator()
        theme_menu = tk.Menu(space_menu, tearoff=0)
        theme_menu.add_command(label="Dunkel", command=lambda: self.apply_theme("dark"))
        theme_menu.add_command(label="Beige", command=lambda: self.apply_theme("beige"))
        space_menu.add_cascade(label="Theme", menu=theme_menu)

        help_menu = tk.Menu(self.menubar, tearoff=0)
        help_menu.add_command(label="Über", command=self.show_about_dialog)

        self.menubar.add_cascade(label="Datei", menu=file_menu)
        self.menubar.add_cascade(label="Bearbeiten", menu=edit_menu)
        self.menubar.add_cascade(label="Space", menu=space_menu)
        self.menubar.add_cascade(label="Hilfe", menu=help_menu)
        self.root.config(menu=self.menubar)

    def _build_toolbar(self):
        self.toolbar = ttk.Frame(self.main_frame, padding=(24, 16), style="Toolbar.TFrame")
        self.toolbar.grid(row=0, column=0, columnspan=2, sticky=tk.EW)
        self.toolbar.columnconfigure(0, weight=1)

        title_wrapper = ttk.Frame(self.toolbar, style="Toolbar.TFrame")
        title_wrapper.grid(row=0, column=0, sticky=tk.W)
        ttk.Label(title_wrapper, text="Karteisystem 2.0", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(title_wrapper, text="Wähle links Space und Tabelle, um Karten zu fokussieren.", style="Subtitle.TLabel").pack(anchor=tk.W)

    def _build_navigation(self):
        self.nav_frame = ttk.Frame(self.main_frame, padding=(24, 16, 12, 24), style="Nav.TFrame")
        self.nav_frame.grid(row=1, column=0, sticky=tk.NS)
        self.nav_frame.rowconfigure(1, weight=1)

        ttk.Label(self.nav_frame, text="Sammlungen", style="Subtitle.TLabel").grid(row=0, column=0, sticky=tk.W, pady=(0, 6))

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
        self.file_path_label.config(text=f"Aktueller Speicherpfad: {SAVE_FILE}")

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

    def _compute_row_header_color(self, row_color: str, ratio: float, expected_grade: float) -> str:
        safe_ratio = max(0.0, min(ratio or 0.0, 1.0))
        grade_factor = 0.0
        if expected_grade:
            grade_factor = max(0.0, min((expected_grade - 4.0) / 2.0, 1.0))
        softened_color = interpolate_color(row_color, "#ffffff", 0.5 * grade_factor)
        return interpolate_color("#ffffff", softened_color, safe_ratio)

    def render_row_header(self, row_name: str, ratio: float, expected_grade: float, row_color: str, total_cards: int):
        header_color = self._compute_row_header_color(row_color, ratio, expected_grade)
        header = tk.Frame(self.table, bg=header_color, padx=18, pady=18, highlightthickness=0)
        header.grid(row=0, column=0, sticky=tk.EW, pady=(0, 18))
        header.columnconfigure(0, weight=1)

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
        self.row_header_widgets = {
            "row_name": row_name,
            "frame": header,
            "title": title,
            "subtitle": subtitle,
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

            title = tk.Label(
                card_frame,
                text=card_front,
                font=("Segoe UI", 11, "bold"),
                bg=base_bg,
                fg=TEXT_PRIMARY,
                wraplength=220,
                justify=tk.LEFT,
            )
            title.pack(fill=tk.X)
            self._bind_card_widget(title, row_name, col_name, card_front)

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

        ratio = calculate_ratio(row_name, cards, columns)
        expected_grade = calculate_expected_grade(row_name, cards, columns)
        row_color = row_colors.get(row_name, self.primary_accent)
        header_color = self._compute_row_header_color(row_color, ratio, expected_grade if expected_grade else 0)

        header = header_info.get("frame")
        title = header_info.get("title")
        subtitle = header_info.get("subtitle")
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
        return True

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
            "Speicheroption",
            "Möchten Sie eine neue JSON-Datei erstellen oder eine bestehende Datei laden?",
            icon='question',
            default='yes',
        )

        if user_choice == 'yes':
            new_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                title="Neuen Speicherort für die JSON-Datei wählen",
            )
            if new_path:
                SAVE_FILE = new_path
                print(f"Neuer Speicherort gesetzt: {SAVE_FILE}")

        elif user_choice == 'no':
            existing_path = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json")],
                title="Vorhandene JSON-Datei auswählen",
            )
            if existing_path:
                SAVE_FILE = existing_path
                print(f"Bestehende Datei geladen: {SAVE_FILE}")
        else:
            print("Keine Änderung vorgenommen.")

        self.update_file_path_label()
        self.refresh_bin_history()

    def choose_save_directory(self):
        global SAVE_FILE
        directory = filedialog.askdirectory(title="Speicherordner wählen")
        if not directory:
            return

        filename = os.path.basename(SAVE_FILE) if SAVE_FILE else "Tabellenspeicher_neu.json"
        SAVE_FILE = os.path.join(directory, filename)
        self.save_data()
        messagebox.showinfo("Erfolg", f"Speicherpfad gesetzt auf:\n{SAVE_FILE}")
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

    def on_card_right_click(self, row_name: str, col_name: str, card_front: str):
        table_data = self.get_current_table_data()
        if table_data is None:
            return
        cards = table_data["cards"]

        card_dict = find_card(cards[row_name][col_name], card_front)
        if not card_dict:
            return

        editor = tk.Toplevel(self.root)
        editor.title(f"Kartenrückseite - {card_front}")
        editor.geometry("400x300")

        text_widget = tk.Text(editor, wrap="word")
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", card_dict["back"])

        def close_editor(_event=None):
            self._record_history()
            card_dict["back"] = text_widget.get("1.0", "end-1c")
            self.save_data()
            editor.destroy()

        editor.bind("<Escape>", close_editor)

    def on_card_press(self, event, row_name: str, col_name: str, card_front: str):
        if self.delete_mode or self.mark_mode:
            self.drag_data = None
            return
        card_frame = self._get_card_frame(event.widget)
        self.drag_data = {
            "row": row_name,
            "col": col_name,
            "card": card_front,
            "start": (event.x_root, event.y_root),
            "widget": card_frame,
            "moved": False,
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
            self.create_drag_preview(self.drag_data["card"])
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

    def create_drag_preview(self, text: str):
        self.destroy_drag_preview()
        self.drag_preview = tk.Toplevel(self.root)
        self.drag_preview.overrideredirect(True)
        self.drag_preview.attributes("-topmost", True)
        label = tk.Label(
            self.drag_preview,
            text=text,
            bg=CARD_BG,
            fg=TEXT_PRIMARY,
            font=("Segoe UI", 10, "bold"),
            padx=12,
            pady=8,
            bd=1,
            relief="solid",
        )
        label.pack()

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
    app = CardApp(root)
    root.after(100, app.load_data_or_create_new_table)
    root.mainloop()
