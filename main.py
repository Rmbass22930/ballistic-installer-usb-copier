import ctypes
import json
import os
import queue
import shutil
import stat
import sys
import threading
from dataclasses import dataclass
from pathlib import Path

def _ensure_tk_runtime_env() -> None:
    tcl_env = os.environ.get("TCL_LIBRARY")
    tk_env = os.environ.get("TK_LIBRARY")
    if tcl_env and tk_env and Path(tcl_env).exists() and Path(tk_env).exists():
        return

    candidate_roots = []
    for raw_root in (
        getattr(sys, "_MEIPASS", None),
        sys.base_prefix,
        sys.base_exec_prefix,
        Path(sys.executable).resolve().parent,
    ):
        if not raw_root:
            continue
        root = Path(raw_root)
        if root not in candidate_roots:
            candidate_roots.append(root)

    for root in candidate_roots:
        tcl_root = root / "tcl"
        if not tcl_root.exists():
            continue
        tcl_candidates = sorted(path for path in tcl_root.glob("tcl8*") if path.is_dir())
        tk_candidates = sorted(path for path in tcl_root.glob("tk8*") if path.is_dir())
        if tcl_candidates and tk_candidates:
            os.environ.setdefault("TCL_LIBRARY", str(tcl_candidates[-1]))
            os.environ.setdefault("TK_LIBRARY", str(tk_candidates[-1]))
            return

_ensure_tk_runtime_env()

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


DEFAULT_SOURCE = Path(r"F:\ballistic target")
DEFAULT_DESTINATION = "."
SETTINGS_DIR = Path.home() / "AppData" / "Local" / "BallisticInstallerUsbCopier"
SETTINGS_PATH = SETTINGS_DIR / "settings.json"
DRIVE_REFRESH_MS = 2000

DRIVE_UNKNOWN = 0
DRIVE_NO_ROOT_DIR = 1
DRIVE_REMOVABLE = 2
DRIVE_FIXED = 3
DRIVE_REMOTE = 4
DRIVE_CDROM = 5
DRIVE_RAMDISK = 6

THEME = {
    "window": "#ebe8e1",
    "panel": "#f3f0ea",
    "panel_alt": "#e2ddd3",
    "field": "#faf8f3",
    "accent": "#aab2a5",
    "accent_active": "#b7beb2",
    "accent_strong": "#6a7464",
    "button_bg": "#7b8576",
    "button_bg_active": "#697361",
    "button_text": "#f7f6f2",
    "button_text_disabled": "#d9ddd5",
    "border": "#a3aa9d",
    "text": "#2c322a",
    "muted": "#667062",
    "log_bg": "#f7f4ee",
}
UI_FONT = ("Segoe UI", 10, "bold")
UI_FONT_BOLD = ("Segoe UI", 10, "bold")
UI_FONT_LARGE = ("Segoe UI", 11, "bold")
UI_FONT_TITLE = ("Segoe UI", 17, "bold")
UI_FONT_BUTTON = ("Segoe UI", 11, "bold")
UI_FONT_ACTION = ("Segoe UI", 13, "bold")


@dataclass(frozen=True)
class DriveInfo:
    root: str
    label: str
    drive_type: int
    free_bytes: int
    total_bytes: int

    @property
    def drive_letter(self) -> str:
        return self.root[:2]

    @property
    def type_name(self) -> str:
        return {
            DRIVE_REMOVABLE: "USB",
            DRIVE_FIXED: "Drive",
            DRIVE_REMOTE: "Network",
            DRIVE_CDROM: "CD/DVD",
            DRIVE_RAMDISK: "RAM",
        }.get(self.drive_type, "Drive")

    @property
    def display_name(self) -> str:
        label_part = f" {self.label}" if self.label else ""
        return f"{self.drive_letter} {self.type_name}{label_part} ({format_size(self.free_bytes)} free)"


@dataclass(frozen=True)
class CopyPlan:
    selected_files: list[Path]
    drive_roots: list[str]
    folder_name: str
    total_bytes: int
    wipe_before_copy: bool


def format_size(size_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    value = float(size_bytes)

    for unit in units:
        if value < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.1f} {unit}"
        value /= 1024

    return f"{size_bytes} B"


def get_drive_label(root: str) -> str:
    volume_name = ctypes.create_unicode_buffer(261)
    file_system_name = ctypes.create_unicode_buffer(261)
    serial_number = ctypes.c_uint(0)
    max_component_length = ctypes.c_uint(0)
    file_system_flags = ctypes.c_uint(0)

    success = ctypes.windll.kernel32.GetVolumeInformationW(
        ctypes.c_wchar_p(root),
        volume_name,
        len(volume_name),
        ctypes.byref(serial_number),
        ctypes.byref(max_component_length),
        ctypes.byref(file_system_flags),
        file_system_name,
        len(file_system_name),
    )
    if not success:
        return ""
    return volume_name.value.strip()


def get_drive_space(root: str) -> tuple[int, int]:
    free_bytes = ctypes.c_ulonglong(0)
    total_bytes = ctypes.c_ulonglong(0)
    total_free_bytes = ctypes.c_ulonglong(0)

    success = ctypes.windll.kernel32.GetDiskFreeSpaceExW(
        ctypes.c_wchar_p(root),
        ctypes.byref(free_bytes),
        ctypes.byref(total_bytes),
        ctypes.byref(total_free_bytes),
    )
    if not success:
        return 0, 0

    return int(free_bytes.value), int(total_bytes.value)


def list_candidate_drives(excluded_roots: set[str]) -> list[DriveInfo]:
    drive_mask = ctypes.windll.kernel32.GetLogicalDrives()
    detected: list[DriveInfo] = []

    for index, letter in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ"):
        if not (drive_mask & (1 << index)):
            continue

        root = f"{letter}:\\"
        if root in excluded_roots or not Path(root).exists():
            continue

        drive_type = ctypes.windll.kernel32.GetDriveTypeW(ctypes.c_wchar_p(root))
        if drive_type != DRIVE_REMOVABLE:
            continue

        free_bytes, total_bytes = get_drive_space(root)
        detected.append(
            DriveInfo(
                root=root,
                label=get_drive_label(root),
                drive_type=drive_type,
                free_bytes=free_bytes,
                total_bytes=total_bytes,
            )
        )

    detected.sort(key=lambda item: item.root)
    return detected


def prepare_target_file_for_overwrite(target_file: Path) -> bool:
    if not target_file.exists():
        return False

    try:
        target_file.chmod(target_file.stat().st_mode | stat.S_IWRITE)
    except OSError:
        pass

    return True


def get_reclaimable_bytes(selected_files: list[Path], source_root: Path, destination_root: Path) -> int:
    reclaimable = 0
    for source_file in selected_files:
        relative_path = source_file.relative_to(source_root)
        target_file = destination_root / relative_path
        if target_file.exists():
            reclaimable += target_file.stat().st_size
    return reclaimable


def get_net_needed_bytes(selected_files: list[Path], source_root: Path, destination_root: Path) -> int:
    total_bytes = sum(file_path.stat().st_size for file_path in selected_files)
    reclaimable_bytes = get_reclaimable_bytes(selected_files, source_root, destination_root)
    return max(0, total_bytes - reclaimable_bytes)


def clear_drive_root(drive_root: Path) -> tuple[int, list[str]]:
    deleted_entries = 0
    skipped_entries: list[str] = []

    for child in drive_root.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child, onerror=_remove_readonly)
            else:
                try:
                    child.chmod(child.stat().st_mode | stat.S_IWRITE)
                except OSError:
                    pass
                child.unlink()
            deleted_entries += 1
        except OSError:
            skipped_entries.append(child.name)

    return deleted_entries, skipped_entries


def list_drive_root_entries(drive_root: Path) -> list[str]:
    try:
        return sorted(child.name for child in drive_root.iterdir())
    except OSError:
        return []


def _remove_readonly(func, path, _exc_info) -> None:
    try:
        Path(path).chmod(Path(path).stat().st_mode | stat.S_IWRITE)
    except OSError:
        pass
    func(path)


def center_window(window: tk.Misc, width: int, height: int) -> None:
    window.update_idletasks()
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_pos = max(0, (screen_width - width) // 2)
    y_pos = max(0, (screen_height - height) // 2)
    window.geometry(f"{width}x{height}+{x_pos}+{y_pos}")


class ScrollableCheckboxList(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self._vars: list[tuple[tk.BooleanVar, Path]] = []
        self.on_change = None

        self.canvas = tk.Canvas(
            self,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            borderwidth=0,
            height=250,
            background=THEME["field"],
        )
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas, background=THEME["field"])

        self.inner.bind(
            "<Configure>",
            lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self._bind_mousewheel(self.canvas)
        self._bind_mousewheel(self.inner)

    def set_items(self, base_path: Path, files: list[Path]) -> None:
        for child in self.inner.winfo_children():
            child.destroy()

        self._vars.clear()

        if not files:
            empty_label = tk.Label(
                self.inner,
                text="No files loaded.",
                anchor="w",
                justify="left",
                background=THEME["field"],
                foreground=THEME["muted"],
                font=UI_FONT_LARGE,
            )
            empty_label.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
            self._bind_mousewheel(empty_label)
            self._notify_change()
            return

        for row_index, file_path in enumerate(files):
            relative_path = file_path.relative_to(base_path)
            var = tk.BooleanVar(value=True)
            checkbox = tk.Checkbutton(
                self.inner,
                text=str(relative_path),
                variable=var,
                command=self._notify_change,
                anchor="w",
                justify="left",
                background=THEME["field"],
                foreground=THEME["text"],
                activebackground=THEME["field"],
                activeforeground=THEME["text"],
                selectcolor=THEME["field"],
                relief="flat",
                font=UI_FONT_LARGE,
            )
            checkbox.grid(row=row_index, column=0, sticky="w", padx=4, pady=1)
            self._bind_mousewheel(checkbox)
            self._vars.append((var, file_path))

        self._notify_change()

    def select_all(self, value: bool) -> None:
        for var, _ in self._vars:
            var.set(value)
        self._notify_change()

    def select_matching_suffixes(self, suffixes: set[str]) -> None:
        for var, file_path in self._vars:
            var.set(file_path.suffix.lower() in suffixes)
        self._notify_change()

    def set_selected_relative_paths(self, base_path: Path, relative_paths: set[str]) -> None:
        if not relative_paths:
            return

        normalized = {path.lower() for path in relative_paths}
        for var, file_path in self._vars:
            relative_path = str(file_path.relative_to(base_path)).lower()
            var.set(relative_path in normalized)
        self._notify_change()

    def get_selected(self) -> list[Path]:
        return [file_path for var, file_path in self._vars if var.get()]

    def get_selected_relative_paths(self, base_path: Path) -> list[str]:
        return [str(file_path.relative_to(base_path)) for var, file_path in self._vars if var.get()]

    @property
    def count(self) -> int:
        return len(self._vars)

    def _notify_change(self) -> None:
        if self.on_change:
            self.on_change()

    def _bind_mousewheel(self, widget: tk.Misc) -> None:
        widget.bind("<MouseWheel>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event) -> str:
        self.canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"


class DriveSelector(ttk.Frame):
    def __init__(self, master: tk.Misc) -> None:
        super().__init__(master)
        self._vars: dict[str, tk.BooleanVar] = {}
        self._drives: dict[str, DriveInfo] = {}
        self.on_change = None

        self.canvas = tk.Canvas(
            self,
            highlightthickness=1,
            highlightbackground=THEME["border"],
            borderwidth=0,
            height=125,
            background=THEME["panel"],
        )
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style="CardSurface.TFrame")

        self.inner.bind(
            "<Configure>",
            lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self._bind_mousewheel(self.canvas)
        self._bind_mousewheel(self.inner)

    def set_drives(self, drives: list[DriveInfo], selected_roots: set[str], auto_select_new_removable: bool) -> None:
        previous_roots = set(self._drives)
        self._drives = {drive.root: drive for drive in drives}

        for child in self.inner.winfo_children():
            child.destroy()

        next_vars: dict[str, tk.BooleanVar] = {}
        for row_index, drive in enumerate(drives):
            existing_var = self._vars.get(drive.root)
            if existing_var is None:
                default_value = drive.root in selected_roots or (
                    auto_select_new_removable and drive.drive_type == DRIVE_REMOVABLE and drive.root not in previous_roots
                )
                existing_var = tk.BooleanVar(value=default_value)
            else:
                existing_var.set(drive.root in selected_roots or existing_var.get())

            card = ttk.Frame(self.inner, padding=(6, 4), style="CardSurface.TFrame")
            card.grid(row=row_index, column=0, sticky="ew", padx=4, pady=3)
            card.columnconfigure(1, weight=1)
            self._bind_mousewheel(card)

            checkbox = ttk.Checkbutton(card, variable=existing_var, command=self._notify_change)
            checkbox.grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 8))
            self._bind_mousewheel(checkbox)

            title = f"{drive.drive_letter}  {drive.type_name}"
            if drive.label:
                title += f"  {drive.label}"
            title_label = ttk.Label(card, text=title, style="Muted.TLabel")
            title_label.grid(row=0, column=1, sticky="w")
            self._bind_mousewheel(title_label)

            details = f"{format_size(drive.free_bytes)} free of {format_size(drive.total_bytes)}"
            details_label = ttk.Label(card, text=details, style="Hint.TLabel")
            details_label.grid(row=1, column=1, sticky="w")
            self._bind_mousewheel(details_label)
            next_vars[drive.root] = existing_var

        self._vars = next_vars
        self._notify_change()

    def get_selected_roots(self) -> list[str]:
        return [root for root, var in self._vars.items() if var.get()]

    def select_all(self, value: bool) -> None:
        for var in self._vars.values():
            var.set(value)
        self._notify_change()

    def clear(self) -> None:
        self.select_all(False)

    def set_selected_roots(self, roots: set[str]) -> None:
        for root, var in self._vars.items():
            var.set(root in roots)
        self._notify_change()

    @property
    def count(self) -> int:
        return len(self._vars)

    def _notify_change(self) -> None:
        if self.on_change:
            self.on_change()

    def _bind_mousewheel(self, widget: tk.Misc) -> None:
        widget.bind("<MouseWheel>", self._on_mousewheel, add="+")

    def _on_mousewheel(self, event) -> str:
        self.canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"


def load_settings() -> dict:
    try:
        if not SETTINGS_PATH.exists():
            return {}
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(data: dict) -> None:
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass


class App(tk.Tk):
    def __init__(self) -> None:
        _ensure_tk_runtime_env()
        super().__init__()
        self.title("Ballistic Target USB Copier")
        self.minsize(980, 720)
        center_window(self, 1120, 820)

        self.settings = load_settings()
        initial_source = self.settings.get("source_path", str(DEFAULT_SOURCE))
        initial_destination = "."

        self.source_var = tk.StringVar(value=initial_source)
        self.status_var = tk.StringVar(value="Ready to prepare a USB copy job.")
        self.action_var = tk.StringVar(value="Choose installer files and target drives, then review the preview before copying.")
        self.progress_var = tk.DoubleVar(value=0)
        self.selection_var = tk.StringVar(value="No files loaded.")
        self.detected_drives_var = tk.StringVar(value="No USB drives detected yet.")
        self.selected_drives_var = tk.StringVar(value="No target drives selected.")
        self.wipe_drive_var = tk.BooleanVar(value=self.settings.get("wipe_before_copy", True))
        self.wipe_status_var = tk.StringVar(value="")
        self.preview_status_var = tk.StringVar(value="Load installer files and select target drives to build a copy preview.")
        self.preview_status_var = tk.StringVar(value="Load installer files and select target drives to build a copy preview.")

        self.source_path = Path(initial_source)
        self.files: list[Path] = []
        self.detected_drives: list[DriveInfo] = []
        self.copy_queue: queue.Queue[tuple[str, str | float]] = queue.Queue()
        self.copy_thread: threading.Thread | None = None
        self.last_drive_signature: tuple[str, ...] = ()
        self.saved_drive_roots = set(self.settings.get("selected_drive_roots", []))

        self._build_ui()
        self.refresh_files()
        self.detect_drives(quiet=True)
        self.after(DRIVE_REFRESH_MS, self.auto_refresh_drives)
        self.after(100, self.process_queue)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        self._configure_styles()
        self.configure(bg=THEME["window"])
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        main = ttk.Frame(self, padding=22)
        main.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        header = ttk.Frame(main, style="Hero.TFrame", padding=(18, 16))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="USB Installer Copier", style="HeroTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Copy the selected installer files from this computer to one or more USB drives with a quick review step first.",
            style="HeroBody.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        source_frame = ttk.LabelFrame(main, text="Source")
        source_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        source_frame.columnconfigure(1, weight=1)

        ttk.Label(source_frame, text="Source folder on this computer:", style="Source.TLabel").grid(
            row=0, column=0, sticky="w", padx=(14, 8), pady=12
        )
        self.source_entry = ttk.Entry(source_frame, textvariable=self.source_var, style="Source.TEntry", state="readonly")
        self.source_entry.grid(row=0, column=1, sticky="ew", pady=12)
        self.source_entry.bind("<Button-1>", self.open_source_picker, add="+")
        self.source_entry.bind("<Return>", self.open_source_picker, add="+")
        ttk.Button(source_frame, text="Choose Source Folder...", command=self.browse_source).grid(row=0, column=2, padx=8, pady=12)
        ttk.Button(source_frame, text="Refresh", command=self.refresh_files).grid(row=0, column=3, padx=(0, 14), pady=12)

        file_actions = ttk.Frame(main)
        file_actions.grid(row=2, column=0, sticky="ew", pady=(4, 8))
        file_actions.columnconfigure(0, weight=1)

        ttk.Label(file_actions, text="Choose which installer files should be copied to the selected USB drives.").grid(row=0, column=0, sticky="w")
        ttk.Button(file_actions, text="Choose Files...", command=self.open_file_selection_dialog).grid(row=0, column=1, padx=(8, 4))
        ttk.Button(file_actions, text="Select All", command=lambda: self.file_list.select_all(True)).grid(row=0, column=2, padx=4)
        ttk.Button(file_actions, text="Clear Selection", command=lambda: self.file_list.select_all(False)).grid(row=0, column=3)
        ttk.Label(file_actions, textvariable=self.selection_var, style="Summary.TLabel").grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))

        files_frame = ttk.LabelFrame(main, text="Installer Files")
        files_frame.grid(row=3, column=0, sticky="nsew")
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)

        self.file_list = ScrollableCheckboxList(files_frame)
        self.file_list.on_change = self.update_selection_summary
        self.file_list.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        options_frame = ttk.LabelFrame(main, text="USB Targets")
        options_frame.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        options_frame.columnconfigure(1, weight=1)
        options_frame.rowconfigure(1, weight=1)

        ttk.Label(options_frame, text="Detected:").grid(row=0, column=0, sticky="nw", padx=(14, 8), pady=(14, 8))
        ttk.Label(options_frame, textvariable=self.detected_drives_var, wraplength=680, justify="left").grid(
            row=0, column=1, columnspan=2, sticky="w", padx=(0, 14), pady=(14, 8)
        )

        ttk.Label(options_frame, text="Target drives:").grid(row=1, column=0, sticky="nw", padx=(14, 8), pady=(4, 8))

        drive_picker_frame = ttk.Frame(options_frame)
        drive_picker_frame.grid(row=1, column=1, sticky="nsew", padx=(0, 14), pady=(4, 8))
        drive_picker_frame.columnconfigure(0, weight=1)

        self.drive_selector = DriveSelector(drive_picker_frame)
        self.drive_selector.on_change = self.update_selected_drives_summary
        self.drive_selector.grid(row=0, column=0, sticky="nsew")

        drive_buttons = ttk.Frame(options_frame)
        drive_buttons.grid(row=1, column=2, sticky="nw", padx=(0, 14), pady=(4, 8))
        drive_buttons.columnconfigure(0, weight=1)
        ttk.Button(
            drive_buttons,
            text="Detect Drives",
            command=lambda: self.detect_drives(quiet=False, force_select_removable=True),
        ).grid(row=0, column=0, sticky="ew", pady=(0, 3))
        ttk.Button(drive_buttons, text="Select All", command=lambda: self.drive_selector.select_all(True)).grid(
            row=1, column=0, sticky="ew", pady=3
        )
        ttk.Button(drive_buttons, text="Clear Selection", command=self.drive_selector.clear).grid(row=2, column=0, sticky="ew", pady=3)
        self.wipe_toggle_button = ttk.Button(
            drive_buttons,
            text="Erase Before Copy",
            command=self.toggle_wipe_option,
            width=16,
        )
        self.wipe_toggle_button.grid(row=3, column=0, sticky="ew", pady=(3, 0))
        ttk.Label(drive_buttons, textvariable=self.wipe_status_var, style="Hint.TLabel").grid(row=4, column=0, sticky="w", pady=(3, 0))
        self._update_wipe_button_appearance()

        ttk.Label(options_frame, textvariable=self.selected_drives_var, style="Summary.TLabel", wraplength=760, justify="left").grid(
            row=2, column=1, columnspan=2, sticky="w", padx=(0, 14), pady=(4, 6)
        )

        ttk.Label(options_frame, text="Files will be copied to the root of each selected USB drive.").grid(
            row=3, column=0, columnspan=3, sticky="w", padx=(14, 8), pady=(6, 12)
        )

        preview_frame = ttk.LabelFrame(main, text="Copy Preview")
        preview_frame.grid(row=5, column=0, sticky="nsew", pady=(14, 0))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)

        ttk.Label(
            preview_frame,
            textvariable=self.preview_status_var,
            style="Summary.TLabel",
            wraplength=980,
            justify="left",
        ).grid(row=0, column=0, sticky="w", padx=8, pady=(8, 0))

        self.preview_box = tk.Text(
            preview_frame,
            height=10,
            wrap="word",
            state="disabled",
            background=THEME["field"],
            foreground=THEME["text"],
            insertbackground=THEME["text"],
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            font=UI_FONT,
        )
        self.preview_box.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        self.preview_box.bind("<MouseWheel>", self._on_preview_mousewheel, add="+")

        preview_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_box.yview)
        preview_scroll.grid(row=1, column=1, sticky="ns", pady=8, padx=(0, 8))
        self.preview_box.configure(yscrollcommand=preview_scroll.set)

        progress_frame = ttk.Frame(main)
        progress_frame.grid(row=6, column=0, sticky="ew", pady=(14, 0))
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew")
        ttk.Label(progress_frame, textvariable=self.status_var).grid(row=1, column=0, sticky="w", pady=(6, 0))

        log_frame = ttk.LabelFrame(main, text="Copy Log")
        log_frame.grid(row=7, column=0, sticky="nsew", pady=(14, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main.rowconfigure(7, weight=1)

        self.log = tk.Text(
            log_frame,
            height=8,
            wrap="word",
            state="disabled",
            background=THEME["log_bg"],
            foreground=THEME["text"],
            insertbackground=THEME["text"],
            relief="solid",
            borderwidth=1,
            highlightthickness=0,
            font=UI_FONT,
        )
        self.log.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.log.bind("<MouseWheel>", self._on_log_mousewheel, add="+")

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        log_scroll.grid(row=0, column=1, sticky="ns", pady=8, padx=(0, 8))
        self.log.configure(yscrollcommand=log_scroll.set)

        self.bind("<Control-Return>", lambda _: self.start_copy())
        self.bind("<KP_Enter>", lambda _: self.start_copy())

        bottom_bar = ttk.Frame(self, padding=(22, 10, 22, 20))
        bottom_bar.grid(row=1, column=0, sticky="ew")
        bottom_bar.columnconfigure(0, weight=1)

        ttk.Label(bottom_bar, textvariable=self.action_var, style="Summary.TLabel", wraplength=780, justify="left").grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.preview_button = ttk.Button(bottom_bar, text="Review Preview", command=self.preview_copy)
        self.preview_button.grid(row=0, column=1, sticky="e", padx=(12, 0))

        self.copy_now_button = tk.Button(
            bottom_bar,
            text="Start Copy",
            command=self.start_copy,
            font=UI_FONT_ACTION,
            bg=THEME["accent_strong"],
            fg=THEME["button_text"],
            activebackground=THEME["button_bg_active"],
            activeforeground=THEME["button_text"],
            disabledforeground=THEME["button_text_disabled"],
            relief="raised",
            bd=2,
            cursor="hand2",
            pady=10,
            padx=18,
        )
        self.copy_now_button.grid(row=0, column=2, sticky="e", padx=(12, 0))
        ttk.Label(bottom_bar, text="Shortcut: Ctrl+Enter").grid(row=1, column=0, sticky="w")

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=THEME["window"], foreground=THEME["text"])
        style.configure("TFrame", background=THEME["window"])
        style.configure("TLabelFrame", background=THEME["panel"], borderwidth=1, relief="groove")
        style.configure("TLabelFrame.Label", background=THEME["panel"], foreground=THEME["text"])
        style.configure("TLabel", background=THEME["window"], foreground=THEME["text"], font=UI_FONT_LARGE)
        style.configure("Hero.TFrame", background=THEME["panel_alt"], relief="flat")
        style.configure("HeroTitle.TLabel", background=THEME["panel_alt"], foreground=THEME["text"], font=UI_FONT_TITLE)
        style.configure("HeroBody.TLabel", background=THEME["panel_alt"], foreground=THEME["muted"], font=UI_FONT_LARGE)
        style.configure("Source.TLabel", background=THEME["panel"], foreground=THEME["text"], font=UI_FONT_LARGE)
        style.configure("Muted.TLabel", background=THEME["panel"], foreground=THEME["text"], font=UI_FONT_BOLD)
        style.configure("Hint.TLabel", background=THEME["panel"], foreground=THEME["muted"], font=UI_FONT)
        style.configure("Summary.TLabel", background=THEME["window"], foreground=THEME["muted"], font=UI_FONT_BOLD)
        style.configure(
            "TButton",
            background=THEME["button_bg"],
            foreground=THEME["button_text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["button_bg"],
            darkcolor=THEME["button_bg"],
            padding=(8, 5),
            font=UI_FONT_LARGE,
        )
        style.map(
            "TButton",
            background=[("active", THEME["button_bg_active"]), ("pressed", THEME["accent_strong"])],
            foreground=[("active", THEME["button_text"]), ("pressed", THEME["button_text"]), ("disabled", THEME["button_text_disabled"])],
        )
        style.configure(
            "TEntry",
            fieldbackground=THEME["field"],
            background=THEME["field"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
            insertcolor=THEME["text"],
            font=UI_FONT_LARGE,
        )
        style.configure(
            "Source.TEntry",
            fieldbackground=THEME["field"],
            background=THEME["field"],
            foreground=THEME["text"],
            bordercolor=THEME["border"],
            lightcolor=THEME["border"],
            darkcolor=THEME["border"],
            insertcolor=THEME["text"],
            font=UI_FONT_LARGE,
        )
        style.configure(
            "CardSurface.TFrame",
            background=THEME["panel"],
        )
        style.configure("TCheckbutton", background=THEME["panel"], foreground=THEME["text"], font=UI_FONT_LARGE)
        style.map("TCheckbutton", background=[("active", THEME["panel"])])

    def browse_source(self) -> None:
        initial_dir = self.source_var.get() or str(DEFAULT_SOURCE)
        chosen = filedialog.askdirectory(initialdir=initial_dir)
        if chosen:
            self.source_var.set(chosen)
            self.persist_settings()
            self.refresh_files()

    def open_source_picker(self, _event=None) -> str:
        self.browse_source()
        return "break"

    def open_file_selection_dialog(self) -> None:
        if not self.files:
            messagebox.showerror("Choose Files", "No source files are loaded.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Choose Files To Copy")
        dialog.minsize(640, 420)
        dialog.transient(self)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)
        center_window(dialog, 760, 520)

        header = ttk.Frame(dialog, padding=12)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(
            header,
            text="Check the files you want copied to the USB drives, then click Use Selected Files.",
        ).grid(row=0, column=0, sticky="w")

        chooser = ScrollableCheckboxList(dialog)
        chooser.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        chooser.set_items(self.source_path, self.files)
        chooser.set_selected_relative_paths(self.source_path, set(self.file_list.get_selected_relative_paths(self.source_path)))

        footer = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)

        ttk.Button(footer, text="Select All", command=lambda: chooser.select_all(True)).grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="Clear", command=lambda: chooser.select_all(False)).grid(row=0, column=1, sticky="w", padx=(8, 0))

        def apply_selection() -> None:
            self.file_list.set_selected_relative_paths(self.source_path, set(chooser.get_selected_relative_paths(self.source_path)))
            self.update_selection_summary()
            dialog.destroy()

        ttk.Button(footer, text="Cancel", command=dialog.destroy).grid(row=0, column=2, sticky="e", padx=(0, 8))
        ttk.Button(footer, text="Use Selected Files", command=apply_selection).grid(row=0, column=3, sticky="e")

    def refresh_files(self) -> None:
        try:
            source = Path(self.source_var.get()).expanduser()
            if not source.exists() or not source.is_dir():
                raise FileNotFoundError("Source folder does not exist.")

            files = sorted([path for path in source.rglob("*") if path.is_file()])
            if not files:
                raise FileNotFoundError("No files were found in the source folder.")

            self.source_path = source
            self.files = files
            saved_relative_paths = set(self.settings.get("selected_files", []))
            self.file_list.set_items(source, files)
            self.file_list.set_selected_relative_paths(source, saved_relative_paths)
            self.update_selection_summary()
            self.persist_settings()
            self.set_status(f"Loaded {len(files)} file(s) from {source}")
            self.append_log(f"Loaded source folder: {source}")
            self.refresh_preview()
        except Exception as exc:
            self.files = []
            self.file_list.set_items(Path("."), [])
            self.selection_var.set("No files loaded.")
            self.set_status("Unable to load source files.")
            self.refresh_preview()
            messagebox.showerror("Source Folder Error", str(exc))

    def update_selection_summary(self) -> None:
        selected = self.file_list.get_selected()
        if not self.files:
            self.selection_var.set("No files loaded.")
            self.refresh_preview()
            return

        selected_size = sum(file_path.stat().st_size for file_path in selected)
        total_size = sum(file_path.stat().st_size for file_path in self.files)
        self.selection_var.set(
            f"Selected {len(selected)} of {len(self.files)} file(s) | "
            f"{format_size(selected_size)} of {format_size(total_size)}"
        )
        self.persist_settings()
        self.refresh_preview()

    def detect_drives(self, quiet: bool = False, force_select_removable: bool = False) -> None:
        excluded = {
            f"{Path.home().drive}\\",
            f"{self.source_path.drive}\\",
        }
        selected_roots = set(self.drive_selector.get_selected_roots()) or set(self.saved_drive_roots)
        previous_signature = self.last_drive_signature
        self.detected_drives = list_candidate_drives(excluded)
        current_signature = tuple(drive.root for drive in self.detected_drives)
        auto_select_new = force_select_removable or bool(previous_signature) or not selected_roots
        self.drive_selector.set_drives(self.detected_drives, selected_roots, auto_select_new)
        if force_select_removable:
            removable_roots = {drive.root for drive in self.detected_drives if drive.drive_type == DRIVE_REMOVABLE}
            self.drive_selector.set_selected_roots(removable_roots)
        self.last_drive_signature = current_signature

        if self.detected_drives:
            self.detected_drives_var.set(" | ".join(drive.display_name for drive in self.detected_drives))
            selected_count = len(self.drive_selector.get_selected_roots())
            self.set_action_message(
                f"Detected {len(self.detected_drives)} USB drive(s). Choose the targets, review the preview, then start the copy."
            )
            self.set_status(f"Detected {len(self.detected_drives)} USB drive(s); {selected_count} selected.")
        else:
            self.detected_drives_var.set("No USB drives detected.")
            self.set_action_message("No target drives detected. Insert USB drives and run Detect Drives again.")
            self.set_status("No USB drives detected.")
        self.update_selected_drives_summary()

        if not quiet:
            if current_signature != previous_signature:
                self.append_log("Drive detection refreshed. Drive list changed.")
            else:
                self.append_log("Drive detection refreshed.")

    def update_selected_drives_summary(self) -> None:
        selected = self.drive_selector.get_selected_roots()
        self.saved_drive_roots = set(selected)

        if selected:
            labels = []
            for root in selected:
                drive = next((item for item in self.detected_drives if item.root == root), None)
                labels.append(drive.display_name if drive else root)
            self.selected_drives_var.set("Selected targets: " + " | ".join(labels))
        else:
            self.selected_drives_var.set("No target drives selected.")

        self.persist_settings()
        self.refresh_preview()
        self.refresh_preview()

    def on_wipe_option_changed(self) -> None:
        mode = "Erase selected USB drives before copying." if self.wipe_drive_var.get() else "Keep existing USB drive contents and overwrite matching selected files only."
        self._update_wipe_button_appearance()
        self.set_status(mode)
        self.set_action_message(f"{mode} Review the plan if needed, then start the copy.")
        self.persist_settings()
        self.refresh_preview()
        self.refresh_preview()

    def toggle_wipe_option(self) -> None:
        self.wipe_drive_var.set(not self.wipe_drive_var.get())
        self.on_wipe_option_changed()

    def _update_wipe_button_appearance(self) -> None:
        if not hasattr(self, "wipe_toggle_button"):
            return
        is_enabled = bool(self.wipe_drive_var.get())
        self.wipe_toggle_button.configure(text="Erase Before Copy")
        self.wipe_status_var.set(f"Status: {'On' if is_enabled else 'Off'}")

    def auto_refresh_drives(self) -> None:
        if not (self.copy_thread and self.copy_thread.is_alive()):
            self.detect_drives(quiet=True)
        self.after(DRIVE_REFRESH_MS, self.auto_refresh_drives)

    def set_status(self, message: str) -> None:
        self.status_var.set(message)

    def set_action_message(self, message: str) -> None:
        self.action_var.set(message)

    def append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"{message}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_preview_text(self, message: str) -> None:
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", message)
        self.preview_box.see("1.0")
        self.preview_box.configure(state="disabled")

    def _on_log_mousewheel(self, event) -> str:
        self.log.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _on_preview_mousewheel(self, event) -> str:
        self.preview_box.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def refresh_preview(self) -> None:
        if not self.files:
            self.preview_status_var.set("Load installer files and select target drives to build a copy preview.")
            self._set_preview_text("Preview unavailable.\n\nNo installer files are currently loaded from the source folder.")
            return

        selected_files = self.file_list.get_selected()
        if not selected_files:
            self.preview_status_var.set("Select at least one installer file to build the copy plan.")
            self._set_preview_text("Preview unavailable.\n\nSelect one or more installer files from the list above.")
            return

        selected_drives = self.drive_selector.get_selected_roots()
        if not selected_drives:
            self.preview_status_var.set("Select at least one target drive to see where the chosen files will go.")
            self._set_preview_text("Preview unavailable.\n\nNo target drives are currently selected.")
            return

        try:
            plan = self._validate_copy_plan()
        except Exception as exc:
            self.preview_status_var.set("Preview blocked until the current selection issues are fixed.")
            self._set_preview_text(f"Preview unavailable.\n\n{exc}")
            return

        self.preview_status_var.set(
            f"Preview ready for {len(plan.selected_files)} file(s) across {len(plan.drive_roots)} drive(s)."
        )
        self._set_preview_text(self._build_copy_preview_text(plan))

    def set_copy_buttons_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.preview_button.configure(state=state)
        self.copy_now_button.configure(state=state)

    def _drive_display_name(self, root: str) -> str:
        drive = next((item for item in self.detected_drives if item.root == root), None)
        return drive.display_name if drive else root

    def _destination_root(self, drive_root: str, folder_name: str) -> Path:
        drive_path = Path(drive_root)
        return drive_path if folder_name == "." else drive_path / folder_name

    def _validate_copy_plan(self) -> CopyPlan:
        if not self.files:
            raise ValueError("No source files are loaded.")

        selected_files = self.file_list.get_selected()
        if not selected_files:
            raise ValueError("Select at least one file.")

        drive_roots = self.drive_selector.get_selected_roots()
        if not drive_roots:
            raise ValueError("Select at least one target drive.")
        if self.source_path.drive and any(root.startswith(self.source_path.drive.upper()) for root in drive_roots):
            raise ValueError("Do not copy back to the source drive.")

        folder_name = DEFAULT_DESTINATION
        if folder_name != "." and any(sep in folder_name for sep in ('*', '?', '"', '<', '>', '|', ':', '/', '\\')):
            raise ValueError("Destination folder contains invalid characters.")

        wipe_before_copy = bool(self.wipe_drive_var.get())
        total_bytes = sum(file_path.stat().st_size for file_path in selected_files)
        low_space_roots = []
        for drive_root in drive_roots:
            free_bytes, total_capacity_bytes = get_drive_space(drive_root)
            destination_root = self._destination_root(drive_root, folder_name)
            needed_bytes = total_bytes if wipe_before_copy else get_net_needed_bytes(selected_files, self.source_path, destination_root)
            available_bytes = total_capacity_bytes if wipe_before_copy else free_bytes
            if available_bytes and available_bytes < needed_bytes:
                mode_label = "total capacity" if wipe_before_copy else "free space"
                low_space_roots.append(
                    f"{self._drive_display_name(drive_root)} ({format_size(available_bytes)} {mode_label}, needs {format_size(needed_bytes)})"
                )

        if low_space_roots:
            shortage_label = "even after wiping the drive first" if wipe_before_copy else "without erasing the drive first"
            raise ValueError(
                "Not enough space on: "
                + ", ".join(low_space_roots)
                + f". The selected files total about {format_size(total_bytes)} {shortage_label}."
            )

        return CopyPlan(
            selected_files=selected_files,
            drive_roots=drive_roots,
            folder_name=folder_name,
            total_bytes=total_bytes,
            wipe_before_copy=wipe_before_copy,
        )

    def _build_copy_preview_text(self, plan: CopyPlan) -> str:
        relative_paths = [str(path.relative_to(self.source_path)) for path in plan.selected_files]
        lines = [
            f"Mode: {'Erase and copy' if plan.wipe_before_copy else 'Copy without wiping'}",
            f"Files selected: {len(plan.selected_files)}",
            f"Total size: {format_size(plan.total_bytes)}",
            f"Target drives: {len(plan.drive_roots)}",
            "",
            "Files to copy:",
        ]
        lines.extend(f"- {relative_path}" for relative_path in relative_paths)

        for drive_root in plan.drive_roots:
            destination_root = self._destination_root(drive_root, plan.folder_name)
            lines.extend(
                [
                    "",
                    f"{self._drive_display_name(drive_root)}",
                    f"Destination: {destination_root}",
                ]
            )
            if plan.wipe_before_copy:
                existing_entries = list_drive_root_entries(Path(drive_root))
                if existing_entries:
                    lines.append("Top-level items that will be deleted:")
                    lines.extend(f"- {entry}" for entry in existing_entries)
                else:
                    lines.append("Drive root is already empty.")
            else:
                lines.append("Existing files and folders will be kept.")
                overlapping = []
                for source_file in plan.selected_files:
                    relative_path = source_file.relative_to(self.source_path)
                    if (destination_root / relative_path).exists():
                        overlapping.append(str(relative_path))
                if overlapping:
                    lines.append("Existing selected files that will be overwritten:")
                    lines.extend(f"- {relative_path}" for relative_path in overlapping)
                else:
                    lines.append("No selected files will be overwritten.")

        return "\n".join(lines)

    def _confirm_copy_plan(self, plan: CopyPlan) -> bool:
        drive_lines = "\n".join(f"- {self._drive_display_name(root)}" for root in plan.drive_roots)
        mode_line = "This will erase each selected USB drive before copying." if plan.wipe_before_copy else (
            "This will keep existing USB contents and overwrite matching selected files only."
        )
        return messagebox.askyesno(
            "Confirm USB Copy",
            f"{mode_line}\n\nTarget drives:\n{drive_lines}\n\nFiles selected: {len(plan.selected_files)}\nTotal size: {format_size(plan.total_bytes)}\n\nDo you want to continue?",
            icon="warning" if plan.wipe_before_copy else "question",
        )

    def preview_copy(self) -> None:
        try:
            plan = self._validate_copy_plan()
            preview_text = self._build_copy_preview_text(plan)
            self.preview_status_var.set(
                f"Preview ready for {len(plan.selected_files)} file(s) across {len(plan.drive_roots)} drive(s)."
            )
            self._set_preview_text(preview_text)
            self.append_log(
                f"Previewed copy plan: {len(plan.selected_files)} file(s), {len(plan.drive_roots)} drive(s), "
                f"mode {'wipe' if plan.wipe_before_copy else 'keep'}."
            )
            self.set_action_message("Preview updated in the Copy Preview panel.")
        except Exception as exc:
            self.preview_status_var.set("Preview blocked until the current selection issues are fixed.")
            self._set_preview_text(f"Preview unavailable.\n\n{exc}")
            self.set_action_message(str(exc))
            self.append_log(f"Preview error: {exc}")
            self.set_status("Preview error.")

    def start_copy(self) -> None:
        if self.copy_thread and self.copy_thread.is_alive():
            return

        try:
            self.set_action_message("Copy requested. Validating selections...")
            self.append_log("Copy requested.")

            plan = self._validate_copy_plan()
            confirmation = self._confirm_copy_plan(plan)
            if not confirmation:
                self.set_action_message("Copy cancelled.")
                self.append_log("Copy cancelled by user.")
                return

            self.set_copy_buttons_enabled(False)
            self.progress_var.set(0)
            self.persist_settings()
            self.set_status(f"Copying {len(plan.selected_files)} file(s) to {len(plan.drive_roots)} drive(s)...")
            self.set_action_message(f"Copying {len(plan.selected_files)} file(s) to {len(plan.drive_roots)} drive(s)...")
            self.append_log(
                f"Starting copy job: {len(plan.selected_files)} file(s), {format_size(plan.total_bytes)} total, "
                f"targets {', '.join(plan.drive_roots)}, mode {'wipe' if plan.wipe_before_copy else 'keep'}"
            )

            self.copy_thread = threading.Thread(
                target=self.copy_files,
                args=(plan.selected_files, plan.drive_roots, plan.folder_name, plan.wipe_before_copy),
                daemon=True,
            )
            self.copy_thread.start()
        except Exception as exc:
            self.set_action_message(str(exc))
            self.append_log(f"Validation error: {exc}")
            messagebox.showerror("Copy Error", str(exc))

    def copy_files(self, selected_files: list[Path], drive_roots: list[str], folder_name: str, wipe_before_copy: bool = True) -> None:
        total_steps = len(selected_files) * len(drive_roots)
        completed_steps = 0
        total_bytes = sum(file_path.stat().st_size for file_path in selected_files)
        overwritten_count = 0

        try:
            for drive_root in drive_roots:
                drive_path = Path(drive_root)
                destination_root = self._destination_root(drive_root, folder_name)
                if wipe_before_copy:
                    deleted_entries, skipped_entries = clear_drive_root(drive_path)
                    self.copy_queue.put(("log", f"Cleared {deleted_entries} existing item(s) from {drive_root}"))
                    if skipped_entries:
                        self.copy_queue.put(("log", f"Skipped undeletable item(s) on {drive_root}: {', '.join(skipped_entries)}"))
                else:
                    self.copy_queue.put(("log", f"Keeping existing files on {drive_root}"))
                destination_root.mkdir(parents=True, exist_ok=True)

                self.copy_queue.put(("log", f"Copying to {destination_root}"))

                for source_file in selected_files:
                    relative_path = source_file.relative_to(self.source_path)
                    target_file = destination_root / relative_path
                    target_file.parent.mkdir(parents=True, exist_ok=True)
                    was_existing = prepare_target_file_for_overwrite(target_file)
                    if was_existing:
                        target_file.unlink()
                    shutil.copy2(source_file, target_file)

                    completed_steps += 1
                    percent = (completed_steps / total_steps) * 100
                    if was_existing:
                        overwritten_count += 1
                        self.copy_queue.put(("status", f"Overwrote {relative_path} on {drive_root}"))
                        self.copy_queue.put(("log", f"Overwrote {relative_path} at {target_file}"))
                    else:
                        self.copy_queue.put(("status", f"Copied {relative_path} to {drive_root}"))
                        self.copy_queue.put(("log", f"Copied {relative_path} to {target_file}"))
                    self.copy_queue.put(("progress", percent))

            overwrite_suffix = f" including {overwritten_count} overwrite(s)" if overwritten_count else ""
            self.copy_queue.put(
                ("done", f"Finished copying {len(selected_files)} file(s) totaling {format_size(total_bytes)}{overwrite_suffix}.")
            )
        except Exception as exc:
            self.copy_queue.put(("error", str(exc)))

    def process_queue(self) -> None:
        try:
            while True:
                kind, payload = self.copy_queue.get_nowait()

                if kind == "log":
                    self.append_log(str(payload))
                elif kind == "status":
                    self.set_status(str(payload))
                elif kind == "progress":
                    self.progress_var.set(float(payload))
                elif kind == "done":
                    self.progress_var.set(100)
                    self.set_status(str(payload))
                    self.set_action_message(str(payload))
                    self.append_log(str(payload))
                    self.set_copy_buttons_enabled(True)
                    messagebox.showinfo("Complete", str(payload))
                elif kind == "error":
                    self.set_status("Copy failed.")
                    self.set_action_message(f"Copy failed: {payload}")
                    self.append_log(f"Error: {payload}")
                    self.set_copy_buttons_enabled(True)
                    messagebox.showerror("Copy Failed", str(payload))
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue)

    def persist_settings(self) -> None:
        selected_files = [str(path.relative_to(self.source_path)) for path in self.file_list.get_selected()] if self.files else []
        data = {
            "source_path": self.source_var.get().strip() or str(DEFAULT_SOURCE),
            "destination_folder": DEFAULT_DESTINATION,
            "wipe_before_copy": bool(self.wipe_drive_var.get()),
            "selected_drive_roots": sorted(self.saved_drive_roots),
            "selected_files": selected_files,
        }
        if data != self.settings:
            save_settings(data)
            self.settings = data

    def on_close(self) -> None:
        self.persist_settings()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
