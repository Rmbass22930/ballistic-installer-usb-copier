import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import tkinter as tk


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import main


class SettingsTests(unittest.TestCase):
    def test_save_and_load_settings_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            settings_dir = Path(temp_dir) / "settings"
            settings_path = settings_dir / "settings.json"
            payload = {
                "source_path": r"F:\ballistic target",
                "destination_folder": ".",
                "wipe_before_copy": False,
                "selected_drive_roots": [r"G:\\"],
                "selected_files": ["BallisticTargetSetup.exe"],
            }

            with mock.patch.object(main, "SETTINGS_DIR", settings_dir), mock.patch.object(main, "SETTINGS_PATH", settings_path):
                main.save_settings(payload)
                loaded = main.load_settings()

            self.assertEqual(payload, loaded)


class WidgetTests(unittest.TestCase):
    def test_checkbox_list_restores_saved_relative_paths(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            base_path = Path(temp_dir)
            exe_file = base_path / "BallisticTargetSetup.exe"
            txt_file = base_path / "notes.txt"
            exe_file.write_text("exe", encoding="utf-8")
            txt_file.write_text("txt", encoding="utf-8")

            root = tk.Tk()
            root.withdraw()
            try:
                widget = main.ScrollableCheckboxList(root)
                widget.set_items(base_path, [exe_file, txt_file])
                widget.set_selected_relative_paths(base_path, {"BallisticTargetSetup.exe"})

                selected = widget.get_selected()
                self.assertEqual([exe_file], selected)
            finally:
                root.destroy()


class AppTests(unittest.TestCase):
    def test_refresh_files_loads_nested_source_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            nested_dir = source_dir / "nested" / "payload"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            nested_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")
            (nested_dir / "README.txt").write_text("readme", encoding="utf-8")
            (nested_dir / "InstallBallistic.exe").write_text("installer", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "selected_drive_roots": [],
                        "selected_files": [],
                    }
                ),
                encoding="utf-8",
            )

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=[]),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    loaded = [str(path.relative_to(source_dir)) for path in app.files]
                    self.assertEqual(
                        [
                            "BallisticTargetSetup.exe",
                            str(Path("nested") / "payload" / "InstallBallistic.exe"),
                            str(Path("nested") / "payload" / "README.txt"),
                        ],
                        loaded,
                    )
                finally:
                    app.on_close()

    def test_app_restores_saved_selection_and_detected_drives(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")
            (source_dir / "README.txt").write_text("txt", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [r"G:\\"],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(root=r"G:\\", label="TestUSB", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024, total_bytes=2048),
                main.DriveInfo(root=r"H:\\", label="Backup", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024, total_bytes=2048),
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()

                    self.assertEqual([r"G:\\"], app.drive_selector.get_selected_roots())
                    selected_names = [path.name for path in app.file_list.get_selected()]
                    self.assertEqual(["BallisticTargetSetup.exe"], selected_names)
                    self.assertFalse(app.wipe_drive_var.get())
                finally:
                    app.on_close()

    def test_do_it_button_invokes_handler_path(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": True,
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=[]),
                mock.patch.object(main.messagebox, "showerror") as showerror,
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.copy_now_button.invoke()
                    self.assertTrue(showerror.called)
                    self.assertIn("Select at least one target drive.", showerror.call_args.args[1])
                finally:
                    app.on_close()

    def test_copy_now_button_starts_copy_thread_with_valid_choices(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [r"G:\\"],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(root=r"G:\\", label="TestUSB", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024 * 1024, total_bytes=2048 * 1024)
            ]
            created_threads = []

            class FakeThread:
                def __init__(self, target=None, args=(), daemon=None):
                    self.target = target
                    self.args = args
                    self.daemon = daemon
                    self.started = False
                    created_threads.append(self)

                def start(self):
                    self.started = True

                def is_alive(self):
                    return self.started

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
                mock.patch.object(main, "get_drive_space", return_value=(1024 * 1024, 2048 * 1024)),
                mock.patch.object(main.threading, "Thread", FakeThread),
                mock.patch.object(main.messagebox, "askyesno", return_value=True),
                mock.patch.object(main.messagebox, "showerror") as showerror,
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.copy_now_button.invoke()

                    self.assertFalse(showerror.called)
                    self.assertEqual(1, len(created_threads))
                    self.assertTrue(created_threads[0].started)
                    self.assertEqual(app.copy_files, created_threads[0].target)
                    self.assertEqual([installer], list(created_threads[0].args[0]))
                    self.assertEqual([r"G:\\"], list(created_threads[0].args[1]))
                    self.assertEqual(".", created_threads[0].args[2])
                    self.assertFalse(created_threads[0].args[3])
                finally:
                    app.on_close()

    def test_copy_confirmation_includes_selected_drives_and_mode(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [r"G:\\"],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(root=r"G:\\", label="TestUSB", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024 * 1024, total_bytes=2048 * 1024)
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
                mock.patch.object(main, "get_drive_space", return_value=(1024 * 1024, 2048 * 1024)),
                mock.patch.object(main.threading, "Thread"),
                mock.patch.object(main.messagebox, "askyesno", return_value=False) as askyesno,
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.start_copy()

                    self.assertTrue(askyesno.called)
                    self.assertIn("keep existing USB contents", askyesno.call_args.args[1])
                    self.assertIn("G: USB TestUSB", askyesno.call_args.args[1])
                finally:
                    app.on_close()

    def test_browse_source_updates_source_path_from_picker(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            alt_source_dir = temp_root / "alt-source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            alt_source_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")
            (alt_source_dir / "AltSetup.exe").write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=[]),
                mock.patch.object(main.filedialog, "askdirectory", return_value=str(alt_source_dir)) as askdirectory,
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.browse_source()

                    self.assertEqual(str(alt_source_dir), app.source_var.get())
                    self.assertEqual(alt_source_dir, app.source_path)
                    self.assertEqual(["AltSetup.exe"], [path.name for path in app.files])
                    self.assertEqual(str(source_dir), askdirectory.call_args.kwargs["initialdir"])
                finally:
                    app.on_close()

    def test_erase_first_button_toggles_wipe_mode(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": True,
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=[]),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    self.assertTrue(app.wipe_drive_var.get())
                    self.assertEqual("Erase Before Copy", app.wipe_toggle_button.cget("text"))
                    self.assertEqual("Status: On", app.wipe_status_var.get())

                    app.wipe_toggle_button.invoke()

                    self.assertFalse(app.wipe_drive_var.get())
                    self.assertEqual("Erase Before Copy", app.wipe_toggle_button.cget("text"))
                    self.assertEqual("Status: Off", app.wipe_status_var.get())
                finally:
                    app.on_close()

    def test_preview_builder_shows_keep_mode_and_overwrite_details(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            target_root = temp_root / "usb"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            target_root.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            readme = source_dir / "README.txt"
            installer.write_text("exe", encoding="utf-8")
            readme.write_text("readme", encoding="utf-8")
            (target_root / "BallisticTargetSetup.exe").write_text("old", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [str(target_root)],
                        "selected_files": ["BallisticTargetSetup.exe", "README.txt"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(
                    root=str(target_root),
                    label="TestUSB",
                    drive_type=main.DRIVE_REMOVABLE,
                    free_bytes=1024 * 1024,
                    total_bytes=2048 * 1024,
                )
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    plan = main.CopyPlan(
                        selected_files=[installer, readme],
                        drive_roots=[str(target_root)],
                        folder_name=".",
                        total_bytes=installer.stat().st_size + readme.stat().st_size,
                        wipe_before_copy=False,
                    )
                    preview_text = app._build_copy_preview_text(plan)
                    self.assertIn("Mode: Copy without wiping", preview_text)
                    self.assertIn("Existing files and folders will be kept.", preview_text)
                    self.assertIn("Existing selected files that will be overwritten:", preview_text)
                    self.assertIn("- BallisticTargetSetup.exe", preview_text)
                    self.assertIn("- README.txt", preview_text)
                finally:
                    app.on_close()

    def test_preview_button_updates_inline_preview_panel(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("exe", encoding="utf-8")
            target_root = r"G:\\"

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [str(target_root)],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(
                    root=str(target_root),
                    label="TestUSB",
                    drive_type=main.DRIVE_REMOVABLE,
                    free_bytes=1024 * 1024,
                    total_bytes=2048 * 1024,
                )
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
                mock.patch.object(main, "get_drive_space", return_value=(1024 * 1024, 2048 * 1024)),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.preview_button.invoke()

                    preview_text = app.preview_box.get("1.0", "end")
                    self.assertIn("Preview ready for 1 file(s) across 1 drive(s).", app.preview_status_var.get())
                    self.assertIn("Mode: Copy without wiping", preview_text)
                    self.assertIn("Destination: G:\\", preview_text)
                finally:
                    app.on_close()

    def test_copy_confirmation_lists_selected_drives_and_keep_mode(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [r"G:\\"],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(root=r"G:\\", label="TestUSB", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024 * 1024, total_bytes=2048 * 1024)
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
                mock.patch.object(main, "get_drive_space", return_value=(1024 * 1024, 2048 * 1024)),
                mock.patch.object(main.threading, "Thread"),
                mock.patch.object(main.messagebox, "askyesno", return_value=False) as askyesno,
                mock.patch.object(main.messagebox, "showerror") as showerror,
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.copy_now_button.invoke()

                    self.assertFalse(showerror.called)
                    self.assertTrue(askyesno.called)
                    confirm_text = askyesno.call_args.args[1]
                    self.assertIn("keep existing USB contents", confirm_text)
                    self.assertIn("G: USB TestUSB", confirm_text)
                finally:
                    app.on_close()

    def test_detect_drives_updates_action_text_and_selects_removable(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": True,
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(root=r"G:\\", label="TestUSB", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024, total_bytes=2048),
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    app.drive_selector.clear()
                    app.detect_drives(quiet=False, force_select_removable=True)

                    self.assertEqual([r"G:\\"], app.drive_selector.get_selected_roots())
                    self.assertIn("start the copy", app.action_var.get().lower())
                    self.assertIn("Detected 1 USB drive", app.action_var.get())
                finally:
                    app.on_close()

    def test_drive_selector_handles_thirteen_usb_drives(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            (source_dir / "BallisticTargetSetup.exe").write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(
                    root=fr"{chr(ord('G') + index)}:\\",
                    label=f"USB{index + 1}",
                    drive_type=main.DRIVE_REMOVABLE,
                    free_bytes=1024 * 1024,
                    total_bytes=2048 * 1024,
                )
                for index in range(13)
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    self.assertEqual(13, app.drive_selector.count)
                    self.assertEqual(13, len(app.drive_selector.get_selected_roots()))
                finally:
                    app.on_close()

    def test_preview_text_lists_copy_and_delete_details(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("exe", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": True,
                        "selected_drive_roots": [r"G:\\"],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            fake_drives = [
                main.DriveInfo(root=r"G:\\", label="TestUSB", drive_type=main.DRIVE_REMOVABLE, free_bytes=1024 * 1024, total_bytes=2048 * 1024)
            ]

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=fake_drives),
                mock.patch.object(main, "list_drive_root_entries", return_value=["old.exe", "logs"]),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.update_idletasks()
                    plan = main.CopyPlan([installer], [r"G:\\"], ".", installer.stat().st_size, True)
                    preview = app._build_copy_preview_text(plan)

                    self.assertIn("Mode: Erase and copy", preview)
                    self.assertIn("Files to copy:", preview)
                    self.assertIn("- BallisticTargetSetup.exe", preview)
                    self.assertIn("Top-level items that will be deleted:", preview)
                    self.assertIn("- old.exe", preview)
                finally:
                    app.on_close()

    def test_copy_files_overwrites_existing_target_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            target_root = temp_root / "usb"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            target_root.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("new-version", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            existing_target = target_root / "BallisticTargetSetup.exe"
            existing_target.write_text("old-version", encoding="utf-8")

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=[]),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.copy_files([installer], [str(target_root)], ".")

                    self.assertEqual("new-version", existing_target.read_text(encoding="utf-8"))
                    queued = []
                    while not app.copy_queue.empty():
                        queued.append(app.copy_queue.get_nowait())
                    self.assertTrue(any(kind == "log" and "Cleared" in str(payload) for kind, payload in queued))
                    self.assertTrue(any(kind == "done" and "Finished copying" in str(payload) for kind, payload in queued))
                finally:
                    app.on_close()

    def test_copy_files_can_keep_existing_nonselected_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            target_root = temp_root / "usb"
            settings_dir = temp_root / "settings"
            settings_path = settings_dir / "settings.json"
            source_dir.mkdir(parents=True)
            target_root.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("new-version", encoding="utf-8")
            retained_file = target_root / "KeepMe.txt"
            retained_file.write_text("retain", encoding="utf-8")
            existing_target = target_root / "BallisticTargetSetup.exe"
            existing_target.write_text("old-version", encoding="utf-8")

            settings_dir.mkdir(parents=True)
            settings_path.write_text(
                json.dumps(
                    {
                        "source_path": str(source_dir),
                        "destination_folder": ".",
                        "wipe_before_copy": False,
                        "selected_drive_roots": [],
                        "selected_files": ["BallisticTargetSetup.exe"],
                    }
                ),
                encoding="utf-8",
            )

            with (
                mock.patch.object(main, "SETTINGS_DIR", settings_dir),
                mock.patch.object(main, "SETTINGS_PATH", settings_path),
                mock.patch.object(main, "list_candidate_drives", return_value=[]),
            ):
                app = main.App()
                app.withdraw()
                try:
                    app.copy_files([installer], [str(target_root)], ".", wipe_before_copy=False)

                    self.assertEqual("new-version", existing_target.read_text(encoding="utf-8"))
                    self.assertEqual("retain", retained_file.read_text(encoding="utf-8"))
                    queued = []
                    while not app.copy_queue.empty():
                        queued.append(app.copy_queue.get_nowait())
                    self.assertTrue(any(kind == "log" and "Keeping existing files" in str(payload) for kind, payload in queued))
                    self.assertFalse(any(kind == "log" and "Cleared" in str(payload) for kind, payload in queued))
                finally:
                    app.on_close()

    def test_clear_drive_root_deletes_existing_entries(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            drive_root = Path(temp_dir)
            (drive_root / "old.txt").write_text("old", encoding="utf-8")
            nested_dir = drive_root / "nested"
            nested_dir.mkdir()
            (nested_dir / "child.txt").write_text("child", encoding="utf-8")

            deleted_entries, skipped_entries = main.clear_drive_root(drive_root)

            self.assertEqual(2, deleted_entries)
            self.assertEqual([], skipped_entries)
            self.assertEqual([], list(drive_root.iterdir()))

    def test_net_needed_bytes_allows_for_existing_target_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_root = Path(temp_dir)
            source_dir = temp_root / "source"
            target_root = temp_root / "usb"
            source_dir.mkdir(parents=True)
            target_root.mkdir(parents=True)
            installer = source_dir / "BallisticTargetSetup.exe"
            installer.write_text("new-version", encoding="utf-8")

            existing_target = target_root / "BallisticTargetSetup.exe"
            existing_target.write_text("old-version", encoding="utf-8")

            net_needed = main.get_net_needed_bytes([installer], source_dir, target_root)
            self.assertEqual(0, net_needed)


if __name__ == "__main__":
    unittest.main()
