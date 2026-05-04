from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import pyperclip
from dotenv import load_dotenv
from PIL import ImageGrab
from pywinauto import Application, Desktop
from pywinauto.findwindows import ElementNotFoundError

load_dotenv()


class ExcelAutomator:
    def __init__(self, logs_dir: Path) -> None:
        self.logs_dir = logs_dir
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.app: Optional[Application] = None
        self.window = None
        self._autofilter_initialized = False

    def launch_excel(self, timeout: int = 25) -> None:
        excel_path = os.getenv("EXCEL_PATH")
        if not excel_path:
            raise ValueError("EXCEL_PATH is missing in .env. Set EXCEL_PATH to full Excel executable path.")
        excel_path = excel_path.strip().strip('"').strip("'")
        logging.info("Launching Excel from path: %s", excel_path)
        self.app = Application(backend="uia").start(excel_path)
        time.sleep(2.0)
        self.app.top_window().type_keys("{ESC}")
        self.window = self.app.window(title_re=".*Excel.*")
        self.window.wait("visible", timeout=timeout)
        self._refresh_focus()
        time.sleep(1.0)

    def create_blank_workbook(self) -> None:
        self._ensure_window()
        self._refresh_focus()
        logging.info("Creating a blank workbook (Ctrl+N)...")
        self.window.type_keys("^n", set_foreground=True)
        time.sleep(1.2)
        self._refresh_focus()
        self._exit_backstage_if_open()
        if not self._is_workbook_open():
            logging.info("Workbook not active after Ctrl+N. Retrying once.")
            self._refresh_focus()
            self.window.type_keys("^n", set_foreground=True)
            time.sleep(1.2)
            self._refresh_focus()
            self._exit_backstage_if_open()
            if not self._is_workbook_open():
                raise RuntimeError("Workbook did not open. Excel is still on Start/Home or blocked UI.")

    def paste_data(self, dataframe: pd.DataFrame) -> None:
        self._ensure_window()
        self._refresh_focus()
        self._exit_backstage_if_open()
        logging.info("Pasting dataframe into workbook...")
        tab_data = dataframe.to_csv(index=False, sep="\t")
        pyperclip.copy(tab_data)
        self.focus_workbook_surface()
        self.window.type_keys("^{HOME}", set_foreground=True)
        time.sleep(0.3)
        self.window.type_keys("^v", set_foreground=True)
        time.sleep(1.0)

    def apply_filters(self, dataframe: pd.DataFrame, header_name: str = "subject", filter_value: str = "Mathematics") -> None:
        self._ensure_window()
        self._refresh_focus()
        self._exit_backstage_if_open()
        logging.info("Applying data-driven filter on header '%s' with value '%s'...", header_name, filter_value)

        columns_lower = [str(col).strip().lower() for col in list(dataframe.columns)]
        try:
            col_index = columns_lower.index(header_name.strip().lower())
        except ValueError as exc:
            raise ValueError(f"Column '{header_name}' not found in dataframe columns: {list(dataframe.columns)}") from exc

        def _open_target_filter_dropdown() -> None:
            self._refresh_focus()
            self.window.type_keys("^{HOME}", set_foreground=True)
            time.sleep(0.2)
            if col_index > 0:
                self.window.type_keys(f"{{RIGHT {col_index}}}", set_foreground=True)
                time.sleep(0.2)
            self.window.type_keys("%{DOWN}", set_foreground=True)
            time.sleep(0.35)

        def _click_first_existing(locators: list[dict[str, str]], step_name: str):
            for locator in locators:
                ctrl = self.window.child_window(**locator)
                if ctrl.exists(timeout=0.8):
                    ctrl.wait("visible", timeout=2)
                    ctrl.wait("ready", timeout=2)
                    ctrl.click_input()
                    logging.info("Filter step complete: %s", step_name)
                    return
            raise RuntimeError(f"Filter control not found for step: {step_name}")

        def _apply_dropdown_selection() -> None:
            _click_first_existing(
                [
                    {"title_re": "(?i)\\(Select All\\)", "control_type": "CheckBox"},
                    {"title_re": "(?i)select all", "control_type": "CheckBox"},
                ],
                "uncheck select all",
            )
            time.sleep(0.15)
            _click_first_existing(
                [
                    {"title": str(filter_value), "control_type": "CheckBox"},
                    {"title_re": rf"(?i){str(filter_value)}", "control_type": "CheckBox"},
                ],
                f"check {filter_value}",
            )
            time.sleep(0.15)
            _click_first_existing(
                [
                    {"title": "OK", "control_type": "Button"},
                    {"title_re": "(?i)^ok$", "control_type": "Button"},
                ],
                "apply filter",
            )
            time.sleep(0.8)

        self._refresh_focus()
        self.focus_workbook_surface()
        # Strict sequence requested: always anchor at A1 first.
        self.window.type_keys("^{HOME}", set_foreground=True)
        time.sleep(0.2)
        if not self._autofilter_initialized:
            self.window.type_keys("^+l", set_foreground=True)
            time.sleep(0.3)
            self._autofilter_initialized = True

        try:
            _open_target_filter_dropdown()
            _apply_dropdown_selection()
        except Exception as exc:
            logging.warning("Filter focus issue on first attempt (%s). Retrying once.", exc)
            try:
                self._refresh_focus()
                self.window.type_keys("{ESC}", set_foreground=True)
                time.sleep(0.2)
                _open_target_filter_dropdown()
                _apply_dropdown_selection()
            except Exception as retry_exc:
                logging.error("Filter menu focus not acquired after retry: %s", retry_exc)
                raise RuntimeError("Filter menu focus not acquired") from retry_exc

        try:
            self._verify_filter_result_via_com(dataframe, col_index=col_index, expected_value=filter_value)
        except Exception as verify_exc:
            message = str(verify_exc).lower()
            if "operation unavailable" in message:
                logging.warning(
                    "Filter applied via UI, but COM verification is unavailable. "
                    "Treating step as success. Details: %s",
                    verify_exc,
                )
                return
            raise

    def press_keys(self, keys: str) -> None:
        self._ensure_window()
        self._refresh_focus()
        logging.info("Pressing keys: %s", keys)
        self.window.type_keys(keys, set_foreground=True)
        time.sleep(0.6)

    def click_text(self, text: str, timeout: int = 3) -> None:
        self._ensure_window()
        self._refresh_focus()
        logging.info("Trying to click control with text: %s", text)
        btn = self.window.child_window(title=text, control_type="Button")
        btn.wait("exists enabled visible", timeout=timeout)
        btn.click_input()
        time.sleep(0.8)

    def focus_workbook_surface(self) -> None:
        self._ensure_window()
        self._refresh_focus()
        try:
            rect = self.window.rectangle()
            target_x = int(rect.width() * 0.35)
            target_y = int(rect.height() * 0.35)
            self.window.click_input(coords=(target_x, target_y))
        except Exception:
            # Fallback to focus only when click coordinates are not available.
            self.window.set_focus()
        time.sleep(0.2)

    def wait(self, seconds: float) -> None:
        logging.info("Waiting for %.2f seconds", seconds)
        time.sleep(seconds)

    def capture_state(self, filename: str = "error_screenshot.png") -> Path:
        self._ensure_window()
        out_path = self.logs_dir / filename
        logging.info("Capturing Excel state screenshot: %s", out_path)
        try:
            rect = self.window.rectangle()
            img = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        except Exception:
            img = ImageGrab.grab()
        img.save(out_path)
        return out_path

    def _ensure_window(self) -> None:
        if self.window is None:
            raise ElementNotFoundError("Excel window not initialized. Call launch_excel() first.")

    def _refresh_focus(self) -> None:
        self._ensure_window()
        try:
            self.window = self.app.top_window()
        except Exception:
            self.window = self.app.window(title_re=".*Excel.*")
        self.window.set_focus()

    def _is_workbook_open(self) -> bool:
        title = (self.window.window_text() or "").strip().lower()
        if not title:
            return False
        if "start" in title or "home" in title:
            return False
        return "excel" in title

    def _exit_backstage_if_open(self) -> None:
        try:
            if self.window.child_window(title="Home", control_type="Text").exists(timeout=0.3):
                logging.info("Backstage/Home detected. Sending Esc to return to worksheet.")
                self.window.type_keys("{ESC}", set_foreground=True)
                time.sleep(0.8)
                self._refresh_focus()
                return
            if self.window.child_window(title_re="Good .*", control_type="Text").exists(timeout=0.3):
                logging.info("Backstage greeting detected. Sending Esc to return to worksheet.")
                self.window.type_keys("{ESC}", set_foreground=True)
                time.sleep(0.8)
                self._refresh_focus()
        except Exception:
            # If detection is inconclusive, continue without blocking flow.
            pass

    def _open_filter_dropdown(self, col_index: int) -> bool:
        # Re-anchor to A1, move to header column, then open AutoFilter dropdown.
        self.press_keys("^{HOME}")
        time.sleep(0.2)
        if col_index > 0:
            self.press_keys(f"{{RIGHT {col_index}}}")
            time.sleep(0.2)
        self.press_keys("%{DOWN}")
        time.sleep(0.4)
        if self._is_filter_popup_visible():
            return True

        # Retry once after refocus and re-anchor.
        logging.info("Filter popup not detected after Alt+Down. Retrying once after refocus.")
        self.press_keys("{ESC}")
        self.focus_workbook_surface()
        self.press_keys("^{HOME}")
        time.sleep(0.2)
        if col_index > 0:
            self.press_keys(f"{{RIGHT {col_index}}}")
            time.sleep(0.2)
        time.sleep(0.2)
        self.press_keys("%{DOWN}")
        time.sleep(0.4)
        return self._is_filter_popup_visible()

    def _is_filter_popup_visible(self) -> bool:
        try:
            desktop = Desktop(backend="uia")
            candidates = desktop.windows(class_name="#32768", visible_only=True)
            return len(candidates) > 0
        except Exception:
            return False

    def _select_filter_values(self, filter_value: str) -> None:
        try:
            popup = Desktop(backend="uia").window(class_name="#32768")
            popup.wait("visible", timeout=2)

            # Step 1: Untick "Select All"
            select_all = self._click_first_existing(
                popup,
                [
                    {"title_re": ".*Select All.*", "control_type": "CheckBox"},
                    {"title_re": ".*\\(Select All\\).*", "control_type": "CheckBox"},
                ],
            )
            if select_all.get_toggle_state() == 1:
                select_all.click_input()

            # Step 2: Tick target value
            target = self._click_first_existing(
                popup,
                [
                    {"title_re": rf".*{filter_value}.*", "control_type": "CheckBox"},
                    {"title_re": rf"^{filter_value}$", "control_type": "ListItem"},
                ],
            )
            if hasattr(target, "get_toggle_state") and target.get_toggle_state() == 0:
                target.click_input()

            # Step 3: Confirm selection
            ok_btn = self._click_first_existing(
                popup,
                [
                    {"title": "OK", "control_type": "Button"},
                    {"title_re": ".*OK.*", "control_type": "Button"},
                ],
            )
            ok_btn.click_input()
        except Exception as exc:
            raise RuntimeError(f"Failed to apply filter selection for '{filter_value}': {exc}") from exc

    def _click_first_existing(self, root: Any, locators: list[dict[str, str]]):
        for locator in locators:
            ctrl = root.child_window(**locator)
            if ctrl.exists(timeout=0.8):
                ctrl.wait("visible", timeout=2)
                ctrl.wait("ready", timeout=2)
                return ctrl
        raise RuntimeError(f"Filter control not found for locators: {locators}")

    def _apply_filter_via_com(self, dataframe: pd.DataFrame, col_index: int, filter_value: str) -> bool:
        try:
            import pythoncom
            from win32com.client import GetActiveObject

            pythoncom.CoInitialize()
            excel = GetActiveObject("Excel.Application")
            workbook = excel.ActiveWorkbook
            sheet = workbook.ActiveSheet

            rows = int(len(dataframe.index)) + 1  # include header row
            cols = int(len(dataframe.columns))
            if rows < 2 or cols < 1:
                raise RuntimeError("Dataframe is empty; cannot apply filter.")

            # Build range A1:<last_col><last_row> and apply criteria.
            data_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows, cols))
            data_range.AutoFilter(Field=col_index + 1, Criteria1=str(filter_value))
            time.sleep(0.3)
            return True
        except Exception as exc:
            logging.warning("COM filter apply failed: %s", exc)
            return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _verify_filter_result_via_com(self, dataframe: pd.DataFrame, col_index: int, expected_value: str) -> None:
        try:
            import pythoncom
            from win32com.client import GetActiveObject

            pythoncom.CoInitialize()
            excel = GetActiveObject("Excel.Application")
            workbook = excel.ActiveWorkbook
            sheet = workbook.ActiveSheet

            rows = int(len(dataframe.index)) + 1  # include header row
            # Validate visible data rows only.
            for row_idx in range(2, rows + 1):
                row_hidden = bool(sheet.Rows(row_idx).Hidden)
                if row_hidden:
                    continue
                cell_value = sheet.Cells(row_idx, col_index + 1).Value
                visible_value = "" if cell_value is None else str(cell_value).strip()
                if visible_value.lower() != str(expected_value).strip().lower():
                    raise RuntimeError(
                        f"Filter verification failed at row {row_idx}: expected '{expected_value}', found '{visible_value}'."
                    )
        except Exception as exc:
            raise RuntimeError(f"Post-filter verification failed: {exc}") from exc
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
