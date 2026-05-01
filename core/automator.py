from __future__ import annotations

import logging
import os
import time
from pathlib import Path
from typing import Optional

import pandas as pd
import pyperclip
from dotenv import load_dotenv
from PIL import ImageGrab
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError

load_dotenv()


class ExcelAutomator:
    def __init__(self, logs_dir: Path) -> None:
        self.logs_dir = logs_dir
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.app: Optional[Application] = None
        self.window = None

    def launch_excel(self, timeout: int = 25) -> None:
        excel_path = os.getenv("EXCEL_PATH")
        if not excel_path:
            raise ValueError("EXCEL_PATH is missing in .env. Set EXCEL_PATH to full Excel executable path.")
        excel_path = excel_path.strip().strip('"').strip("'")
        logging.info("Launching Excel from path: %s", excel_path)
        self.app = Application(backend="uia").start(excel_path)
        self.window = self.app.window(best_match="Excel")
        self.window.wait("visible", timeout=timeout)
        self.window.set_focus()
        time.sleep(1.0)

    def create_blank_workbook(self) -> None:
        self._ensure_window()
        logging.info("Creating a blank workbook (Alt+F, N, L)...")
        self.window.type_keys("%F", set_foreground=True)
        time.sleep(0.5)
        self.window.type_keys("N")
        time.sleep(0.5)
        self.window.type_keys("L")
        time.sleep(1.2)

    def paste_data(self, dataframe: pd.DataFrame) -> None:
        self._ensure_window()
        logging.info("Pasting dataframe into workbook...")
        tab_data = dataframe.to_csv(index=False, sep="\t")
        pyperclip.copy(tab_data)
        self.window.set_focus()
        self.window.type_keys("^v", set_foreground=True)
        time.sleep(1.0)

    def apply_filters(self) -> None:
        self._ensure_window()
        logging.info("Applying filters (Ctrl+Shift+L)...")
        self.window.set_focus()
        self.window.type_keys("^a", set_foreground=True)
        time.sleep(0.3)
        self.window.type_keys("^+l")
        time.sleep(1.0)

    def press_keys(self, keys: str) -> None:
        self._ensure_window()
        logging.info("Pressing keys: %s", keys)
        self.window.set_focus()
        self.window.type_keys(keys, set_foreground=True)
        time.sleep(0.6)

    def click_text(self, text: str, timeout: int = 3) -> None:
        self._ensure_window()
        logging.info("Trying to click control with text: %s", text)
        btn = self.window.child_window(title=text, control_type="Button")
        btn.wait("exists enabled visible", timeout=timeout)
        btn.click_input()
        time.sleep(0.8)

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
