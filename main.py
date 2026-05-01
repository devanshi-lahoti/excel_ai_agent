from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd

from core.automator import ExcelAutomator
from core.brain import get_ai_decision
from core.vision import encode_image_to_data_url

BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs"
DATA_DIR = BASE_DIR / "data"


def setup_logging() -> None:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOGS_DIR / "agent.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def build_user_goal() -> str:
    return (
        "Open Excel, create a blank workbook, paste sample_data.csv into Sheet1, "
        "and apply filters on the pasted range."
    )


def execute_step(automator: ExcelAutomator, step: Dict[str, Any], dataframe: pd.DataFrame) -> None:
    action = step.get("action")
    args = step.get("args", {}) or {}

    if action == "launch_excel":
        automator.launch_excel()
    elif action == "create_blank_workbook":
        automator.create_blank_workbook()
    elif action == "paste_data":
        automator.paste_data(dataframe)
    elif action == "apply_filters":
        automator.apply_filters()
    elif action == "press_keys":
        automator.press_keys(args.get("keys", "{ESC}"))
    elif action == "click_text":
        text = args.get("text", "Close")
        automator.click_text(text)
    elif action == "wait":
        automator.wait(float(args.get("seconds", 1)))
    else:
        raise ValueError(f"Unsupported action: {action}")


def recover_and_retry(
    automator: ExcelAutomator,
    failed_step: Dict[str, Any],
    dataframe: pd.DataFrame,
    err: Exception,
) -> bool:
    logging.warning("Step failed: %s | Error: %s", failed_step, err)

    screenshot_path = automator.capture_state("error_screenshot.png")
    image_data_url = encode_image_to_data_url(screenshot_path)

    recovery_prompt = (
        "The previous RPA step failed in Excel desktop. "
        f"Failed step: {json.dumps(failed_step)}. "
        f"Error: {str(err)}. "
        "Analyze the screenshot and provide one fix action."
    )

    fixes = get_ai_decision(recovery_prompt, image=image_data_url)
    if not fixes:
        return False

    fix_step = fixes[0]
    logging.info("Applying recovery step: %s", fix_step)
    execute_step(automator, fix_step, dataframe)

    logging.info("Retrying original failed step after recovery...")
    execute_step(automator, failed_step, dataframe)
    return True


def run_agent() -> None:
    setup_logging()

    csv_path = DATA_DIR / "sample_data.csv"
    if not csv_path.exists():
        raise FileNotFoundError(f"Missing input file: {csv_path}")

    dataframe = pd.read_csv(csv_path)
    automator = ExcelAutomator(logs_dir=LOGS_DIR)

    user_goal = build_user_goal()
    planning_prompt = (
        "Create Excel automation plan for this user request. "
        "Return JSON array of executable steps only.\n"
        f"User request: {user_goal}"
    )

    steps = get_ai_decision(planning_prompt)
    logging.info("Planned %d step(s)", len(steps))

    for idx, step in enumerate(steps, start=1):
        logging.info("Executing step %d/%d: %s", idx, len(steps), step)
        try:
            execute_step(automator, step, dataframe)
        except Exception as e:
            try:
                recovered = recover_and_retry(automator, step, dataframe, e)
                if not recovered:
                    raise RuntimeError("Recovery returned no action")
            except Exception as recovery_error:
                logging.exception("Recovery failed: %s", recovery_error)
                raise

    logging.info("Workflow completed successfully.")


if __name__ == "__main__":
    run_agent()
