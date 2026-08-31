# Excel AI Agent

An LLM-assisted RPA agent that drives the **desktop** Microsoft Excel application on Windows. An AI model (any OpenAI-compatible provider) plans a sequence of high-level actions (launch Excel, create a workbook, paste data, apply filters), a deterministic automation layer executes each step with `pywinauto`, and — if a step fails — the agent takes a screenshot, asks the model to diagnose the failure, and applies a safe fix before retrying.

## How it works

1. **Plan** — `main.py` sends the user's goal to the model (`core/brain.py`, using the [`prompts/planner.txt`](prompts/planner.txt) system prompt). The model must respond with a JSON array of high-level steps such as `launch_excel`, `create_blank_workbook`, `paste_data`, `apply_filters`.
2. **Sanitize** — Any planner step outside that allow-list (`PLANNER_ALLOWED_ACTIONS` in [main.py](main.py)) is dropped. The planner is intentionally not allowed to emit low-level keyboard/mouse steps.
3. **Execute** — Each step is dispatched to [`core/automator.py`](core/automator.py)'s `ExcelAutomator`, which drives the real Excel window via `pywinauto` (UIA backend): launching Excel, creating a blank workbook, pasting a DataFrame through the clipboard, and applying an AutoFilter to a given column/value.
4. **Verify** — After a filter is applied, the automator cross-checks the visible rows through the Excel COM object model (`win32com`) to confirm the filter actually took effect.
5. **Recover** — If a step throws, `main.py` captures a screenshot of the Excel window (`core/vision.py` base64-encodes it), sends it to the model with [`prompts/recovery.txt`](prompts/recovery.txt), and applies the single suggested fix (`press_keys`, `click_text`, `retry`, or `abort`) before retrying the original step. If the LLM suggests something authentication-related (sign in, login, activate, account), that suggestion is blocked and replaced with a safe deterministic recovery chain (Esc → refocus workbook → Ctrl+Home). A second failure after LLM-guided recovery falls back to that same safe chain plus one more retry.
6. **Log** — Every run writes to `logs/agent.log`, and failure screenshots are saved to `logs/error_screenshot.png`.

## Project structure

```
excel_ai_agent/
├── main.py                # Orchestrator: build goal → plan → execute → recover → retry
├── core/
│   ├── automator.py       # ExcelAutomator: pywinauto-driven Excel control (+ COM verification)
│   ├── brain.py           # OpenAI-compatible client: planning & recovery decisions
│   └── vision.py          # Screenshot → base64 data URL helper
├── prompts/
│   ├── planner.txt        # System prompt constraining the planner to high-level actions
│   └── recovery.txt       # System prompt constraining the recovery assistant's fix suggestions
├── data/                  # Not committed — put sample_data.csv here (see Setup)
├── logs/                  # Created at runtime — agent.log + error screenshots
└── requirements.txt
```

## Prerequisites

- Windows, with **Microsoft Excel (desktop)** installed
- Python 3.10+
- An API key for an OpenAI-compatible provider, with a deployed multimodal-capable model (used for both planning and screenshot-based recovery)

## Setup

1. **Install dependencies**

   ```powershell
   pip install -r requirements.txt
   ```

2. **Create a `.env` file** in the project root:

   ```env
   EXCEL_PATH=your-excel-path
   API_KEY=your-api-key
   MODEL_NAME=your-model-name
   BASE_URL=
   ```

   - `EXCEL_PATH` must point to the actual `EXCEL.EXE` on your machine.
   - `API_KEY` and `MODEL_NAME` are required; both are validated at call time (`core/brain.py`) — the agent raises a clear error if either is missing.
   - `BASE_URL` is optional — leave it blank to hit the default OpenAI API, or set it to point at a different OpenAI-compatible endpoint (self-hosted, gateway, alternate provider, etc.).

3. **Add input data**: create `data/sample_data.csv` with a `subject` column (the default goal filters this column to `Mathematics` — see `build_user_goal()` in [main.py](main.py)). Both `data/` and `.env` are gitignored.

## Usage

```powershell
python main.py
```

This runs the default goal: open Excel, create a blank workbook, paste `data/sample_data.csv` into Sheet1, apply an AutoFilter, and filter the `subject` column to `Mathematics`. Progress and errors stream to the console and to `logs/agent.log`.

To automate a different task, edit `build_user_goal()` in [main.py](main.py), or extend `PLANNER_ALLOWED_ACTIONS` / `execute_step()` and the corresponding methods on `ExcelAutomator` to support new deterministic actions.

## Safety notes

- The planner is restricted to a fixed action allow-list; anything else is logged and ignored.
- Recovery suggestions containing authentication-related intent (sign in, login, activate, account, etc.) are blocked outright and replaced with a deterministic, non-destructive recovery sequence — the agent will never attempt to click through a sign-in/activation prompt on the model's suggestion.

## Limitations

- Windows + Excel desktop only (relies on `pywinauto`'s UIA backend and Excel's COM object model).
- Single-sheet, single-filter workflow out of the box (`apply_filters` targets one header/value pair); extend `ExcelAutomator` for more complex sheets or multi-column filters.
- The planning/recovery model calls depend on an available OpenAI-compatible API; no offline/local-model fallback is included.
