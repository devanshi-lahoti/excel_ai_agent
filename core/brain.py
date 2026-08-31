from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
from openai import NotFoundError
from openai import OpenAI

load_dotenv()

API_KEY = os.getenv("API_KEY", "API_KEY")
MODEL_NAME = os.getenv("MODEL_NAME", "MODEL_NAME")
BASE_URL = os.getenv("BASE_URL")  # optional: only needed for a custom/self-hosted OpenAI-compatible endpoint

_BASE_DIR = Path(__file__).resolve().parents[1]
_PROMPTS_DIR = _BASE_DIR / "prompts"


def _read_prompt(name: str) -> str:
    return (_PROMPTS_DIR / name).read_text(encoding="utf-8")


def _client() -> OpenAI:
    if BASE_URL:
        return OpenAI(api_key=API_KEY, base_url=BASE_URL)
    return OpenAI(api_key=API_KEY)


def _ensure_config() -> None:
    missing = []
    for key, value in {
        "API_KEY": API_KEY,
        "MODEL_NAME": MODEL_NAME,
    }.items():
        if not value or value == key:
            missing.append(key)

    if missing:
        raise ValueError(f"Missing config in .env: {', '.join(missing)}")


def get_ai_decision(prompt: str, image: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Returns a JSON array of steps (planning mode) or a list with one recovery step.
    image: optional data URL for screenshot-based recovery.
    """
    _ensure_config()
    system_prompt = _read_prompt("recovery.txt" if image else "planner.txt")

    logging.info("Requesting AI decision (%s mode)", "recovery" if image else "planner")
    client = _client()

    user_content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
    if image:
        user_content.append({"type": "input_image", "image_url": image})

    text = ""
    try:
        # Preferred path for multimodal + structured outputs.
        resp = client.responses.create(
            model=MODEL_NAME,
            input=[
                {"role": "system", "content": [{"type": "text", "text": system_prompt}]},
                {"role": "user", "content": user_content},
            ],
            temperature=0,
        )
        text = getattr(resp, "output_text", "") or ""
    except NotFoundError as e:
        # Fallback for resources/api-versions where /responses is unavailable.
        logging.warning("Responses API not available (%s). Falling back to chat.completions.", e)
        if image:
            # If image is a data URL, Chat Completions can still consume it via image_url.
            user_message = [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": image}},
            ]
        else:
            user_message = prompt

        chat = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message},
            ],
            temperature=0,
        )
        text = (chat.choices[0].message.content or "").strip()

    if not text:
        raise ValueError("Empty response from AI provider")

    data = json.loads(text)

    if isinstance(data, dict):
        return [
            {
                "action": data.get("fix_action", "retry"),
                "args": data.get("args", {}),
                "reason": data.get("reason", "recovery suggestion"),
            }
        ]

    if not isinstance(data, list):
        raise ValueError("LLM output is not a JSON array/object")

    return data
