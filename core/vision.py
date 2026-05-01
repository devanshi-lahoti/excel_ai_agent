from __future__ import annotations

import base64
from pathlib import Path


def encode_image_to_data_url(image_path: Path) -> str:
    b = image_path.read_bytes()
    encoded = base64.b64encode(b).decode("utf-8")
    return f"data:image/png;base64,{encoded}"
