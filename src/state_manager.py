"""Persistent generation state for outgoing numbers and run IDs."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path


@dataclass
class GenerationState:
    last_committed_out_number: int
    last_generation_id: int


def load_generation_state(path: Path, default_start_out_number: int) -> GenerationState:
    """Load state from disk or create initial state from START_OUT_NUMBER."""
    if not path.exists():
        state = GenerationState(
            last_committed_out_number=default_start_out_number - 1,
            last_generation_id=0,
        )
        save_generation_state(path, state)
        return state

    raw = json.loads(path.read_text(encoding="utf-8"))

    if "last_committed_out_number" not in raw and "last_used_out_number" in raw:
        raw["last_committed_out_number"] = raw["last_used_out_number"]

    return GenerationState(
        last_committed_out_number=int(raw.get("last_committed_out_number", default_start_out_number - 1)),
        last_generation_id=int(raw.get("last_generation_id", 0)),
    )


def save_generation_state(path: Path, state: GenerationState) -> None:
    """Save state atomically via a temporary file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "last_committed_out_number": state.last_committed_out_number,
        "last_generation_id": state.last_generation_id,
    }
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    tmp_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    os.replace(tmp_path, path)


def get_next_out_number(state: GenerationState) -> int:
    return state.last_committed_out_number + 1


def mark_out_number_committed(path: Path, state: GenerationState, out_number: int) -> None:
    state.last_committed_out_number = out_number
    save_generation_state(path, state)


def increment_generation_id(path: Path, state: GenerationState) -> int:
    """Increment and persist generation run ID. Returns the new ID."""
    state.last_generation_id += 1
    save_generation_state(path, state)
    return state.last_generation_id
