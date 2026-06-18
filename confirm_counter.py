"""Show and manually adjust the last committed outgoing number."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.config import Settings
from src.state_manager import (
    get_next_out_number,
    load_generation_state,
    save_generation_state,
)


def main() -> int:
    settings = Settings(ROOT)
    state_path = ROOT / "state" / "generation_state.json"
    state = load_generation_state(state_path, settings.env.start_out_number)

    next_number = get_next_out_number(state)
    print(f"Последний зафиксированный исходящий номер: {state.last_committed_out_number}")
    print(f"Следующий финальный исходящий номер будет: {next_number}")
    print()
    user_input = input(
        "Введите новый последний зафиксированный номер или нажмите Enter, "
        "чтобы оставить как есть: "
    ).strip()

    if not user_input:
        print("Номер не изменён.")
        return 0

    try:
        new_value = int(user_input)
    except ValueError:
        print("Ошибка: введите целое число.")
        return 1

    state.last_committed_out_number = new_value
    save_generation_state(state_path, state)
    print(f"Сохранено. Последний зафиксированный номер: {new_value}")
    print(f"Следующий финальный исходящий номер будет: {get_next_out_number(state)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
