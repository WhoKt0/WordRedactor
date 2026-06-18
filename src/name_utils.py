"""Greeting name normalization for Word template placeholders."""

from __future__ import annotations

import re

SERVICE_WORDS = {"перепроверь", "проверь", "уточнить", "уточни"}

_TRAILING_PUNCT = re.compile(r"[,.;]+$")


class GreetingNameError(ValueError):
    """Raised when a safe two-word greeting name cannot be built."""


def _clean_name(text: str) -> str:
    cleaned = " ".join(text.split())
    cleaned = _TRAILING_PUNCT.sub("", cleaned).strip()
    return cleaned


def _word_count(text: str) -> int:
    return len(text.split()) if text else 0


def _filter_service_words(text: str) -> str:
    words = text.split()
    filtered = [word for word in words if word.lower() not in SERVICE_WORDS]
    return " ".join(filtered)


def build_greeting_name(greeting_name: str, chair_full_name: str) -> str:
    """
    Return a safe GREETING_NAME without a trailing comma.

    Rules:
    - result must contain at least 2 words;
    - if greeting_name already has 2+ words, use it after cleanup;
    - if greeting_name has 1 word, try to restore the second word from chair_full_name;
    - raise GreetingNameError if restoration is impossible.
    """
    cleaned = _clean_name(greeting_name)
    if _word_count(cleaned) >= 2:
        return cleaned

    if _word_count(cleaned) < 1:
        raise GreetingNameError(
            "невозможно сформировать корректное обращение после «Уважаемый/Уважаемая» "
            "минимум из двух слов: greeting_name пустой."
        )

    greeting_word = cleaned
    chair_clean = _filter_service_words(_clean_name(chair_full_name))
    chair_words = chair_clean.split()

    if not chair_words:
        raise GreetingNameError(
            "невозможно сформировать корректное обращение после «Уважаемый/Уважаемая» "
            "минимум из двух слов."
        )

    if len(chair_words) >= 3:
        result = f"{chair_words[1]} {chair_words[2]}"
    elif len(chair_words) == 2:
        if greeting_word.lower() == chair_words[1].lower():
            result = f"{chair_words[1]} {chair_words[0]}"
        elif greeting_word.lower() == chair_words[0].lower():
            result = f"{chair_words[0]} {chair_words[1]}"
        else:
            result = f"{chair_words[1]} {chair_words[0]}"
    else:
        result = greeting_word

    result = _clean_name(result)
    if _word_count(result) < 2:
        raise GreetingNameError(
            "невозможно сформировать корректное обращение после «Уважаемый/Уважаемая» "
            "минимум из двух слов."
        )
    return result
