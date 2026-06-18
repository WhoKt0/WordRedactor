"""Generation run performance metrics."""

from __future__ import annotations

import time
from dataclasses import dataclass, field


@dataclass
class RunPerformance:
    """Wall-clock and phase timings for one generation run."""

    _started_at: float = field(default_factory=time.perf_counter)
    word_startup_seconds: float | None = None
    docx_total_seconds: float = 0.0
    docx_count: int = 0
    pdf_total_seconds: float = 0.0
    pdf_count: int = 0
    pdf_retries: int = 0

    def mark_started(self) -> None:
        self._started_at = time.perf_counter()

    def record_word_startup(self, seconds: float) -> None:
        self.word_startup_seconds = seconds

    def add_docx(self, seconds: float) -> None:
        self.docx_total_seconds += seconds
        self.docx_count += 1

    def add_pdf_success(self, seconds: float) -> None:
        self.pdf_total_seconds += seconds
        self.pdf_count += 1

    def add_pdf_retry(self) -> None:
        self.pdf_retries += 1

    def total_seconds(self) -> float:
        return time.perf_counter() - self._started_at

    def pdf_avg_seconds(self) -> float:
        if self.pdf_count == 0:
            return 0.0
        return self.pdf_total_seconds / self.pdf_count

    def print_summary(self) -> None:
        print()
        print("Performance:")
        if self.word_startup_seconds is not None:
            print(f"Word startup: {self.word_startup_seconds:.1f}s")
        print(f"DOCX total: {self.docx_total_seconds:.1f}s")
        print(f"PDF total: {self.pdf_total_seconds:.1f}s")
        if self.pdf_count:
            print(f"PDF avg: {self.pdf_avg_seconds():.2f}s")
        else:
            print("PDF avg: —")
        print(f"PDF retries: {self.pdf_retries}")
        print(f"Total: {self.total_seconds():.1f}s")
