from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable


@dataclass
class Diagnostics:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return bool(self.errors)

    def add_error(self, message: str) -> None:
        text = str(message or "").strip()
        if text:
            self.errors.append(text)

    def add_warning(self, message: str) -> None:
        text = str(message or "").strip()
        if text:
            self.warnings.append(text)

    def extend(
        self,
        *,
        errors: Iterable[str] = (),
        warnings: Iterable[str] = (),
    ) -> None:
        for message in errors:
            self.add_error(str(message))
        for message in warnings:
            self.add_warning(str(message))

    def to_dict(self) -> dict[str, list[str]]:
        return {
            "errors": list(self.errors),
            "warnings": list(self.warnings),
        }
