from __future__ import annotations

from pathlib import Path

from mini_layout_engine.models.design_family import DesignFamilyContext

from .spec_loader import iter_spec_paths, load_spec_file


class FamilyRegistry:
    def __init__(self, families: dict[str, DesignFamilyContext] | None = None):
        self._families: dict[str, DesignFamilyContext] = dict(families or {})

    @classmethod
    def from_specs_dir(cls, specs_dir: Path) -> "FamilyRegistry":
        families: dict[str, DesignFamilyContext] = {}
        for path in iter_spec_paths(specs_dir):
            spec = load_spec_file(path)
            context = DesignFamilyContext.from_spec(spec)
            if context.family_id in families:
                raise ValueError(
                    f"Duplicate family_id '{context.family_id}' in specs: {path}"
                )
            families[context.family_id] = context
        return cls(families)

    def get(self, family_id: str) -> DesignFamilyContext | None:
        return self._families.get(str(family_id or "").strip())

    def require(self, family_id: str) -> DesignFamilyContext:
        context = self.get(family_id)
        if context is None:
            raise KeyError(f"Unknown family_id '{family_id}'.")
        return context

    def all_family_ids(self) -> list[str]:
        return sorted(self._families.keys())
