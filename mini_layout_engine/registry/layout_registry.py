from __future__ import annotations

from pathlib import Path

from mini_layout_engine.models.layout_contract import LayoutArchetypeContract

from .spec_loader import iter_spec_paths, load_spec_file


class LayoutRegistry:
    def __init__(self, contracts: dict[str, LayoutArchetypeContract] | None = None):
        self._contracts: dict[str, LayoutArchetypeContract] = dict(contracts or {})

    @classmethod
    def from_specs_dir(cls, specs_dir: Path) -> "LayoutRegistry":
        contracts: dict[str, LayoutArchetypeContract] = {}
        for path in iter_spec_paths(specs_dir):
            spec = load_spec_file(path)
            contract = LayoutArchetypeContract.from_spec(spec)
            if contract.layout_id in contracts:
                raise ValueError(
                    f"Duplicate layout_id '{contract.layout_id}' in specs: {path}"
                )
            contracts[contract.layout_id] = contract
        return cls(contracts)

    def get(self, layout_id: str) -> LayoutArchetypeContract | None:
        return self._contracts.get(str(layout_id or "").strip())

    def require(self, layout_id: str) -> LayoutArchetypeContract:
        contract = self.get(layout_id)
        if contract is None:
            raise KeyError(f"Unknown layout_id '{layout_id}'.")
        return contract

    def all_layout_ids(self) -> list[str]:
        return sorted(self._contracts.keys())
