from dataclasses import dataclass
from typing import Any


@dataclass
class _PersistenceFacade:
    database_enabled: bool = False

    def resolve_topic_node(self, *args: Any, **kwargs: Any) -> dict[str, Any] | None:
        return None

    def get_topic_script(self, *args: Any, **kwargs: Any) -> dict[str, Any] | None:
        return None

    def register_local_file(self, *args: Any, **kwargs: Any) -> dict[str, Any] | None:
        return None

    def upsert_slide_deck(self, *args: Any, **kwargs: Any) -> dict[str, Any] | None:
        return None

    def get_slide_deck(self, *args: Any, **kwargs: Any) -> dict[str, Any] | None:
        return None

    def list_slide_decks(self, *args: Any, **kwargs: Any) -> list[dict[str, Any]]:
        return []


def get_persistence_facade(settings: Any = None) -> _PersistenceFacade:
    return _PersistenceFacade(database_enabled=False)

