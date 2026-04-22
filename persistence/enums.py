from enum import Enum


class StoredFileRole(str, Enum):
    SLIDE_JSON = "SLIDE_JSON"
    SLIDE_PPTX = "SLIDE_PPTX"

