from __future__ import annotations

import math
from pathlib import Path


EMU_PER_INCH = 914400
_TARGET_PPI = 220


def _to_inches(value_emu: int) -> float:
    return float(value_emu) / float(EMU_PER_INCH)


def _target_pixels(*, target_width_emu: int, target_height_emu: int) -> tuple[int, int]:
    target_width_in = _to_inches(int(target_width_emu))
    target_height_in = _to_inches(int(target_height_emu))
    if target_width_in <= 0 or target_height_in <= 0:
        return (0, 0)
    target_width_px = max(80, int(round(target_width_in * _TARGET_PPI)))
    target_height_px = max(80, int(round(target_height_in * _TARGET_PPI)))
    return (target_width_px, target_height_px)


def _resampling_lanczos(image_module) -> int:
    resampling = getattr(image_module, "Resampling", None)
    if resampling is not None:
        return int(resampling.LANCZOS)
    return int(image_module.LANCZOS)


def prepare_image_for_cover_box(
    image_path: str,
    *,
    target_width_emu: int,
    target_height_emu: int,
    output_path: str | None = None,
) -> str:
    """
    Create a centered cover-cropped image matching target box aspect ratio.
    """
    source_path = Path(image_path)
    if not source_path.is_file():
        return image_path

    target_width_px, target_height_px = _target_pixels(
        target_width_emu=target_width_emu,
        target_height_emu=target_height_emu,
    )
    if target_width_px <= 0 or target_height_px <= 0:
        return image_path

    try:
        from PIL import Image
    except Exception as exc:
        raise RuntimeError("Pillow is required for cover image preprocessing.") from exc

    lanczos = _resampling_lanczos(Image)
    with Image.open(source_path) as image:
        source_width, source_height = image.size
        if source_width <= 0 or source_height <= 0:
            return image_path

        cover_scale = max(
            float(target_width_px) / float(source_width),
            float(target_height_px) / float(source_height),
        )
        resized_width = max(target_width_px, int(math.ceil(source_width * cover_scale)))
        resized_height = max(target_height_px, int(math.ceil(source_height * cover_scale)))
        resized = image.resize((resized_width, resized_height), lanczos)

        crop_left = max(0, int(round((resized_width - target_width_px) / 2.0)))
        crop_top = max(0, int(round((resized_height - target_height_px) / 2.0)))
        crop_box = (
            crop_left,
            crop_top,
            crop_left + target_width_px,
            crop_top + target_height_px,
        )
        prepared = resized.crop(crop_box)
        if prepared.mode not in ("RGB", "RGBA"):
            prepared = prepared.convert("RGB")

        if output_path:
            prepared_path = Path(output_path)
        else:
            prepared_name = (
                f"{source_path.stem}__cover_{target_width_px}x{target_height_px}.png"
            )
            prepared_path = source_path.with_name(prepared_name)
        prepared.save(prepared_path, format="PNG", optimize=True)

    return str(prepared_path)


def prepare_image_for_contain_box(
    image_path: str,
    *,
    target_width_emu: int,
    target_height_emu: int,
    output_path: str | None = None,
) -> str:
    """
    Create a centered fit-inside image with transparent padding.
    """
    source_path = Path(image_path)
    if not source_path.is_file():
        return image_path

    target_width_px, target_height_px = _target_pixels(
        target_width_emu=target_width_emu,
        target_height_emu=target_height_emu,
    )
    if target_width_px <= 0 or target_height_px <= 0:
        return image_path

    try:
        from PIL import Image
    except Exception as exc:
        raise RuntimeError("Pillow is required for contain image preprocessing.") from exc

    lanczos = _resampling_lanczos(Image)
    with Image.open(source_path) as image:
        source_width, source_height = image.size
        if source_width <= 0 or source_height <= 0:
            return image_path

        fit_scale = min(
            float(target_width_px) / float(source_width),
            float(target_height_px) / float(source_height),
        )
        resized_width = max(1, int(round(source_width * fit_scale)))
        resized_height = max(1, int(round(source_height * fit_scale)))

        resized = image.resize((resized_width, resized_height), lanczos)
        if resized.mode != "RGBA":
            resized = resized.convert("RGBA")

        canvas = Image.new("RGBA", (target_width_px, target_height_px), (0, 0, 0, 0))
        paste_left = max(0, int(round((target_width_px - resized_width) / 2.0)))
        paste_top = max(0, int(round((target_height_px - resized_height) / 2.0)))
        canvas.paste(resized, (paste_left, paste_top), resized)

        if output_path:
            prepared_path = Path(output_path)
        else:
            prepared_name = (
                f"{source_path.stem}__contain_{target_width_px}x{target_height_px}.png"
            )
            prepared_path = source_path.with_name(prepared_name)
        canvas.save(prepared_path, format="PNG", optimize=True)

    return str(prepared_path)


def prepare_image_for_fit_mode(
    image_path: str,
    *,
    target_width_emu: int,
    target_height_emu: int,
    fit_mode: str = "stretch",
    output_path: str | None = None,
) -> str:
    """
    Fit-mode dispatcher for later use.
    Current default behavior for unknown/stretch modes is passthrough.
    """
    mode = str(fit_mode or "stretch").strip().lower()
    if mode == "cover":
        return prepare_image_for_cover_box(
            image_path,
            target_width_emu=target_width_emu,
            target_height_emu=target_height_emu,
            output_path=output_path,
        )
    if mode == "contain":
        return prepare_image_for_contain_box(
            image_path,
            target_width_emu=target_width_emu,
            target_height_emu=target_height_emu,
            output_path=output_path,
        )
    return image_path
