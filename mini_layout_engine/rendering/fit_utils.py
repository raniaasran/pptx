"""Shared text fitting estimates for mini layout renderers.

These helpers are intentionally layout-agnostic. They estimate text wrapping,
height usage, and font sizes without mutating PowerPoint shapes. Layout renderers
can use the returned metrics later when they need adaptive behavior.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
import textwrap
from typing import Any


EMU_PER_INCH = 914400
POINTS_PER_INCH = 72.0
DEFAULT_AVERAGE_CHAR_WIDTH_FACTOR = 0.52
DEFAULT_LINE_HEIGHT_FACTOR = 1.2


@dataclass(frozen=True)
class TextBoxMetrics:
    """Estimated text usage for a rectangular text area."""

    width_in: float
    height_in: float
    font_size_pt: float
    line_spacing: float
    chars_per_line: int
    estimated_lines: int
    estimated_height_in: float
    fits: bool
    overflow_lines: int
    overflow_height_in: float


@dataclass(frozen=True)
class FontFitResult:
    """Best font candidate found for a text box."""

    font_size_pt: float
    metrics: TextBoxMetrics
    candidates_tested: int


def emu_to_inches(value_emu: int | float | None) -> float:
    """Convert an EMU value from python-pptx into inches."""

    if value_emu is None:
        return 0.0
    return float(value_emu) / EMU_PER_INCH


def inches_to_emu(value_in: int | float | None) -> int:
    """Convert inches into EMUs for python-pptx APIs."""

    if value_in is None:
        return 0
    return int(round(float(value_in) * EMU_PER_INCH))


def normalize_text_lines(value: Any) -> list[str]:
    """Return non-empty logical text lines while preserving explicit breaks."""

    if value is None:
        return []
    normalized = re.sub(r"\r\n?", "\n", str(value))
    return [line.strip() for line in normalized.split("\n") if line.strip()]


def compact_text(value: Any) -> str:
    """Collapse whitespace inside text for width/line estimates."""

    return re.sub(r"\s+", " ", str(value or "").strip())


def estimate_chars_per_line(
    width_in: float,
    font_size_pt: float,
    *,
    average_char_width_factor: float = DEFAULT_AVERAGE_CHAR_WIDTH_FACTOR,
    min_chars: int = 1,
) -> int:
    """Estimate how many average-width characters fit on one line.

    The estimate uses a simple typographic approximation:
    average character width ~= font size * factor.
    """

    safe_width = max(0.0, float(width_in))
    safe_font = max(1.0, float(font_size_pt))
    safe_factor = max(0.1, float(average_char_width_factor))
    char_width_in = (safe_font * safe_factor) / POINTS_PER_INCH
    return max(int(min_chars), int(math.floor(safe_width / char_width_in)))


def estimate_wrapped_lines(text: Any, chars_per_line: int) -> int:
    """Estimate wrapped line count for one logical paragraph."""

    compact = compact_text(text)
    if not compact:
        return 0

    width = max(1, int(chars_per_line))
    wrapped = textwrap.wrap(
        compact,
        width=width,
        break_long_words=True,
        break_on_hyphens=False,
    )
    return max(1, len(wrapped))


def estimate_multiline_wrapped_lines(value: Any, chars_per_line: int) -> int:
    """Estimate wrapped lines across explicit newline-separated text."""

    logical_lines = normalize_text_lines(value)
    if not logical_lines:
        return estimate_wrapped_lines(value, chars_per_line)
    return sum(max(1, estimate_wrapped_lines(line, chars_per_line)) for line in logical_lines)


def estimate_line_height_in(
    font_size_pt: float,
    *,
    line_spacing: float = 1.0,
    line_height_factor: float = DEFAULT_LINE_HEIGHT_FACTOR,
) -> float:
    """Estimate rendered line height in inches."""

    safe_font = max(1.0, float(font_size_pt))
    safe_spacing = max(0.1, float(line_spacing))
    safe_factor = max(0.1, float(line_height_factor))
    return (safe_font / POINTS_PER_INCH) * safe_spacing * safe_factor


def estimate_text_height_in(
    text: Any,
    *,
    width_in: float,
    font_size_pt: float,
    line_spacing: float = 1.0,
    paragraph_spacing_in: float = 0.0,
    average_char_width_factor: float = DEFAULT_AVERAGE_CHAR_WIDTH_FACTOR,
    line_height_factor: float = DEFAULT_LINE_HEIGHT_FACTOR,
) -> float:
    """Estimate the vertical space needed by text in a shape."""

    chars_per_line = estimate_chars_per_line(
        width_in,
        font_size_pt,
        average_char_width_factor=average_char_width_factor,
    )
    logical_lines = normalize_text_lines(text)
    estimated_lines = estimate_multiline_wrapped_lines(text, chars_per_line)
    line_height = estimate_line_height_in(
        font_size_pt,
        line_spacing=line_spacing,
        line_height_factor=line_height_factor,
    )
    paragraph_gaps = max(0, len(logical_lines) - 1) * max(0.0, float(paragraph_spacing_in))
    return (estimated_lines * line_height) + paragraph_gaps


def measure_text_box(
    text: Any,
    *,
    width_in: float,
    height_in: float,
    font_size_pt: float,
    line_spacing: float = 1.0,
    paragraph_spacing_in: float = 0.0,
    vertical_padding_in: float = 0.0,
    average_char_width_factor: float = DEFAULT_AVERAGE_CHAR_WIDTH_FACTOR,
    line_height_factor: float = DEFAULT_LINE_HEIGHT_FACTOR,
) -> TextBoxMetrics:
    """Estimate whether text fits inside a text box."""

    chars_per_line = estimate_chars_per_line(
        width_in,
        font_size_pt,
        average_char_width_factor=average_char_width_factor,
    )
    estimated_lines = estimate_multiline_wrapped_lines(text, chars_per_line)
    line_height = estimate_line_height_in(
        font_size_pt,
        line_spacing=line_spacing,
        line_height_factor=line_height_factor,
    )
    logical_lines = normalize_text_lines(text)
    paragraph_gaps = max(0, len(logical_lines) - 1) * max(0.0, float(paragraph_spacing_in))
    estimated_height = (
        estimated_lines * line_height
        + paragraph_gaps
        + max(0.0, float(vertical_padding_in))
    )
    available_height = max(0.0, float(height_in))
    fits = estimated_height <= available_height + 1e-6
    estimated_capacity_lines = int(math.floor(available_height / line_height)) if line_height else 0

    return TextBoxMetrics(
        width_in=float(width_in),
        height_in=float(height_in),
        font_size_pt=float(font_size_pt),
        line_spacing=float(line_spacing),
        chars_per_line=chars_per_line,
        estimated_lines=estimated_lines,
        estimated_height_in=estimated_height,
        fits=fits,
        overflow_lines=max(0, estimated_lines - estimated_capacity_lines),
        overflow_height_in=max(0.0, estimated_height - available_height),
    )


def text_fits_box(text: Any, **kwargs: Any) -> bool:
    """Convenience wrapper around measure_text_box(...).fits."""

    return measure_text_box(text, **kwargs).fits


def find_largest_fitting_font(
    text: Any,
    *,
    width_in: float,
    height_in: float,
    max_font_pt: float,
    min_font_pt: float,
    line_spacing: float = 1.0,
    step_pt: float = 1.0,
    paragraph_spacing_in: float = 0.0,
    vertical_padding_in: float = 0.0,
    average_char_width_factor: float = DEFAULT_AVERAGE_CHAR_WIDTH_FACTOR,
    line_height_factor: float = DEFAULT_LINE_HEIGHT_FACTOR,
) -> FontFitResult | None:
    """Find the largest discrete font size that is estimated to fit.

    This is the generic best-fit search: it may return any size between
    min_font_pt and max_font_pt, including a size larger than a template's
    original font if max_font_pt allows it.

    The search uses binary search over step-sized candidates. It assumes smaller
    font sizes are no worse than larger sizes for fitting, which is true for this
    estimator.
    """

    safe_step = max(0.1, float(step_pt))
    low = float(min_font_pt)
    high = float(max_font_pt)
    if high < low:
        low, high = high, low

    candidate_count = int(math.floor((high - low) / safe_step)) + 1
    best: FontFitResult | None = None
    tested = 0
    left = 0
    right = max(0, candidate_count - 1)

    while left <= right:
        mid = (left + right) // 2
        font_size = low + (mid * safe_step)
        metrics = measure_text_box(
            text,
            width_in=width_in,
            height_in=height_in,
            font_size_pt=font_size,
            line_spacing=line_spacing,
            paragraph_spacing_in=paragraph_spacing_in,
            vertical_padding_in=vertical_padding_in,
            average_char_width_factor=average_char_width_factor,
            line_height_factor=line_height_factor,
        )
        tested += 1

        if metrics.fits:
            best = FontFitResult(
                font_size_pt=font_size,
                metrics=metrics,
                candidates_tested=tested,
            )
            left = mid + 1
        else:
            right = mid - 1

    if best is None:
        return None
    return FontFitResult(
        font_size_pt=best.font_size_pt,
        metrics=best.metrics,
        candidates_tested=tested,
    )


def find_shrink_to_fit_font(
    text: Any,
    *,
    width_in: float,
    height_in: float,
    template_font_pt: float,
    min_font_pt: float,
    line_spacing: float = 1.0,
    step_pt: float = 1.0,
    paragraph_spacing_in: float = 0.0,
    vertical_padding_in: float = 0.0,
    average_char_width_factor: float = DEFAULT_AVERAGE_CHAR_WIDTH_FACTOR,
    line_height_factor: float = DEFAULT_LINE_HEIGHT_FACTOR,
) -> FontFitResult | None:
    """Shrink from the template font size only when text does not fit.

    This is the preferred strategy for template-driven layouts where the
    template's typography is the visual source of truth. The template font size
    is tested first and returned unchanged when it fits. If it does not fit, the
    search only moves downward toward min_font_pt; it never grows above the
    template size.
    """

    template_size = float(template_font_pt)
    floor_size = min(float(min_font_pt), template_size)
    template_metrics = measure_text_box(
        text,
        width_in=width_in,
        height_in=height_in,
        font_size_pt=template_size,
        line_spacing=line_spacing,
        paragraph_spacing_in=paragraph_spacing_in,
        vertical_padding_in=vertical_padding_in,
        average_char_width_factor=average_char_width_factor,
        line_height_factor=line_height_factor,
    )

    if template_metrics.fits:
        return FontFitResult(
            font_size_pt=template_size,
            metrics=template_metrics,
            candidates_tested=1,
        )

    fit = find_largest_fitting_font(
        text,
        width_in=width_in,
        height_in=height_in,
        max_font_pt=template_size,
        min_font_pt=floor_size,
        line_spacing=line_spacing,
        step_pt=step_pt,
        paragraph_spacing_in=paragraph_spacing_in,
        vertical_padding_in=vertical_padding_in,
        average_char_width_factor=average_char_width_factor,
        line_height_factor=line_height_factor,
    )

    if fit is None:
        return None
    return FontFitResult(
        font_size_pt=fit.font_size_pt,
        metrics=fit.metrics,
        candidates_tested=fit.candidates_tested + 1,
    )


def try_builtin_fit_text(
    text_frame: Any,
    *,
    font_family: str = "Calibri",
    max_size: int = 18,
    bold: bool = False,
    italic: bool = False,
    font_file: str | None = None,
) -> bool:
    """Best-effort wrapper for python-pptx TextFrame.fit_text().

    Unlike the other helpers in this module, this mutates the provided text
    frame. Keep it as an explicit fallback for experiments, not as the default
    fitting strategy.
    """

    try:
        text_frame.fit_text(
            font_family=font_family,
            max_size=max_size,
            bold=bold,
            italic=italic,
            font_file=font_file,
        )
    except Exception:
        return False
    return True
