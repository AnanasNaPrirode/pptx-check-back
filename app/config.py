"""Пороги и лимиты. Меняются правкой этого модуля."""

import os

MAX_FILE_BYTES = 5 * 1024 * 1024
# Запас на multipart-обёртку поверх файла в 5 МБ.
MAX_CONTENT_LENGTH_BYTES = MAX_FILE_BYTES + 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 64 * 1024 * 1024
MAX_ZIP_ENTRIES = 2_000

CHECK_TIMEOUT_SECONDS = 170
RATE_LIMIT_SECONDS = 30

ASPECT_TARGET = 16 / 9
ASPECT_TOLERANCE = 0.02

MIN_FONT_SIZE_PT = 18
BOTTOM_MARGIN_RATIO = 0.06
NUMBER_GROUPING_MAX_DIGITS = 4
ENGLISH_RATIO_MAX = 0.35
OVERLAP_MIN_RATIO = 0.15

CONTRAST_NORMAL = 4.5
CONTRAST_LARGE = 4.0
CONTRAST_LARGE_FONT_PT = 18

# Верхняя зона и «крупный» шрифт для title_heuristic.
TITLE_ZONE_RATIO = 0.20
TITLE_HEURISTIC_MIN_PT = 28

TIMEOUT_MESSAGE = "Презентация слишком большая. Разделите её на несколько файлов."

# Короткое слово в конце строки: предлоги и союзы не длиннее 3 букв.
HANGING_WORDS = frozenset(
    {
        "а",
        "и",
        "но",
        "да",
        "ни",
        "в",
        "во",
        "на",
        "по",
        "за",
        "из",
        "к",
        "ко",
        "с",
        "со",
        "у",
        "о",
        "об",
        "от",
        "до",
        "без",
        "для",
        "над",
        "под",
        "при",
        "про",
    }
)

MONOSPACE_FONTS = frozenset(
    {
        "consolas",
        "courier",
        "courier new",
        "lucida console",
        "monaco",
        "menlo",
        "source code pro",
        "fira code",
        "jetbrains mono",
        "cascadia mono",
        "cascadia code",
        "roboto mono",
        "ubuntu mono",
        "pt mono",
        "dejavu sans mono",
        "noto sans mono",
        "sarasa mono",
    }
)

# Шрифты, которые не требуют ppt/fonts.
STANDARD_FONTS = frozenset(
    {
        "arial",
        "arial black",
        "arial narrow",
        "aptos",
        "aptos display",
        "aptos narrow",
        "bahnschrift",
        "book antiqua",
        "calibri",
        "calibri light",
        "cambria",
        "cambria math",
        "candara",
        "century gothic",
        "comic sans ms",
        "consolas",
        "constantia",
        "corbel",
        "courier new",
        "franklin gothic",
        "franklin gothic medium",
        "garamond",
        "georgia",
        "impact",
        "lucida console",
        "lucida sans unicode",
        "microsoft sans serif",
        "palatino linotype",
        "segoe ui",
        "segoe ui light",
        "segoe ui semibold",
        "symbol",
        "tahoma",
        "times new roman",
        "trebuchet ms",
        "verdana",
        "webdings",
        "wingdings",
    }
)

ALLOWED_MIME_TYPES = frozenset(
    {
        "",
        "application/octet-stream",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "application/x-zip-compressed",
        "application/zip",
    }
)


def frontend_origins() -> list[str]:
    raw = os.environ.get("FRONTEND_URL", "")
    return [item.strip() for item in raw.split(",") if item.strip()]


def xff_index() -> int:
    """Какой адрес из X-Forwarded-For считать клиентским. 0 — левый, -1 — правый."""
    raw = os.environ.get("XFF_INDEX", "0")
    try:
        return int(raw)
    except ValueError:
        return 0
