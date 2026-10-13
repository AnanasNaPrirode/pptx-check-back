from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass

from app import config
from app.models import ElementModel, Issue, PresentationModel, SlideModel


@dataclass(frozen=True)
class Check:
    id: str
    name: str
    description: str
    level: str
    scope: str
    default_enabled: bool = True

    def as_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "level": self.level,
            "scope": self.scope,
            "default_enabled": self.default_enabled,
        }


def catalog() -> list[Check]:
    return list(CHECKS)


def by_id() -> dict[str, Check]:
    return {item.id: item for item in CHECKS}


CHECKS: list[Check] = []


def _register(check: Check) -> Check:
    CHECKS.append(check)
    return check


ASPECT_16_9 = _register(
    Check(
        id="aspect_16_9",
        name="Соотношение сторон 16:9",
        description="Соотношение сторон слайда 16:9, допуск ±2%.",
        level="error",
        scope="presentation",
    )
)
EMBEDDED_FONTS = _register(
    Check(
        id="embedded_fonts",
        name="Встроенные шрифты",
        description="Использован нестандартный шрифт, а папки ppt/fonts в архиве нет.",
        level="warning",
        scope="presentation",
    )
)
SLIDE_DUPLICATES = _register(
    Check(
        id="slide_duplicates",
        name="Дубли слайдов",
        description="Слайды с одинаковым текстом и одинаковым набором типов объектов.",
        level="warning",
        scope="presentation",
    )
)
MISSING_TITLE = _register(
    Check(
        id="missing_title",
        name="Заголовок слайда",
        description="Есть title-плейсхолдер и в нём непустой текст.",
        level="error",
        scope="slide",
    )
)
TITLE_DOT = _register(
    Check(
        id="title_dot",
        name="Точка в заголовке",
        description="Заголовок не заканчивается точкой.",
        level="warning",
        scope="slide",
    )
)
HAS_VIDEO = _register(
    Check(
        id="has_video",
        name="Встроенное видео",
        description="На слайде есть встроенное видео.",
        level="warning",
        scope="slide",
    )
)
QUOTE_GUILLEMETS = _register(
    Check(
        id="quote_guillemets",
        name="Кавычки-ёлочки",
        description="Прямые кавычки вместо «ёлочек».",
        level="warning",
        scope="slide",
    )
)
EM_DASH = _register(
    Check(
        id="em_dash",
        name="Длинное тире",
        description="Дефис с пробелами между словами или в числовом диапазоне вместо «—».",
        level="warning",
        scope="slide",
    )
)
ORDINALS = _register(
    Check(
        id="ordinals",
        name="Наращения числительных",
        description="Буквенные наращения вида «2-го» вместо словесной записи.",
        level="warning",
        scope="slide",
    )
)
LIST_PERIODS = _register(
    Check(
        id="list_periods",
        name="Точки в списке",
        description="Пункты одного списка по-разному заканчиваются точкой.",
        level="warning",
        scope="slide",
    )
)
TITLE_HEURISTIC = _register(
    Check(
        id="title_heuristic",
        name="Заголовок без плейсхолдера",
        description="Заголовок определяется по позиции в верхней зоне и крупному шрифту. Для файлов без плейсхолдеров.",
        level="warning",
        scope="slide",
        default_enabled=False,
    )
)
FONT_SIZE_MIN = _register(
    Check(
        id="font_size_min",
        name="Минимальный размер шрифта",
        description="Текст мельче 18 pt. Футер, номер слайда и дата не учитываются.",
        level="warning",
        scope="slide",
    )
)
BOTTOM_MARGIN = _register(
    Check(
        id="bottom_margin",
        name="Нижний отступ",
        description="Контент ближе 6% высоты слайда к нижнему краю.",
        level="warning",
        scope="slide",
    )
)
NUMBER_GROUPING = _register(
    Check(
        id="number_grouping",
        name="Краткая запись чисел",
        description="Числа длиннее 4 цифр: лучше краткая запись (тыс., млн).",
        level="warning",
        scope="slide",
    )
)
ENGLISH_RATIO = _register(
    Check(
        id="english_ratio",
        name="Доля латиницы",
        description="Латиницы больше 35% букв относительно кириллицы; код не учитывается.",
        level="warning",
        scope="slide",
    )
)
SUPERSCRIPT_UNITS = _register(
    Check(
        id="superscript_units",
        name="Верхний индекс в единицах",
        description="Степени (м², км³) записаны обычными цифрами, а не верхним индексом.",
        level="warning",
        scope="slide",
    )
)
OVERLAP = _register(
    Check(
        id="overlap",
        name="Пересечение текста",
        description="Текстовые блоки пересекаются на 15% и больше.",
        level="warning",
        scope="slide",
    )
)
CONTRAST_FILL = _register(
    Check(
        id="contrast_fill",
        name="Контраст текста и заливки",
        description="Контраст текста и сплошной заливки ниже WCAG: 4.5:1 для обычного текста, 4:1 для крупного (от 18 pt).",
        level="warning",
        scope="slide",
    )
)
HANGING_PREPOSITIONS = _register(
    Check(
        id="hanging_prepositions",
        name="Висячие предлоги",
        description="Короткое слово в конце строки перед явным разрывом.",
        level="warning",
        scope="slide",
    )
)


_WORD_HYPHEN = re.compile(r"[A-Za-zА-Яа-яЁё]{2,}[ \t]+-[ \t]+[A-Za-zА-Яа-яЁё]{2,}")
_NUM_RANGE = re.compile(r"(?<!\d)(\d+)\s*-\s*(\d+)(?!\d)")
_ORDINAL = re.compile(
    r"(?<!\d)(\d+)-(го|му|ми|ом|ой|ый|ая|ое|ые|ым|ых|ою|ею|ти|мя|й|я|ю|е|х|м)\b",
    re.IGNORECASE,
)
_LONG_NUMBER = re.compile(r"(?<![\dA-Za-zА-Яа-яЁё])(\d{5,})(?!\d)")
_GROUPED_NUMBER = re.compile(r"(?<!\d)(\d{1,3}(?:[ \u00a0]\d{3})+)(?!\d)")
_UNIT = re.compile(r"(км|см|мм|дм|м)([23])", re.IGNORECASE)
_LATIN = re.compile(r"[A-Za-z]")
_CYRILLIC = re.compile(r"[А-Яа-яЁё]")
_STRAIGHT_QUOTE = '"'


def run_check(check: Check, presentation: PresentationModel) -> None:
    if check.scope == "presentation":
        presentation.presentation_issues.extend(_presentation_issues(check, presentation))
        return
    for slide in presentation.slides:
        slide.issues.extend(_slide_issues(check, slide))


def _presentation_issues(check: Check, presentation: PresentationModel) -> list[Issue]:
    if check.id == "aspect_16_9":
        return _aspect(presentation)
    if check.id == "embedded_fonts":
        return _embedded_fonts(presentation)
    if check.id == "slide_duplicates":
        return _duplicates(presentation)
    return []


def _slide_issues(check: Check, slide: SlideModel) -> list[Issue]:
    runners = {
        "missing_title": _missing_title,
        "title_dot": _title_dot,
        "has_video": _has_video,
        "quote_guillemets": _quotes,
        "em_dash": _em_dash,
        "ordinals": _ordinals,
        "list_periods": _list_periods,
        "title_heuristic": _title_heuristic,
        "font_size_min": _font_size,
        "bottom_margin": _bottom_margin,
        "number_grouping": _numbers,
        "english_ratio": _english,
        "superscript_units": _units,
        "overlap": _overlap,
        "contrast_fill": _contrast,
        "hanging_prepositions": _hanging,
    }
    runner = runners.get(check.id)
    if runner is None:
        return []
    return runner(slide)


def _issue(check_id: str, message: str, element_id: str | None = None, slide_numbers: list[int] | None = None) -> Issue:
    level = by_id()[check_id].level
    return Issue(check_id, level, message, element_id, slide_numbers)


def _aspect(presentation: PresentationModel) -> list[Issue]:
    if presentation.height_emu <= 0:
        return []
    ratio = presentation.width_emu / presentation.height_emu
    if abs(ratio - config.ASPECT_TARGET) / config.ASPECT_TARGET <= config.ASPECT_TOLERANCE:
        return []
    return [_issue("aspect_16_9", "Соотношение сторон слайда не 16:9.")]


def _embedded_fonts(presentation: PresentationModel) -> list[Issue]:
    if presentation.archive.has_fonts_dir:
        return []
    unusual: list[str] = []
    seen: set[str] = set()
    for slide in presentation.slides:
        for element in slide.elements:
            for paragraph in element.paragraphs:
                for line in paragraph.lines:
                    for run in line.runs:
                        name = run.font_name
                        if not name:
                            continue
                        key = name.casefold()
                        if key in seen or key in config.STANDARD_FONTS:
                            continue
                        seen.add(key)
                        unusual.append(name)
    if not unusual:
        return []
    shown = ", ".join(f"«{name}»" for name in unusual[:3])
    return [
        _issue(
            "embedded_fonts",
            f"Использован нестандартный шрифт {shown}, но в файле нет встроенных шрифтов.",
        )
    ]


def _duplicates(presentation: PresentationModel) -> list[Issue]:
    groups: dict[tuple, list[int]] = defaultdict(list)
    for slide in presentation.slides:
        ordered = sorted(slide.elements, key=lambda item: (item.y, item.x, item.id))
        text = re.sub(r"\s+", " ", "\n".join(element.text for element in ordered)).strip()
        types = tuple(sorted(element.type for element in slide.elements))
        groups[(text, types)].append(slide.number)
    issues: list[Issue] = []
    for numbers in groups.values():
        if len(numbers) < 2:
            continue
        numbers = sorted(numbers)
        issues.append(
            _issue(
                "slide_duplicates",
                f"Слайды {_join_numbers(numbers)} совпадают по тексту и набору объектов.",
                slide_numbers=numbers,
            )
        )
    return issues


def _join_numbers(numbers: list[int]) -> str:
    rendered = [str(number) for number in numbers]
    if len(rendered) == 2:
        return f"{rendered[0]} и {rendered[1]}"
    return ", ".join(rendered[:-1]) + f" и {rendered[-1]}"


def _titles(slide: SlideModel) -> list[ElementModel]:
    return [element for element in slide.elements if element.placeholder_kind in {"title", "center_title", "vertical_title"}]


def _missing_title(slide: SlideModel) -> list[Issue]:
    titles = _titles(slide)
    if not titles:
        return [_issue("missing_title", "На слайде нет заголовка.")]
    if any(element.text.strip() for element in titles):
        return []
    return [_issue("missing_title", "Заголовок пустой.", titles[0].id)]


def _title_dot(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in _titles(slide):
        text = element.text.strip()
        if not text:
            continue
        if text.endswith("."):
            issues.append(_issue("title_dot", "Заголовок заканчивается точкой.", element.id))
    return issues


def _has_video(slide: SlideModel) -> list[Issue]:
    if not slide.has_video:
        return []
    video = next((element for element in slide.elements if element.type == "video"), None)
    return [_issue("has_video", "На слайде есть встроенное видео.", video.id if video else None)]


def _text_elements(slide: SlideModel) -> list[ElementModel]:
    return [element for element in slide.elements if element.text.strip()]


def _quotes(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in _text_elements(slide):
        if _STRAIGHT_QUOTE in element.text:
            issues.append(
                _issue(
                    "quote_guillemets",
                    "Прямые кавычки лучше заменить на ёлочки «...»." ,
                    element.id,
                )
            )
    return issues


def _em_dash(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in _text_elements(slide):
        word = _WORD_HYPHEN.search(element.text)
        number = _NUM_RANGE.search(element.text)
        if word:
            issues.append(
                _issue("em_dash", f"Дефис в «{word.group(0)}» лучше заменить на тире «—».", element.id)
            )
        elif number:
            issues.append(
                _issue(
                    "em_dash",
                    f"Диапазон «{number.group(0)}» лучше записать через тире «—».",
                    element.id,
                )
            )
    return issues


def _ordinals(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in _text_elements(slide):
        match = _ORDINAL.search(element.text)
        if match:
            issues.append(
                _issue(
                    "ordinals",
                    f"Наращение «{match.group(0)}» лучше заменить словесной записью.",
                    element.id,
                )
            )
    return issues


def _list_periods(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in slide.elements:
        items: list[str] = []
        current_level: int | None = None
        for paragraph in element.paragraphs:
            if not (paragraph.bullet or paragraph.numbered):
                _flush_list(items, element, issues)
                items = []
                current_level = None
                continue
            if current_level is None:
                current_level = paragraph.level
            if paragraph.level != current_level:
                _flush_list(items, element, issues)
                items = []
                current_level = paragraph.level
            text = paragraph.text.strip()
            if text:
                items.append(text)
        _flush_list(items, element, issues)
    return issues


def _flush_list(items: list[str], element: ElementModel, issues: list[Issue]) -> None:
    if len(items) < 2:
        return
    ended = [item.endswith(".") for item in items]
    if any(ended) and not all(ended):
        issues.append(_issue("list_periods", "Пункты списка по-разному заканчиваются точкой.", element.id))


def _title_heuristic(slide: SlideModel) -> list[Issue]:
    if slide.has_title_placeholder:
        return []
    zone = slide.height_emu * config.TITLE_ZONE_RATIO
    for element in slide.elements:
        if element.is_decoration or not element.text.strip():
            continue
        if element.y >= zone:
            continue
        sizes = [
            run.font_size_pt
            for paragraph in element.paragraphs
            for line in paragraph.lines
            for run in line.runs
            if run.font_size_pt is not None and run.text.strip()
        ]
        if sizes and max(sizes) >= config.TITLE_HEURISTIC_MIN_PT:
            return []
    return [_issue("title_heuristic", "Заголовок не найден в верхней зоне слайда.")]


def _font_size(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in slide.elements:
        if element.is_decoration:
            continue
        small: list[float] = []
        for paragraph in element.paragraphs:
            for line in paragraph.lines:
                for run in line.runs:
                    if not run.text.strip() or run.font_size_pt is None:
                        continue
                    if run.font_size_pt < config.MIN_FONT_SIZE_PT:
                        small.append(run.font_size_pt)
        if not small:
            continue
        size = min(small)
        issues.append(
            _issue(
                "font_size_min",
                f"Размер шрифта {_format_pt(size)} pt меньше минимального ({_format_pt(config.MIN_FONT_SIZE_PT)} pt).",
                element.id,
            )
        )
    return issues


def _format_pt(value: float) -> str:
    if abs(value - round(value)) < 0.05:
        return str(int(round(value)))
    return f"{value:.1f}"


def _bottom_margin(slide: SlideModel) -> list[Issue]:
    if slide.height_emu <= 0:
        return []
    limit = slide.height_emu * config.BOTTOM_MARGIN_RATIO
    issues: list[Issue] = []
    for element in slide.elements:
        if element.is_decoration or element.w <= 0 or element.h <= 0:
            continue
        gap = slide.height_emu - (element.y + element.h)
        if gap < limit:
            issues.append(
                _issue("bottom_margin", "Контент ближе 6% высоты слайда к нижнему краю.", element.id)
            )
    return issues


def _numbers(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in _text_elements(slide):
        found: list[str] = []
        for match in _LONG_NUMBER.finditer(element.text):
            found.append(match.group(1))
        for match in _GROUPED_NUMBER.finditer(element.text):
            digits = re.sub(r"\D", "", match.group(1))
            if len(digits) > config.NUMBER_GROUPING_MAX_DIGITS:
                found.append(digits)
        if not found:
            continue
        suggestion = _compact_number(found[0])
        issues.append(
            _issue(
                "number_grouping",
                f"Число «{found[0]}» лучше записать кратко ({suggestion}).",
                element.id,
            )
        )
    return issues


def _compact_number(digits: str) -> str:
    value = int(digits)
    if value >= 1_000_000_000:
        return _scaled(value / 1_000_000_000, "млрд")
    if value >= 1_000_000:
        return _scaled(value / 1_000_000, "млн")
    if value >= 1_000:
        return _scaled(value / 1_000, "тыс.")
    return digits


def _scaled(value: float, suffix: str) -> str:
    if abs(value - round(value)) < 0.05:
        return f"{int(round(value))} {suffix}"
    return f"{value:.1f}".replace(".", ",") + f" {suffix}"


def _english(slide: SlideModel) -> list[Issue]:
    latin = 0
    cyrillic = 0
    for element in slide.elements:
        for paragraph in element.paragraphs:
            for line in paragraph.lines:
                for run in line.runs:
                    if run.monospace:
                        continue
                    latin += len(_LATIN.findall(run.text))
                    cyrillic += len(_CYRILLIC.findall(run.text))
    total = latin + cyrillic
    if total == 0:
        return []
    ratio = latin / total
    if ratio <= config.ENGLISH_RATIO_MAX:
        return []
    return [
        _issue(
            "english_ratio",
            f"Доля латиницы {round(ratio * 100)}% превышает 35%.",
        )
    ]


def _units(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in slide.elements:
        chars: list[tuple[str, bool]] = []
        for paragraph in element.paragraphs:
            for line in paragraph.lines:
                for run in line.runs:
                    for char in run.text:
                        chars.append((char, run.superscript))
        if not chars:
            continue
        plain = "".join(char for char, _ in chars)
        for match in _UNIT.finditer(plain):
            digit_index = match.start(2)
            if chars[digit_index][1]:
                continue
            unit = match.group(1)
            digit = match.group(2)
            pretty = {"2": "²", "3": "³"}[digit]
            issues.append(
                _issue(
                    "superscript_units",
                    f"Единица «{unit}{digit}» лучше записать как «{unit}{pretty}».",
                    element.id,
                )
            )
            break
    return issues


def _overlap(slide: SlideModel) -> list[Issue]:
    blocks = [
        element
        for element in slide.elements
        if element.text.strip() and not element.is_decoration and element.w > 0 and element.h > 0
    ]
    issues: list[Issue] = []
    for index, left in enumerate(blocks):
        for right in blocks[index + 1 :]:
            ratio = _overlap_ratio(left, right)
            if ratio + 1e-9 >= config.OVERLAP_MIN_RATIO:
                percent = round(ratio * 100)
                issues.append(
                    _issue(
                        "overlap",
                        f"Текстовые блоки пересекаются на {percent}%.",
                        right.id,
                    )
                )
    return issues


def _overlap_ratio(left: ElementModel, right: ElementModel) -> float:
    x1 = max(left.x, right.x)
    y1 = max(left.y, right.y)
    x2 = min(left.x + left.w, right.x + right.w)
    y2 = min(left.y + left.h, right.y + right.h)
    if x2 <= x1 or y2 <= y1:
        return 0
    intersection = (x2 - x1) * (y2 - y1)
    smallest = min(left.w * left.h, right.w * right.h)
    if smallest <= 0:
        return 0
    return intersection / smallest


def _contrast(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in slide.elements:
        if element.is_decoration or element.fill_type != "solid":
            continue
        if element.fill_transparent or element.fill_rgb is None:
            continue
        worst: tuple[float, float] | None = None
        for paragraph in element.paragraphs:
            for line in paragraph.lines:
                for run in line.runs:
                    if not run.text.strip() or run.font_size_pt is None or run.color_rgb is None:
                        continue
                    if run.text_transparent:
                        continue
                    ratio = _contrast_ratio(run.color_rgb, element.fill_rgb)
                    required = (
                        config.CONTRAST_LARGE
                        if run.font_size_pt >= config.CONTRAST_LARGE_FONT_PT
                        else config.CONTRAST_NORMAL
                    )
                    if ratio + 1e-9 < required and (worst is None or ratio < worst[0]):
                        worst = (ratio, required)
        if worst is None:
            continue
        ratio, required = worst
        issues.append(
            _issue(
                "contrast_fill",
                f"Контраст текста и фона {ratio:.2f}:1 ниже требуемого {required:.1f}:1.",
                element.id,
            )
        )
    return issues


def _channel(value: int) -> float:
    channel = value / 255
    if channel <= 0.04045:
        return channel / 12.92
    return ((channel + 0.055) / 1.055) ** 2.4


def _contrast_ratio(foreground: tuple[int, int, int], background: tuple[int, int, int]) -> float:
    def luminance(rgb: tuple[int, int, int]) -> float:
        red, green, blue = rgb
        return 0.2126 * _channel(red) + 0.7152 * _channel(green) + 0.0722 * _channel(blue)

    lighter = max(luminance(foreground), luminance(background))
    darker = min(luminance(foreground), luminance(background))
    return (lighter + 0.05) / (darker + 0.05)


def _hanging(slide: SlideModel) -> list[Issue]:
    issues: list[Issue] = []
    for element in slide.elements:
        if element.is_decoration:
            continue
        reported = False
        for paragraph in element.paragraphs:
            for line in paragraph.lines:
                if not line.explicit_break_after:
                    continue
                word = _last_word(line.text)
                if word is None:
                    continue
                folded = word.casefold()
                if len(folded) <= 3 and folded in config.HANGING_WORDS:
                    issues.append(
                        _issue(
                            "hanging_prepositions",
                            f"Короткое слово «{folded}» стоит в конце строки перед разрывом.",
                            element.id,
                        )
                    )
                    reported = True
                    break
            if reported:
                break
    return issues


def _last_word(text: str) -> str | None:
    stripped = text.strip()
    if not stripped:
        return None
    token = stripped.split()[-1].strip(".,;:!?»«\"()[]")
    if not token:
        return None
    return token
