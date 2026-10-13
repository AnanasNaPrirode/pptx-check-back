from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class RunModel:
    text: str
    font_name: str | None
    font_size_pt: float | None
    color_rgb: tuple[int, int, int] | None
    superscript: bool
    monospace: bool
    text_transparent: bool


@dataclass
class LineModel:
    runs: list[RunModel]
    explicit_break_after: bool

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)


@dataclass
class ParagraphModel:
    level: int
    bullet: bool
    numbered: bool
    lines: list[LineModel]

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)


@dataclass
class ElementModel:
    id: str
    type: str
    x: int
    y: int
    w: int
    h: int
    text: str
    is_title: bool
    placeholder_kind: str | None
    is_decoration: bool
    paragraphs: list[ParagraphModel]
    fill_type: str | None
    fill_rgb: tuple[int, int, int] | None
    fill_transparent: bool

    def public(self) -> dict:
        payload: dict = {
            "id": self.id,
            "type": self.type,
            "x": self.x,
            "y": self.y,
            "w": self.w,
            "h": self.h,
        }
        if self.type in {"text", "placeholder", "table"} or self.text:
            payload["text"] = self.text
        if self.is_title:
            payload["is_title"] = True
        return payload


@dataclass
class Issue:
    check_id: str
    level: str
    message: str
    element_id: str | None = None
    slide_numbers: list[int] | None = None

    def slide_public(self) -> dict:
        return {
            "check_id": self.check_id,
            "level": self.level,
            "message": self.message,
            "element_id": self.element_id,
        }

    def presentation_public(self) -> dict:
        payload = {
            "check_id": self.check_id,
            "level": self.level,
            "message": self.message,
        }
        if self.slide_numbers is not None:
            payload["slide_numbers"] = self.slide_numbers
        return payload


@dataclass
class SlideModel:
    number: int
    width_emu: int
    height_emu: int
    elements: list[ElementModel]
    has_video: bool
    has_title_placeholder: bool
    issues: list[Issue] = field(default_factory=list)

    def public(self) -> dict:
        return {
            "number": self.number,
            "width_emu": self.width_emu,
            "height_emu": self.height_emu,
            "elements": [element.public() for element in self.elements],
            "issues": [issue.slide_public() for issue in self.issues],
        }


@dataclass
class ArchiveFacts:
    has_fonts_dir: bool


@dataclass
class PresentationModel:
    width_emu: int
    height_emu: int
    slides: list[SlideModel]
    archive: ArchiveFacts
    presentation_issues: list[Issue] = field(default_factory=list)
