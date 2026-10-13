from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import PurePosixPath

from app import config
from app.errors import ApiError


def validate_upload(filename: str | None, content_type: str | None, data: bytes) -> list[str]:
    if not filename or not filename.lower().endswith(".pptx"):
        raise ApiError(415, "UNSUPPORTED_TYPE", "Допустим только формат .pptx.")
    mime = (content_type or "").split(";", 1)[0].strip().lower()
    if mime not in config.ALLOWED_MIME_TYPES:
        raise ApiError(415, "UNSUPPORTED_TYPE", "Неверный тип файла.")
    if len(data) > config.MAX_FILE_BYTES:
        raise ApiError(413, "FILE_TOO_LARGE", "Файл больше 5 МБ.")
    if len(data) < 4 or data[:2] != b"PK":
        raise ApiError(400, "INVALID_FILE", "Файл повреждён или не является презентацией .pptx.")
    try:
        archive = zipfile.ZipFile(BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ApiError(400, "INVALID_FILE", "Файл повреждён или не является презентацией .pptx.") from exc
    infos = archive.infolist()
    if len(infos) > config.MAX_ZIP_ENTRIES:
        raise ApiError(400, "ARCHIVE_TOO_LARGE", "Архив презентации слишком большой.")
    total = 0
    names: list[str] = []
    for info in infos:
        name = info.filename.replace("\\", "/")
        parts = PurePosixPath(name).parts
        if name.startswith("/") or ".." in parts:
            raise ApiError(400, "INVALID_FILE", "Файл повреждён или не является презентацией .pptx.")
        if info.flag_bits & 0x1:
            raise ApiError(400, "INVALID_FILE", "Файл повреждён или не является презентацией .pptx.")
        total += max(info.file_size, 0)
        if total > config.MAX_UNCOMPRESSED_BYTES:
            raise ApiError(400, "ARCHIVE_TOO_LARGE", "Архив презентации слишком большой.")
        base = name.rsplit("/", 1)[-1]
        if base == "vbaProject.bin":
            raise ApiError(400, "MACROS_NOT_ALLOWED", "Файлы с макросами не принимаются.")
        names.append(name)
    if "[Content_Types].xml" not in names or "ppt/presentation.xml" not in names:
        raise ApiError(400, "INVALID_FILE", "Файл повреждён или не является презентацией .pptx.")
    return names
