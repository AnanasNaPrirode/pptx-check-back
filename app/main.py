from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app import config
from app.checks import catalog
from app.config import TIMEOUT_MESSAGE
from app.errors import ApiError
from app.logging_setup import log
from app.pipeline import build_report, parse_enabled_checks
from app.rate_limit import RateLimiter


def create_app() -> FastAPI:
    app = FastAPI(title="Presentation Checker", version="0.1.0")
    app.state.limiter = RateLimiter(config.RATE_LIMIT_SECONDS)
    origins = config.frontend_origins()
    if origins:
        app.add_middleware(
            CORSMiddleware,
            allow_origins=origins,
            allow_credentials=True,
            allow_methods=["*"],
            allow_headers=["*"],
        )

    @app.exception_handler(ApiError)
    def handle_api_error(_request: Request, exc: ApiError) -> JSONResponse:
        return JSONResponse(status_code=exc.status, content=exc.body(), headers=exc.headers)

    @app.exception_handler(Exception)
    def handle_unexpected(_request: Request, exc: Exception) -> JSONResponse:
        log.error("internal_error", extra={"error_code": type(exc).__name__})
        return JSONResponse(
            status_code=500,
            content={"error": {"code": "INTERNAL_ERROR", "message": "Не удалось проверить файл."}},
        )

    @app.get("/api/health")
    def health() -> dict:
        return {"status": "ok"}

    @app.get("/api/checks")
    def list_checks() -> list[dict]:
        return [item.as_dict() for item in catalog()]

    @app.post("/api/check")
    def check_presentation(
        request: Request,
        file: UploadFile | None = File(None),
        enabled_checks: str | None = Form(None),
    ) -> JSONResponse:
        started = time.perf_counter()
        status = 500
        error_code: str | None = None
        file_size = 0
        slides_count: int | None = None
        applied: list[str] | None = None
        try:
            ip = _client_ip(request)
            retry_after = request.app.state.limiter.retry_after(ip)
            if retry_after is not None:
                raise ApiError(
                    429,
                    "RATE_LIMITED",
                    "Слишком частые запросы. Повторите позже.",
                    headers={"Retry-After": str(retry_after)},
                )
            _reject_huge_content_length(request)
            if file is None:
                raise ApiError(400, "INVALID_FILE", "Файл не передан.")
            data = file.file.read(config.MAX_FILE_BYTES + 1)
            file_size = len(data)
            if file_size > config.MAX_FILE_BYTES:
                raise ApiError(413, "FILE_TOO_LARGE", "Файл больше 5 МБ.")
            # Список проверок разбираем до тяжёлой работы, чтобы неизвестный id
            # отвечал 422 и без разбора презентации.
            applied = [item.id for item in parse_enabled_checks(enabled_checks)]
            report = _run_bounded(file.filename, file.content_type, data, enabled_checks)
            slides_count = report["slides_count"]
            applied = report["applied_checks"]
            status = 200
            return JSONResponse(status_code=200, content=report)
        except ApiError as exc:
            status = exc.status
            error_code = exc.code
            raise
        finally:
            duration_ms = round((time.perf_counter() - started) * 1000, 1)
            log.info(
                "request",
                extra={
                    "endpoint": "/api/check",
                    "status": status,
                    "duration_ms": duration_ms,
                    "file_size": file_size,
                    "slides_count": slides_count,
                    "enabled_checks": applied,
                    "error_code": error_code,
                },
            )

    return app


def _run_bounded(filename: str | None, content_type: str | None, data: bytes, enabled_raw: str | None) -> dict:
    executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="pptx-check")
    future = executor.submit(build_report, filename, content_type, data, enabled_raw)
    try:
        return future.result(timeout=config.CHECK_TIMEOUT_SECONDS)
    except TimeoutError as exc:
        raise ApiError(504, "TIMEOUT", TIMEOUT_MESSAGE) from exc
    except ApiError:
        raise
    except Exception as exc:
        raise ApiError(500, "INTERNAL_ERROR", "Не удалось проверить файл.") from exc
    finally:
        executor.shutdown(wait=False, cancel_futures=True)


def _reject_huge_content_length(request: Request) -> None:
    raw = request.headers.get("content-length")
    if raw is None:
        return
    try:
        length = int(raw)
    except ValueError:
        return
    if length > config.MAX_CONTENT_LENGTH_BYTES:
        raise ApiError(413, "FILE_TOO_LARGE", "Файл больше 5 МБ.")


def _client_ip(request: Request) -> str:
    forwarded = request.headers.get("x-forwarded-for", "")
    parts = [item.strip() for item in forwarded.split(",") if item.strip()]
    if parts:
        index = config.xff_index()
        if index < 0:
            return parts[index]
        if index < len(parts):
            return parts[index]
        return parts[-1]
    if request.client is not None and request.client.host:
        return request.client.host
    return "unknown"


app = create_app()


def main() -> None:
    import uvicorn

    port = int(os.environ.get("PORT", "8080"))
    uvicorn.run("app.main:app", host="0.0.0.0", port=port)


if __name__ == "__main__":
    main()
