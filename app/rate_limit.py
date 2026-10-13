from __future__ import annotations

import math
import threading
import time


class RateLimiter:
    def __init__(self, interval_seconds: float) -> None:
        self.interval = interval_seconds
        self._last: dict[str, float] = {}
        self._lock = threading.Lock()

    def retry_after(self, key: str) -> int | None:
        now = time.monotonic()
        with self._lock:
            previous = self._last.get(key)
            if previous is not None and now - previous < self.interval:
                return max(1, math.ceil(self.interval - (now - previous)))
            self._last[key] = now
            return None
