from __future__ import annotations

from threading import Event, Lock
from typing import Callable


class CancelledRequest(RuntimeError):
    def __init__(self, request_id: str):
        self.request_id = request_id
        super().__init__("요청이 취소되었습니다.")


_lock = Lock()
_events: dict[str, Event] = {}


def register_request(request_id: str) -> None:
    if not request_id:
        return
    with _lock:
        _events[request_id] = Event()


def cancel_request(request_id: str) -> bool:
    if not request_id:
        return False
    with _lock:
        event = _events.get(request_id)
    if event is None:
        return False
    event.set()
    return True


def unregister_request(request_id: str) -> None:
    if not request_id:
        return
    with _lock:
        _events.pop(request_id, None)


def is_cancelled(request_id: str) -> bool:
    if not request_id:
        return False
    with _lock:
        event = _events.get(request_id)
    return bool(event and event.is_set())


def throw_if_cancelled(request_id: str) -> None:
    if is_cancelled(request_id):
        raise CancelledRequest(request_id)


def cancel_checker(request_id: str) -> Callable[[], None]:
    def check() -> None:
        throw_if_cancelled(request_id)

    return check
