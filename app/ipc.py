from __future__ import annotations

import json
import socket
import threading
from typing import Callable

IPC_HOST = "127.0.0.1"
IPC_PORT = 48221


def send_message(message: dict, timeout: float = 1.5) -> bool:
    payload = (json.dumps(message) + "\n").encode("utf-8")
    try:
        with socket.create_connection((IPC_HOST, IPC_PORT), timeout=timeout) as conn:
            conn.sendall(payload)
        return True
    except OSError:
        return False


class IpcServer:
    def __init__(self, on_message: Callable[[dict], None]) -> None:
        self._on_message = on_message
        self._sock: socket.socket | None = None
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()

    def start(self) -> bool:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind((IPC_HOST, IPC_PORT))
            sock.listen(8)
        except OSError:
            sock.close()
            return False

        self._sock = sock
        self._thread = threading.Thread(target=self._serve_loop, daemon=True)
        self._thread.start()
        return True

    def stop(self) -> None:
        self._stop_event.set()
        if self._sock:
            try:
                self._sock.close()
            except OSError:
                pass
            self._sock = None

    def _serve_loop(self) -> None:
        assert self._sock is not None
        while not self._stop_event.is_set():
            try:
                conn, _ = self._sock.accept()
            except OSError:
                break
            with conn:
                data = bytearray()
                while True:
                    chunk = conn.recv(8192)
                    if not chunk:
                        break
                    data.extend(chunk)

                for line in data.decode("utf-8", errors="ignore").splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        payload = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    self._on_message(payload)
