"""
IPCServer — TCP socket server that receives JSON requests from the Qt frontend
and dispatches them to SparkEngine.

Protocol (line-delimited JSON over TCP):
  Request:  { "id": <int>, "action": "<string>", ...params... }
  Response: { "id": <int>, ...result or "error": "<string>"... }

The Qt side opens a persistent connection on localhost:PORT, sends one
newline-terminated JSON request, and reads one newline-terminated JSON response
per request.  The connection stays open until Qt closes it.

Default port: 9876  (configurable via --port or OPENSHEET_IPC_PORT env var)
"""
from __future__ import annotations

import json
import logging
import os
import socket
import threading
from typing import Any

from .spark_engine import SparkEngine

logger = logging.getLogger(__name__)

DEFAULT_PORT = int(os.environ.get("OPENSHEET_IPC_PORT", "9876"))


class IPCServer:
    def __init__(self, port: int = DEFAULT_PORT) -> None:
        self.port = port
        self.engine = SparkEngine()
        self._running = False
        self._server_sock: socket.socket | None = None

    # ── Start / stop ───────────────────────────────────────────────────────────

    def start(self) -> None:
        """Block and serve forever (call from a dedicated thread)."""
        self._running = True
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as srv:
            self._server_sock = srv
            srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            srv.bind(("127.0.0.1", self.port))
            srv.listen(8)
            logger.info("OpenSheet IPC server listening on port %d", self.port)
            while self._running:
                try:
                    srv.settimeout(1.0)
                    conn, addr = srv.accept()
                    t = threading.Thread(
                        target=self._handle_client,
                        args=(conn,),
                        daemon=True,
                    )
                    t.start()
                except socket.timeout:
                    continue
                except OSError:
                    break

    def stop(self) -> None:
        self._running = False
        self.engine.shutdown()
        if self._server_sock:
            try:
                self._server_sock.close()
            except OSError:
                pass

    # ── Client handler ─────────────────────────────────────────────────────────

    def _handle_client(self, conn: socket.socket) -> None:
        buf = b""
        with conn:
            while True:
                try:
                    chunk = conn.recv(65536)
                    if not chunk:
                        break
                    buf += chunk
                    while b"\n" in buf:
                        line, buf = buf.split(b"\n", 1)
                        if not line.strip():
                            continue
                        response = self._dispatch(line)
                        conn.sendall(json.dumps(response).encode() + b"\n")
                except (OSError, ConnectionResetError):
                    break

    def _dispatch(self, raw: bytes) -> dict:
        try:
            req: dict = json.loads(raw)
        except json.JSONDecodeError as exc:
            return {"error": f"Invalid JSON: {exc}"}

        req_id = req.get("id", 0)
        action = req.get("action", "")

        try:
            if action == "ping":
                result: dict[str, Any] = {"pong": True}
            elif action == "load_csv":
                result = self.engine.load_csv(
                    req["path"],
                    chunk=req.get("chunk", 0),
                    chunk_size=req.get("chunk_size", 1000),
                )
            elif action == "load_parquet":
                result = self.engine.load_parquet(
                    req["path"],
                    chunk=req.get("chunk", 0),
                    chunk_size=req.get("chunk_size", 1000),
                )
            elif action == "get_chunk":
                result = self.engine.get_chunk(
                    req["df_id"],
                    chunk=req["chunk"],
                    chunk_size=req.get("chunk_size", 1000),
                )
            elif action == "aggregate":
                result = self.engine.aggregate(req["df_id"], req["operations"])
            elif action == "filter":
                result = self.engine.filter_rows(
                    req["df_id"],
                    column=req["column"],
                    operator=req["operator"],
                    value=str(req["value"]),
                    chunk=req.get("chunk", 0),
                    chunk_size=req.get("chunk_size", 1000),
                )
            elif action == "release":
                result = self.engine.release(req["df_id"])
            elif action == "shutdown":
                self.stop()
                result = {"ok": True}
            else:
                result = {"error": f"Unknown action: {action}"}
        except KeyError as exc:
            result = {"error": f"Missing parameter: {exc}"}
        except Exception as exc:
            logger.exception("Dispatch error for action '%s'", action)
            result = {"error": str(exc)}

        result["id"] = req_id
        return result


# ── Entry point ────────────────────────────────────────────────────────────────

def main() -> None:
    import argparse

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )

    parser = argparse.ArgumentParser(description="OpenSheet IPC Data Engine")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT,
                        help=f"TCP port to listen on (default: {DEFAULT_PORT})")
    args = parser.parse_args()

    server = IPCServer(port=args.port)
    try:
        server.start()
    except KeyboardInterrupt:
        server.stop()


if __name__ == "__main__":
    main()
