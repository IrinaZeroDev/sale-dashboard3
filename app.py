from __future__ import annotations

import io
import json
import mimetypes
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd


ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
MAX_UPLOAD = 25 * 1024 * 1024
ALIASES = {
    "date": ["дата", "date", "дата продажи", "sale_date"],
    "region": ["регион", "region"],
    "store": ["магазин", "store", "точка", "shop"],
    "product": ["продукт", "product", "товар", "наименование"],
    "sales": ["продажи", "выручка", "sales", "revenue", "сумма"],
    "transaction": ["транзакция", "transaction", "чек", "transaction_id", "order_id"],
}


def normalize(value: object) -> str:
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def detect_columns(columns: list[str]) -> dict[str, str | None]:
    normalized = {normalize(column): column for column in columns}
    return {
        role: next((normalized[name] for name in names if name in normalized), None)
        for role, names in ALIASES.items()
    }


def read_table(payload: bytes, filename: str) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    source = io.BytesIO(payload)
    if suffix == ".csv":
        return pd.read_csv(source, sep=None, engine="python")
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(source)
    raise ValueError("Поддерживаются файлы XLSX, XLSM и CSV")


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "SalesDashboard/1.0"

    def send_json(self, body: dict, status: int = 200) -> None:
        data = json.dumps(body, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        route = urlparse(self.path).path
        if route == "/api/health":
            self.send_json({"status": "ok"})
            return
        relative = "index.html" if route == "/" else route.lstrip("/")
        target = (STATIC / relative).resolve()
        if STATIC.resolve() not in target.parents and target != STATIC.resolve():
            self.send_error(403)
            return
        if not target.is_file():
            self.send_error(404)
            return
        data = target.read_bytes()
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", f"{content_type}; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/upload":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_UPLOAD:
                raise ValueError("Файл пуст или превышает лимит 25 МБ")
            filename = self.headers.get("X-Filename", "sales.xlsx")
            frame = read_table(self.rfile.read(length), filename)
            frame = frame.dropna(how="all")
            if frame.empty:
                raise ValueError("В файле нет строк с данными")
            frame.columns = [str(column).strip() for column in frame.columns]
            sample = frame.head(8000).where(pd.notna(frame.head(8000)), None)
            self.send_json({
                "filename": filename,
                "rowCount": int(len(frame)),
                "columns": frame.columns.tolist(),
                "mapping": detect_columns(frame.columns.tolist()),
                "rows": sample.to_dict(orient="records"),
                "isTruncated": len(frame) > len(sample),
            })
        except Exception as error:
            self.send_json({"error": str(error)}, 400)

    def log_message(self, format: str, *args: object) -> None:
        print(f"[{self.log_date_time_string()}] {format % args}")


if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    address = (host, port)
    print(f"Дашборд запущен: http://{host}:{port}")
    ThreadingHTTPServer(address, DashboardHandler).serve_forever()
