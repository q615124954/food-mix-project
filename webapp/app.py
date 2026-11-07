from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse

from foodmix.errors import AppError
from foodmix.service import Workspace

app = FastAPI(title="电子账单入账", version="1.0.0")
workspace = Workspace()
INDEX = Path(__file__).resolve().parent / "static" / "index.html"


@app.exception_handler(AppError)
async def handle_app_error(_request, exc: AppError):
    return JSONResponse({"detail": exc.message}, status_code=exc.status)


@app.get("/")
def index():
    return FileResponse(INDEX, headers={"Cache-Control": "no-store"})


@app.get("/api/health")
def health():
    return workspace.health()


@app.get("/api/settings")
def get_settings():
    return workspace.get_settings()


@app.post("/api/settings")
def update_settings(body: dict):
    return workspace.update_settings(body.get("input_dir") or "", body.get("material_file") or "")


@app.post("/api/scan")
def scan():
    return workspace.scan()


@app.get("/api/invoices")
def list_invoices():
    return workspace.list_invoices()


@app.get("/api/invoices/{invoice_id}")
def get_invoice(invoice_id: str):
    return workspace.get_invoice(invoice_id)


@app.get("/api/invoices/{invoice_id}/image")
def invoice_image(invoice_id: str):
    return FileResponse(workspace.image_file(invoice_id))


@app.post("/api/invoices/{invoice_id}/recognize")
def recognize(invoice_id: str):
    return workspace.recognize(invoice_id)


@app.post("/api/invoices/{invoice_id}/remap")
def remap(invoice_id: str):
    return workspace.remap(invoice_id)


@app.put("/api/invoices/{invoice_id}")
def save_invoice(invoice_id: str, body: dict):
    return workspace.save(invoice_id, body.get("date"), body.get("lines") or [])


@app.post("/api/invoices/{invoice_id}/confirm")
def confirm_invoice(invoice_id: str, body: dict | None = None):
    # 确认只锁定复核，不写物料明细。
    body = body or {}
    lines = body.get("lines") if "lines" in body else None
    return workspace.confirm(invoice_id, body.get("date"), lines)


@app.post("/api/invoices/{invoice_id}/unconfirm")
def unconfirm_invoice(invoice_id: str):
    return workspace.unconfirm(invoice_id)


@app.post("/api/sync")
def sync():
    return workspace.sync()


@app.get("/api/sync/last")
def last_sync():
    return workspace.last_sync()


@app.get("/api/catalog")
def catalog():
    return workspace.catalog_public()
