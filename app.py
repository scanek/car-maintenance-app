import os
import json
import shutil
import tempfile
from pathlib import Path
from typing import List, Optional, Dict, Any
from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = FastAPI(title="Changan Auto Maintenance Dashboard")

DATA_DIR = Path(__file__).parent / "data"
DB_FILE = DATA_DIR / "db.json"
EXAMPLE_DB_FILE = DATA_DIR / "db.example.json"
STATIC_DIR = Path(__file__).parent / "static"

def load_db():
    if not DB_FILE.exists():
        if EXAMPLE_DB_FILE.exists():
            shutil.copy2(EXAMPLE_DB_FILE, DB_FILE)
        else:
            raise HTTPException(status_code=500, detail="Database file not found")
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

class VehicleUpdate(BaseModel):
    current_km: int
    current_engine_hours: int
    brand: Optional[str] = "Changan"
    model: Optional[str] = "CS55 Plus"

class MaintenanceRecord(BaseModel):
    to_tag: str
    date: str
    engine_hours: int
    mileage: int
    category: str
    item_name: str
    brand: Optional[str] = ""
    article: Optional[str] = ""
    quantity: float = 1.0
    unit: Optional[str] = "шт"
    price_per_unit: float
    total_price: Optional[float] = None
    interval_km: int = 7500
    interval_hours: int = 250
    note: Optional[str] = "Плановая замена"
    store: Optional[str] = "Ozon"
    url: Optional[str] = ""

class TrackerSetting(BaseModel):
    id: Optional[str] = None
    name: str
    category: str
    match: str
    interval_km: int
    interval_hours: int = 0
    warn_km: int = 1500
    warn_hours: int = 30
    spec: Optional[str] = ""
    article: Optional[str] = ""
    brand: Optional[str] = ""
    icon: Optional[str] = "wrench"
    enabled: bool = True

class PartItem(BaseModel):
    id: Optional[int] = None
    category: str
    item_name: str
    brand: Optional[str] = ""
    article: Optional[str] = ""
    quantity: float = 1.0
    unit: Optional[str] = "шт"
    price_per_unit: float
    total_price: Optional[float] = None
    interval_km: int = 7500
    interval_hours: int = 250
    note: Optional[str] = "Плановая замена"
    store: Optional[str] = "Ozon"
    url: Optional[str] = ""

class TOGroupPayload(BaseModel):
    original_to_tag: Optional[str] = None
    to_tag: str
    date: str
    mileage: int
    engine_hours: int
    parts: List[PartItem]

@app.get("/")
def get_index():
    return FileResponse(STATIC_DIR / "index.html")

@app.get("/api/status")
def get_status():
    db = load_db()
    vehicle = db["vehicle"]
    records = db.get("maintenance_records", [])
    trackers = db.get("trackers", [])
    
    current_km = vehicle.get("current_km", 0)
    current_hours = vehicle.get("current_engine_hours", 0)
    
    total_spent = sum(r.get("total_price", 0) for r in records)
    cost_per_km = round(total_spent / current_km, 2) if current_km > 0 else 0
    avg_speed = round(current_km / current_hours, 1) if current_hours > 0 else 0
    
    consumables_status = []
    
    for tracker in trackers:
        if not tracker.get("enabled", True):
            continue
            
        keyword = tracker.get("match", "").lower()
        matching = [r for r in records if keyword in r["item_name"].lower()]
        matching.sort(key=lambda x: (x["mileage"], x["date"]))
        latest = matching[-1] if matching else None
        
        interval_km = tracker.get("interval_km", 7500)
        interval_h = tracker.get("interval_hours", 250)
        warn_km = tracker.get("warn_km", 1500)
        warn_h = tracker.get("warn_hours", 30)
        
        if latest:
            last_km = latest["mileage"]
            last_h = latest["engine_hours"]
            # Use record's interval if present or tracker's default
            rec_interval_km = latest.get("interval_km") or interval_km
            rec_interval_h = latest.get("interval_hours") or interval_h
            
            next_km = latest.get("next_km", last_km + rec_interval_km)
            next_h = latest.get("next_hours", last_h + rec_interval_h if rec_interval_h > 0 else 0)
            
            rem_km = next_km - current_km
            rem_h = (next_h - current_hours) if next_h > 0 else None
            
            used_km = current_km - last_km
            wear_ratio = max(0.0, min(1.0, used_km / rec_interval_km)) if rec_interval_km > 0 else 0.0
            wear_percent = round(wear_ratio * 100, 1)
            
            if rem_km <= 0 or (rem_h is not None and rem_h <= 0):
                status_code = "danger"
                status_text = "Требуется замена"
            elif rem_km <= warn_km or (rem_h is not None and rem_h <= warn_h):
                status_code = "warning"
                status_text = "Скоро замена"
            else:
                status_code = "ok"
                status_text = "В норме"
                
            consumables_status.append({
                "tracker_id": tracker.get("id"),
                "item_name": tracker.get("name"),
                "icon": tracker.get("icon", "wrench"),
                "category": tracker.get("category", "Прочее"),
                "last_date": latest["date"],
                "last_km": last_km,
                "last_hours": last_h,
                "next_km": next_km,
                "next_hours": next_h,
                "rem_km": rem_km,
                "rem_hours": rem_h,
                "wear_percent": wear_percent,
                "status_code": status_code,
                "status_text": status_text,
                "brand": latest.get("brand") or tracker.get("brand", ""),
                "article": latest.get("article") or tracker.get("article", ""),
                "spec": tracker.get("spec", ""),
                "to_tag": latest.get("to_tag", "")
            })
        else:
            # Not yet serviced / No record found
            consumables_status.append({
                "tracker_id": tracker.get("id"),
                "item_name": tracker.get("name"),
                "icon": tracker.get("icon", "wrench"),
                "category": tracker.get("category", "Прочее"),
                "last_date": "-",
                "last_km": 0,
                "last_hours": 0,
                "next_km": interval_km,
                "next_hours": interval_h,
                "rem_km": interval_km - current_km,
                "rem_hours": (interval_h - current_hours) if interval_h > 0 else None,
                "wear_percent": 0.0,
                "status_code": "warning" if current_km >= interval_km else "ok",
                "status_text": "Нет записей ТО",
                "brand": tracker.get("brand", ""),
                "article": tracker.get("article", ""),
                "spec": tracker.get("spec", ""),
                "to_tag": "-"
            })
    
    attention_count = sum(1 for c in consumables_status if c["status_code"] in ["danger", "warning"])
    
    return {
        "vehicle": vehicle,
        "kpi": {
            "current_km": current_km,
            "current_hours": current_hours,
            "total_spent": total_spent,
            "cost_per_km": cost_per_km,
            "avg_speed": avg_speed,
            "total_records": len(records),
            "attention_count": attention_count
        },
        "consumables": consumables_status
    }

@app.get("/api/records")
def get_records():
    db = load_db()
    return db.get("maintenance_records", [])

@app.post("/api/records")
def add_record(record: MaintenanceRecord):
    db = load_db()
    records = db.setdefault("maintenance_records", [])
    new_id = (max([r["id"] for r in records]) + 1) if records else 1
    
    total_price = record.total_price if record.total_price is not None else (record.price_per_unit * record.quantity)
    next_km = record.mileage + record.interval_km
    next_hours = (record.engine_hours + record.interval_hours) if record.interval_hours > 0 else 0
    
    new_rec = {
        "id": new_id,
        "to_tag": record.to_tag,
        "date": record.date,
        "engine_hours": record.engine_hours,
        "mileage": record.mileage,
        "category": record.category,
        "item_name": record.item_name,
        "brand": record.brand or "",
        "article": record.article or "",
        "quantity": record.quantity,
        "unit": record.unit or "шт",
        "price_per_unit": record.price_per_unit,
        "total_price": total_price,
        "interval_km": record.interval_km,
        "interval_hours": record.interval_hours,
        "next_km": next_km,
        "next_hours": next_hours,
        "note": record.note or "",
        "store": record.store or "",
        "url": record.url or ""
    }
    
    records.append(new_rec)
    
    if record.mileage > db["vehicle"].get("current_km", 0):
        db["vehicle"]["current_km"] = record.mileage
    if record.engine_hours > db["vehicle"].get("current_engine_hours", 0):
        db["vehicle"]["current_engine_hours"] = record.engine_hours
        
    save_db(db)
    return {"status": "success", "record": new_rec}

@app.put("/api/records/{record_id}")
def update_record(record_id: int, record: MaintenanceRecord):
    db = load_db()
    records = db.get("maintenance_records", [])
    idx = next((i for i, r in enumerate(records) if r["id"] == record_id), None)
    if idx is None:
        raise HTTPException(status_code=404, detail="Запись не найдена")
    
    total_price = record.total_price if record.total_price is not None else (record.price_per_unit * record.quantity)
    next_km = record.mileage + record.interval_km
    next_hours = (record.engine_hours + record.interval_hours) if record.interval_hours > 0 else 0
    
    updated_rec = {
        "id": record_id,
        "to_tag": record.to_tag,
        "date": record.date,
        "engine_hours": record.engine_hours,
        "mileage": record.mileage,
        "category": record.category,
        "item_name": record.item_name,
        "brand": record.brand or "",
        "article": record.article or "",
        "quantity": record.quantity,
        "unit": record.unit or "шт",
        "price_per_unit": record.price_per_unit,
        "total_price": total_price,
        "interval_km": record.interval_km,
        "interval_hours": record.interval_hours,
        "next_km": next_km,
        "next_hours": next_hours,
        "note": record.note or "",
        "store": record.store or "",
        "url": record.url or ""
    }
    
    records[idx] = updated_rec
    save_db(db)
    return {"status": "success", "record": updated_rec}

@app.delete("/api/records/{record_id}")
def delete_record(record_id: int):
    db = load_db()
    records = db.get("maintenance_records", [])
    db["maintenance_records"] = [r for r in records if r["id"] != record_id]
    save_db(db)
    return {"status": "deleted", "id": record_id}

# --- TO Group Endpoints (Batch TO Event Management) ---
@app.get("/api/to-groups")
def get_to_groups():
    db = load_db()
    records = db.get("maintenance_records", [])
    
    groups: Dict[str, Any] = {}
    for r in records:
        tag = r.get("to_tag") or "Без метки"
        if tag not in groups:
            groups[tag] = {
                "to_tag": tag,
                "date": r["date"],
                "mileage": r["mileage"],
                "engine_hours": r["engine_hours"],
                "total_cost": 0,
                "parts_count": 0,
                "parts": []
            }
        groups[tag]["total_cost"] += r.get("total_price", 0)
        groups[tag]["parts_count"] += 1
        groups[tag]["parts"].append(r)
        
    return list(groups.values())

@app.post("/api/to-groups")
def save_to_group(payload: TOGroupPayload):
    db = load_db()
    records = db.setdefault("maintenance_records", [])
    
    # If editing existing group, remove old parts
    if payload.original_to_tag:
        db["maintenance_records"] = [r for r in records if r.get("to_tag") != payload.original_to_tag]
        records = db["maintenance_records"]
        
    current_max_id = max([r["id"] for r in records]) if records else 0
    
    for part in payload.parts:
        current_max_id += 1
        total_p = part.total_price if part.total_price is not None else (part.price_per_unit * part.quantity)
        next_k = payload.mileage + part.interval_km
        next_h = (payload.engine_hours + part.interval_hours) if part.interval_hours > 0 else 0
        
        new_item = {
            "id": current_max_id,
            "to_tag": payload.to_tag,
            "date": payload.date,
            "engine_hours": payload.engine_hours,
            "mileage": payload.mileage,
            "category": part.category,
            "item_name": part.item_name,
            "brand": part.brand or "",
            "article": part.article or "",
            "quantity": part.quantity,
            "unit": part.unit or "шт",
            "price_per_unit": part.price_per_unit,
            "total_price": total_p,
            "interval_km": part.interval_km,
            "interval_hours": part.interval_hours,
            "next_km": next_k,
            "next_hours": next_h,
            "note": part.note or "",
            "store": part.store or "",
            "url": part.url or ""
        }
        records.append(new_item)
        
    if payload.mileage > db["vehicle"].get("current_km", 0):
        db["vehicle"]["current_km"] = payload.mileage
    if payload.engine_hours > db["vehicle"].get("current_engine_hours", 0):
        db["vehicle"]["current_engine_hours"] = payload.engine_hours
        
    save_db(db)
    return {"status": "success", "to_tag": payload.to_tag, "parts_added": len(payload.parts)}

@app.delete("/api/to-groups/{to_tag}")
def delete_to_group(to_tag: str):
    db = load_db()
    records = db.get("maintenance_records", [])
    db["maintenance_records"] = [r for r in records if r.get("to_tag") != to_tag]
    save_db(db)
    return {"status": "deleted", "to_tag": to_tag}

# --- Settings & Trackers Endpoints ---
@app.get("/api/settings")
def get_settings():
    db = load_db()
    return {
        "vehicle": db.get("vehicle", {}),
        "trackers": db.get("trackers", [])
    }

@app.post("/api/settings/tracker")
def save_tracker(tracker: TrackerSetting):
    db = load_db()
    trackers = db.setdefault("trackers", [])
    
    t_id = tracker.id or tracker.name.lower().replace(" ", "_").replace("(", "").replace(")", "")
    
    idx = next((i for i, t in enumerate(trackers) if t.get("id") == t_id), None)
    
    tracker_dict = {
        "id": t_id,
        "name": tracker.name,
        "category": tracker.category,
        "match": tracker.match,
        "interval_km": tracker.interval_km,
        "interval_hours": tracker.interval_hours,
        "warn_km": tracker.warn_km,
        "warn_hours": tracker.warn_hours,
        "spec": tracker.spec or "",
        "article": tracker.article or "",
        "brand": tracker.brand or "",
        "icon": tracker.icon or "wrench",
        "enabled": tracker.enabled
    }
    
    if idx is not None:
        trackers[idx] = tracker_dict
    else:
        trackers.append(tracker_dict)
        
    save_db(db)
    return {"status": "success", "tracker": tracker_dict}

@app.delete("/api/settings/tracker/{tracker_id}")
def delete_tracker(tracker_id: str):
    db = load_db()
    trackers = db.get("trackers", [])
    db["trackers"] = [t for t in trackers if t.get("id") != tracker_id]
    save_db(db)
    return {"status": "deleted", "id": tracker_id}

@app.post("/api/vehicle")
def update_vehicle(v: VehicleUpdate):
    db = load_db()
    db["vehicle"]["current_km"] = v.current_km
    db["vehicle"]["current_engine_hours"] = v.current_engine_hours
    if v.brand:
        db["vehicle"]["brand"] = v.brand
    if v.model:
        db["vehicle"]["model"] = v.model
    save_db(db)
    return {"status": "success", "vehicle": db["vehicle"]}

@app.get("/api/export-excel")
def export_excel():
    db = load_db()
    vehicle = db.get("vehicle", {})
    records = db.get("maintenance_records", [])
    trackers = db.get("trackers", [])
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    NAVY_HEADER = "1E293B"
    WHITE = "FFFFFF"
    BORDER_COLOR = "CBD5E1"
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    thick_bottom = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='medium', color="2563EB")
    )
    header_font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    
    # Sheet 1: Dashboard
    ws_dash = wb.create_sheet(title="Дашборд & Статус")
    ws_dash.views.sheetView[0].showGridLines = True
    ws_dash["B2"] = "УЧЕТ И СТАТУС ОБСЛУЖИВАНИЯ АВТОМОБИЛЯ"
    ws_dash["B2"].font = Font(name="Segoe UI", size=15, bold=True, color="0F172A")
    ws_dash["B3"] = f"Автомобиль: {vehicle.get('brand', 'Changan')} {vehicle.get('model', 'CS55 Plus')}"
    ws_dash["B3"].font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
    
    ws_dash["B5"] = "Текущий пробег (км):"
    ws_dash["B5"].font = Font(name="Segoe UI", size=11, bold=True)
    ws_dash["C5"] = vehicle.get("current_km", 0)
    ws_dash["C5"].number_format = '#,##0 "км"'
    ws_dash["C5"].font = Font(name="Segoe UI", size=12, bold=True)
    ws_dash["C5"].fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    ws_dash["C5"].border = thin_border
    
    ws_dash["D5"] = "Текущие моточасы (м/ч):"
    ws_dash["D5"].font = Font(name="Segoe UI", size=11, bold=True)
    ws_dash["E5"] = vehicle.get("current_engine_hours", 0)
    ws_dash["E5"].number_format = '#,##0 "м/ч"'
    ws_dash["E5"].font = Font(name="Segoe UI", size=12, bold=True)
    ws_dash["E5"].fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    ws_dash["E5"].border = thin_border
    
    # Sheet 2: Log
    ws_log = wb.create_sheet(title="Журнал обслуживания")
    ws_log.views.sheetView[0].showGridLines = True
    log_headers = [
        "ТО", "№", "Дата", "Моточасы (м/ч)", "Пробег (км)", "Категория", "Наименование расходника",
        "Марка / Модель", "Артикул", "Кол-во", "Ед.", "Цена за ед.", "Сумма (₽)",
        "Интервал (км)", "Интервал (м/ч)", "След. замена (км)", "След. замена (м/ч)", "Примечание", "Где куплено", "Ссылка"
    ]
    for c_idx, title in enumerate(log_headers, start=1):
        cell = ws_log.cell(row=2, column=c_idx, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thick_bottom
        
    for r_idx, r in enumerate(records, start=3):
        ws_log.cell(row=r_idx, column=1, value=r.get("to_tag", ""))
        ws_log.cell(row=r_idx, column=2, value=r_idx-2)
        ws_log.cell(row=r_idx, column=3, value=r.get("date", ""))
        ws_log.cell(row=r_idx, column=4, value=r.get("engine_hours", 0))
        ws_log.cell(row=r_idx, column=5, value=r.get("mileage", 0))
        ws_log.cell(row=r_idx, column=6, value=r.get("category", ""))
        ws_log.cell(row=r_idx, column=7, value=r.get("item_name", ""))
        ws_log.cell(row=r_idx, column=8, value=r.get("brand", ""))
        ws_log.cell(row=r_idx, column=9, value=r.get("article", ""))
        ws_log.cell(row=r_idx, column=10, value=r.get("quantity", 1))
        ws_log.cell(row=r_idx, column=11, value=r.get("unit", "шт"))
        ws_log.cell(row=r_idx, column=12, value=r.get("price_per_unit", 0))
        ws_log.cell(row=r_idx, column=13, value=r.get("total_price", 0))
        ws_log.cell(row=r_idx, column=14, value=r.get("interval_km", 0))
        ws_log.cell(row=r_idx, column=15, value=r.get("interval_hours", 0))
        ws_log.cell(row=r_idx, column=16, value=r.get("next_km", 0))
        ws_log.cell(row=r_idx, column=17, value=r.get("next_hours", 0))
        ws_log.cell(row=r_idx, column=18, value=r.get("note", ""))
        ws_log.cell(row=r_idx, column=19, value=r.get("store", ""))
        ws_log.cell(row=r_idx, column=20, value=r.get("url", ""))
        
        for c in range(1, 21):
            cell = ws_log.cell(row=r_idx, column=c)
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 10, 11, 14, 15, 16, 17, 19]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
    # Sheet 3: Settings / Reference
    ws_set = wb.create_sheet(title="Настройки регламентов")
    ws_set.views.sheetView[0].showGridLines = True
    set_headers = ["ID", "Расходник", "Категория", "Интервал (км)", "Интервал (м/ч)", "Спецификация", "Бренд", "Артикул"]
    for c_idx, title in enumerate(set_headers, start=1):
        cell = ws_set.cell(row=2, column=c_idx, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thick_bottom
        
    for r_idx, t in enumerate(trackers, start=3):
        ws_set.cell(row=r_idx, column=1, value=t.get("id", ""))
        ws_set.cell(row=r_idx, column=2, value=t.get("name", ""))
        ws_set.cell(row=r_idx, column=3, value=t.get("category", ""))
        ws_set.cell(row=r_idx, column=4, value=t.get("interval_km", 0))
        ws_set.cell(row=r_idx, column=5, value=t.get("interval_hours", 0))
        ws_set.cell(row=r_idx, column=6, value=t.get("spec", ""))
        ws_set.cell(row=r_idx, column=7, value=t.get("brand", ""))
        ws_set.cell(row=r_idx, column=8, value=t.get("article", ""))
        for c in range(1, 9):
            ws_set.cell(row=r_idx, column=c).border = thin_border
    
    for ws in [ws_dash, ws_log, ws_set]:
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.close()
    wb.save(temp_file.name)
    
    return FileResponse(
        temp_file.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Учет_обслуживания_авто.xlsx"
    )

@app.get("/api/export-json")
def export_json():
    db = load_db()
    return JSONResponse(
        content=db,
        headers={"Content-Disposition": "attachment; filename=maintenance_backup.json"}
    )

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"🚀 Starting Auto Maintenance Server on http://localhost:{port}")
    uvicorn.run("app:app", host=host, port=port, reload=False)
