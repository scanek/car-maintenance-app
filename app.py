import os
import json
from pathlib import Path
from typing import List, Optional
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
STATIC_DIR = Path(__file__).parent / "static"

def load_db():
    if not DB_FILE.exists():
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

@app.get("/")
def get_index():
    return FileResponse(STATIC_DIR / "index.html")

@app.get("/api/status")
def get_status():
    db = load_db()
    vehicle = db["vehicle"]
    records = db["maintenance_records"]
    
    current_km = vehicle.get("current_km", 0)
    current_hours = vehicle.get("current_engine_hours", 0)
    
    total_spent = sum(r.get("total_price", 0) for r in records)
    cost_per_km = round(total_spent / current_km, 2) if current_km > 0 else 0
    avg_speed = round(current_km / current_hours, 1) if current_hours > 0 else 0
    
    # Calculate status for each tracked consumable
    # Group by item_name or track canonical items
    tracked_items = [
        {"name": "Моторное масло (0W-20)", "match": "масло", "default_km": 7500, "default_h": 250, "icon": "droplet"},
        {"name": "Масляный фильтр", "match": "фильтр масляный", "default_km": 7500, "default_h": 250, "icon": "disc"},
        {"name": "Воздушный фильтр", "match": "фильтр воздушный", "default_km": 9000, "default_h": 250, "icon": "wind"},
        {"name": "Салонный фильтр", "match": "фильтр салонный", "default_km": 9000, "default_h": 250, "icon": "fan"},
        {"name": "Кольцо сливной пробки", "match": "кольцо", "default_km": 7500, "default_h": 250, "icon": "circle"},
        {"name": "Свечи зажигания", "match": "свечи", "default_km": 30000, "default_h": 0, "icon": "zap"},
        {"name": "Антифриз (ОЖ)", "match": "антифриз", "default_km": 50000, "default_h": 0, "icon": "thermometer"},
    ]
    
    consumables_status = []
    
    for tracker in tracked_items:
        # Find latest replacement
        matching = [r for r in records if tracker["match"] in r["item_name"].lower()]
        matching.sort(key=lambda x: (x["mileage"], x["date"]))
        latest = matching[-1] if matching else None
        
        if latest:
            last_km = latest["mileage"]
            last_h = latest["engine_hours"]
            interval_km = latest.get("interval_km", tracker["default_km"])
            interval_h = latest.get("interval_hours", tracker["default_h"])
            next_km = latest.get("next_km", last_km + interval_km)
            next_h = latest.get("next_hours", last_h + interval_h if interval_h > 0 else 0)
            
            rem_km = next_km - current_km
            rem_h = (next_h - current_hours) if next_h > 0 else None
            
            used_km = current_km - last_km
            wear_ratio = max(0.0, min(1.0, used_km / interval_km)) if interval_km > 0 else 0.0
            wear_percent = round(wear_ratio * 100, 1)
            
            if rem_km <= 0 or (rem_h is not None and rem_h <= 0):
                status_code = "danger"
                status_text = "Требуется замена"
            elif rem_km <= 1500 or (rem_h is not None and rem_h <= 30):
                status_code = "warning"
                status_text = "Скоро замена"
            else:
                status_code = "ok"
                status_text = "В норме"
                
            consumables_status.append({
                "item_name": tracker["name"],
                "icon": tracker["icon"],
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
                "brand": latest.get("brand", ""),
                "article": latest.get("article", ""),
                "to_tag": latest.get("to_tag", "")
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
    return db["maintenance_records"]

@app.post("/api/records")
def add_record(record: MaintenanceRecord):
    db = load_db()
    records = db["maintenance_records"]
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
    
    # Also update current vehicle mileage if higher
    if record.mileage > db["vehicle"].get("current_km", 0):
        db["vehicle"]["current_km"] = record.mileage
    if record.engine_hours > db["vehicle"].get("current_engine_hours", 0):
        db["vehicle"]["current_engine_hours"] = record.engine_hours
        
    save_db(db)
    return {"status": "success", "record": new_rec}

@app.delete("/api/records/{record_id}")
def delete_record(record_id: int):
    db = load_db()
    records = db["maintenance_records"]
    db["maintenance_records"] = [r for r in records if r["id"] != record_id]
    save_db(db)
    return {"status": "deleted", "id": record_id}

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

@app.get("/api/reference")
def get_reference():
    db = load_db()
    return db.get("reference_intervals", [])

# Mount static files
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"🚀 Starting Auto Maintenance Server on http://localhost:{port}")
    uvicorn.run("app:app", host=host, port=port, reload=False)
