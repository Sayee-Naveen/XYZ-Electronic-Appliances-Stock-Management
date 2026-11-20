import os
import sys
import io
import shutil
from fastapi import FastAPI, HTTPException, Header, Depends, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from typing import Optional, Dict, Any, List

from models import Product, ProductUpdate, StockAdjust, LoginRequest
from excel_service import ExcelService, EXCEL_FILE_PATH
from audit_service import AuditService
from auth import authenticate_user, verify_token, logout_token, USERS, has_permission

app = FastAPI(
    title="XYZ Electronic Appliances - Stock Management System",
    description="Enterprise Inventory & Stock Management Portal connecting directly with Excel spreadsheets.",
    version="2.4.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Dependency for authentication
def get_current_user(authorization: Optional[str] = Header(None)) -> Dict[str, Any]:
    if not authorization:
        raise HTTPException(status_code=401, detail="Authentication required. Please log in.")
    user = verify_token(authorization)
    if not user:
        raise HTTPException(status_code=401, detail="Session expired or invalid token. Please log in again.")
    return user

# ----------------- AUTH ENDPOINTS -----------------
@app.post("/api/auth/login")
def login(creds: LoginRequest):
    user = authenticate_user(creds.email, creds.password)
    if not user:
        raise HTTPException(status_code=401, detail="Invalid email or password.")
    
    AuditService.log_event(
        user_name=user["name"],
        user_role=user["role"],
        action="USER_LOGIN",
        sku="AUTH",
        product_name="System Access",
        details=f"User {user['name']} ({user['role']}) logged in successfully."
    )
    return user

@app.get("/api/auth/me")
def get_me(user: Dict[str, Any] = Depends(get_current_user)):
    return user

@app.post("/api/auth/logout")
def logout(authorization: Optional[str] = Header(None)):
    if authorization:
        logout_token(authorization)
    return {"message": "Logged out successfully"}

# ----------------- EXCEL & STOCK ENDPOINTS -----------------
@app.get("/api/excel-info")
def get_excel_info():
    """Retrieve metadata about the active Excel spreadsheet."""
    return ExcelService.get_excel_info()

@app.get("/api/stock")
def get_stock(
    search: Optional[str] = None,
    category: Optional[str] = None,
    brand: Optional[str] = None,
    status: Optional[str] = None
):
    """Retrieve all stock items along with computed KPIs."""
    products = ExcelService.load_all()
    kpis = ExcelService.calculate_kpis()
    
    # Filter in-memory
    filtered = products
    if search:
        s = search.lower().strip()
        filtered = [
            p for p in filtered
            if s in p["sku"].lower() or s in p["product_name"].lower() or s in p["model_number"].lower() or s in p["brand"].lower()
        ]
    if category and category != "All":
        filtered = [p for p in filtered if p["category"].lower() == category.lower()]
    if brand and brand != "All":
        filtered = [p for p in filtered if p["brand"].lower() == brand.lower()]
    if status and status != "All":
        filtered = [p for p in filtered if p["status"].lower() == status.lower()]
        
    return {
        "products": filtered,
        "total_count": len(filtered),
        "kpis": kpis,
        "excel_info": ExcelService.get_excel_info()
    }

@app.get("/api/stock/{sku}")
def get_single_product(sku: str):
    products = ExcelService.load_all()
    for p in products:
        if p["sku"].upper() == sku.upper():
            return p
    raise HTTPException(status_code=404, detail=f"Appliance SKU '{sku}' not found.")

@app.put("/api/stock/{sku}")
def update_product_details(sku: str, payload: ProductUpdate, user: Dict[str, Any] = Depends(get_current_user)):
    """Update fields of an appliance and commit directly to the Excel file."""
    if not has_permission(user, "stock_edit"):
        raise HTTPException(
            status_code=403, 
            detail="Permission Denied: Only Admin1 and Store Manager can edit appliance specifications."
        )

    update_data = payload.dict(exclude_unset=True)
    if not update_data:
        raise HTTPException(status_code=400, detail="No updates provided.")
    
    updated = ExcelService.update_product(sku, update_data)
    if not updated:
        raise HTTPException(status_code=404, detail=f"Appliance SKU '{sku}' not found.")
    
    AuditService.log_event(
        user_name=user["name"],
        user_role=user["role"],
        action="UPDATE_DETAILS",
        sku=updated["sku"],
        product_name=updated["product_name"],
        details=f"Updated fields: {', '.join(update_data.keys())} in Excel sheet."
    )
    
    return {"message": "Appliance updated and synced with Excel successfully.", "product": updated}

@app.post("/api/stock/{sku}/adjust")
def adjust_stock_quantity(sku: str, payload: StockAdjust, user: Dict[str, Any] = Depends(get_current_user)):
    """Increment/decrement appliance stock quantity and write directly to Excel."""
    if not has_permission(user, "stock_adjust"):
        raise HTTPException(status_code=403, detail="Permission Denied: Stock adjustment not permitted.")

    delta = payload.delta
    if delta == 0:
        raise HTTPException(status_code=400, detail="Stock adjustment delta cannot be 0.")
    
    updated = ExcelService.adjust_stock(sku, delta)
    if not updated:
        raise HTTPException(status_code=404, detail=f"Appliance SKU '{sku}' not found.")
    
    sign = f"+{delta}" if delta > 0 else str(delta)
    details_str = f"Stock adjusted by {sign} (Reason: {payload.reason}). New stock: {updated['stock_quantity']} units."
    
    AuditService.log_event(
        user_name=user["name"],
        user_role=user["role"],
        action="STOCK_ADJUSTMENT",
        sku=updated["sku"],
        product_name=updated["product_name"],
        details=details_str
    )
    
    return {
        "message": f"Stock adjusted by {sign} and saved to Excel.",
        "product": updated
    }

@app.post("/api/stock")
def add_new_product(payload: Product, user: Dict[str, Any] = Depends(get_current_user)):
    """Add a new appliance model to inventory and write directly to Excel."""
    if not has_permission(user, "stock_add"):
        raise HTTPException(
            status_code=403, 
            detail="Permission Denied: Godown Staff cannot add new SKUs. Please request Store Manager or Admin1."
        )

    try:
        new_prod = ExcelService.add_product(payload.dict())
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    
    AuditService.log_event(
        user_name=user["name"],
        user_role=user["role"],
        action="ADD_APPLIANCE",
        sku=new_prod["sku"],
        product_name=new_prod["product_name"],
        details=f"New appliance created with {new_prod['stock_quantity']} units in Excel sheet."
    )
    
    return {"message": "New appliance added and recorded in Excel.", "product": new_prod}

@app.delete("/api/stock/{sku}")
def delete_product(sku: str, user: Dict[str, Any] = Depends(get_current_user)):
    """Remove an appliance from inventory and rewrite the Excel file (Admin only)."""
    if not has_permission(user, "stock_delete"):
        raise HTTPException(
            status_code=403, 
            detail="Permission Denied: Only Admin1 can delete appliances from the inventory database."
        )

    products = ExcelService.load_all()
    target_name = "Appliance"
    for p in products:
        if p["sku"].upper() == sku.upper():
            target_name = p["product_name"]
            break
            
    success = ExcelService.delete_product(sku)
    if not success:
        raise HTTPException(status_code=404, detail=f"Appliance SKU '{sku}' not found.")
    
    AuditService.log_event(
        user_name=user["name"],
        user_role=user["role"],
        action="DELETE_APPLIANCE",
        sku=sku,
        product_name=target_name,
        details=f"Appliance SKU {sku} deleted from inventory Excel."
    )
    
    return {"message": f"Appliance SKU '{sku}' deleted and removed from Excel."}

# ----------------- EXCEL DOWNLOAD & UPLOAD -----------------
@app.get("/api/download-excel")
def download_excel():
    """Download the live active Excel file."""
    if not os.path.exists(EXCEL_FILE_PATH):
        raise HTTPException(status_code=404, detail="Excel file does not exist.")
    return FileResponse(
        path=EXCEL_FILE_PATH,
        filename=os.path.basename(EXCEL_FILE_PATH),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...), user: Dict[str, Any] = Depends(get_current_user)):
    """Upload a replacement or updated Excel spreadsheet (Admin only)."""
    if not has_permission(user, "excel_upload"):
        raise HTTPException(
            status_code=403, 
            detail="Permission Denied: Only Admin1 has authority to overwrite or import the master Excel sheet."
        )

    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx) are supported.")
    
    contents = await file.read()
    
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(contents), data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows or len(rows) < 2:
            raise ValueError("The uploaded Excel sheet contains no data rows.")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid Excel file format: {str(e)}")
    
    from excel_service import ensure_backup
    ensure_backup()
    
    with open(EXCEL_FILE_PATH, "wb") as f:
        f.write(contents)
        
    count = len(ExcelService.load_all())
    
    AuditService.log_event(
        user_name=user["name"],
        user_role=user["role"],
        action="EXCEL_IMPORT",
        sku="EXCEL",
        product_name=file.filename,
        details=f"Imported custom Excel file with {count} items."
    )
    
    return {
        "message": f"Successfully loaded Excel sheet '{file.filename}'. Found {count} appliances.",
        "item_count": count
    }

# ----------------- REORDER / PURCHASE ORDER ASSISTANT -----------------
@app.get("/api/reorder-recommendations")
def get_reorder_recommendations(authorization: Optional[str] = Header(None)):
    """Smart Reorder Assistant: calculates replenishment needs for low stock items."""
    # Anyone authenticated or browsing the portal can view reorder needs
    products = ExcelService.load_all()
    recommendations = []
    total_estimated_cost = 0.0
    
    for p in products:
        qty = p["stock_quantity"]
        thresh = p["min_threshold"]
        if qty <= thresh:
            deficit = max(1, (thresh * 2) - qty)
            cost = deficit * p["cost_price"]
            total_estimated_cost += cost
            urgency = "CRITICAL" if qty == 0 else "HIGH"
            
            recommendations.append({
                "sku": p["sku"],
                "product_name": p["product_name"],
                "category": p["category"],
                "brand": p["brand"],
                "current_stock": qty,
                "min_threshold": thresh,
                "suggested_order_qty": deficit,
                "supplier": p["supplier"],
                "unit_cost": p["cost_price"],
                "estimated_po_cost": round(cost, 2),
                "urgency": urgency
            })
            
    recommendations.sort(key=lambda x: (0 if x["urgency"] == "CRITICAL" else 1, x["current_stock"]))
    
    return {
        "items": recommendations,
        "total_items_to_reorder": len(recommendations),
        "total_estimated_po_cost": round(total_estimated_cost, 2)
    }

# ----------------- AUDIT LOGS -----------------
@app.get("/api/audit-log")
def get_audit_logs():
    return AuditService.get_recent(60)

# ----------------- ROOT WEB PAGE -----------------
@app.get("/", response_class=HTMLResponse)
def index_page():
    index_file = os.path.join(TEMPLATES_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>XYZ Electronic Appliances Portal is initializing...</h1>"

if __name__ == "__main__":
    import uvicorn
    print("\n" + "="*70)
    print("  XYZ Electronic Appliances - Stock Management Portal")
    print("  Excel Database : " + EXCEL_FILE_PATH)
    print("  Local Portal   : http://127.0.0.1:8000")
    print("="*70 + "\n")
    uvicorn.run(app, host="127.0.0.1", port=8000)
