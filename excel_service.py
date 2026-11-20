import os
import sys
import shutil
import threading
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from typing import List, Dict, Any, Optional

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "data", "XYZ_Stock_Inventory.xlsx")
BACKUP_DIR = os.path.join(BASE_DIR, "data", "backups")

# Thread lock for safe concurrent Excel access
excel_lock = threading.Lock()

HEADERS = [
    "SKU", "Product_Name", "Category", "Brand", "Model_Number",
    "Stock_Quantity", "Min_Threshold", "Cost_Price", "Selling_Price",
    "Warehouse_Location", "Energy_Rating", "Warranty", "Supplier",
    "Status", "Last_Updated"
]

def determine_status(qty: int, min_thresh: int) -> str:
    if qty <= 0:
        return "Out of Stock"
    elif qty <= min_thresh:
        return "Low Stock"
    return "In Stock"

def ensure_backup():
    """Create a backup of the current Excel file before modification."""
    if os.path.exists(EXCEL_FILE_PATH):
        os.makedirs(BACKUP_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(BACKUP_DIR, f"XYZ_Stock_Inventory_{timestamp}.xlsx")
        shutil.copy2(EXCEL_FILE_PATH, backup_path)
        
        # Keep only latest 10 backups to prevent bloat
        backups = sorted([os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR) if f.endswith(".xlsx")])
        if len(backups) > 10:
            for old_b in backups[:-10]:
                try:
                    os.remove(old_b)
                except Exception:
                    pass

class ExcelService:
    @staticmethod
    def get_excel_info() -> Dict[str, Any]:
        """Return information about the current Excel file."""
        if not os.path.exists(EXCEL_FILE_PATH):
            return {
                "file_name": os.path.basename(EXCEL_FILE_PATH),
                "exists": False,
                "file_size_kb": 0,
                "last_modified": "N/A",
                "total_rows": 0
            }
        
        stat = os.stat(EXCEL_FILE_PATH)
        last_mod = datetime.fromtimestamp(stat.st_mtime).strftime("%d-%b-%Y %I:%M %p")
        size_kb = round(stat.st_size / 1024, 1)
        
        products = ExcelService.load_all()
        return {
            "file_name": os.path.basename(EXCEL_FILE_PATH),
            "file_path": os.path.abspath(EXCEL_FILE_PATH),
            "relative_path": os.path.join("data", os.path.basename(EXCEL_FILE_PATH)),
            "exists": True,
            "file_size_kb": size_kb,
            "last_modified": last_mod,
            "total_rows": len(products)
        }

    @staticmethod
    def load_all() -> List[Dict[str, Any]]:
        """Load all products directly from the Excel sheet."""
        with excel_lock:
            if not os.path.exists(EXCEL_FILE_PATH):
                return []
            
            wb = openpyxl.load_workbook(EXCEL_FILE_PATH, data_only=True)
            ws = wb.active
            
            rows = list(ws.iter_rows(values_only=True))
            if not rows or len(rows) < 2:
                return []
            
            header_row = [str(c).strip() if c is not None else "" for c in rows[0]]
            # Map column names to index
            col_map = {name.lower(): idx for idx, name in enumerate(header_row)}
            
            def get_val(row, field, default=None):
                idx = col_map.get(field.lower())
                if idx is not None and idx < len(row):
                    val = row[idx]
                    return val if val is not None else default
                return default

            products = []
            for r in rows[1:]:
                # Skip empty rows
                if not any(r):
                    continue
                
                sku = str(get_val(r, "SKU", "")).strip()
                if not sku:
                    continue
                
                name = str(get_val(r, "Product_Name", "Unnamed Product")).strip()
                cat = str(get_val(r, "Category", "General")).strip()
                brand = str(get_val(r, "Brand", "Generic")).strip()
                model = str(get_val(r, "Model_Number", "")).strip()
                
                try:
                    qty = int(float(get_val(r, "Stock_Quantity", 0)))
                except (ValueError, TypeError):
                    qty = 0
                
                try:
                    thresh = int(float(get_val(r, "Min_Threshold", 5)))
                except (ValueError, TypeError):
                    thresh = 5
                    
                try:
                    cost = float(get_val(r, "Cost_Price", 0.0))
                except (ValueError, TypeError):
                    cost = 0.0
                    
                try:
                    price = float(get_val(r, "Selling_Price", 0.0))
                except (ValueError, TypeError):
                    price = 0.0
                    
                loc = str(get_val(r, "Warehouse_Location", "Bay A-01")).strip()
                energy = str(get_val(r, "Energy_Rating", "N/A")).strip()
                warranty = str(get_val(r, "Warranty", "1 Year")).strip()
                supplier = str(get_val(r, "Supplier", "Standard Distributor")).strip()
                status = determine_status(qty, thresh)
                
                last_updated = get_val(r, "Last_Updated")
                if isinstance(last_updated, datetime):
                    last_updated_str = last_updated.strftime("%Y-%m-%d %H:%M")
                else:
                    last_updated_str = str(last_updated or datetime.now().strftime("%Y-%m-%d %H:%M"))

                products.append({
                    "sku": sku,
                    "product_name": name,
                    "category": cat,
                    "brand": brand,
                    "model_number": model,
                    "stock_quantity": qty,
                    "min_threshold": thresh,
                    "cost_price": cost,
                    "selling_price": price,
                    "warehouse_location": loc,
                    "energy_rating": energy,
                    "warranty": warranty,
                    "supplier": supplier,
                    "status": status,
                    "last_updated": last_updated_str
                })
            
            return products

    @staticmethod
    def save_all(products: List[Dict[str, Any]]) -> bool:
        """Write the updated product list back into the Excel spreadsheet with professional styling."""
        with excel_lock:
            ensure_backup()
            
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Appliance_Stock"

            header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            data_font = Font(name="Calibri", size=10)
            center_align = Alignment(horizontal="center", vertical="center")
            left_align = Alignment(horizontal="left", vertical="center")
            right_align = Alignment(horizontal="right", vertical="center")
            
            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            # Write header
            for col_idx, header in enumerate(HEADERS, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border
            ws.row_dimensions[1].height = 28

            # Write data rows
            for row_idx, p in enumerate(products, 2):
                qty = int(p.get("stock_quantity", 0))
                thresh = int(p.get("min_threshold", 5))
                status = determine_status(qty, thresh)

                row_values = [
                    p.get("sku", ""),
                    p.get("product_name", ""),
                    p.get("category", ""),
                    p.get("brand", ""),
                    p.get("model_number", ""),
                    qty,
                    thresh,
                    float(p.get("cost_price", 0.0)),
                    float(p.get("selling_price", 0.0)),
                    p.get("warehouse_location", ""),
                    p.get("energy_rating", "N/A"),
                    p.get("warranty", "1 Year"),
                    p.get("supplier", ""),
                    status,
                    p.get("last_updated", datetime.now().strftime("%Y-%m-%d %H:%M"))
                ]

                for col_idx, val in enumerate(row_values, 1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=val)
                    cell.font = data_font
                    cell.border = thin_border

                    if col_idx in [1, 5, 10, 11, 12, 14, 15]:
                        cell.alignment = center_align
                    elif col_idx in [6, 7, 8, 9]:
                        cell.alignment = right_align
                        if col_idx in [8, 9]:
                            cell.number_format = '₹#,##0'
                        else:
                            cell.number_format = '#,##0'
                    else:
                        cell.alignment = left_align

                    if col_idx == 14: # Status styling
                        if status == "In Stock":
                            cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                            cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                        elif status == "Low Stock":
                            cell.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
                            cell.font = Font(name="Calibri", size=10, bold=True, color="854D0E")
                        elif status == "Out of Stock":
                            cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                            cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")

                ws.row_dimensions[row_idx].height = 20

            # Auto fit widths
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            os.makedirs(os.path.dirname(EXCEL_FILE_PATH), exist_ok=True)
            wb.save(EXCEL_FILE_PATH)
            return True

    @staticmethod
    def update_product(sku: str, updates: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Update fields of an existing product and write to Excel."""
        products = ExcelService.load_all()
        target_idx = -1
        for idx, p in enumerate(products):
            if p["sku"].upper() == sku.upper():
                target_idx = idx
                break
        
        if target_idx == -1:
            return None
        
        product = products[target_idx]
        for key, val in updates.items():
            if val is not None and key in product:
                product[key] = val
        
        product["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        product["status"] = determine_status(int(product["stock_quantity"]), int(product["min_threshold"]))
        products[target_idx] = product
        
        ExcelService.save_all(products)
        return product

    @staticmethod
    def adjust_stock(sku: str, delta: int) -> Optional[Dict[str, Any]]:
        """Increment or decrement stock quantity directly."""
        products = ExcelService.load_all()
        for p in products:
            if p["sku"].upper() == sku.upper():
                new_qty = max(0, int(p["stock_quantity"]) + delta)
                p["stock_quantity"] = new_qty
                p["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                p["status"] = determine_status(new_qty, int(p["min_threshold"]))
                ExcelService.save_all(products)
                return p
        return None

    @staticmethod
    def add_product(new_prod: Dict[str, Any]) -> Dict[str, Any]:
        """Add a new appliance to inventory and save to Excel."""
        products = ExcelService.load_all()
        
        # Check if SKU already exists
        sku = new_prod.get("sku", "").strip().upper()
        for p in products:
            if p["sku"].upper() == sku:
                raise ValueError(f"An appliance with SKU '{sku}' already exists.")
        
        qty = int(new_prod.get("stock_quantity", 0))
        thresh = int(new_prod.get("min_threshold", 5))
        
        record = {
            "sku": sku,
            "product_name": new_prod.get("product_name", "New Appliance"),
            "category": new_prod.get("category", "General"),
            "brand": new_prod.get("brand", "Generic"),
            "model_number": new_prod.get("model_number", "N/A"),
            "stock_quantity": qty,
            "min_threshold": thresh,
            "cost_price": float(new_prod.get("cost_price", 0.0)),
            "selling_price": float(new_prod.get("selling_price", 0.0)),
            "warehouse_location": new_prod.get("warehouse_location", "Bay A-01"),
            "energy_rating": new_prod.get("energy_rating", "N/A"),
            "warranty": new_prod.get("warranty", "1 Year"),
            "supplier": new_prod.get("supplier", "Standard Distributor"),
            "status": determine_status(qty, thresh),
            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M")
        }
        
        products.insert(0, record)  # Add at top
        ExcelService.save_all(products)
        return record

    @staticmethod
    def delete_product(sku: str) -> bool:
        """Remove an appliance from inventory and update Excel."""
        products = ExcelService.load_all()
        original_len = len(products)
        products = [p for p in products if p["sku"].upper() != sku.upper()]
        if len(products) < original_len:
            ExcelService.save_all(products)
            return True
        return False

    @staticmethod
    def calculate_kpis() -> Dict[str, Any]:
        """Calculate high-level metrics for dashboard analytics."""
        products = ExcelService.load_all()
        
        total_skus = len(products)
        total_units = sum(p["stock_quantity"] for p in products)
        total_cost_val = sum(p["stock_quantity"] * p["cost_price"] for p in products)
        total_retail_val = sum(p["stock_quantity"] * p["selling_price"] for p in products)
        projected_margin = total_retail_val - total_cost_val
        margin_pct = round((projected_margin / total_retail_val * 100), 1) if total_retail_val > 0 else 0
        
        in_stock_count = sum(1 for p in products if p["status"] == "In Stock")
        low_stock_count = sum(1 for p in products if p["status"] == "Low Stock")
        out_of_stock_count = sum(1 for p in products if p["status"] == "Out of Stock")

        # Category Breakdown
        categories: Dict[str, Dict[str, Any]] = {}
        for p in products:
            c = p["category"]
            if c not in categories:
                categories[c] = {"count": 0, "units": 0, "value": 0.0}
            categories[c]["count"] += 1
            categories[c]["units"] += p["stock_quantity"]
            categories[c]["value"] += p["stock_quantity"] * p["selling_price"]

        # Brand Breakdown
        brands: Dict[str, int] = {}
        for p in products:
            b = p["brand"]
            brands[b] = brands.get(b, 0) + p["stock_quantity"]

        return {
            "total_skus": total_skus,
            "total_units": total_units,
            "total_cost_valuation": round(total_cost_val, 2),
            "total_retail_valuation": round(total_retail_val, 2),
            "projected_margin": round(projected_margin, 2),
            "margin_percentage": margin_pct,
            "in_stock_count": in_stock_count,
            "low_stock_count": low_stock_count,
            "out_of_stock_count": out_of_stock_count,
            "category_breakdown": categories,
            "brand_breakdown": brands
        }
