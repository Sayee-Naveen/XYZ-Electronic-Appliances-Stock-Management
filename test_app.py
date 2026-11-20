import sys
import os
import io
import openpyxl

# Test local imports
from excel_service import ExcelService, EXCEL_FILE_PATH
from auth import authenticate_user, verify_token
from models import Product, ProductUpdate
from audit_service import AuditService

def run_tests():
    print("[*] Testing ExcelService.load_all()...")
    products = ExcelService.load_all()
    print(f" -> Found {len(products)} products in {EXCEL_FILE_PATH}")
    assert len(products) > 0, "No products loaded from Excel!"
    
    first = products[0]
    print(f" -> First product: SKU={first['sku']}, Name={first['product_name']}, Stock={first['stock_quantity']}")
    
    print("\n[*] Testing KPI calculations...")
    kpis = ExcelService.calculate_kpis()
    print(f" -> Total SKUs: {kpis['total_skus']}")
    print(f" -> Total Units: {kpis['total_units']}")
    print(f" -> Retail Valuation: ₹{kpis['total_retail_valuation']:,.2f}")
    print(f" -> Low Stock Count: {kpis['low_stock_count']}")
    print(f" -> Out of Stock Count: {kpis['out_of_stock_count']}")
    
    print("\n[*] Testing Authentication...")
    user = authenticate_user("admin@xyzappliances.com", "admin123")
    assert user is not None, "Admin authentication failed!"
    print(f" -> Logged in as: {user['name']} ({user['role']}), Token={user['token'][:8]}...")
    
    verified = verify_token(user['token'])
    assert verified is not None, "Token verification failed!"
    print(" -> Token successfully verified!")
    
    print("\n[*] Testing Stock Adjustment on Excel...")
    initial_stock = first["stock_quantity"]
    sku = first["sku"]
    print(f" -> Adjusting stock for {sku} (+2 units)...")
    updated = ExcelService.adjust_stock(sku, 2)
    assert updated["stock_quantity"] == initial_stock + 2, "Stock adjust mismatch!"
    
    # Reload from disk to verify it actually wrote to the Excel file
    reloaded_products = ExcelService.load_all()
    reloaded_prod = next(p for p in reloaded_products if p["sku"] == sku)
    print(f" -> Verified on disk in Excel: new stock = {reloaded_prod['stock_quantity']}")
    assert reloaded_prod["stock_quantity"] == initial_stock + 2, "Disk sync check failed!"
    
    # Revert back
    ExcelService.adjust_stock(sku, -2)
    print(f" -> Reverted stock back to {initial_stock}")
    
    print("\n[*] Testing AuditService...")
    AuditService.log_event("Test Runner", "Admin", "TEST_ACTION", sku, first["product_name"], "Automated test adjustment")
    logs = AuditService.get_recent(5)
    assert len(logs) > 0, "Audit log empty!"
    print(f" -> Recent audit event: [{logs[0]['action']}] {logs[0]['details']}")

    print("\n" + "="*50)
    print(" ALL BACKEND & EXCEL TESTS PASSED SUCCESSFULLY! ")
    print("="*50 + "\n")

if __name__ == "__main__":
    run_tests()
