# ⚡ XYZ Electronic Appliances — Stock Management Portal

> **Enterprise Inventory & Live Excel Stock Synchronization System**  
> Standalone Zero-Dependency Distribution for XYZ Electronic Appliances Pvt. Ltd.

---

## 📌 Executive Summary

This application connects directly with the stock inventory Excel file (`data/XYZ_Stock_Inventory.xlsx`), providing a real-time, interactive, and beautifully designed web application where stock managers, administrators, and warehouse godown staff can edit, monitor, restock, and manage electronic appliances.

Every change made through the web portal is written **directly back to the Excel spreadsheet** with full formula preservation, number formatting (INR ₹), status highlighting, and automatic timestamped backups.

---

## 🚀 1-Click Launch (Zero Dependencies Required)

### Standalone Executable (No Python or Pip Required):
Simply **double-click** either:
- **`XYZ_Stock_Portal.exe`** OR
- **`start_app.bat`**

The application boots immediately, binds to the local port, connects with `XYZ_Stock_Inventory.xlsx`, and opens your default web browser to:
👉 **`http://127.0.0.1:8000`**

*(No manual setup, no command line, and no external installation required!)*

---

## 🔐 Role-Based Access Control (RBAC) & Permissions Matrix

The portal features **1-Click Quick Demo Login** pills right on the login screen so you can switch between roles with a single tap:

| Role | Username | Email | Password | Permissions & Capabilities |
| :--- | :--- | :--- | :--- | :--- |
| **👑 Admin1** | `Admin1` | `admin@xyzappliances.com` | `admin123` | **Full Authority**: Edit all specs, quick adjust stock, add new SKUs, delete appliances, upload/replace Excel sheet, generate Reorder POs, view audit logs. |
| **🏬 Store Manager** | `Store Manager` | `manager@xyzappliances.com` | `manager123` | **Store & Sales Ops**: Edit appliance specs & prices, quick adjust stock, add new SKUs, download Excel, generate Reorder Purchase Orders. *(Deletion and Excel database overwrite restricted)*. |
| **📦 Godown Staff** | `Godown Staff` | `godown@xyzappliances.com` | `godown123` | **Warehouse & Logistics**: Real-time stock lookup, in-line quick stock adjust (`+` / `-` shipments), print warehouse bin barcode labels, view warehouse shelf locations. *(Prices/specs editing, addition, deletion, and DB overwrite restricted)*. |

---

## ✨ Key Enterprise Features

1. **Direct Bidirectional Excel Connection (`openpyxl`)**:
   - Live sync with `data/XYZ_Stock_Inventory.xlsx`.
   - Edits update the spreadsheet on disk in real time.
   - Automatic timestamped backups saved in `data/backups/`.
   - Instant Excel Download & Upload buttons.

2. **In-Line Quick Stock Adjuster**:
   - Single-click `[-]` and `[+]` buttons directly on the inventory table row or batch adjustment (`±`).

3. **Executive Visual Analytics & KPIs (Chart.js)**:
   - Real-time KPIs: Retail Valuation, Cost Valuation, Profit Margin %, Total Stock Units, Low Stock and Out of Stock alerts.
   - Interactive Category Doughnut Chart & Brand Volume Bar Chart.

4. **Smart Reorder & Purchase Order (PO) Advisor**:
   - Auto-detects items below safety threshold and computes suggested replenishment quantities (2x safety stock formula).

5. **Warehouse Barcode & Bin Tag Generator**:
   - Generates scannable barcode label tags with SKU, bay location, model, and retail price.

6. **Live Audit Trail**:
   - Records every stock change, user role, timestamp, and modification diff.

7. **Printable Inventory Audit Sheet**:
   - Physical checklist mode formatted for clipboard audits with tick boxes.

8. **Dark Mode & Light Mode**:
   - Enterprise high-contrast theme toggle.

---

## 📂 Project Architecture

```
D:\XYZ_Appliances_Stock_App\
│
├── XYZ_Stock_Portal.exe              <-- Self-contained standalone executable (Runs with 0 dependencies)
├── start_app.bat                     <-- 1-Click launcher
│
├── data\
│   ├── XYZ_Stock_Inventory.xlsx      <-- Live Excel database
│   ├── audit_log.json                <-- Transaction audit history
│   └── backups\                      <-- Automatic safety backups of Excel
│
├── static\
│   └── app.js                        <-- Client application logic & RBAC UI handling
│
├── templates\
│   └── index.html                    <-- Modern responsive UI (Tailwind CSS, FontAwesome, Chart.js)
│
├── server.py                         <-- FastAPI backend API
├── excel_service.py                  <-- Bidirectional openpyxl read/write engine
├── auth.py                           <-- Role-based authentication & permissions engine
├── models.py                         <-- Pydantic schemas
├── audit_service.py                  <-- Audit logging service
└── README.md                         <-- Project documentation
```
