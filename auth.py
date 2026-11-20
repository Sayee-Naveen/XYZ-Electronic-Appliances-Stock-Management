from typing import Optional, Dict, Any, List
import secrets

# Preconfigured role-based users for the XYZ Electronic Appliances platform
USERS = {
    "admin1@xyzappliances.com": {
        "name": "Admin1",
        "role": "Administrator",
        "password": "admin123",
        "avatar": "👑",
        "department": "Executive Management",
        "permissions": [
            "all", "stock_view", "stock_edit", "stock_adjust",
            "stock_add", "stock_delete", "excel_upload",
            "excel_download", "reorder_po", "audit_view"
        ]
    },
    "admin@xyzappliances.com": {
        "name": "Admin1",
        "role": "Administrator",
        "password": "admin123",
        "avatar": "👑",
        "department": "Executive Management",
        "permissions": [
            "all", "stock_view", "stock_edit", "stock_adjust",
            "stock_add", "stock_delete", "excel_upload",
            "excel_download", "reorder_po", "audit_view"
        ]
    },
    "manager@xyzappliances.com": {
        "name": "Store Manager",
        "role": "Store Manager",
        "password": "manager123",
        "avatar": "🏬",
        "department": "Retail & Store Operations",
        "permissions": [
            "stock_view", "stock_edit", "stock_adjust",
            "stock_add", "excel_download", "reorder_po", "audit_view"
        ]
    },
    "godown@xyzappliances.com": {
        "name": "Godown Staff",
        "role": "Godown Staff",
        "password": "godown123",
        "avatar": "📦",
        "department": "Warehouse & Godown Logistics",
        "permissions": [
            "stock_view", "stock_adjust", "barcode_view"
        ]
    },
    "staff@xyzappliances.com": {
        "name": "Godown Staff",
        "role": "Godown Staff",
        "password": "godown123",
        "avatar": "📦",
        "department": "Warehouse & Godown Logistics",
        "permissions": [
            "stock_view", "stock_adjust", "barcode_view"
        ]
    }
}

ACTIVE_SESSIONS: Dict[str, Dict[str, Any]] = {}

def authenticate_user(email: str, password: str) -> Optional[Dict[str, Any]]:
    clean_email = email.strip().lower()
    user = USERS.get(clean_email)
    if user and user["password"] == password.strip():
        token = secrets.token_hex(24)
        session_data = {
            "token": token,
            "email": clean_email,
            "name": user["name"],
            "role": user["role"],
            "avatar": user["avatar"],
            "department": user["department"],
            "permissions": user["permissions"]
        }
        ACTIVE_SESSIONS[token] = session_data
        return session_data
    return None

def verify_token(token: Optional[str]) -> Optional[Dict[str, Any]]:
    if not token:
        return None
    if token.startswith("Bearer "):
        token = token[7:]
    return ACTIVE_SESSIONS.get(token)

def logout_token(token: str):
    if token.startswith("Bearer "):
        token = token[7:]
    if token in ACTIVE_SESSIONS:
        del ACTIVE_SESSIONS[token]

def has_permission(user: Dict[str, Any], permission: str) -> bool:
    perms = user.get("permissions", [])
    return "all" in perms or permission in perms
