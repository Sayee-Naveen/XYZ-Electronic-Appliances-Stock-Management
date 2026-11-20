import json
import os
import sys
import uuid
from datetime import datetime
from typing import List, Dict, Any

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
AUDIT_FILE = os.path.join(BASE_DIR, "data", "audit_log.json")

class AuditService:
    @staticmethod
    def _load_logs() -> List[Dict[str, Any]]:
        if not os.path.exists(AUDIT_FILE):
            return []
        try:
            with open(AUDIT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []

    @staticmethod
    def _save_logs(logs: List[Dict[str, Any]]):
        os.makedirs(os.path.dirname(AUDIT_FILE), exist_ok=True)
        with open(AUDIT_FILE, "w", encoding="utf-8") as f:
            json.dump(logs[:200], f, indent=2)  # retain last 200 logs

    @staticmethod
    def log_event(user_name: str, user_role: str, action: str, sku: str, product_name: str, details: str):
        logs = AuditService._load_logs()
        entry = {
            "id": str(uuid.uuid4())[:8],
            "timestamp": datetime.now().strftime("%d-%b-%Y %I:%M:%S %p"),
            "user_name": user_name,
            "user_role": user_role,
            "action": action,
            "sku": sku,
            "product_name": product_name,
            "details": details
        }
        logs.insert(0, entry)
        AuditService._save_logs(logs)

    @staticmethod
    def get_recent(limit: int = 50) -> List[Dict[str, Any]]:
        logs = AuditService._load_logs()
        return logs[:limit]
