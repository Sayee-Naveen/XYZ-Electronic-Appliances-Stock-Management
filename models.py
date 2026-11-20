from pydantic import BaseModel, Field
from typing import Optional, List
from datetime import datetime

class Product(BaseModel):
    sku: str
    product_name: str
    category: str
    brand: str
    model_number: str
    stock_quantity: int = Field(ge=0)
    min_threshold: int = Field(ge=0, default=5)
    cost_price: float = Field(ge=0)
    selling_price: float = Field(ge=0)
    warehouse_location: str = "Main Warehouse"
    energy_rating: str = "N/A"
    warranty: str = "1 Year"
    supplier: str = "Standard Distributor"
    status: Optional[str] = None
    last_updated: Optional[str] = None

class ProductUpdate(BaseModel):
    product_name: Optional[str] = None
    category: Optional[str] = None
    brand: Optional[str] = None
    model_number: Optional[str] = None
    stock_quantity: Optional[int] = Field(ge=0, default=None)
    min_threshold: Optional[int] = Field(ge=0, default=None)
    cost_price: Optional[float] = Field(ge=0, default=None)
    selling_price: Optional[float] = Field(ge=0, default=None)
    warehouse_location: Optional[str] = None
    energy_rating: Optional[str] = None
    warranty: Optional[str] = None
    supplier: Optional[str] = None

class StockAdjust(BaseModel):
    delta: int  # can be positive (e.g. +5) or negative (e.g. -2)
    reason: Optional[str] = "Manual Stock Adjustment"

class LoginRequest(BaseModel):
    email: str
    password: str

class UserResponse(BaseModel):
    email: str
    name: str
    role: str
    token: str

class AuditEntry(BaseModel):
    id: str
    timestamp: str
    user_name: str
    user_role: str
    action: str
    sku: str
    product_name: str
    details: str
