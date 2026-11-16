"""
security.py — Compatibility Layer
Delegates to app_security.py to prevent any third-party namespace collisions.
"""
from app_security import (
    get_odoo_credentials,
    get_database_credentials,
    decrypt_pin_secret,
    fetch_dailysales_passwords,
    verify_dailysales_credentials
)

__all__ = [
    "get_odoo_credentials",
    "get_database_credentials",
    "decrypt_pin_secret",
    "fetch_dailysales_passwords",
    "verify_dailysales_credentials"
]
