import base64
import json
import os
from typing import Optional, Dict, Tuple

# Obfuscated key parts (split into byte segments to prevent simple binary string extraction)
_K_PART1 = b"RD4rpXJk1OegsLt"
_K_PART2 = b"7ZvZTFgpqiCAPcv"
_K_PART3 = b"sjbKtM9mZrbeM="

# Encrypted credentials payload (AES-CBC + HMAC via Fernet)
_ENC_BLOB = b"gAAAAABqlWmgdEcjCgIilioEACQmA37zJOdYety5w6H2Q9g1ChYcK8TFY7P3rXxExpn9wPfIWhjjszxBRdyXmmduqn38LJRe6rNUNGWgnOlTJ_TSQGNSXVQN5nldhOCHNmwy4r9rmoEt3b4dR1t2bSio00qvgiUBOctUWqjSxiGT4A1EBjRbi-VA3WMYXkvfzensXf70EGHERiOhNUaMWhetkaaVmxy8sh_iFKlGYTBERb2jxsoPst8f26R6YWvIfcePTJpuUhTOfDwlXuMMR65uRCfXu6VMDEdNdrCZaSi2P9L_Zwbr2oMccMrbY0k9sLW3CiJSZ8vJ"

# Pure Python zero-dependency fallback payload
_FALLBACK_KEY = b"AlAssali2026SecureOdooXmlRpcKey"
_FALLBACK_BLOB = b"KRg1AwBbQ0ZQUV9ZPFdTRUJILggOHDkBBTEfTSQBFi5CIhweHRAVUFFfWTxXU0VCSC4IDhw5AQUxH047FxYlGSIHGg4CRAMIAAdkUlFMDhkzBQ0LLQEAMxhNICUYLQ0yABINBUdcVUZKLxlUFEJdeQFfCjlfCjBBVy0BG3APdRJBUVteBFIBUmRcWhRBU3tRWQs7"

# Encrypted Database Payload (Supabase Cloud PostgreSQL)
_ENC_DB_BLOB = b"gAAAAABq7uJFwTn2fvFe8moH_G1c56_fwljUPh3ukAhxTXZYH_7fE9V7o27ZavA9CEQmA7ZJVttwonzvSZkoIen92SlvYq7EIRLRA41lXEPXDFxlaZob-o2JWHz8wjQJtkymtIaL1QGw6L3NqfgMxvJZ_BA7zQpLTFdwO0iQn67HGXFZ-jRdvBDx4rjvLrtmRIWaFa-vw8GoEae32uFXspioyNEQmXd_l62hTePLYUV6zN8qG9j1vl5HEpR0RqiUXz397D4bF6qj"
_FALLBACK_DB_BLOB = b"MQMyBxQTCRpDXBlGIBwAGgICfV5AQCgCHyYXES4WVzsYIB8bFB0QV0FWQyAdDgAZCSITVS46CUlmQFV4V093WXIzEhYfRAIdU0Z+CwwHBg0qBRwbdVxCIh8MJwALbx80AxIDDRpXHlFZPl9WQUFXYBQAHCwKHjcDXDgWFSwDJRZOEwkYR1lAUw=="

def get_odoo_credentials():
    """
    Decrypts and returns the Odoo API credentials securely in memory.
    Credentials are never written to disk or shown in plain text.
    Resilient to missing third-party crypto packages in minimal environments.
    """
    # 1. Primary: Fernet decryption if cryptography is installed
    try:
        from cryptography.fernet import Fernet
        combined_key = _K_PART1 + _K_PART2 + _K_PART3
        cipher = Fernet(combined_key)
        decrypted_bytes = cipher.decrypt(_ENC_BLOB)
        data = json.loads(decrypted_bytes.decode('utf-8'))
        return data['url'], data['db'], data['user'], data['pwd']
    except Exception:
        pass

    # 2. Resilient Pure-Python Fallback (guaranteed to work in zero-dependency environments)
    try:
        raw_bytes = base64.b64decode(_FALLBACK_BLOB)
        decrypted = bytes([b ^ _FALLBACK_KEY[i % len(_FALLBACK_KEY)] for i, b in enumerate(raw_bytes)]).decode('utf-8')
        u, d, us, p = decrypted.split('|||')
        return u, d, us, p
    except Exception as e:
        raise RuntimeError("Security verification failed: Unable to decode credentials.") from e

def get_database_credentials():
    """
    Decrypts and returns the Supabase Cloud PostgreSQL database URL securely in memory.
    Credentials are never written to disk or shown in plain text.
    """
    # 1. Primary: Fernet decryption if cryptography is installed
    try:
        from cryptography.fernet import Fernet
        combined_key = _K_PART1 + _K_PART2 + _K_PART3
        cipher = Fernet(combined_key)
        decrypted_bytes = cipher.decrypt(_ENC_DB_BLOB)
        return decrypted_bytes.decode('utf-8')
    except Exception:
        pass

    # 2. Resilient Pure-Python Fallback
    try:
        raw_bytes = base64.b64decode(_FALLBACK_DB_BLOB)
        return bytes([b ^ _FALLBACK_KEY[i % len(_FALLBACK_KEY)] for i, b in enumerate(raw_bytes)]).decode('utf-8')
    except Exception as e:
        raise RuntimeError("Security verification failed: Unable to decode database credentials.") from e

_AES_PIN_KEY = b"AlAssali2026SecureOdooKeySecret!"
_AES_PIN_IV  = b"AlAssaliIV123456"

def decrypt_pin_secret(enc_str: str, default: str = "") -> str:
    """Decrypts AES-256-CBC encrypted PIN or password stored in Odoo."""
    if not enc_str:
        return default
    try:
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        from cryptography.hazmat.primitives import padding
        ct = base64.b64decode(enc_str.strip())
        cipher = Cipher(algorithms.AES(_AES_PIN_KEY), modes.CBC(_AES_PIN_IV))
        decryptor = cipher.decryptor()
        padded_data = decryptor.update(ct) + decryptor.finalize()
        unpadder = padding.PKCS7(128).unpadder()
        return (unpadder.update(padded_data) + unpadder.finalize()).decode('utf-8')
    except Exception:
        return default

def fetch_dailysales_passwords() -> Dict[str, str]:
    """
    Fetches the configured Editor and Viewer passwords for DailySalesApp.
    Reads from Odoo ir.config_parameter with AES-256 decryption.
    Falls back gracefully to defaults:
      - editor: 'assali2026'
      - viewer: 'reports2026'
    """
    env_editor = os.environ.get("APP_PASSWORD") or os.environ.get("DAILYSALES_EDITOR_PASSWORD")
    env_viewer = os.environ.get("DAILYSALES_VIEWER_PASSWORD")

    editor_pwd = env_editor or "assali2026"
    viewer_pwd = env_viewer or "reports2026"

    try:
        import xmlrpc.client
        url, db, user, pwd = get_odoo_credentials()
        common = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/common", allow_none=True)
        uid = common.authenticate(db, user, pwd, {})
        if uid:
            models = xmlrpc.client.ServerProxy(f"{url}/xmlrpc/2/object", allow_none=True)
            ed_raw = models.execute_kw(db, uid, pwd, "ir.config_parameter", "get_param", ["dailysales.editor_password"])
            vi_raw = models.execute_kw(db, uid, pwd, "ir.config_parameter", "get_param", ["dailysales.viewer_password"])
            if ed_raw and ed_raw is not False:
                editor_pwd = decrypt_pin_secret(str(ed_raw), editor_pwd)
            if vi_raw and vi_raw is not False:
                viewer_pwd = decrypt_pin_secret(str(vi_raw), viewer_pwd)
    except Exception:
        pass

    return {
        "editor": editor_pwd,
        "viewer": viewer_pwd
    }

def verify_dailysales_credentials(input_pwd: str) -> Optional[str]:
    """
    Verifies input password and returns role:
      - 'editor': Full permissions (Daily sales, adjustments, closures, reports).
      - 'viewer': Restricted to Periodic & Comparative Reports screen only.
      - None: Invalid credentials.
    """
    if not input_pwd:
        return None
    inp = input_pwd.strip()
    pwds = fetch_dailysales_passwords()
    if inp == pwds.get("editor", "assali2026"):
        return "editor"
    if inp == pwds.get("viewer", "reports2026"):
        return "viewer"
    return None

if __name__ == "__main__":
    url, db, user, pwd = get_odoo_credentials()
    print("Odoo verified:", url, db, user, "Password length:", len(pwd))
    db_url = get_database_credentials()
    print("Database verified:", "postgres" in db_url)
