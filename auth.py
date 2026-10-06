"""Autentikasi lokal berbasis SQLite tanpa menyimpan password plaintext."""
import hashlib
import hmac
import os
import re
from datetime import datetime, timezone


PBKDF2_ITERATIONS = 310_000
USERNAME_RE = re.compile(r"^[A-Za-z0-9_.-]{3,32}$")


def init_auth_db(conn):
    conn.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE COLLATE NOCASE,
        password_hash TEXT NOT NULL,
        password_salt TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('admin', 'user')),
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    )""")
    conn.commit()


def _validate_username(username):
    if not isinstance(username, str) or not USERNAME_RE.fullmatch(username):
        raise ValueError("Username harus 3-32 karakter: huruf, angka, titik, garis bawah, atau tanda hubung.")


def _validate_password(password):
    if not isinstance(password, str) or not password:
        raise ValueError("Password tidak boleh kosong.")


def _hash_password(password, salt=None):
    salt = salt or os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, PBKDF2_ITERATIONS)
    return digest.hex(), salt.hex()


def create_user(conn, username, password, role="user"):
    _validate_username(username)
    _validate_password(password)
    if role not in {"admin", "user"}:
        raise ValueError("Role harus admin atau user.")
    password_hash, salt = _hash_password(password)
    conn.execute(
        "INSERT INTO users(username,password_hash,password_salt,role,created_at) VALUES (?,?,?,?,?)",
        (username, password_hash, salt, role, datetime.now(timezone.utc).isoformat(timespec="seconds")),
    )
    conn.commit()


def authenticate(conn, username, password):
    row = conn.execute(
        "SELECT id, username, password_hash, password_salt, role FROM users "
        "WHERE username=? AND active=1", (username.strip(),)
    ).fetchone()
    if not row:
        return None
    password_hash, _ = _hash_password(password, bytes.fromhex(row[3]))
    if not hmac.compare_digest(password_hash, row[2]):
        return None
    return {"id": row[0], "username": row[1], "role": row[4]}


def list_users(conn):
    return conn.execute(
        "SELECT id, username, role, active, created_at FROM users ORDER BY username"
    ).fetchall()


def set_user_active(conn, user_id, active):
    conn.execute("UPDATE users SET active=? WHERE id=?", (int(bool(active)), user_id))
    conn.commit()


def user_count(conn):
    return conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
