import hashlib
import secrets
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path


AUTH_ENABLED = False
DB_PATH = Path(__file__).resolve().parent / "reclamacoes.db"
SESSION_DAYS = 7


def hash_password(password):
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def create_user(conn, nome, email, senha, perfil="consulta", regional=None, analista_sac=None, ativo=1):
    conn.execute(
        """
        INSERT INTO users (nome, email, senha_hash, perfil, regional, analista_sac, ativo, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (nome, email.lower().strip(), hash_password(senha), perfil, regional, analista_sac, ativo, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    )


def authenticate(conn, email, password):
    row = conn.execute(
        "SELECT id, nome, email, perfil, regional, analista_sac, ativo, senha_hash FROM users WHERE email = ?",
        (email.lower().strip(),),
    ).fetchone()
    if not row or not row["ativo"] or row["senha_hash"] != hash_password(password):
        return None
    token = secrets.token_urlsafe(32)
    user = {key: row[key] for key in ["id", "nome", "email", "perfil", "regional", "analista_sac", "ativo"]}
    expires_at = (datetime.now() + timedelta(days=SESSION_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    conn.execute(
        """
        INSERT INTO user_sessions (token, user_id, created_at, expires_at)
        VALUES (?, ?, ?, ?)
        """,
        (token, user["id"], datetime.now().strftime("%Y-%m-%d %H:%M:%S"), expires_at),
    )
    return token, user


def logout(token):
    if token:
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("DELETE FROM user_sessions WHERE token = ?", (token,))


def token_from_headers(headers):
    cookie = headers.get("Cookie", "")
    for part in cookie.split(";"):
        name, _, value = part.strip().partition("=")
        if name == "mms_session":
            return value
    auth = headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        return auth[7:]
    return None


def current_user(headers=None):
    token = token_from_headers(headers or {})
    if token:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                """
                SELECT u.id, u.nome, u.email, u.perfil, u.regional, u.analista_sac, u.ativo
                FROM user_sessions s
                JOIN users u ON u.id = s.user_id
                WHERE s.token = ? AND s.expires_at >= ? AND u.ativo = 1
                """,
                (token, datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ).fetchone()
            if row:
                return {key: row[key] for key in ["id", "nome", "email", "perfil", "regional", "analista_sac", "ativo"]}
    if not AUTH_ENABLED:
        return {"id": 0, "perfil": "admin", "nome": "Acesso aberto", "email": None, "regional": None, "analista_sac": None}
    return None


def require_user(headers=None):
    user = current_user(headers)
    if not user:
        raise PermissionError("Usuario nao autenticado.")
    return user
