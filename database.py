from app import DB_PATH, connect, init_db
from models_sqlalchemy import SQLALCHEMY_AVAILABLE, DEFAULT_DATABASE_URL, get_session, create_all_models


def database_path():
    return DB_PATH


def database_url():
    return DEFAULT_DATABASE_URL


def sqlalchemy_status():
    return {"available": SQLALCHEMY_AVAILABLE, "databaseUrl": DEFAULT_DATABASE_URL}
