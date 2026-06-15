import argparse
import csv
import io
import hashlib
import json
import re
import sqlite3
import unicodedata
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
UPLOAD_DIR = APP_DIR / "uploads"
DB_PATH = APP_DIR / "reclamacoes.db"
CITIES_PATH = STATIC_DIR / "data" / "cidades_brasil.json"
CONFIG_PATH = APP_DIR / "config.json"

DEFAULT_CONFIG = {
    "daily_closed_goal": 10,
    "sla_business_days_target": 5,
    "date_lag_warning_days": 45,
    "efficiency_score": {
        "closed_weight": 1.0,
        "sla_weight": 1.0,
        "backlog_penalty": 0.75,
    },
}


def load_config():
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if CONFIG_PATH.exists():
        try:
            loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            for key, value in loaded.items():
                if isinstance(value, dict) and isinstance(config.get(key), dict):
                    config[key].update(value)
                else:
                    config[key] = value
        except Exception:
            pass
    return config


CONFIG = load_config()
DAILY_CLOSED_GOAL = int(CONFIG.get("daily_closed_goal", DEFAULT_CONFIG["daily_closed_goal"]))
SLA_BUSINESS_DAYS_TARGET = int(CONFIG.get("sla_business_days_target", DEFAULT_CONFIG["sla_business_days_target"]))
DATE_LAG_WARNING_DAYS = int(CONFIG.get("date_lag_warning_days", DEFAULT_CONFIG["date_lag_warning_days"]))
EFFICIENCY_SCORE_CONFIG = CONFIG.get("efficiency_score", DEFAULT_CONFIG["efficiency_score"])


def slug(value):
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")


def find_col(columns, *needles):
    for col in columns:
        normalized = slug(col)
        if all(needle in normalized for needle in needles):
            return col
    raise KeyError(f"Campo nao encontrado: {needles}")


def find_optional_col(columns, *needles):
    try:
        return find_col(columns, *needles)
    except KeyError:
        return None


def clean(value, fallback):
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return fallback
    return text


def clean_date(value):
    if pd.isna(value):
        return None
    date = pd.to_datetime(value, errors="coerce")
    if pd.isna(date):
        return None
    return date.strftime("%Y-%m-%d")


def business_days_between(start, end):
    if pd.isna(start) or pd.isna(end):
        return None
    start = pd.to_datetime(start, errors="coerce")
    end = pd.to_datetime(end, errors="coerce")
    if pd.isna(start) or pd.isna(end) or end.date() < start.date():
        return None
    return len(pd.bdate_range(start=start.date(), end=end.date()))


def pct_value(part, total):
    return part / total if total else 0


def status_flag(value, *needles):
    normalized = slug(value)
    return any(needle in normalized for needle in needles)


def add_period_filter(where, args, field, month_field, params, max_key):
    period = params.get("period", "all")
    if period == "last7":
        max_date = params.get(max_key)
        if max_date:
            where.append(f"{field} >= date(?, '-7 day')")
            args.append(max_date)
    elif period == "last30":
        max_date = params.get(max_key)
        if max_date:
            where.append(f"{field} >= date(?, '-30 day')")
            args.append(max_date)
    elif period == "last90":
        max_date = params.get(max_key)
        if max_date:
            where.append(f"{field} >= date(?, '-90 day')")
            args.append(max_date)
    elif period == "last12":
        max_date = params.get(max_key)
        if max_date:
            where.append(f"{field} >= date(?, '-12 month')")
            args.append(max_date)
    elif period == "ytd":
        max_date = params.get(max_key)
        if max_date:
            where.append(f"{field} >= date(?, 'start of year')")
            args.append(max_date)
    elif period == "month" and params.get("month"):
        where.append(f"{month_field} = ?")
        args.append(params["month"])
    elif period == "range":
        if params.get("start"):
            where.append(f"{field} >= ?")
            args.append(params["start"])
        if params.get("end"):
            where.append(f"{field} <= ?")
            args.append(params["end"])


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    with connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS import_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                mode TEXT NOT NULL,
                rows_total INTEGER NOT NULL,
                rows_inserted INTEGER NOT NULL,
                procedentes INTEGER NOT NULL,
                date_min TEXT,
                date_max TEXT
            );

            CREATE TABLE IF NOT EXISTS complaints (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                record_hash TEXT NOT NULL,
                data_reclamacao TEXT,
                dt_fechamento TEXT,
                mes TEXT,
                mes_fechamento TEXT,
                cliente TEXT,
                cidade TEXT,
                uf TEXT,
                regiao TEXT,
                hub TEXT,
                servico TEXT,
                status_reclamacao TEXT,
                validacao TEXT,
                motivo TEXT,
                grupo_motivo TEXT,
                prestador TEXT,
                analista_sac TEXT,
                dias_uteis_fechamento INTEGER,
                sla_ok INTEGER,
                procedente INTEGER NOT NULL,
                source_file TEXT NOT NULL,
                imported_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_complaints_date ON complaints(data_reclamacao);
            CREATE INDEX IF NOT EXISTS idx_complaints_month ON complaints(mes);
            CREATE INDEX IF NOT EXISTS idx_complaints_proc ON complaints(procedente);
            CREATE INDEX IF NOT EXISTS idx_complaints_dims ON complaints(regiao, hub, cliente, cidade);
            CREATE INDEX IF NOT EXISTS idx_complaints_analyst ON complaints(analista_sac);

            CREATE TABLE IF NOT EXISTS operation_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                mode TEXT NOT NULL,
                rows_total INTEGER NOT NULL,
                rows_inserted INTEGER NOT NULL,
                concluidas INTEGER NOT NULL,
                date_min TEXT,
                date_max TEXT
            );

            CREATE TABLE IF NOT EXISTS operations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                record_hash TEXT NOT NULL,
                data TEXT,
                mes TEXT,
                numero_assistencia TEXT,
                filial TEXT,
                categoria TEXT,
                cidade TEXT,
                estado TEXT,
                regional TEXT,
                area_trabalho TEXT,
                hub TEXT,
                posto TEXT,
                tipo_atividade TEXT,
                status TEXT,
                recurso TEXT,
                nome_cliente TEXT,
                cliente TEXT,
                motivo_nao_realizada TEXT,
                id_montador TEXT,
                concluida INTEGER NOT NULL,
                cancelada INTEGER NOT NULL,
                frustrada INTEGER NOT NULL,
                reagendada INTEGER NOT NULL,
                improdutiva INTEGER NOT NULL,
                perda INTEGER NOT NULL,
                source_file TEXT NOT NULL,
                imported_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_operations_date ON operations(data);
            CREATE INDEX IF NOT EXISTS idx_operations_month ON operations(mes);
            CREATE INDEX IF NOT EXISTS idx_operations_dims ON operations(regional, hub, cidade, cliente);
            CREATE INDEX IF NOT EXISTS idx_operations_resource ON operations(recurso);
            """
        )
        cols = [row["name"] for row in conn.execute("PRAGMA table_info(complaints)").fetchall()]
        for name, definition in {
            "analista_sac": "TEXT",
            "dt_fechamento": "TEXT",
            "mes_fechamento": "TEXT",
            "dias_uteis_fechamento": "INTEGER",
            "sla_ok": "INTEGER",
        }.items():
            if name not in cols:
                conn.execute(f"ALTER TABLE complaints ADD COLUMN {name} {definition}")


def row_hash(row):
    parts = [
        row.get("id_feedback", ""),
        row.get("id_assistservico", ""),
        row.get("id_assistencia", ""),
        row.get("date", ""),
        row.get("cliente", ""),
        row.get("cidade", ""),
        row.get("hub", ""),
        row.get("motivo", ""),
        row.get("row_index", ""),
    ]
    return hashlib.sha256("||".join(str(part) for part in parts).encode("utf-8")).hexdigest()


def operation_row_hash(row):
    parts = [
        row.get("numero_assistencia", ""),
        row.get("date", ""),
        row.get("cliente", ""),
        row.get("cidade", ""),
        row.get("hub", ""),
        row.get("recurso", ""),
        row.get("status", ""),
        row.get("motivo", ""),
        row.get("row_index", ""),
    ]
    return hashlib.sha256("||".join(str(part) for part in parts).encode("utf-8")).hexdigest()


def excel_sheet_candidates(path):
    candidates = []
    with pd.ExcelFile(path) as workbook:
        for sheet_name in workbook.sheet_names:
            preview = pd.read_excel(workbook, sheet_name=sheet_name, nrows=10)
            columns = {slug(col) for col in preview.columns}
            operation_score = sum(
                1
                for expected in [
                    "numero_da_assistencia",
                    "tipo_atividade",
                    "recurso",
                    "status",
                    "data",
                    "motivo_nao_realizada",
                ]
                if expected in columns
            )
            complaint_score = 0
            for col in columns:
                if col in {"data_reclamacao", "dt_fechamento", "analistasac", "hub"}:
                    complaint_score += 2
                if "reclamacao" in col or "procedencia" in col or "feedback" in col:
                    complaint_score += 1
            sheet_type = "operation" if operation_score >= 3 and operation_score >= complaint_score else "complaints"
            score = max(operation_score, complaint_score)
            candidates.append(
                {
                    "sheet": sheet_name,
                    "type": sheet_type,
                    "score": score,
                    "operationScore": operation_score,
                    "complaintScore": complaint_score,
                }
            )
    candidates.sort(key=lambda item: item["score"], reverse=True)
    return candidates


def read_best_excel_sheet(path):
    candidates = excel_sheet_candidates(path)
    if not candidates or candidates[0]["score"] <= 0:
        sheets = ", ".join(item["sheet"] for item in candidates) or "nenhuma aba encontrada"
        raise ValueError(f"Nao encontrei uma aba compativel para importar. Abas disponiveis: {sheets}.")
    best = candidates[0]
    df = pd.read_excel(path, sheet_name=best["sheet"])
    return df, best


def operation_flags(status, reason):
    combined = f"{status} {reason}"
    concluida = 1 if status_flag(status, "concluida", "concluido", "finalizada", "realizada") else 0
    cancelada = 1 if status_flag(combined, "cancelada", "cancelado", "cancelamento") else 0
    frustrada = 1 if status_flag(combined, "frustrada", "frustrado", "frustracao") else 0
    reagendada = 1 if status_flag(combined, "reagendada", "reagendado", "reagendamento") else 0
    improdutiva = 1 if status_flag(combined, "improdutiva", "improdutivo", "improdutividade") else 0
    perda = 1 if any([cancelada, frustrada, reagendada, improdutiva]) else 0
    return concluida, cancelada, frustrada, reagendada, improdutiva, perda


def import_operation_workbook(path, mode="replace", df=None, sheet_name=None):
    init_db()
    path = Path(path)
    if df is None:
        df, best = read_best_excel_sheet(path)
        sheet_name = best["sheet"]
    df["_data"] = pd.to_datetime(df["data"], errors="coerce")
    df["_mes"] = df["_data"].dt.to_period("M").astype(str).replace("NaT", "Sem data")

    imported_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for idx, source in df.iterrows():
        status = clean(source.get("status"), "Sem status")
        reason = clean(source.get("motivo_nao_realizada"), "Sem motivo")
        concluida, cancelada, frustrada, reagendada, improdutiva, perda = operation_flags(status, reason)
        record = {
            "row_index": idx,
            "date": clean_date(source["_data"]),
            "month": clean(source["_mes"], "Sem data"),
            "numero_assistencia": clean(source.get("numero_da_assistencia"), ""),
            "filial": clean(source.get("filial"), "Sem filial"),
            "categoria": clean(source.get("categoria"), "Sem categoria"),
            "cidade": clean(source.get("cidade"), "Sem cidade"),
            "estado": clean(source.get("estado"), "Sem UF"),
            "regional": clean(source.get("celula"), "Sem regional"),
            "area_trabalho": clean(source.get("area_de_trabalho"), "Sem area"),
            "hub": clean(source.get("hub"), "Sem HUB"),
            "posto": clean(source.get("posto"), "Sem posto"),
            "tipo_atividade": clean(source.get("tipo_atividade"), "Sem atividade"),
            "status": status,
            "recurso": clean(source.get("recurso"), "Sem recurso"),
            "nome_cliente": clean(source.get("nome_cliente"), "Sem nome"),
            "cliente": clean(source.get("cliente"), "Sem cliente"),
            "motivo": reason,
            "id_montador": clean(source.get("id_montador"), "Sem ID"),
            "concluida": concluida,
            "cancelada": cancelada,
            "frustrada": frustrada,
            "reagendada": reagendada,
            "improdutiva": improdutiva,
            "perda": perda,
            "source_file": path.name,
            "imported_at": imported_at,
        }
        record["record_hash"] = operation_row_hash(record)
        rows.append(record)

    with connect() as conn:
        if mode == "replace":
            conn.execute("DELETE FROM operations")
            conn.execute("DELETE FROM operation_batches")

        before = conn.execute("SELECT COUNT(*) AS total FROM operations").fetchone()["total"]
        conn.executemany(
            """
            INSERT INTO operations (
                record_hash, data, mes, numero_assistencia, filial, categoria, cidade, estado, regional,
                area_trabalho, hub, posto, tipo_atividade, status, recurso, nome_cliente, cliente,
                motivo_nao_realizada, id_montador, concluida, cancelada, frustrada, reagendada,
                improdutiva, perda, source_file, imported_at
            ) VALUES (
                :record_hash, :date, :month, :numero_assistencia, :filial, :categoria, :cidade, :estado, :regional,
                :area_trabalho, :hub, :posto, :tipo_atividade, :status, :recurso, :nome_cliente, :cliente,
                :motivo, :id_montador, :concluida, :cancelada, :frustrada, :reagendada,
                :improdutiva, :perda, :source_file, :imported_at
            )
            """,
            rows,
        )
        after = conn.execute("SELECT COUNT(*) AS total FROM operations").fetchone()["total"]
        inserted = after - before
        dates = [item["date"] for item in rows if item["date"]]
        conn.execute(
            """
            INSERT INTO operation_batches (
                file_name, imported_at, mode, rows_total, rows_inserted, concluidas, date_min, date_max
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                path.name,
                imported_at,
                mode,
                len(rows),
                inserted,
                sum(item["concluida"] for item in rows),
                min(dates) if dates else None,
                max(dates) if dates else None,
            ),
        )
    return {
        "type": "operation",
        "file": path.name,
        "sheet": sheet_name,
        "mode": mode,
        "rowsTotal": len(rows),
        "rowsInserted": inserted,
        "concluidas": sum(item["concluida"] for item in rows),
        "dateMin": min(dates) if dates else None,
        "dateMax": max(dates) if dates else None,
    }


def import_workbook(path, mode="replace"):
    init_db()
    path = Path(path)
    df, best = read_best_excel_sheet(path)
    if best["type"] == "operation":
        return import_operation_workbook(path, mode=mode, df=df, sheet_name=best["sheet"])

    try:
        proc_col = find_col(df.columns, "tipo", "reclamacao", "csr")
        consolidated_col = find_col(df.columns, "procedencia", "consolidado")
    except KeyError as exc:
        columns = ", ".join(str(col) for col in df.columns)
        raise ValueError(
            f"A aba '{best['sheet']}' parece ser de reclamacoes, mas faltam campos obrigatorios. "
            f"Campo esperado: {exc}. Colunas encontradas: {columns}"
        ) from exc

    status_norm = df[proc_col].astype("string").str.strip().str.upper()
    consolidated = pd.to_numeric(df[consolidated_col], errors="coerce").fillna(0)
    df["_procedente"] = (status_norm.eq("PROCEDENTE") | consolidated.eq(1)).fillna(False)
    df["_data"] = pd.to_datetime(df["data_reclamacao"], errors="coerce")
    df["_dt_fechamento"] = pd.to_datetime(df["dt_fechamento"], errors="coerce")
    df["_mes"] = df["_data"].dt.to_period("M").astype(str).replace("NaT", "Sem data")
    df["_mes_fechamento"] = df["_dt_fechamento"].dt.to_period("M").astype(str).replace("NaT", "Sem data")

    imported_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []
    for idx, source in df.iterrows():
        record = {
            "row_index": idx,
            "id_feedback": clean(source.get("id_feedback"), ""),
            "id_assistservico": clean(source.get("id_assistservico"), ""),
            "id_assistencia": clean(source.get("id_assistencia"), ""),
            "date": clean_date(source["_data"]),
            "close_date": clean_date(source["_dt_fechamento"]),
            "month": clean(source["_mes"], "Sem data"),
            "close_month": clean(source["_mes_fechamento"], "Sem data"),
            "cliente": clean(source.get("Cliente"), "Sem cliente"),
            "cidade": clean(source.get("desc_cidade"), "Sem cidade"),
            "uf": clean(source.get("sigla"), "Sem UF"),
            "regiao": clean(source.get("regiao"), "Sem região"),
            "hub": clean(source.get("hub"), "Sem HUB"),
            "servico": clean(source.get("desc_servico"), "Sem serviço"),
            "status": clean(source.get("Statusreclamacao"), "Sem status"),
            "validacao": clean(source.get(proc_col), "Sem validação"),
            "motivo": clean(source.get("Motivo_Feedback"), "Sem motivo"),
            "grupo": clean(source.get("GrupoMotivoFeedback"), "Sem grupo"),
            "prestador": clean(source.get("Prestador"), "Sem prestador"),
            "analista": clean(source.get("analistasac"), "Sem analista"),
            "business_days_close": business_days_between(source["_data"], source["_dt_fechamento"]),
            "procedente": 1 if bool(source["_procedente"] is True or source["_procedente"] == 1) else 0,
            "source_file": path.name,
            "imported_at": imported_at,
        }
        record["sla_ok"] = (
            1
            if record["business_days_close"] is not None
            and record["business_days_close"] <= SLA_BUSINESS_DAYS_TARGET
            else 0
        )
        record["record_hash"] = row_hash(record)
        rows.append(record)

    with connect() as conn:
        if mode == "replace":
            conn.execute("DELETE FROM complaints")
            conn.execute("DELETE FROM import_batches")

        before = conn.execute("SELECT COUNT(*) AS total FROM complaints").fetchone()["total"]
        conn.executemany(
            """
            INSERT INTO complaints (
                record_hash, data_reclamacao, dt_fechamento, mes, mes_fechamento, cliente, cidade, uf, regiao, hub, servico,
                status_reclamacao, validacao, motivo, grupo_motivo, prestador, analista_sac, dias_uteis_fechamento, sla_ok, procedente,
                source_file, imported_at
            ) VALUES (
                :record_hash, :date, :close_date, :month, :close_month, :cliente, :cidade, :uf, :regiao, :hub, :servico,
                :status, :validacao, :motivo, :grupo, :prestador, :analista, :business_days_close, :sla_ok, :procedente,
                :source_file, :imported_at
            )
            """,
            rows,
        )
        after = conn.execute("SELECT COUNT(*) AS total FROM complaints").fetchone()["total"]
        inserted = after - before
        dates = [item["date"] for item in rows if item["date"]]
        conn.execute(
            """
            INSERT INTO import_batches (
                file_name, imported_at, mode, rows_total, rows_inserted, procedentes, date_min, date_max
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                path.name,
                imported_at,
                mode,
                len(rows),
                inserted,
                sum(item["procedente"] for item in rows),
                min(dates) if dates else None,
                max(dates) if dates else None,
            ),
        )
    return {
        "type": "complaints",
        "file": path.name,
        "sheet": best["sheet"],
        "mode": mode,
        "rowsTotal": len(rows),
        "rowsInserted": inserted,
        "procedentes": sum(item["procedente"] for item in rows),
        "dateMin": min(dates) if dates else None,
        "dateMax": max(dates) if dates else None,
    }


def build_where(params, procedentes_only=False):
    where = []
    args = []
    if procedentes_only:
        where.append("procedente = 1")
    if params.get("regiao"):
        where.append("regiao = ?")
        args.append(params["regiao"])
    if params.get("hub"):
        where.append("hub = ?")
        args.append(params["hub"])
    if params.get("cliente"):
        where.append("cliente = ?")
        args.append(params["cliente"])
    if params.get("cidade"):
        where.append("cidade LIKE ?")
        args.append(f"%{params['cidade']}%")
    if params.get("analista"):
        where.append("analista_sac = ?")
        args.append(params["analista"])

    add_period_filter(where, args, "data_reclamacao", "mes", params, "_max_date")

    clause = " WHERE " + " AND ".join(where) if where else ""
    return clause, args


def build_close_where(params, closed_only=True):
    where = []
    args = []
    if closed_only:
        where.append("UPPER(status_reclamacao) = 'FECHADA'")
    where.append("dt_fechamento IS NOT NULL")
    if params.get("regiao"):
        where.append("regiao = ?")
        args.append(params["regiao"])
    if params.get("hub"):
        where.append("hub = ?")
        args.append(params["hub"])
    if params.get("cliente"):
        where.append("cliente = ?")
        args.append(params["cliente"])
    if params.get("cidade"):
        where.append("cidade LIKE ?")
        args.append(f"%{params['cidade']}%")
    if params.get("analista"):
        where.append("analista_sac = ?")
        args.append(params["analista"])

    if "_max_close_date" not in params and "_max_date" in params:
        params["_max_close_date"] = params["_max_date"]
    add_period_filter(where, args, "dt_fechamento", "mes_fechamento", params, "_max_close_date")

    clause = " WHERE " + " AND ".join(where) if where else ""
    return clause, args


def count_by(conn, params, field, limit=12):
    where, args = build_where(params, procedentes_only=params.get("scope", "procedentes") == "procedentes")
    rows = conn.execute(
        f"""
        SELECT {field} AS label, COUNT(*) AS value
        FROM complaints
        {where}
        GROUP BY {field}
        ORDER BY value DESC, label ASC
        LIMIT ?
        """,
        [*args, limit],
    ).fetchall()
    return [dict(row) for row in rows]


def table_combo(conn, params):
    where, args = build_where(params, procedentes_only=True)
    rows = conn.execute(
        f"""
        SELECT regiao, hub, cliente, cidade, COUNT(*) AS procedentes
        FROM complaints
        {where}
        GROUP BY regiao, hub, cliente, cidade
        ORDER BY procedentes DESC, regiao, hub, cliente, cidade
        LIMIT 50
        """,
        args,
    ).fetchall()
    return [dict(row) for row in rows]


def table_group_region(conn, params):
    where, args = build_where(params, procedentes_only=True)
    rows = conn.execute(
        f"""
        SELECT regiao, grupo_motivo AS grupo, COUNT(*) AS procedentes
        FROM complaints
        {where}
        GROUP BY regiao, grupo_motivo
        ORDER BY procedentes DESC, regiao, grupo
        LIMIT 40
        """,
        args,
    ).fetchall()
    return [dict(row) for row in rows]


def table_providers(conn, params):
    where, args = build_where(params, procedentes_only=True)
    rows = conn.execute(
        f"""
        SELECT prestador, hub, cidade, COUNT(*) AS procedentes
        FROM complaints
        {where}
        GROUP BY prestador, hub, cidade
        ORDER BY procedentes DESC, prestador
        LIMIT 40
        """,
        args,
    ).fetchall()
    return [dict(row) for row in rows]


def api_productivity(query):
    init_db()
    params = {key: values[0] for key, values in query.items() if values and values[0]}
    closed_weight = float(EFFICIENCY_SCORE_CONFIG.get("closed_weight", 1.0))
    sla_weight = float(EFFICIENCY_SCORE_CONFIG.get("sla_weight", 1.0))
    backlog_penalty = float(EFFICIENCY_SCORE_CONFIG.get("backlog_penalty", 0.75))
    with connect() as conn:
        max_date = conn.execute("SELECT MAX(data_reclamacao) AS value FROM complaints").fetchone()["value"]
        max_close_date = conn.execute("SELECT MAX(dt_fechamento) AS value FROM complaints").fetchone()["value"]
        params["_max_date"] = max_date
        params["_max_close_date"] = max_close_date
        where, args = build_where(params, procedentes_only=False)
        close_where, close_args = build_close_where(params, closed_only=True)

        total = conn.execute(f"SELECT COUNT(*) AS value FROM complaints {where}", args).fetchone()["value"]
        analysts = conn.execute(
            f"SELECT COUNT(DISTINCT analista_sac) AS value FROM complaints {where}", args
        ).fetchone()["value"]
        days = conn.execute(
            f"SELECT COUNT(DISTINCT data_reclamacao) AS value FROM complaints {where}", args
        ).fetchone()["value"]
        months = conn.execute(
            f"SELECT COUNT(DISTINCT CASE WHEN mes != 'Sem data' THEN mes END) AS value FROM complaints {where}",
            args,
        ).fetchone()["value"]
        closed_total = conn.execute(
            f"SELECT COUNT(*) AS value FROM complaints {close_where}", close_args
        ).fetchone()["value"]
        closed_days = conn.execute(
            f"SELECT COUNT(DISTINCT dt_fechamento) AS value FROM complaints {close_where}", close_args
        ).fetchone()["value"]
        closed_analysts = conn.execute(
            f"SELECT COUNT(DISTINCT analista_sac) AS value FROM complaints {close_where}", close_args
        ).fetchone()["value"]
        goal_expected = closed_analysts * closed_days * DAILY_CLOSED_GOAL
        sla = conn.execute(
            f"""
            SELECT
              COUNT(*) AS fechadas_com_prazo,
              SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) AS dentro_sla,
              ROUND(AVG(dias_uteis_fechamento), 2) AS media_dias_uteis
            FROM complaints
            {close_where}
            """,
            close_args,
        ).fetchone()
        backlog_rows = conn.execute(
            f"""
            SELECT status_reclamacao AS label, COUNT(*) AS value
            FROM complaints
            {where + (" AND UPPER(status_reclamacao) != 'FECHADA'" if where else " WHERE UPPER(status_reclamacao) != 'FECHADA'")}
            GROUP BY status_reclamacao
            ORDER BY value DESC, status_reclamacao ASC
            """
        , args).fetchall()
        status = conn.execute(
            f"""
            SELECT
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              SUM(CASE WHEN UPPER(validacao) = 'IMPROCEDENTE' THEN 1 ELSE 0 END) AS improcedentes,
              SUM(CASE WHEN validacao = 'Sem validacao' OR validacao = 'Sem validação' OR validacao IS NULL THEN 1 ELSE 0 END) AS sem_status,
              SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END) AS fechadas
            FROM complaints
            {where}
            """,
            args,
        ).fetchone()
        lead = conn.execute(
            f"""
            SELECT analista_sac AS label, COUNT(*) AS value
            FROM complaints
            {where}
            GROUP BY analista_sac
            ORDER BY value DESC, analista_sac ASC
            LIMIT 1
            """,
            args,
        ).fetchone()

        analyst_rows = conn.execute(
            f"""
            SELECT
              analista_sac AS analista,
              COUNT(*) AS total,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              SUM(CASE WHEN UPPER(validacao) = 'IMPROCEDENTE' THEN 1 ELSE 0 END) AS improcedentes,
              SUM(CASE WHEN validacao = 'Sem validacao' OR validacao = 'Sem validação' OR validacao IS NULL THEN 1 ELSE 0 END) AS sem_status,
              SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END) AS fechadas,
              COUNT(DISTINCT data_reclamacao) AS dias_ativos,
              COUNT(DISTINCT CASE WHEN mes != 'Sem data' THEN mes END) AS meses_ativos,
              ROUND(COUNT(*) * 1.0 / NULLIF(COUNT(DISTINCT data_reclamacao), 0), 2) AS media_dia,
              ROUND(COUNT(*) * 1.0 / NULLIF(COUNT(DISTINCT CASE WHEN mes != 'Sem data' THEN mes END), 0), 2) AS media_mes,
              ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS taxa_procedencia
            FROM complaints
            {where}
            GROUP BY analista_sac
            ORDER BY total DESC, analista_sac ASC
            LIMIT 80
            """,
            args,
        ).fetchall()

        monthly_where = close_where + " AND mes_fechamento != 'Sem data'"
        monthly = conn.execute(
            f"""
            SELECT mes_fechamento AS label, COUNT(*) AS value
            FROM complaints
            {monthly_where}
            GROUP BY mes_fechamento
            ORDER BY mes_fechamento ASC
            """,
            close_args,
        ).fetchall()
        daily = conn.execute(
            f"""
            SELECT dt_fechamento AS label, COUNT(*) AS value
            FROM complaints
            {close_where}
            GROUP BY dt_fechamento
            ORDER BY dt_fechamento ASC
            LIMIT 120
            """,
            close_args,
        ).fetchall()
        status_by_analyst = conn.execute(
            f"""
            SELECT
              analista_sac AS analista,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              SUM(CASE WHEN UPPER(validacao) = 'IMPROCEDENTE' THEN 1 ELSE 0 END) AS improcedentes,
              SUM(CASE WHEN validacao = 'Sem validacao' OR validacao = 'Sem validação' OR validacao IS NULL THEN 1 ELSE 0 END) AS sem_status,
              SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END) AS fechadas,
              COUNT(*) AS total
            FROM complaints
            {where}
            GROUP BY analista_sac
            ORDER BY total DESC, analista_sac ASC
            LIMIT 30
            """,
            args,
        ).fetchall()
        closed_by_analyst = conn.execute(
            f"""
            SELECT analista_sac AS label, COUNT(*) AS value
            FROM complaints
            {close_where}
            GROUP BY analista_sac
            ORDER BY value DESC, analista_sac ASC
            LIMIT 15
            """,
            close_args,
        ).fetchall()
        closed_daily = conn.execute(
            f"""
            SELECT dt_fechamento AS label, COUNT(*) AS value
            FROM complaints
            {close_where}
            GROUP BY dt_fechamento
            ORDER BY dt_fechamento ASC
            LIMIT 120
            """,
            close_args,
        ).fetchall()
        daily_by_analyst = conn.execute(
            f"""
            SELECT dt_fechamento AS data, analista_sac AS analista, COUNT(*) AS total
            FROM complaints
            {close_where}
            GROUP BY dt_fechamento, analista_sac
            ORDER BY dt_fechamento DESC, total DESC, analista_sac ASC
            LIMIT 120
            """,
            close_args,
        ).fetchall()
        goal_by_analyst = conn.execute(
            f"""
            SELECT
              analista_sac AS analista,
              COUNT(*) AS fechadas,
              COUNT(DISTINCT dt_fechamento) AS dias_fechamento,
              COUNT(DISTINCT dt_fechamento) * {DAILY_CLOSED_GOAL} AS meta,
              ROUND(COUNT(*) * 100.0 / NULLIF(COUNT(DISTINCT dt_fechamento) * {DAILY_CLOSED_GOAL}, 0), 1) AS atingimento_meta
            FROM complaints
            {close_where}
            GROUP BY analista_sac
            ORDER BY atingimento_meta DESC, fechadas DESC, analista_sac ASC
            LIMIT 40
            """,
            close_args,
        ).fetchall()
        sla_by_analyst = conn.execute(
            f"""
            SELECT
              analista_sac AS analista,
              COUNT(*) AS fechadas,
              SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) AS dentro_sla,
              ROUND(SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_sla,
              ROUND(AVG(dias_uteis_fechamento), 2) AS media_dias_uteis
            FROM complaints
            {close_where}
            GROUP BY analista_sac
            ORDER BY taxa_sla DESC, fechadas DESC, analista_sac ASC
            LIMIT 40
            """,
            close_args,
        ).fetchall()
        quality_volume = conn.execute(
            f"""
            SELECT
              analista_sac AS analista,
              COUNT(*) AS total,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_procedencia,
              SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END) AS fechadas,
              ROUND(SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END), 0), 1) AS taxa_sla
            FROM complaints
            {where}
            GROUP BY analista_sac
            HAVING total >= 10
            ORDER BY total DESC, taxa_procedencia DESC
            LIMIT 40
            """,
            args,
        ).fetchall()
        open_where = where + (" AND UPPER(status_reclamacao) != 'FECHADA'" if where else " WHERE UPPER(status_reclamacao) != 'FECHADA'")
        backlog_aging = conn.execute(
            f"""
            SELECT faixa AS label, COUNT(*) AS value
            FROM (
              SELECT CASE
                WHEN data_reclamacao IS NULL THEN 'Sem data'
                WHEN julianday(?) - julianday(data_reclamacao) <= 5 THEN '0 a 5 dias'
                WHEN julianday(?) - julianday(data_reclamacao) <= 10 THEN '6 a 10 dias'
                WHEN julianday(?) - julianday(data_reclamacao) <= 20 THEN '11 a 20 dias'
                ELSE 'Mais de 20 dias'
              END AS faixa
              FROM complaints
              {open_where}
            )
            GROUP BY faixa
            ORDER BY CASE faixa
              WHEN '0 a 5 dias' THEN 1
              WHEN '6 a 10 dias' THEN 2
              WHEN '11 a 20 dias' THEN 3
              WHEN 'Mais de 20 dias' THEN 4
              ELSE 5
            END
            """,
            [max_date, max_date, max_date, *args],
        ).fetchall()
        backlog_by_analyst = conn.execute(
            f"""
            SELECT
              analista_sac AS analista,
              COUNT(*) AS backlog,
              SUM(CASE WHEN data_reclamacao IS NOT NULL AND julianday(?) - julianday(data_reclamacao) > {SLA_BUSINESS_DAYS_TARGET} THEN 1 ELSE 0 END) AS vencidas,
              SUM(CASE WHEN data_reclamacao IS NULL OR julianday(?) - julianday(data_reclamacao) <= {SLA_BUSINESS_DAYS_TARGET} THEN 1 ELSE 0 END) AS dentro_prazo
            FROM complaints
            {open_where}
            GROUP BY analista_sac
            ORDER BY backlog DESC, analista_sac ASC
            LIMIT 60
            """,
            [max_date, max_date, *args],
        ).fetchall()
        efficiency = conn.execute(
            f"""
            WITH base AS (
              SELECT
                analista_sac AS analista,
                COUNT(*) AS total,
                SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END) AS fechadas,
                SUM(CASE WHEN UPPER(status_reclamacao) != 'FECHADA' THEN 1 ELSE 0 END) AS backlog,
                SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) AS dentro_sla,
                SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
                ROUND(AVG(CASE WHEN dias_uteis_fechamento IS NOT NULL THEN dias_uteis_fechamento END), 2) AS media_dias_uteis
              FROM complaints
              {where}
              GROUP BY analista_sac
            )
            SELECT
              analista,
              total,
              fechadas,
              backlog,
              dentro_sla,
              ROUND(dentro_sla * 100.0 / NULLIF(fechadas, 0), 1) AS taxa_sla,
              ROUND(procedentes * 100.0 / NULLIF(total, 0), 1) AS taxa_procedencia,
              media_dias_uteis,
              ROUND(
                fechadas * {closed_weight}
                + COALESCE(dentro_sla * 100.0 / NULLIF(fechadas, 0), 0) * {sla_weight}
                - backlog * {backlog_penalty},
                1
              ) AS score_eficiencia
            FROM base
            WHERE total > 0
            ORDER BY score_eficiencia DESC, fechadas DESC, analista ASC
            LIMIT 80
            """,
            args,
        ).fetchall()
        sla_by_client = conn.execute(
            f"""
            SELECT
              cliente,
              COUNT(*) AS fechadas,
              SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) AS dentro_sla,
              ROUND(SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_sla,
              ROUND(AVG(dias_uteis_fechamento), 2) AS media_dias_uteis
            FROM complaints
            {close_where}
            GROUP BY cliente
            HAVING fechadas >= 5
            ORDER BY taxa_sla ASC, fechadas DESC, cliente ASC
            LIMIT 50
            """,
            close_args,
        ).fetchall()
        sla_by_reason = conn.execute(
            f"""
            SELECT
              motivo,
              COUNT(*) AS fechadas,
              SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) AS dentro_sla,
              ROUND(SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_sla,
              ROUND(AVG(dias_uteis_fechamento), 2) AS media_dias_uteis
            FROM complaints
            {close_where}
            GROUP BY motivo
            HAVING fechadas >= 5
            ORDER BY media_dias_uteis DESC, fechadas DESC, motivo ASC
            LIMIT 50
            """,
            close_args,
        ).fetchall()
        concentration_client = conn.execute(
            f"""
            SELECT analista_sac AS analista, cliente AS dimensao, COUNT(*) AS total
            FROM complaints
            {where}
            GROUP BY analista_sac, cliente
            ORDER BY total DESC, analista_sac ASC
            LIMIT 30
            """,
            args,
        ).fetchall()
        concentration_hub = conn.execute(
            f"""
            SELECT analista_sac AS analista, hub AS dimensao, COUNT(*) AS total
            FROM complaints
            {where}
            GROUP BY analista_sac, hub
            ORDER BY total DESC, analista_sac ASC
            LIMIT 30
            """,
            args,
        ).fetchall()
        concentration_reason = conn.execute(
            f"""
            SELECT analista_sac AS analista, motivo AS dimensao, COUNT(*) AS total
            FROM complaints
            {where}
            GROUP BY analista_sac, motivo
            ORDER BY total DESC, analista_sac ASC
            LIMIT 30
            """,
            args,
        ).fetchall()
        detail_rows = conn.execute(
            f"""
            SELECT
              data_reclamacao, dt_fechamento, cliente, cidade, uf, regiao, hub,
              status_reclamacao, validacao, motivo, analista_sac,
              dias_uteis_fechamento, sla_ok
            FROM complaints
            {where}
            ORDER BY COALESCE(dt_fechamento, data_reclamacao) DESC, analista_sac ASC
            LIMIT 200
            """,
            args,
        ).fetchall()
        compare_params = dict(params)
        for key in ["period", "month", "start", "end"]:
            compare_params.pop(key, None)
        compare_params["_max_close_date"] = max_close_date
        compare_where, compare_args = build_close_where(compare_params, closed_only=True)
        comparison = conn.execute(
            f"""
            SELECT
              SUM(CASE WHEN mes_fechamento = strftime('%Y-%m', date(?, 'start of month')) THEN 1 ELSE 0 END) AS mes_atual,
              SUM(CASE WHEN mes_fechamento = strftime('%Y-%m', date(?, 'start of month', '-1 month')) THEN 1 ELSE 0 END) AS mes_anterior
            FROM complaints
            {compare_where}
            """,
            [max_close_date, max_close_date, *compare_args],
        ).fetchone()
        top_backlog = dict(backlog_rows[0]) if backlog_rows else None
        executive_summary = [
            f"Total fechado no recorte: {closed_total} reclamacoes.",
            f"Meta diaria: {DAILY_CLOSED_GOAL} fechadas por analista; atingimento geral de {round((closed_total / goal_expected * 100), 1) if goal_expected else 0}%.",
            f"SLA de ate {SLA_BUSINESS_DAYS_TARGET} dias uteis: {round(((sla['dentro_sla'] or 0) / (sla['fechadas_com_prazo'] or 1)) * 100, 1) if sla else 0}% dentro do prazo.",
            f"Backlog principal: {top_backlog['label']} com {top_backlog['value']} reclamacoes." if top_backlog else "Sem backlog no recorte.",
        ]

        return {
            "kpis": {
                "total": total,
                "analysts": analysts,
                "days": days,
                "months": months,
                "avgPerAnalyst": total / analysts if analysts else 0,
                "avgPerDay": total / days if days else 0,
                "avgPerMonth": total / months if months else 0,
                "procedentes": status["procedentes"] or 0,
                "improcedentes": status["improcedentes"] or 0,
                "semStatus": status["sem_status"] or 0,
                "fechadas": closed_total,
                "dailyGoal": DAILY_CLOSED_GOAL,
                "goalExpected": goal_expected,
                "goalPct": closed_total / goal_expected if goal_expected else 0,
                "slaTarget": SLA_BUSINESS_DAYS_TARGET,
                "slaOk": sla["dentro_sla"] or 0,
                "slaTotal": sla["fechadas_com_prazo"] or 0,
                "slaPct": (sla["dentro_sla"] or 0) / (sla["fechadas_com_prazo"] or 1),
                "avgBusinessDaysToClose": sla["media_dias_uteis"] or 0,
                "backlog": sum(row["value"] for row in backlog_rows),
                "comparison": {
                    "currentMonth": comparison["mes_atual"] or 0,
                    "previousMonth": comparison["mes_anterior"] or 0,
                    "delta": (comparison["mes_atual"] or 0) - (comparison["mes_anterior"] or 0),
                    "deltaPct": ((comparison["mes_atual"] or 0) - (comparison["mes_anterior"] or 0)) / (comparison["mes_anterior"] or 1),
                },
                "leadAnalyst": dict(lead) if lead else None,
            },
            "executiveSummary": executive_summary,
            "scoreConfig": {
                "closedWeight": closed_weight,
                "slaWeight": sla_weight,
                "backlogPenalty": backlog_penalty,
                "formula": (
                    f"Score eficiencia = fechadas x {closed_weight:g} + % SLA x {sla_weight:g} "
                    f"- backlog x {backlog_penalty:g}."
                ),
            },
            "charts": {
                "analystVolume": [{"label": row["analista"], "value": row["total"]} for row in analyst_rows[:15]],
                "monthly": [dict(row) for row in monthly],
                "daily": [dict(row) for row in daily],
                "closedDaily": [dict(row) for row in closed_daily],
                "closedByAnalyst": [dict(row) for row in closed_by_analyst],
                "backlog": [dict(row) for row in backlog_rows],
                "backlogAging": [dict(row) for row in backlog_aging],
                "statusByAnalyst": [dict(row) for row in status_by_analyst],
            },
            "tables": {
                "analysts": [dict(row) for row in analyst_rows],
                "dailyByAnalyst": [dict(row) for row in daily_by_analyst],
                "goalByAnalyst": [dict(row) for row in goal_by_analyst],
                "slaByAnalyst": [dict(row) for row in sla_by_analyst],
                "qualityVolume": [dict(row) for row in quality_volume],
                "efficiency": [dict(row) for row in efficiency],
                "backlogByAnalyst": [dict(row) for row in backlog_by_analyst],
                "slaByClient": [dict(row) for row in sla_by_client],
                "slaByReason": [dict(row) for row in sla_by_reason],
                "concentrationClient": [dict(row) for row in concentration_client],
                "concentrationHub": [dict(row) for row in concentration_hub],
                "concentrationReason": [dict(row) for row in concentration_reason],
                "details": [dict(row) for row in detail_rows],
            },
        }


def api_summary(query):
    init_db()
    params = {key: values[0] for key, values in query.items() if values and values[0]}
    with connect() as conn:
        max_date = conn.execute("SELECT MAX(data_reclamacao) AS value FROM complaints").fetchone()["value"]
        params["_max_date"] = max_date
        all_where, all_args = build_where(params, procedentes_only=False)
        proc_where, proc_args = build_where(params, procedentes_only=True)

        total = conn.execute(f"SELECT COUNT(*) AS value FROM complaints {all_where}", all_args).fetchone()["value"]
        procedentes = conn.execute(f"SELECT COUNT(*) AS value FROM complaints {proc_where}", proc_args).fetchone()["value"]
        cities = conn.execute(f"SELECT COUNT(DISTINCT cidade) AS value FROM complaints {proc_where}", proc_args).fetchone()["value"]
        hubs = conn.execute(f"SELECT COUNT(DISTINCT hub) AS value FROM complaints {proc_where}", proc_args).fetchone()["value"]
        lead = conn.execute(
            f"""
            SELECT cliente AS label, COUNT(*) AS value
            FROM complaints {proc_where}
            GROUP BY cliente
            ORDER BY value DESC, cliente ASC
            LIMIT 1
            """,
            proc_args,
        ).fetchone()
        monthly_where = proc_where + (" AND mes != 'Sem data'" if proc_where else " WHERE mes != 'Sem data'")
        monthly = conn.execute(
            f"""
            SELECT mes AS label, COUNT(*) AS value
            FROM complaints
            {monthly_where}
            GROUP BY mes
            ORDER BY mes ASC
            """,
            proc_args,
        ).fetchall()
        monthly_mix = conn.execute(
            f"""
            SELECT
              mes AS label,
              COUNT(*) AS total,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              SUM(CASE WHEN UPPER(validacao) = 'IMPROCEDENTE' THEN 1 ELSE 0 END) AS improcedentes,
              SUM(CASE WHEN validacao = 'Sem validacao' OR validacao = 'Sem validação' OR validacao IS NULL THEN 1 ELSE 0 END) AS sem_status,
              ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_procedencia
            FROM complaints
            {all_where + (" AND mes != 'Sem data'" if all_where else " WHERE mes != 'Sem data'")}
            GROUP BY mes
            ORDER BY mes ASC
            """,
            all_args,
        ).fetchall()
        validation_mix = conn.execute(
            f"""
            SELECT validacao AS label, COUNT(*) AS value
            FROM complaints
            {all_where}
            GROUP BY validacao
            ORDER BY value DESC, validacao ASC
            """
        , all_args).fetchall()
        status_mix = conn.execute(
            f"""
            SELECT status_reclamacao AS label, COUNT(*) AS value
            FROM complaints
            {all_where}
            GROUP BY status_reclamacao
            ORDER BY value DESC, status_reclamacao ASC
            LIMIT 10
            """
        , all_args).fetchall()
        rate_region = conn.execute(
            f"""
            SELECT regiao AS label, COUNT(*) AS total,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa
            FROM complaints
            {all_where}
            GROUP BY regiao
            HAVING total >= 10
            ORDER BY taxa DESC, total DESC
            LIMIT 15
            """
        , all_args).fetchall()
        rate_hub = conn.execute(
            f"""
            SELECT hub AS label, COUNT(*) AS total,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa
            FROM complaints
            {all_where}
            GROUP BY hub
            HAVING total >= 10
            ORDER BY taxa DESC, total DESC
            LIMIT 20
            """
        , all_args).fetchall()
        rate_client = conn.execute(
            f"""
            SELECT cliente AS label, COUNT(*) AS total,
              SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
              ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa
            FROM complaints
            {all_where}
            GROUP BY cliente
            HAVING total >= 10
            ORDER BY taxa DESC, total DESC
            LIMIT 20
            """
        , all_args).fetchall()
        proc_details = conn.execute(
            f"""
            SELECT data_reclamacao, dt_fechamento, cliente, cidade, uf, regiao, hub, status_reclamacao,
                   validacao, motivo, grupo_motivo, analista_sac, dias_uteis_fechamento, sla_ok
            FROM complaints
            {proc_where}
            ORDER BY data_reclamacao DESC, cliente, hub
            LIMIT 200
            """
        , proc_args).fetchall()
        pareto_reason = conn.execute(
            f"""
            SELECT motivo AS label, COUNT(*) AS value
            FROM complaints
            {proc_where}
            GROUP BY motivo
            ORDER BY value DESC, motivo ASC
            LIMIT 20
            """
        , proc_args).fetchall()
        heatmap_region_reason = conn.execute(
            f"""
            SELECT regiao, motivo, COUNT(*) AS value
            FROM complaints
            {proc_where}
            GROUP BY regiao, motivo
            ORDER BY value DESC, regiao, motivo
            LIMIT 80
            """
        , proc_args).fetchall()
        recurrence_clients = conn.execute(
            f"""
            WITH ordered AS (
              SELECT cliente, data_reclamacao,
                     LAG(data_reclamacao) OVER (PARTITION BY cliente ORDER BY data_reclamacao) AS prev_date
              FROM complaints
              {all_where + (" AND data_reclamacao IS NOT NULL" if all_where else " WHERE data_reclamacao IS NOT NULL")}
            )
            SELECT cliente AS label, COUNT(*) AS reincidencias
            FROM ordered
            WHERE prev_date IS NOT NULL AND julianday(data_reclamacao) - julianday(prev_date) <= 30
            GROUP BY cliente
            ORDER BY reincidencias DESC, cliente ASC
            LIMIT 20
            """
        , all_args).fetchall()
        recurrence_total = len(recurrence_clients)
        provider_critical = conn.execute(
            f"""
            SELECT
              prestador,
              hub,
              cidade,
              COUNT(*) AS procedentes,
              SUM(CASE WHEN sla_ok = 0 AND dias_uteis_fechamento IS NOT NULL THEN 1 ELSE 0 END) AS sla_estourado,
              COUNT(*) + SUM(CASE WHEN sla_ok = 0 AND dias_uteis_fechamento IS NOT NULL THEN 1 ELSE 0 END) AS score
            FROM complaints
            {proc_where}
            GROUP BY prestador, hub, cidade
            HAVING procedentes > 0
            ORDER BY score DESC, procedentes DESC, prestador ASC
            LIMIT 50
            """
        , proc_args).fetchall()
        compare_params = dict(params)
        for key in ["period", "month", "start", "end"]:
            compare_params.pop(key, None)
        compare_params["_max_date"] = max_date
        compare_where, compare_args = build_where(compare_params, procedentes_only=False)
        comparison = conn.execute(
            f"""
            SELECT
              SUM(CASE WHEN mes = strftime('%Y-%m', date(?, 'start of month')) AND procedente = 1 THEN 1 ELSE 0 END) AS mes_atual,
              SUM(CASE WHEN mes = strftime('%Y-%m', date(?, 'start of month', '-1 month')) AND procedente = 1 THEN 1 ELSE 0 END) AS mes_anterior
            FROM complaints
            {compare_where}
            """,
            [max_date, max_date, *compare_args],
        ).fetchone()
        date_bounds = conn.execute(
            "SELECT MIN(data_reclamacao) AS minDate, MAX(data_reclamacao) AS maxDate FROM complaints"
        ).fetchone()
        batches = conn.execute(
            """
            SELECT file_name, imported_at, mode, rows_total, rows_inserted, procedentes, date_min, date_max
            FROM import_batches
            ORDER BY id DESC
            LIMIT 8
            """
        ).fetchall()

        return {
            "kpis": {
                "total": total,
                "procedentes": procedentes,
                "procedenciaPct": procedentes / total if total else 0,
                "improcedentes": sum(row["value"] for row in validation_mix if str(row["label"]).upper() == "IMPROCEDENTE"),
                "semStatus": sum(row["value"] for row in validation_mix if str(row["label"]).lower().startswith("sem")),
                "reincidencias": recurrence_total,
                "reincidenciaPct": recurrence_total / total if total else 0,
                "comparison": {
                    "currentMonth": comparison["mes_atual"] or 0,
                    "previousMonth": comparison["mes_anterior"] or 0,
                    "delta": (comparison["mes_atual"] or 0) - (comparison["mes_anterior"] or 0),
                    "deltaPct": ((comparison["mes_atual"] or 0) - (comparison["mes_anterior"] or 0)) / (comparison["mes_anterior"] or 1),
                },
                "cities": cities,
                "hubs": hubs,
                "leadCliente": dict(lead) if lead else None,
            },
            "dateBounds": dict(date_bounds),
            "charts": {
                "monthly": [dict(row) for row in monthly],
                "regiao": count_by(conn, params, "regiao", 8),
                "hub": count_by(conn, params, "hub", 12),
                "cliente": count_by(conn, params, "cliente", 10),
                "cidade": count_by(conn, params, "cidade", 15),
                "motivo": count_by(conn, params, "motivo", 15),
                "validationMix": [dict(row) for row in validation_mix],
                "statusMix": [dict(row) for row in status_mix],
                "paretoReason": [dict(row) for row in pareto_reason],
                "heatmapRegionReason": [dict(row) for row in heatmap_region_reason],
            },
            "tables": {
                "combo": table_combo(conn, params),
                "groupRegion": table_group_region(conn, params),
                "providers": table_providers(conn, params),
                "monthlyMix": [dict(row) for row in monthly_mix],
                "rateRegion": [dict(row) for row in rate_region],
                "rateHub": [dict(row) for row in rate_hub],
                "rateClient": [dict(row) for row in rate_client],
                "details": [dict(row) for row in proc_details],
                "recurrenceClients": [dict(row) for row in recurrence_clients],
                "providerCritical": [dict(row) for row in provider_critical],
            },
            "imports": [dict(row) for row in batches],
        }


def build_operation_where(params):
    where = []
    args = []
    if params.get("regiao"):
        where.append("regional = ?")
        args.append(params["regiao"])
    if params.get("hub"):
        where.append("hub = ?")
        args.append(params["hub"])
    if params.get("cliente"):
        where.append("cliente = ?")
        args.append(params["cliente"])
    if params.get("cidade"):
        where.append("cidade LIKE ?")
        args.append(f"%{params['cidade']}%")

    add_period_filter(where, args, "data", "mes", params, "_max_operation_date")

    clause = " WHERE " + " AND ".join(where) if where else ""
    return clause, args


def operation_count_by(conn, params, field, value_field="COUNT(*)", limit=12, order_by="value DESC"):
    where, args = build_operation_where(params)
    rows = conn.execute(
        f"""
        SELECT {field} AS label, {value_field} AS value
        FROM operations
        {where}
        GROUP BY {field}
        ORDER BY {order_by}, label ASC
        LIMIT ?
        """,
        [*args, limit],
    ).fetchall()
    return [dict(row) for row in rows]


def operation_performance_table(conn, params, field, limit=40):
    where, args = build_operation_where(params)
    rows = conn.execute(
        f"""
        SELECT
          {field} AS label,
          COUNT(*) AS total,
          SUM(concluida) AS concluidas,
          SUM(cancelada) AS canceladas,
          SUM(frustrada) AS frustradas,
          SUM(reagendada) AS reagendadas,
          SUM(improdutiva) AS improdutivas,
          SUM(perda) AS perdas,
          ROUND(SUM(concluida) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_sucesso,
          ROUND(SUM(perda) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_perda,
          ROUND((SUM(concluida) - SUM(perda)) * 100.0 / NULLIF(COUNT(*), 0), 1) AS score_operacional
        FROM operations
        {where}
        GROUP BY {field}
        HAVING total > 0
        ORDER BY taxa_perda DESC, total DESC, label ASC
        LIMIT ?
        """,
        [*args, limit],
    ).fetchall()
    return [dict(row) for row in rows]


def api_operation(query):
    init_db()
    params = {key: values[0] for key, values in query.items() if values and values[0]}
    with connect() as conn:
        max_date = conn.execute("SELECT MAX(data) AS value FROM operations").fetchone()["value"]
        params["_max_operation_date"] = max_date
        where, args = build_operation_where(params)
        totals = conn.execute(
            f"""
            SELECT
              COUNT(*) AS total,
              SUM(concluida) AS concluidas,
              SUM(cancelada) AS canceladas,
              SUM(frustrada) AS frustradas,
              SUM(reagendada) AS reagendadas,
              SUM(improdutiva) AS improdutivas,
              SUM(perda) AS perdas
            FROM operations
            {where}
            """,
            args,
        ).fetchone()
        total = totals["total"] or 0
        concluidas = totals["concluidas"] or 0
        perdas = totals["perdas"] or 0
        monthly = conn.execute(
            f"""
            SELECT mes AS label, COUNT(*) AS total, SUM(concluida) AS concluidas, SUM(perda) AS perdas,
                   ROUND(SUM(concluida) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_sucesso
            FROM operations
            {where + (" AND mes != 'Sem data'" if where else " WHERE mes != 'Sem data'")}
            GROUP BY mes
            ORDER BY mes ASC
            """,
            args,
        ).fetchall()
        compare_params = dict(params)
        for key in ["period", "month", "start", "end"]:
            compare_params.pop(key, None)
        compare_params["_max_operation_date"] = max_date
        compare_where, compare_args = build_operation_where(compare_params)
        comparison = conn.execute(
            f"""
            SELECT
              SUM(CASE WHEN mes = strftime('%Y-%m', date(?, 'start of month')) THEN concluida ELSE 0 END) AS mes_atual,
              SUM(CASE WHEN mes = strftime('%Y-%m', date(?, 'start of month', '-1 month')) THEN concluida ELSE 0 END) AS mes_anterior
            FROM operations
            {compare_where}
            """,
            [max_date, max_date, *compare_args],
        ).fetchone()
        hub_table = operation_performance_table(conn, params, "hub", 40)
        regional_table = operation_performance_table(conn, params, "regional", 40)
        resource_table = operation_performance_table(conn, params, "recurso", 80)
        city_table = operation_performance_table(conn, params, "cidade", 60)
        service_table = operation_performance_table(conn, params, "tipo_atividade", 40)
        reason_rows = conn.execute(
            f"""
            SELECT motivo_nao_realizada AS motivo, COUNT(*) AS total,
                   SUM(CASE WHEN perda = 1 THEN 1 ELSE 0 END) AS perdas
            FROM operations
            {where + (" AND perda = 1" if where else " WHERE perda = 1")}
            GROUP BY motivo_nao_realizada
            ORDER BY perdas DESC, motivo ASC
            LIMIT 30
            """,
            args,
        ).fetchall()
        top_reason = reason_rows[0] if reason_rows else None
        worst_hub = hub_table[0] if hub_table else None
        critical_regional = regional_table[0] if regional_table else None
        critical_resource = resource_table[0] if resource_table else None
        best_resources = sorted(resource_table, key=lambda row: (-row["taxa_sucesso"], -row["total"], row["label"]))[:20]
        capacity_rows = conn.execute(
            f"""
            SELECT recurso AS label, COUNT(*) AS atendimentos, COUNT(DISTINCT data) AS dias_ativos,
                   ROUND(COUNT(*) * 1.0 / NULLIF(COUNT(DISTINCT data), 0), 2) AS media_dia
            FROM operations
            {where}
            GROUP BY recurso
            ORDER BY atendimentos DESC, label ASC
            LIMIT 60
            """,
            args,
        ).fetchall()
        imports = conn.execute(
            """
            SELECT file_name, imported_at, mode, rows_total, rows_inserted, concluidas, date_min, date_max
            FROM operation_batches
            ORDER BY id DESC
            LIMIT 8
            """
        ).fetchall()
        date_bounds = conn.execute(
            "SELECT MIN(data) AS minDate, MAX(data) AS maxDate FROM operations"
        ).fetchone()
        insights = [
            f"Taxa de sucesso operacional: {round(pct_value(concluidas, total) * 100, 1)}% ({concluidas} concluidas de {total} assistencias).",
            f"Taxa de perda operacional: {round(pct_value(perdas, total) * 100, 1)}% considerando canceladas, frustradas, reagendadas e improdutivas.",
        ]
        if worst_hub:
            insights.append(f"HUB com maior taxa de perda no recorte: {worst_hub['label']} ({worst_hub['taxa_perda']}%).")
        if top_reason:
            insights.append(f"Principal motivo de nao realizacao: {top_reason['motivo']} ({top_reason['perdas']} perdas).")
        return {
            "kpis": {
                "total": total,
                "concluidas": concluidas,
                "canceladas": totals["canceladas"] or 0,
                "frustradas": totals["frustradas"] or 0,
                "reagendadas": totals["reagendadas"] or 0,
                "improdutivas": totals["improdutivas"] or 0,
                "perdas": perdas,
                "successPct": pct_value(concluidas, total),
                "lossPct": pct_value(perdas, total),
                "worstHub": worst_hub,
                "criticalRegional": critical_regional,
                "criticalResource": critical_resource,
                "topReason": dict(top_reason) if top_reason else None,
                "comparison": {
                    "currentMonth": comparison["mes_atual"] or 0,
                    "previousMonth": comparison["mes_anterior"] or 0,
                    "delta": (comparison["mes_atual"] or 0) - (comparison["mes_anterior"] or 0),
                    "deltaPct": ((comparison["mes_atual"] or 0) - (comparison["mes_anterior"] or 0)) / (comparison["mes_anterior"] or 1),
                },
            },
            "dateBounds": dict(date_bounds),
            "insights": insights,
            "charts": {
                "monthly": [{"label": row["label"], "value": row["concluidas"] or 0} for row in monthly],
                "lossMonthly": [{"label": row["label"], "value": row["perdas"] or 0} for row in monthly],
                "status": operation_count_by(conn, params, "status", "COUNT(*)", 12),
                "hub": [{"label": row["label"], "value": row["taxa_perda"] or 0} for row in hub_table[:15]],
                "regional": [{"label": row["label"], "value": row["taxa_perda"] or 0} for row in regional_table[:12]],
                "resource": [{"label": row["label"], "value": row["taxa_perda"] or 0} for row in resource_table[:15]],
                "reason": [{"label": row["motivo"], "value": row["perdas"] or 0} for row in reason_rows],
                "cityLoss": [{"label": row["label"], "value": row["taxa_perda"] or 0} for row in city_table[:20]],
            },
            "tables": {
                "hubPerformance": hub_table,
                "regionalPerformance": regional_table,
                "resourceBest": best_resources,
                "resourceCritical": resource_table[:40],
                "reasonPareto": [dict(row) for row in reason_rows],
                "cityRisk": city_table,
                "servicePerformance": service_table,
                "capacity": [dict(row) for row in capacity_rows],
            },
            "imports": [dict(row) for row in imports],
        }


def risk_status(score):
    if score >= 70:
        return "Vermelho"
    if score >= 40:
        return "Amarelo"
    return "Verde"


def city_key(city, uf):
    return f"{slug(city)}|{slug(uf)}"


def load_city_coordinates():
    if not CITIES_PATH.exists():
        return {}
    try:
        rows = json.loads(CITIES_PATH.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    coords = {}
    for row in rows:
        coords[city_key(row.get("cidade"), row.get("uf"))] = row
    return coords


def map_status(row):
    procedencia = row["taxa_procedencia"] or 0
    perda = row["taxa_perda"] or 0
    sucesso = row["taxa_sucesso"] if row["taxa_sucesso"] is not None else 100
    if procedencia >= 65 or perda >= 35 or sucesso < 60:
        return "critico"
    if procedencia >= 45 or perda >= 20 or sucesso < 75:
        return "atencao"
    return "saudavel"


def api_map(query):
    init_db()
    params = {key: values[0] for key, values in query.items() if values and values[0]}
    with connect() as conn:
        max_date = conn.execute("SELECT MAX(data_reclamacao) AS value FROM complaints").fetchone()["value"]
        max_operation_date = conn.execute("SELECT MAX(data) AS value FROM operations").fetchone()["value"]
        params["_max_date"] = max_date
        params["_max_operation_date"] = max_operation_date
        complaint_where, complaint_args = build_where(params, procedentes_only=False)
        operation_where, operation_args = build_operation_where(params)
        complaint_rows = conn.execute(
            f"""
            SELECT cidade, uf,
                   COUNT(*) AS reclamacoes,
                   SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedencias
            FROM complaints
            {complaint_where}
            GROUP BY cidade, uf
            """,
            complaint_args,
        ).fetchall()
        operation_rows = conn.execute(
            f"""
            SELECT cidade, estado AS uf,
                   COUNT(*) AS assistencias,
                   SUM(frustrada) AS frustracoes,
                   SUM(perda) AS perdas,
                   SUM(concluida) AS concluidas
            FROM operations
            {operation_where}
            GROUP BY cidade, estado
            """,
            operation_args,
        ).fetchall()
    merged = {}
    for row in complaint_rows:
        key = city_key(row["cidade"], row["uf"])
        merged.setdefault(key, {"cidade": row["cidade"], "uf": row["uf"], "reclamacoes": 0, "procedencias": 0, "assistencias": 0, "frustracoes": 0, "perdas": 0, "concluidas": 0})
        merged[key]["reclamacoes"] += row["reclamacoes"] or 0
        merged[key]["procedencias"] += row["procedencias"] or 0
    for row in operation_rows:
        key = city_key(row["cidade"], row["uf"])
        merged.setdefault(key, {"cidade": row["cidade"], "uf": row["uf"], "reclamacoes": 0, "procedencias": 0, "assistencias": 0, "frustracoes": 0, "perdas": 0, "concluidas": 0})
        merged[key]["assistencias"] += row["assistencias"] or 0
        merged[key]["frustracoes"] += row["frustracoes"] or 0
        merged[key]["perdas"] += row["perdas"] or 0
        merged[key]["concluidas"] += row["concluidas"] or 0
    coords = load_city_coordinates()
    points = []
    missing = 0
    for key, item in merged.items():
        coord = coords.get(key)
        if not coord:
            missing += 1
            continue
        item["lat"] = coord["lat"]
        item["lon"] = coord["lon"]
        item["taxa_procedencia"] = round(item["procedencias"] * 100.0 / item["reclamacoes"], 1) if item["reclamacoes"] else 0
        item["taxa_perda"] = round(item["perdas"] * 100.0 / item["assistencias"], 1) if item["assistencias"] else 0
        item["taxa_sucesso"] = round(item["concluidas"] * 100.0 / item["assistencias"], 1) if item["assistencias"] else None
        item["status"] = map_status(item)
        item["volume_total"] = item["reclamacoes"] + item["assistencias"]
        points.append(item)
    points.sort(key=lambda row: row["volume_total"], reverse=True)
    return {"points": points, "missingCoordinates": missing, "citiesFile": str(CITIES_PATH)}


def api_providers(query):
    init_db()
    params = {key: values[0] for key, values in query.items() if values and values[0]}
    provider_filter = params.pop("prestador", None)
    with connect() as conn:
        max_date = conn.execute("SELECT MAX(data_reclamacao) AS value FROM complaints").fetchone()["value"]
        params["_max_date"] = max_date
        where, args = build_where(params, procedentes_only=False)
        if provider_filter:
            where = where + (" AND prestador = ?" if where else " WHERE prestador = ?")
            args.append(provider_filter)
        rows = conn.execute(
            f"""
            WITH base AS (
              SELECT
                prestador,
                hub,
                regiao,
                cidade,
                cliente,
                mes,
                procedente,
                validacao,
                sla_ok,
                dias_uteis_fechamento,
                status_reclamacao
              FROM complaints
              {where}
            ),
            agg AS (
              SELECT
                prestador,
                hub,
                regiao,
                COUNT(*) AS reclamacoes,
                SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
                SUM(CASE WHEN UPPER(validacao) = 'IMPROCEDENTE' THEN 1 ELSE 0 END) AS improcedentes,
                SUM(CASE WHEN sla_ok = 0 AND dias_uteis_fechamento IS NOT NULL THEN 1 ELSE 0 END) AS sla_estourado,
                COUNT(DISTINCT cidade) AS cidades,
                ROUND(SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(*), 0), 1) AS taxa_procedencia,
                ROUND(SUM(CASE WHEN sla_ok = 1 THEN 1 ELSE 0 END) * 100.0 / NULLIF(SUM(CASE WHEN UPPER(status_reclamacao) = 'FECHADA' THEN 1 ELSE 0 END), 0), 1) AS taxa_sla,
                ROUND(AVG(CASE WHEN dias_uteis_fechamento IS NOT NULL THEN dias_uteis_fechamento END), 2) AS media_dias_uteis
              FROM base
              GROUP BY prestador, hub, regiao
            ),
            maxes AS (
              SELECT MAX(reclamacoes) AS max_reclamacoes FROM agg
            )
            SELECT
              agg.*,
              ROUND(
                40.0 * reclamacoes / NULLIF(max_reclamacoes, 0) +
                30.0 * COALESCE(taxa_procedencia, 0) / 100.0 +
                20.0 * (1.0 - COALESCE(taxa_sla, 0) / 100.0) +
                10.0 * CASE WHEN cidades > 1 THEN 1 ELSE 0 END
              , 1) AS score_risco
            FROM agg, maxes
            ORDER BY score_risco DESC, reclamacoes DESC, prestador ASC
            LIMIT 120
            """,
            args,
        ).fetchall()
        providers = []
        for row in rows:
            item = dict(row)
            item["status_risco"] = risk_status(item["score_risco"] or 0)
            providers.append(item)
        detail = None
        if provider_filter:
            monthly = conn.execute(
                f"""
                SELECT mes AS label, COUNT(*) AS reclamacoes,
                       SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes
                FROM complaints
                {where + (" AND mes != 'Sem data'" if where else " WHERE mes != 'Sem data'")}
                GROUP BY mes
                ORDER BY mes ASC
                """,
                args,
            ).fetchall()
            cities = conn.execute(
                f"""
                SELECT cidade AS label, COUNT(*) AS value
                FROM complaints
                {where}
                GROUP BY cidade
                ORDER BY value DESC, cidade ASC
                LIMIT 20
                """,
                args,
            ).fetchall()
            records = conn.execute(
                f"""
                SELECT data_reclamacao, dt_fechamento, cliente, cidade, regiao, hub, status_reclamacao,
                       validacao, motivo, analista_sac, dias_uteis_fechamento, sla_ok
                FROM complaints
                {where}
                ORDER BY COALESCE(dt_fechamento, data_reclamacao) DESC
                LIMIT 100
                """,
                args,
            ).fetchall()
            detail = {
                "provider": provider_filter,
                "monthly": [dict(row) for row in monthly],
                "cities": [dict(row) for row in cities],
                "records": [dict(row) for row in records],
            }
        return {"providers": providers, "detail": detail}


def api_data_freshness():
    init_db()
    with connect() as conn:
        row = conn.execute(
            """
            SELECT
              (SELECT MAX(data_reclamacao) FROM complaints) AS max_complaint_date,
              (SELECT MAX(data) FROM operations) AS max_operation_date,
              (SELECT COUNT(*) FROM complaints) AS complaints_total,
              (SELECT COUNT(*) FROM operations) AS operations_total
            """
        ).fetchone()
    complaint_date = row["max_complaint_date"]
    operation_date = row["max_operation_date"]
    lag_days = None
    if complaint_date and operation_date:
        complaint_ts = pd.to_datetime(complaint_date, errors="coerce")
        operation_ts = pd.to_datetime(operation_date, errors="coerce")
        if not pd.isna(complaint_ts) and not pd.isna(operation_ts):
            lag_days = abs((complaint_ts.date() - operation_ts.date()).days)
    return {
        "maxComplaintDate": complaint_date,
        "maxOperationDate": operation_date,
        "lagDays": lag_days,
        "warningDays": DATE_LAG_WARNING_DAYS,
        "complaintsTotal": row["complaints_total"] or 0,
        "operationsTotal": row["operations_total"] or 0,
    }


def api_alerts(query):
    summary = api_summary(query)
    productivity = api_productivity(query)
    operation = api_operation(query)
    providers = api_providers(query)["providers"]
    freshness = api_data_freshness()
    alerts = []
    if (
        freshness["lagDays"] is not None
        and freshness["lagDays"] > freshness["warningDays"]
        and freshness["complaintsTotal"]
        and freshness["operationsTotal"]
    ):
        alerts.append(
            {
                "severity": "Amarelo",
                "area": "Bases",
                "message": (
                    f"Bases fora de sincronia: reclamacoes ate {freshness['maxComplaintDate']} "
                    f"e operacao ate {freshness['maxOperationDate']} "
                    f"({freshness['lagDays']} dias de diferenca)."
                ),
            }
        )
    comp = summary["kpis"]["comparison"]
    if comp["previousMonth"] and comp["deltaPct"] > 0.20:
        alerts.append({"severity": "Vermelho", "area": "Reclamacoes", "message": f"Procedentes cresceram {round(comp['deltaPct'] * 100, 1)}% contra o mes anterior."})
    if summary["kpis"]["procedenciaPct"] > 0.65:
        alerts.append({"severity": "Vermelho", "area": "Qualidade", "message": f"Taxa de procedencia em {round(summary['kpis']['procedenciaPct'] * 100, 1)}%, acima do limite critico."})
    elif summary["kpis"]["procedenciaPct"] > 0.45:
        alerts.append({"severity": "Amarelo", "area": "Qualidade", "message": f"Taxa de procedencia em {round(summary['kpis']['procedenciaPct'] * 100, 1)}%, exige acompanhamento."})
    if productivity["kpis"]["slaPct"] < 0.50:
        alerts.append({"severity": "Vermelho", "area": "SLA", "message": f"SLA SAC em {round(productivity['kpis']['slaPct'] * 100, 1)}%, abaixo da meta de fechamento em ate {SLA_BUSINESS_DAYS_TARGET} dias uteis."})
    elif productivity["kpis"]["slaPct"] < 0.75:
        alerts.append({"severity": "Amarelo", "area": "SLA", "message": f"SLA SAC em {round(productivity['kpis']['slaPct'] * 100, 1)}%, em atencao."})
    if providers:
        top = providers[0]
        if top["score_risco"] >= 70:
            alerts.append({"severity": "Vermelho", "area": "Prestador", "message": f"Prestador {top['prestador']} com score de risco {top['score_risco']}."})
        elif top["score_risco"] >= 40:
            alerts.append({"severity": "Amarelo", "area": "Prestador", "message": f"Prestador {top['prestador']} em atencao, score {top['score_risco']}."})
    worst_hub = operation["kpis"].get("worstHub")
    if worst_hub and (worst_hub.get("taxa_perda") or 0) >= 30:
        alerts.append({"severity": "Vermelho", "area": "HUB", "message": f"HUB {worst_hub['label']} com {worst_hub['taxa_perda']}% de perda operacional."})
    if not alerts:
        alerts.append({"severity": "Verde", "area": "Geral", "message": "Nenhum alerta critico nos filtros atuais."})
    return {"alerts": alerts}


def api_insights(query):
    summary = api_summary(query)
    productivity = api_productivity(query)
    operation = api_operation(query)
    providers = api_providers(query)["providers"]
    insights = []
    top_hub = summary["charts"]["hub"][0] if summary["charts"]["hub"] else None
    top_reason = summary["charts"]["motivo"][0] if summary["charts"]["motivo"] else None
    top_provider = providers[0] if providers else None
    if top_hub and summary["kpis"]["procedentes"]:
        share = top_hub["value"] / summary["kpis"]["procedentes"]
        insights.append(f"O HUB {top_hub['label']} concentra {round(share * 100, 1)}% das procedencias do recorte.")
    if top_reason and summary["kpis"]["procedentes"]:
        share = top_reason["value"] / summary["kpis"]["procedentes"]
        insights.append(f"O motivo {top_reason['label']} representa {round(share * 100, 1)}% das procedencias.")
    if operation["kpis"].get("worstHub"):
        hub = operation["kpis"]["worstHub"]
        insights.append(f"O HUB {hub['label']} tem a maior taxa de perda operacional: {hub['taxa_perda']}%.")
    if productivity["kpis"].get("leadAnalyst"):
        lead = productivity["kpis"]["leadAnalyst"]
        insights.append(f"Maior volume SAC no recorte: {lead['label']} com {lead['value']} reclamacoes tratadas.")
    if top_provider:
        insights.append(f"Prestador com maior risco: {top_provider['prestador']} com score {top_provider['score_risco']} ({top_provider['status_risco']}).")
    return {"insights": insights}


def export_executive_html(query):
    summary = api_summary(query)
    productivity = api_productivity(query)
    operation = api_operation(query)
    alerts = api_alerts(query)["alerts"]
    html = f"""<!doctype html><html><head><meta charset='utf-8'><title>Resumo Executivo MMS</title>
    <style>body{{font-family:Calibri,Arial;margin:32px;color:#171a45}}h1{{color:#2B3594}}.card{{border:1px solid #d8dbee;padding:12px;margin:10px 0;border-left:6px solid #FF8C28}}</style></head><body>
    <h1>Resumo Executivo MMS</h1>
    <div class='card'>Reclamacoes: {summary['kpis']['total']} | Procedentes: {summary['kpis']['procedentes']} | Taxa: {round(summary['kpis']['procedenciaPct']*100,1)}%</div>
    <div class='card'>SLA SAC: {round(productivity['kpis']['slaPct']*100,1)}% | Fechadas: {productivity['kpis']['fechadas']} | Backlog: {productivity['kpis']['backlog']}</div>
    <div class='card'>Operacao: {operation['kpis']['total']} assistencias | Sucesso: {round(operation['kpis']['successPct']*100,1)}% | Perdas: {operation['kpis']['perdas']}</div>
    <h2>Alertas</h2>{''.join(f"<div class='card'><strong>{a['severity']} - {a['area']}</strong><br>{a['message']}</div>" for a in alerts)}
    <script>window.print()</script></body></html>"""
    return html.encode("utf-8-sig")


def api_filters():
    init_db()
    with connect() as conn:
        def values(field):
            return [
                row["value"]
                for row in conn.execute(
                    f"SELECT DISTINCT {field} AS value FROM complaints WHERE {field} IS NOT NULL ORDER BY value"
                )
            ]
        def union_values(complaint_field, operation_field=None):
            operation_field = operation_field or complaint_field
            rows = conn.execute(
                f"""
                SELECT value FROM (
                    SELECT DISTINCT {complaint_field} AS value FROM complaints WHERE {complaint_field} IS NOT NULL
                    UNION
                    SELECT DISTINCT {operation_field} AS value FROM operations WHERE {operation_field} IS NOT NULL
                )
                ORDER BY value
                """
            ).fetchall()
            return [row["value"] for row in rows]

        return {
            "regiao": union_values("regiao", "regional"),
            "hub": union_values("hub"),
            "cliente": union_values("cliente"),
            "analista": values("analista_sac"),
            "months": union_values("mes"),
        }


def export_csv(query):
    init_db()
    params = {key: values[0] for key, values in query.items() if values and values[0]}
    with connect() as conn:
        max_date = conn.execute("SELECT MAX(data_reclamacao) AS value FROM complaints").fetchone()["value"]
        params["_max_date"] = max_date
        where, args = build_where(params, procedentes_only=params.get("scope") == "procedentes")
        rows = conn.execute(
            f"""
            SELECT
              data_reclamacao, dt_fechamento, mes, mes_fechamento, cliente, cidade, uf, regiao, hub,
              servico, status_reclamacao, validacao, motivo, grupo_motivo, prestador, analista_sac,
              procedente, dias_uteis_fechamento, sla_ok
            FROM complaints
            {where}
            ORDER BY data_reclamacao, analista_sac, cliente
            """,
            args,
        ).fetchall()
    output = io.StringIO()
    writer = csv.writer(output, delimiter=";")
    headers = rows[0].keys() if rows else [
        "data_reclamacao", "dt_fechamento", "mes", "mes_fechamento", "cliente", "cidade",
        "uf", "regiao", "hub", "servico", "status_reclamacao", "validacao", "motivo",
        "grupo_motivo", "prestador", "analista_sac", "procedente", "dias_uteis_fechamento", "sla_ok"
    ]
    writer.writerow(headers)
    for row in rows:
        writer.writerow([row[key] for key in headers])
    return output.getvalue().encode("utf-8-sig")


def parse_multipart_form(headers, rfile):
    content_type = headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("Envio invalido. Use o formulario de upload do dashboard.")
    try:
        content_length = int(headers.get("Content-Length", "0") or 0)
    except ValueError as exc:
        raise ValueError("Tamanho do arquivo invalido.") from exc
    if content_length <= 0:
        raise ValueError("Arquivo vazio ou nao recebido.")

    body = rfile.read(content_length)
    raw_message = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8") + body
    message = BytesParser(policy=default).parsebytes(raw_message)
    fields = {}
    files = {}
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        payload = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            files[name] = {"filename": filename, "content": payload}
        else:
            charset = part.get_content_charset() or "utf-8"
            fields[name] = payload.decode(charset, errors="replace")
    return fields, files


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def send_json(self, payload, status=HTTPStatus.OK):
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_file(self, path, content_type):
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_download(self, data, filename, content_type):
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_static(self, parsed_path):
        rel = parsed_path.lstrip("/")
        if rel.startswith("static/"):
            rel = rel[len("static/"):]
        target = (STATIC_DIR / rel).resolve()
        if not str(target).startswith(str(STATIC_DIR.resolve())) or not target.is_file():
            return self.send_error(HTTPStatus.NOT_FOUND)
        content_type = "application/octet-stream"
        if target.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"
        elif target.suffix == ".json":
            content_type = "application/json; charset=utf-8"
        elif target.suffix == ".css":
            content_type = "text/css; charset=utf-8"
        elif target.suffix == ".html":
            content_type = "text/html; charset=utf-8"
        return self.send_file(target, content_type)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            return self.send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
        if parsed.path.startswith("/data/") or parsed.path.startswith("/static/"):
            return self.send_static(parsed.path)
        if parsed.path == "/api/filters":
            return self.send_json(api_filters())
        if parsed.path in ("/api/summary", "/api/reclamacoes"):
            return self.send_json(api_summary(parse_qs(parsed.query)))
        if parsed.path in ("/api/productivity", "/api/produtividade"):
            return self.send_json(api_productivity(parse_qs(parsed.query)))
        if parsed.path in ("/api/operation", "/api/operacao"):
            return self.send_json(api_operation(parse_qs(parsed.query)))
        if parsed.path == "/api/kpis":
            query = parse_qs(parsed.query)
            return self.send_json({
                "reclamacoes": api_summary(query)["kpis"],
                "produtividade": api_productivity(query)["kpis"],
                "operacao": api_operation(query)["kpis"],
            })
        if parsed.path == "/api/prestadores":
            return self.send_json(api_providers(parse_qs(parsed.query)))
        if parsed.path == "/api/alertas":
            return self.send_json(api_alerts(parse_qs(parsed.query)))
        if parsed.path == "/api/insights":
            return self.send_json(api_insights(parse_qs(parsed.query)))
        if parsed.path == "/api/mapa":
            return self.send_json(api_map(parse_qs(parsed.query)))
        if parsed.path in ("/api/clientes", "/api/hubs", "/api/regionais"):
            summary = api_summary(parse_qs(parsed.query))
            key = {"/api/clientes": "rateClient", "/api/hubs": "rateHub", "/api/regionais": "rateRegion"}[parsed.path]
            return self.send_json({"rows": summary["tables"][key]})
        if parsed.path == "/api/export":
            return self.send_download(export_csv(parse_qs(parsed.query)), "dashboard_reclamacoes_filtrado.csv", "text/csv; charset=utf-8")
        if parsed.path == "/api/export/excel":
            return self.send_download(export_csv(parse_qs(parsed.query)), "dashboard_reclamacoes_filtrado.csv", "text/csv; charset=utf-8")
        if parsed.path == "/api/export/pdf":
            return self.send_download(export_executive_html(parse_qs(parsed.query)), "resumo_executivo_mms.html", "text/html; charset=utf-8")
        if parsed.path == "/api/export/ppt":
            return self.send_download(export_executive_html(parse_qs(parsed.query)), "slides_executivos_mms.html", "text/html; charset=utf-8")
        if parsed.path == "/api/health":
            return self.send_json({"ok": True, "db": str(DB_PATH)})
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path != "/api/upload":
            return self.send_error(HTTPStatus.NOT_FOUND)

        try:
            fields, files = parse_multipart_form(self.headers, self.rfile)
        except ValueError as exc:
            return self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)

        if "file" not in files:
            return self.send_json({"ok": False, "error": "Nenhum arquivo enviado."}, HTTPStatus.BAD_REQUEST)

        item = files["file"]
        original_name = Path(item["filename"] or "relatorio.xlsx").name
        if not original_name.lower().endswith((".xlsx", ".xls")):
            return self.send_json({"ok": False, "error": "Envie uma planilha .xlsx ou .xls."}, HTTPStatus.BAD_REQUEST)

        mode = "append" if fields.get("mode") == "append" else "replace"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saved = UPLOAD_DIR / f"{stamp}_{original_name}"
        saved.write_bytes(item["content"])

        try:
            result = import_workbook(saved, mode=mode)
        except Exception as exc:
            return self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
        return self.send_json({"ok": True, "result": result})


def run(host, port):
    init_db()
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Dashboard disponível em http://{host}:{port}")
    server.serve_forever()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8787)
    parser.add_argument("--import-file")
    parser.add_argument("--mode", choices=["replace", "append"], default="replace")
    parser.add_argument("--serve", action="store_true")
    args = parser.parse_args()

    init_db()
    if args.import_file:
        print(json.dumps(import_workbook(args.import_file, mode=args.mode), ensure_ascii=False, indent=2))
    if args.serve:
        run(args.host, args.port)


if __name__ == "__main__":
    main()
