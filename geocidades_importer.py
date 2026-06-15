import json
import re
import sqlite3
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "reclamacoes.db"
META_COBERTURA_GEOGRAFICA = 95.0
UF_DESCONHECIDA = {"", "NI", "NA", "N/I", "SEM_UF", "SEM UF", "SEM ESTADO", "NAN", "NONE", "NULL"}
STOPWORDS_CIDADE = {"D", "DE", "DA", "DO", "DAS", "DOS", "E"}
DF_REGIOES_ADMINISTRATIVAS = {
    "BRAZLANDIA", "CEILANDIA", "GAMA", "GUARA", "LAGO NORTE", "LAGO SUL", "PARANOA",
    "PLANALTINA", "RECANTO DAS EMAS", "RIACHO FUNDO", "SAMAMBAIA", "SANTA MARIA",
    "SAO SEBASTIAO", "SOBRADINHO", "TAGUATINGA", "VICENTE PIRES",
}


def normalizar_texto(valor):
    texto = unicodedata.normalize("NFKD", str(valor or "")).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[^A-Za-z0-9 ]+", " ", texto.upper())
    return re.sub(r"\s+", " ", texto).strip()


def slug(valor):
    texto = normalizar_texto(valor).lower()
    return re.sub(r"[^a-z0-9]+", "_", texto).strip("_")


def compactar_cidade(valor):
    tokens = [token for token in normalizar_texto(valor).split() if token not in STOPWORDS_CIDADE]
    return "".join(tokens)


def expandir_abreviacoes(valor):
    texto = normalizar_texto(valor)
    regras = [
        (r"^GOV\b", "GOVERNADOR"),
        (r"^FCO\b", "FRANCISCO"),
        (r"^CNEL\b", "CORONEL"),
        (r"^PRESI", "PRESIDENTE "),
        (r"^CONS\b", "CONSELHEIRO"),
        (r"^ALMTE", "ALMIRANTE "),
        (r"^BALN", "BALNEARIO "),
        (r"^CACH", "CACHOEIRO "),
        (r"^CAMPOS\b", "CAMPOS DOS"),
        (r"^CPOS\b", "CAMPOS DOS"),
        (r"^FAZ", "FAZENDA "),
        (r"^JAB", "JABOATAO DOS "),
        (r"^NSRA\b", "NOSSA SENHORA"),
        (r"^R GDE\b", "RIO GRANDE"),
        (r"^SCR", "SANTA CRUZ "),
        (r"^SMIGUEL\b", "SAO MIGUEL"),
        (r"^STO\b", "SANTO"),
        (r"^STA\b", "SANTA"),
        (r"^STANA\b", "SANTANA"),
        (r"^SJOAO\b", "SAO JOAO"),
        (r"^SJOSE\b", "SAO JOSE"),
        (r"^SLOURENCO\b", "SAO LOURENCO"),
        (r"^SPEDRO\b", "SAO PEDRO"),
        (r"^SGONC", "SAO GONCALO "),
        (r"^SCAETANO\b", "SAO CAETANO"),
        (r"^SBERN", "SAO BERN"),
        (r"^SANTDO\b", "SANTANA DO"),
        (r"^SADESCOBERTO\b", "SANTO ANTONIO DO DESCOBERTO"),
    ]
    for pattern, replacement in regras:
        texto = re.sub(pattern, replacement, texto)
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_uf(valor):
    uf = normalizar_texto(valor)
    if uf in UF_DESCONHECIDA:
        return "NI"
    return uf[:2] if len(uf) == 2 else uf


def connect():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout = 30000")
    return conn


def _columns(conn, table):
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def _add_column(conn, table, column, definition):
    if column not in _columns(conn, table):
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def garantir_tabelas_geocidades(conn=None):
    own_conn = conn is None
    conn = conn or connect()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS geocidades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_ibge INTEGER,
                cidade TEXT NOT NULL,
                cidade_normalizada TEXT,
                uf TEXT NOT NULL,
                cidade_key TEXT NOT NULL,
                uf_key TEXT NOT NULL,
                codigo_uf INTEGER,
                estado TEXT,
                regiao TEXT,
                lat REAL,
                lon REAL,
                fonte TEXT NOT NULL DEFAULT 'base',
                precisa_coordenada INTEGER NOT NULL DEFAULT 1,
                uf_corrigida_de TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(cidade_key, uf_key)
            );

            CREATE TABLE IF NOT EXISTS geocidades_pendentes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cidade TEXT NOT NULL,
                uf TEXT,
                uf_original TEXT,
                cidade_normalizada TEXT NOT NULL,
                ocorrencias INTEGER,
                origem TEXT,
                reclamacoes INTEGER NOT NULL DEFAULT 0,
                procedentes INTEGER NOT NULL DEFAULT 0,
                assistencias INTEGER NOT NULL DEFAULT 0,
                ignorada INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(cidade_normalizada, uf, origem)
            );
            """
        )
        for column, definition in {
            "codigo_ibge": "INTEGER",
            "cidade_normalizada": "TEXT",
            "codigo_uf": "INTEGER",
            "estado": "TEXT",
            "regiao": "TEXT",
        }.items():
            _add_column(conn, "geocidades", column, definition)
        for column, definition in {
            "reclamacoes": "INTEGER NOT NULL DEFAULT 0",
            "procedentes": "INTEGER NOT NULL DEFAULT 0",
            "assistencias": "INTEGER NOT NULL DEFAULT 0",
            "ignorada": "INTEGER NOT NULL DEFAULT 0",
            "updated_at": "TEXT",
        }.items():
            _add_column(conn, "geocidades_pendentes", column, definition)
        conn.executescript(
            """
            CREATE INDEX IF NOT EXISTS idx_geocidades_keys ON geocidades(cidade_key, uf_key);
            CREATE INDEX IF NOT EXISTS idx_geocidades_lookup ON geocidades(cidade_normalizada, uf);
            CREATE INDEX IF NOT EXISTS idx_geocidades_missing ON geocidades(precisa_coordenada);
            CREATE INDEX IF NOT EXISTS idx_geocidades_pendentes_lookup
            ON geocidades_pendentes(cidade_normalizada, uf, ignorada);
            """
        )
        if own_conn:
            conn.commit()
    finally:
        if own_conn:
            conn.close()


def _load_json(path):
    with path.open("r", encoding="utf-8-sig") as file:
        data = json.load(file)
    if isinstance(data, dict):
        for key in ("estados", "municipios", "data", "items"):
            if isinstance(data.get(key), list):
                return data[key]
    return data if isinstance(data, list) else []


def _resolve_data_paths():
    candidates = [
        (APP_DIR / "municipios.json", APP_DIR / "estados.json"),
        (APP_DIR / "static" / "data" / "municipios.json", APP_DIR / "static" / "data" / "estados.json"),
        (APP_DIR / "municipios-brasileiros-main" / "json" / "municipios.json",
         APP_DIR / "municipios-brasileiros-main" / "json" / "estados.json"),
    ]
    for municipios_path, estados_path in candidates:
        if municipios_path.exists() and estados_path.exists():
            return municipios_path, estados_path
    expected = "\n".join(f"- {m}\n- {e}" for m, e in candidates)
    raise FileNotFoundError(
        "Arquivos municipios.json e estados.json nao encontrados. Coloque a base em uma destas opcoes:\n"
        f"{expected}"
    )


def _to_number(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _nested_uf_code(row):
    current = row
    for key in ("microrregiao", "mesorregiao", "UF"):
        if not isinstance(current, dict):
            return None
        current = current.get(key)
    if isinstance(current, dict):
        return current.get("id") or current.get("codigo_uf")
    return None


def carregar_estados(estados_path=None):
    if estados_path is None:
        _, estados_path = _resolve_data_paths()
    estados = {}
    for row in _load_json(estados_path):
        codigo_uf = row.get("codigo_uf") or row.get("id") or row.get("codigo")
        uf = row.get("uf") or row.get("sigla")
        regiao = row.get("regiao")
        if isinstance(regiao, dict):
            regiao = regiao.get("nome") or regiao.get("sigla")
        if codigo_uf is None or not uf:
            continue
        estados[int(codigo_uf)] = {
            "codigo_uf": int(codigo_uf),
            "uf": normalizar_uf(uf),
            "estado": row.get("nome") or row.get("estado"),
            "regiao": regiao,
        }
    return estados


def importar_municipios(conn=None):
    own_conn = conn is None
    conn = conn or connect()
    try:
        garantir_tabelas_geocidades(conn)
        municipios_path, estados_path = _resolve_data_paths()
        estados = carregar_estados(estados_path)
        municipios = _load_json(municipios_path)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        imported = 0
        skipped = 0
        for row in municipios:
            nome = row.get("nome") or row.get("cidade") or row.get("municipio")
            codigo_uf = row.get("codigo_uf") or row.get("uf_id") or _nested_uf_code(row)
            estado = estados.get(int(codigo_uf)) if codigo_uf not in (None, "") else None
            uf = normalizar_uf(row.get("uf") or row.get("sigla_uf") or (estado or {}).get("uf"))
            lat = _to_number(row.get("latitude") or row.get("lat"))
            lon = _to_number(row.get("longitude") or row.get("lon") or row.get("lng"))
            if not nome or uf == "NI" or lat is None or lon is None:
                skipped += 1
                continue
            cidade_normalizada = normalizar_texto(nome)
            cidade_key = slug(nome)
            uf_key = slug(uf)
            codigo_ibge = row.get("codigo_ibge") or row.get("id")
            conn.execute(
                """
                INSERT INTO geocidades (
                    codigo_ibge, cidade, cidade_normalizada, uf, cidade_key, uf_key,
                    codigo_uf, estado, regiao, lat, lon, fonte, precisa_coordenada,
                    created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)
                ON CONFLICT(cidade_key, uf_key) DO UPDATE SET
                    codigo_ibge = COALESCE(excluded.codigo_ibge, geocidades.codigo_ibge),
                    cidade = excluded.cidade,
                    cidade_normalizada = excluded.cidade_normalizada,
                    uf = excluded.uf,
                    codigo_uf = COALESCE(excluded.codigo_uf, geocidades.codigo_uf),
                    estado = COALESCE(excluded.estado, geocidades.estado),
                    regiao = COALESCE(excluded.regiao, geocidades.regiao),
                    lat = excluded.lat,
                    lon = excluded.lon,
                    fonte = excluded.fonte,
                    precisa_coordenada = 0,
                    updated_at = excluded.updated_at
                """,
                (
                    codigo_ibge,
                    cidade_normalizada,
                    cidade_normalizada,
                    uf,
                    cidade_key,
                    uf_key,
                    int(codigo_uf) if codigo_uf not in (None, "") else None,
                    (estado or {}).get("estado"),
                    (estado or {}).get("regiao"),
                    lat,
                    lon,
                    "municipios-brasileiros",
                    now,
                    now,
                ),
            )
            imported += 1
        if own_conn:
            conn.commit()
        return {
            "municipios_importados": imported,
            "municipios_ignorados": skipped,
            "municipios_arquivo": str(municipios_path),
            "estados_arquivo": str(estados_path),
        }
    finally:
        if own_conn:
            conn.close()


def coletar_cidades_da_base(conn):
    rows = []
    rows.extend(
        dict(row)
        for row in conn.execute(
            """
            SELECT cidade, uf, COUNT(*) AS reclamacoes,
                   SUM(CASE WHEN procedente = 1 THEN 1 ELSE 0 END) AS procedentes,
                   0 AS assistencias, 'complaints' AS origem
            FROM complaints
            WHERE cidade IS NOT NULL AND cidade != 'Sem cidade'
            GROUP BY cidade, uf
            """
        ).fetchall()
    )
    rows.extend(
        dict(row)
        for row in conn.execute(
            """
            SELECT cidade, estado AS uf, 0 AS reclamacoes, 0 AS procedentes,
                   COUNT(*) AS assistencias, 'operations' AS origem
            FROM operations
            WHERE cidade IS NOT NULL AND cidade != 'Sem cidade'
            GROUP BY cidade, estado
            """
        ).fetchall()
    )
    nps_exists = conn.execute(
        "SELECT COUNT(*) AS total FROM sqlite_master WHERE type='table' AND name='nps'"
    ).fetchone()["total"]
    if nps_exists:
        try:
            rows.extend(
                dict(row)
                for row in conn.execute(
                    """
                    SELECT cidade, uf, 0 AS reclamacoes, 0 AS procedentes,
                           COUNT(*) AS assistencias, 'nps' AS origem
                    FROM nps
                    WHERE cidade IS NOT NULL
                    GROUP BY cidade, uf
                    """
                ).fetchall()
            )
        except sqlite3.Error:
            pass
    return rows


def _geocidade_maps(conn):
    rows = conn.execute(
        """
        SELECT cidade, cidade_key, uf, uf_key, lat, lon, fonte
        FROM geocidades
        WHERE lat IS NOT NULL AND lon IS NOT NULL
        """
    ).fetchall()
    by_pair = {}
    by_city = {}
    for row in rows:
        item = dict(row)
        by_pair[(item["cidade_key"], item["uf_key"])] = item
        by_city.setdefault(item["cidade_key"], []).append(item)
    by_uf = {}
    for item in by_pair.values():
        by_uf.setdefault(item["uf"], []).append(item)
    return by_pair, by_city, by_uf


def _similaridade_cidade(origem, destino):
    origem_expandida = expandir_abreviacoes(origem)
    origem_compacta = compactar_cidade(origem_expandida)
    destino_compacta = compactar_cidade(destino)
    if not origem_compacta or not destino_compacta:
        return 0
    if origem_compacta == destino_compacta:
        return 1.0
    if origem_compacta in destino_compacta or destino_compacta in origem_compacta:
        menor = min(len(origem_compacta), len(destino_compacta))
        maior = max(len(origem_compacta), len(destino_compacta))
        return 0.92 + (menor / maior) * 0.08
    return SequenceMatcher(None, origem_compacta, destino_compacta).ratio()


def _match_geocidade_alias(cidade, uf, by_pair, by_uf):
    uf = normalizar_uf(uf)
    cidade_norm = normalizar_texto(cidade)
    if uf == "DF" and cidade_norm in DF_REGIOES_ADMINISTRATIVAS:
        return by_pair.get((slug("Brasilia"), slug("DF")))
    if uf == "NI":
        return None
    candidates = by_uf.get(uf, [])
    scored = []
    for candidate in candidates:
        score = _similaridade_cidade(cidade, candidate["cidade"])
        if score >= 0.86:
            scored.append((score, candidate))
    if not scored:
        return None
    scored.sort(key=lambda item: item[0], reverse=True)
    best_score, best = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0
    if best_score >= 0.92 or best_score - second_score >= 0.05:
        return best
    return None


def _upsert_alias_geocidade(conn, cidade_alias, uf_alias, matched):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cidade_normalizada = normalizar_texto(cidade_alias)
    uf = normalizar_uf(uf_alias or matched["uf"])
    conn.execute(
        """
        INSERT INTO geocidades (
            cidade, cidade_normalizada, uf, cidade_key, uf_key, lat, lon, fonte,
            precisa_coordenada, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)
        ON CONFLICT(cidade_key, uf_key) DO UPDATE SET
            cidade = excluded.cidade,
            cidade_normalizada = excluded.cidade_normalizada,
            uf = excluded.uf,
            lat = excluded.lat,
            lon = excluded.lon,
            fonte = excluded.fonte,
            precisa_coordenada = 0,
            updated_at = excluded.updated_at
        """,
        (
            cidade_normalizada,
            cidade_normalizada,
            uf,
            slug(cidade_alias),
            slug(uf),
            matched["lat"],
            matched["lon"],
            f"alias-auto:{matched['cidade']}/{matched['uf']}",
            now,
            now,
        ),
    )


def _resolver_uf(row, by_pair, by_city, by_uf):
    cidade_key = slug(row.get("cidade"))
    uf_original = row.get("uf")
    uf = normalizar_uf(uf_original)
    direct = by_pair.get((cidade_key, slug(uf)))
    if direct:
        return direct["uf"], uf_original, direct
    candidates = by_city.get(cidade_key, [])
    if (uf == "NI" or not direct) and len(candidates) == 1:
        return candidates[0]["uf"], uf_original, candidates[0]
    alias = _match_geocidade_alias(row.get("cidade"), uf, by_pair, by_uf)
    if alias:
        return alias["uf"], uf_original, alias
    return uf, uf_original, None


def _upsert_pendente(conn, row, uf, uf_original):
    cidade_normalizada = normalizar_texto(row.get("cidade"))
    if not cidade_normalizada:
        return
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    reclamacoes = int(row.get("reclamacoes") or 0)
    procedentes = int(row.get("procedentes") or 0)
    assistencias = int(row.get("assistencias") or 0)
    ocorrencias = reclamacoes + assistencias
    conn.execute(
        """
        INSERT INTO geocidades_pendentes (
            cidade, uf, uf_original, cidade_normalizada, ocorrencias, origem,
            reclamacoes, procedentes, assistencias, ignorada, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)
        ON CONFLICT(cidade_normalizada, uf, origem) DO UPDATE SET
            cidade = excluded.cidade,
            uf_original = excluded.uf_original,
            ocorrencias = excluded.ocorrencias,
            reclamacoes = excluded.reclamacoes,
            procedentes = excluded.procedentes,
            assistencias = excluded.assistencias,
            updated_at = excluded.updated_at
        """,
        (
            cidade_normalizada,
            uf,
            uf_original,
            cidade_normalizada,
            ocorrencias,
            row.get("origem"),
            reclamacoes,
            procedentes,
            assistencias,
            now,
            now,
        ),
    )


def recalcular_pendencias(conn=None):
    own_conn = conn is None
    conn = conn or connect()
    try:
        garantir_tabelas_geocidades(conn)
        conn.execute("DELETE FROM geocidades_pendentes WHERE ignorada = 0")
        by_pair, by_city, by_uf = _geocidade_maps(conn)
        mapped_keys = set()
        pending_keys = set()
        processed = 0
        corrected_uf = 0
        aliases = 0
        for row in coletar_cidades_da_base(conn):
            fixed_uf, original_uf, geocity = _resolver_uf(row, by_pair, by_city, by_uf)
            key = (slug(row.get("cidade")), slug(fixed_uf))
            processed += 1
            if geocity:
                if (slug(row.get("cidade")), slug(fixed_uf)) not in by_pair:
                    _upsert_alias_geocidade(conn, row.get("cidade"), fixed_uf, geocity)
                    aliases += 1
                mapped_keys.add(key)
                if normalizar_uf(original_uf) != normalizar_uf(fixed_uf):
                    corrected_uf += 1
                continue
            pending_keys.add(key)
            _upsert_pendente(conn, row, fixed_uf, original_uf)
        total = len(mapped_keys | pending_keys)
        mapped = len(mapped_keys)
        pending = len(pending_keys)
        coverage = round((mapped * 100.0 / total), 1) if total else 100.0
        payload = {
            "processadas": processed,
            "mapeadas": mapped,
            "pendentes": pending,
            "cobertura": round(coverage / 100.0, 4),
            "coberturaPercentual": coverage,
            "uf_corrigidas": corrected_uf,
            "aliases_criados": aliases,
            "cobertura_geografica": {
                "cidades_cadastradas": total,
                "cidades_mapeadas": mapped,
                "sem_coordenadas": pending,
                "cobertura_geografica": coverage,
                "meta_cobertura": META_COBERTURA_GEOGRAFICA,
                "aviso": coverage < META_COBERTURA_GEOGRAFICA,
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
        }
        if own_conn:
            conn.commit()
        return payload
    finally:
        if own_conn:
            conn.close()


def atualizar_geocidades():
    with connect() as conn:
        garantir_tabelas_geocidades(conn)
        importacao = importar_municipios(conn)
        cobertura = recalcular_pendencias(conn)
        conn.commit()
    return {"importacao": importacao, **cobertura}


if __name__ == "__main__":
    print(json.dumps(atualizar_geocidades(), ensure_ascii=False, indent=2))
