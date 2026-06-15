import os
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
DEFAULT_DATABASE_URL = f"sqlite:///{APP_DIR / 'reclamacoes.db'}"


try:
    from sqlalchemy import Boolean, Column, DateTime, Float, Integer, String, create_engine
    from sqlalchemy.orm import declarative_base, sessionmaker
except ModuleNotFoundError:
    SQLALCHEMY_AVAILABLE = False
    Base = None
    SessionLocal = None

    def create_all_models():
        return None

    def get_session():
        raise RuntimeError("SQLAlchemy nao esta instalado.")
else:
    SQLALCHEMY_AVAILABLE = True
    Base = declarative_base()
    DATABASE_URL = os.environ.get("DATABASE_URL", DEFAULT_DATABASE_URL)
    engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {})
    SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)


    def create_all_models():
        Base.metadata.create_all(bind=engine)


    def get_session():
        return SessionLocal()

    class Complaint(Base):
        __tablename__ = "complaints"
        id = Column(Integer, primary_key=True)
        record_hash = Column(String, unique=True)
        data_reclamacao = Column(String)
        dt_fechamento = Column(String)
        cliente = Column(String)
        cidade = Column(String)
        uf = Column(String)
        regiao = Column(String)
        hub = Column(String)
        analista_sac = Column(String)
        procedente = Column(Integer)
        dias_uteis_fechamento = Column(Integer)
        sla_ok = Column(Integer)

    class Operation(Base):
        __tablename__ = "operations"
        id = Column(Integer, primary_key=True)
        record_hash = Column(String, unique=True)
        data = Column(String)
        cidade = Column(String)
        estado = Column(String)
        regional = Column(String)
        hub = Column(String)
        recurso = Column(String)
        concluida = Column(Integer)
        perda = Column(Integer)

    class ImportBatch(Base):
        __tablename__ = "import_batches"
        id = Column(Integer, primary_key=True)
        file_name = Column(String)
        imported_at = Column(String)
        mode = Column(String)
        rows_total = Column(Integer)
        rows_inserted = Column(Integer)

    class OperationBatch(Base):
        __tablename__ = "operation_batches"
        id = Column(Integer, primary_key=True)
        file_name = Column(String)
        imported_at = Column(String)
        mode = Column(String)
        rows_total = Column(Integer)
        rows_inserted = Column(Integer)

    class GeoCidade(Base):
        __tablename__ = "geocidades"
        id = Column(Integer, primary_key=True)
        codigo_ibge = Column(Integer)
        cidade = Column(String)
        cidade_normalizada = Column(String)
        uf = Column(String)
        cidade_key = Column(String)
        uf_key = Column(String)
        codigo_uf = Column(Integer)
        estado = Column(String)
        regiao = Column(String)
        lat = Column(Float)
        lon = Column(Float)
        fonte = Column(String)
        precisa_coordenada = Column(Integer)
        uf_corrigida_de = Column(String)
        created_at = Column(String)
        updated_at = Column(String)

    class GeoCidadePendente(Base):
        __tablename__ = "geocidades_pendentes"
        id = Column(Integer, primary_key=True)
        cidade = Column(String)
        uf = Column(String)
        uf_original = Column(String)
        cidade_normalizada = Column(String)
        ocorrencias = Column(Integer)
        origem = Column(String)
        reclamacoes = Column(Integer)
        procedentes = Column(Integer)
        assistencias = Column(Integer)
        ignorada = Column(Integer)
        created_at = Column(String)
        updated_at = Column(String)

    class User(Base):
        __tablename__ = "users"
        id = Column(Integer, primary_key=True)
        nome = Column(String)
        email = Column(String, unique=True)
        senha_hash = Column(String)
        perfil = Column(String)
        regional = Column(String)
        analista_sac = Column(String)
        ativo = Column(Boolean, default=True)
        created_at = Column(DateTime)

    class UserProfile(Base):
        __tablename__ = "user_profiles"
        id = Column(Integer, primary_key=True)
        user_id = Column(Integer)
        perfil = Column(String)
        regional = Column(String)
        analista_sac = Column(String)

    class Config(Base):
        __tablename__ = "configs"
        id = Column(Integer, primary_key=True)
        key = Column(String, unique=True)
        value = Column(String)
