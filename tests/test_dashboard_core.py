import tempfile
import unittest
from pathlib import Path

import pandas as pd

import sys

APP_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(APP_DIR))

import app
import auto_import
import auth
import cache_manager
import config_manager
import geocidades_importer


class DashboardCoreTests(unittest.TestCase):
    def test_business_days_between_counts_weekdays_only(self):
        self.assertEqual(app.business_days_between("2026-06-12", "2026-06-15"), 2)

    def test_operation_flags_mark_operational_losses(self):
        flags = app.operation_flags("Cancelada", "Cliente ausente")
        self.assertEqual(flags[-1], 1)

    def test_config_targets_are_loaded(self):
        self.assertEqual(app.DAILY_CLOSED_GOAL, 10)
        self.assertEqual(app.SLA_BUSINESS_DAYS_TARGET, 5)

    def test_excel_sheet_detection_prefers_operation_sheet(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "teste.xlsx"
            with pd.ExcelWriter(path) as writer:
                pd.DataFrame({"qualquer": [1]}).to_excel(writer, sheet_name="Resumo", index=False)
                pd.DataFrame(
                    {
                        "numero_da_assistencia": [123],
                        "tipo_atividade": ["Entrega"],
                        "recurso": ["Analista"],
                        "status": ["Concluida"],
                        "data": ["2026-06-01"],
                    }
                ).to_excel(writer, sheet_name="Operacao", index=False)

            _, best = app.read_best_excel_sheet(path)
            self.assertEqual(best["sheet"], "Operacao")
            self.assertEqual(best["type"], "operation")

    def test_config_manager_reads_config(self):
        config = config_manager.get_config()
        self.assertIn("daily_closed_goal", config)
        self.assertIn("risk_score", config)

    def test_score_endpoint_shape(self):
        payload = app.api_score_dimension({}, "hubs")
        self.assertIn("rows", payload)
        if payload["rows"]:
            self.assertIn("score_risco", payload["rows"][0])
            self.assertIn("status_risco", payload["rows"][0])

    def test_auto_import_status_shape(self):
        status = auto_import.auto_import_status()
        self.assertIn("pending", status)
        self.assertIn("folder", status)
        self.assertIn("schedulerAvailable", status)

    def test_dependency_registry_contains_install_requirements(self):
        for package in [
            "pandas",
            "numpy",
            "openpyxl",
            "sqlalchemy",
            "cachetools",
            "apscheduler",
            "python-pptx",
            "reportlab",
            "jinja2",
        ]:
            self.assertIn(package, app.REQUIRED_DEPENDENCIES)

    def test_auth_open_mode_returns_admin_user(self):
        auth.AUTH_ENABLED = False
        user = auth.current_user({})
        self.assertEqual(user["perfil"], "admin")

    def test_auth_enabled_requires_session(self):
        auth.AUTH_ENABLED = True
        self.assertIsNone(auth.current_user({}))
        auth.AUTH_ENABLED = False

    def test_auth_login_creates_persistent_session(self):
        app.init_db()
        auth.AUTH_ENABLED = True
        with app.connect() as conn:
            result = auth.authenticate(conn, "admin@mms.local", "admin123")
            self.assertIsNotNone(result)
            token, user = result
        session_user = auth.current_user({"Cookie": f"mms_session={token}"})
        self.assertEqual(session_user["perfil"], "admin")
        auth.logout(token)
        self.assertIsNone(auth.current_user({"Cookie": f"mms_session={token}"}))
        auth.AUTH_ENABLED = False

    def test_cache_manager_reuses_builder_result(self):
        calls = {"total": 0}

        def builder():
            calls["total"] += 1
            return {"value": calls["total"]}

        cache_manager.clear_api_cache()
        first = cache_manager.get_cached("teste", {}, builder)
        second = cache_manager.get_cached("teste", {}, builder)
        self.assertEqual(first, second)
        self.assertEqual(calls["total"], 1)

    def test_alerts_include_actionable_fields(self):
        payload = app.api_alerts({})
        self.assertIn("alerts", payload)
        self.assertTrue(payload["alerts"])
        first = payload["alerts"][0]
        for field in ["titulo", "recomendacao", "acao", "prioridade", "impacto", "urgencia"]:
            self.assertIn(field, first)

    def test_insights_include_executive_cards(self):
        payload = app.api_insights({})
        self.assertIn("insights", payload)
        self.assertIn("cards", payload)
        if payload["cards"]:
            first = payload["cards"][0]
            for field in ["tipo", "titulo", "descricao", "acao", "prioridade"]:
                self.assertIn(field, first)

    def test_map_exposes_geographic_coverage(self):
        payload = app.api_map({})
        self.assertIn("geoCoverage", payload)
        coverage = payload["geoCoverage"]
        for field in ["mappedCities", "missingCoordinates", "totalCities", "coveragePct", "warning", "targetPct"]:
            self.assertIn(field, coverage)
        self.assertEqual(payload["source"], "geocidades")

    def test_geocidades_table_is_synchronized(self):
        app.init_db()
        with app.connect() as conn:
            app.sync_geocidades(conn)
            total = conn.execute("SELECT COUNT(*) AS total FROM geocidades").fetchone()["total"]
        self.assertGreater(total, 0)

    def test_geo_update_and_pending_shape(self):
        result = app.atualizar_geocidades(force=True)
        self.assertIn("processadas", result)
        self.assertIn("cobertura", result)
        pending = app.api_geocidades_pendentes()
        self.assertIn("pendentes", pending)
        if pending["pendentes"]:
            first = pending["pendentes"][0]
            for field in ["cidade", "uf", "origem", "ocorrencias", "reclamacoes", "procedentes", "assistencias"]:
                self.assertIn(field, first)

    def test_geo_manual_coordinate_validation(self):
        with self.assertRaises(ValueError):
            app.api_save_geocidade({"cidade": "Teste", "uf": "SP", "lat": "abc", "lon": -46})

    def test_geocidades_importer_reads_complete_local_base(self):
        municipios_path, estados_path = geocidades_importer._resolve_data_paths()
        self.assertTrue(municipios_path.exists())
        self.assertTrue(estados_path.exists())
        self.assertGreaterEqual(len(geocidades_importer._load_json(municipios_path)), 5570)

    def test_geocidades_importer_schema_columns(self):
        geocidades_importer.garantir_tabelas_geocidades()
        with app.connect() as conn:
            columns = {row["name"] for row in conn.execute("PRAGMA table_info(geocidades)").fetchall()}
        for column in ["codigo_ibge", "cidade_normalizada", "codigo_uf", "estado", "regiao", "lat", "lon"]:
            self.assertIn(column, columns)


if __name__ == "__main__":
    unittest.main()
