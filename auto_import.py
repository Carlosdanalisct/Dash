from pathlib import Path
import shutil


try:
    from apscheduler.schedulers.background import BackgroundScheduler
except ModuleNotFoundError:
    BackgroundScheduler = None


APP_DIR = Path(__file__).resolve().parent
AUTO_IMPORT_DIR = APP_DIR / "auto_import"
PROCESSED_DIR = AUTO_IMPORT_DIR / "processados"
ERROR_DIR = AUTO_IMPORT_DIR / "erro"
SCHEDULER = None
LAST_RUN = None
LAST_RESULT = {"imported": 0, "errors": 0, "message": "Aguardando execucao."}


def ensure_auto_import_dirs():
    for path in [AUTO_IMPORT_DIR, PROCESSED_DIR, ERROR_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def importar_planilhas():
    global LAST_RUN, LAST_RESULT
    from datetime import datetime
    from app import import_workbook
    from cache_manager import clear_api_cache

    ensure_auto_import_dirs()
    imported = 0
    errors = 0
    for path in sorted(AUTO_IMPORT_DIR.glob("*.xls*")):
        target_dir = PROCESSED_DIR
        try:
            import_workbook(path, mode="append")
            imported += 1
            clear_api_cache()
        except Exception:
            errors += 1
            target_dir = ERROR_DIR
        destination = target_dir / path.name
        if destination.exists():
            destination = target_dir / f"{path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{path.suffix}"
        shutil.move(str(path), str(destination))
    LAST_RUN = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    LAST_RESULT = {
        "imported": imported,
        "errors": errors,
        "message": f"{imported} arquivo(s) importado(s), {errors} com erro.",
    }
    return LAST_RESULT


def start_scheduler(minutes=30):
    global SCHEDULER
    ensure_auto_import_dirs()
    if BackgroundScheduler is None:
        return False
    if SCHEDULER and SCHEDULER.running:
        return True
    SCHEDULER = BackgroundScheduler()
    SCHEDULER.add_job(importar_planilhas, "interval", minutes=minutes, id="mms_auto_import", replace_existing=True)
    SCHEDULER.start()
    return True


def auto_import_status():
    ensure_auto_import_dirs()
    pending = sorted(AUTO_IMPORT_DIR.glob("*.xls*"))
    processed = sorted(PROCESSED_DIR.glob("*.xls*"))
    errors = sorted(ERROR_DIR.glob("*.xls*"))
    scheduler_running = bool(SCHEDULER and SCHEDULER.running)
    if scheduler_running:
        message = "Monitor automatico ativo."
    elif BackgroundScheduler is None:
        message = "Monitor preparado. Instale apscheduler para ativar verificacao automatica."
    else:
        message = "Monitor preparado. Ele inicia automaticamente quando o servidor do dashboard e aberto."
    return {
        "enabled": scheduler_running,
        "schedulerAvailable": BackgroundScheduler is not None,
        "mode": "automatico" if scheduler_running else "manual",
        "pending": len(pending),
        "processed": len(processed),
        "errors": len(errors),
        "folder": str(AUTO_IMPORT_DIR),
        "lastRun": LAST_RUN,
        "lastResult": LAST_RESULT,
        "message": message,
    }
