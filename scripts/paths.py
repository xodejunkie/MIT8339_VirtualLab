"""Shared project paths — all scripts resolve relative to the codebase root."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
SEGMENTS = DATA / "segments"
FIGS = DATA / "figs"
OUTPUTS = ROOT / "outputs"
SCRIPTS = ROOT / "scripts"

RAW_CSV = DATA / "BankChurners_raw.csv"
CLEANED_CSV = DATA / "cleaned_dataset.csv"
SCORED_CSV = DATA / "cleaned_dataset_scored.csv"
CLEANING_LOG = DATA / "cleaning_log.json"
QUALITY_CHECKS = DATA / "quality_checks.csv"
KPIS_JSON = DATA / "kpis.json"
DASHBOARD_XLSX = OUTPUTS / "MIT8339_Customer_Retention_Dashboard.xlsx"

for _d in (DATA, SEGMENTS, FIGS, OUTPUTS):
    _d.mkdir(parents=True, exist_ok=True)
