import os
import re
import shutil
import tempfile
import time
import zipfile
import json
import math
import copy
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from main import main as run_ami_optix_analysis, default_converter
from ami_optix.narrator import generate_internal_summary
from ami_optix.report_generator import create_excel_reports
from ami_optix.config_loader import load_config
from ami_optix.solver import find_optimal_scenarios, find_max_revenue_scenario
from ami_optix.rent_calculator import load_rent_schedule, compute_rents_for_assignments

app = Flask(__name__)

# ----------------------------------------------------------------------------
# Basic hardening / safety limits (configurable via env vars)
# ----------------------------------------------------------------------------
# NOTE: These are especially important if file upload endpoints are enabled.
# We keep defaults conservative; you can raise them via Render env vars if needed.

def _get_int_env(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or str(raw).strip() == "":
        return default
    try:
        return int(float(str(raw).strip()))
    except Exception:
        return default


# Request size limit (applies to ALL requests, including /api/analyze uploads).
MAX_UPLOAD_MB = max(1, _get_int_env("AMI_OPTIX_MAX_UPLOAD_MB", 50))
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

# Uploaded report cleanup policy for UPLOADS_DIR (zips from /api/analyze).
UPLOAD_RETENTION_HOURS = max(1, _get_int_env("AMI_OPTIX_UPLOAD_RETENTION_HOURS", 24))
UPLOAD_MAX_FILES = max(1, _get_int_env("AMI_OPTIX_UPLOAD_MAX_FILES", 50))


def _cleanup_uploads_dir() -> None:
    """
    Best-effort cleanup for generated report zips to prevent disk bloat.

    Policy:
    - Delete .zip files older than UPLOAD_RETENTION_HOURS
    - Keep at most UPLOAD_MAX_FILES newest .zip files
    """
    try:
        os.makedirs(UPLOADS_DIR, exist_ok=True)
        now = time.time()
        max_age = float(UPLOAD_RETENTION_HOURS) * 3600.0

        remaining: list[tuple[float, str]] = []
        for entry in os.scandir(UPLOADS_DIR):
            if not entry.is_file():
                continue
            if not entry.name.lower().endswith(".zip"):
                continue
            try:
                mtime = entry.stat().st_mtime
            except Exception:
                continue
            age = now - float(mtime)
            if age > max_age:
                try:
                    os.remove(entry.path)
                except OSError:
                    pass
                continue
            remaining.append((float(mtime), entry.path))

        if len(remaining) > UPLOAD_MAX_FILES:
            remaining.sort(key=lambda x: x[0])  # oldest first
            for _, path in remaining[:-UPLOAD_MAX_FILES]:
                try:
                    os.remove(path)
                except OSError:
                    pass
    except Exception:
        # Never break request flow due to cleanup failures.
        pass

UPLOADS_DIR = os.path.join(os.getcwd(), 'uploads')
DASHBOARD_DIR = os.path.join(os.getcwd(), 'dashboard_static')
os.makedirs(UPLOADS_DIR, exist_ok=True)

# Rent calculator storage - uses Render persistent disk if available
# Set RENT_CALCULATOR_DIR env var to your Render disk mount (e.g., /var/data/rent_calculators)
RENT_CALCULATORS_DIR = os.environ.get('RENT_CALCULATOR_DIR', os.path.join(os.getcwd(), 'rent_calculators'))
ACTIVE_CALCULATOR_FILE = os.path.join(RENT_CALCULATORS_DIR, '.active')
os.makedirs(RENT_CALCULATORS_DIR, exist_ok=True)

# ----------------------------------------------------------------------------
# Rent calculator naming conventions (Excel add-in)
# ----------------------------------------------------------------------------

DEFAULT_RENT_ROLL_YEAR = 2025
DEFAULT_RENT_CALCULATOR_FILENAME = "2025 AMI Rent Calculator Unlocked.xlsx"
RENT_CALC_REMOTE_PREFIX = "AMI_Optix_Rent_Calculator_"
RENT_CALC_REMOTE_SUFFIX = ".xlsx"

# Auto-seed bundled rent calculators into RENT_CALCULATORS_DIR on startup
BUNDLED_RENT_CALCULATORS_DIR = os.path.join(os.path.dirname(__file__), 'tools', 'excel-agent', 'assets', 'rent-roll-years')
if os.path.isdir(BUNDLED_RENT_CALCULATORS_DIR):
    for _year_dir in os.listdir(BUNDLED_RENT_CALCULATORS_DIR):
        _year_path = os.path.join(BUNDLED_RENT_CALCULATORS_DIR, _year_dir)
        if not os.path.isdir(_year_path):
            continue
        _src = os.path.join(_year_path, f"RentCalculator_{_year_dir}.xlsx")
        _dst = os.path.join(RENT_CALCULATORS_DIR, f"{RENT_CALC_REMOTE_PREFIX}{_year_dir}{RENT_CALC_REMOTE_SUFFIX}")
        if os.path.isfile(_src) and not os.path.isfile(_dst):
            try:
                shutil.copy2(_src, _dst)
                print(f"[STARTUP] Seeded rent calculator: {os.path.basename(_dst)}", flush=True)
            except Exception as _e:
                print(f"[STARTUP] Warning: could not seed {os.path.basename(_dst)}: {_e}", flush=True)

# API Key for Excel Add-in authentication
# Set this in environment variable: AMI_OPTIX_API_KEY
API_KEY = os.environ.get('AMI_OPTIX_API_KEY', '')
# Admin key for rent calculator management (optional, defaults to API key)
ADMIN_KEY = os.environ.get('AMI_OPTIX_ADMIN_KEY', API_KEY)
# Backward-compatibility: allow regular API key on admin endpoints.
# Set AMI_OPTIX_ALLOW_API_KEY_FOR_ADMIN=0 to require strict admin-only key.
ALLOW_API_KEY_FOR_ADMIN = str(os.environ.get('AMI_OPTIX_ALLOW_API_KEY_FOR_ADMIN', '1')).strip().lower() in {'1', 'true', 'yes', 'on'}

# ----------------------------------------------------------------------------
# Rent schedule cache (per-process)
# ----------------------------------------------------------------------------

_RENT_SCHEDULE_CACHE: dict[str, tuple[float, object]] = {}


def _timing_log_enabled() -> bool:
    raw = str(os.environ.get("AMI_OPTIX_TIMING_LOG", "")).strip().lower()
    return raw in {"1", "true", "yes", "on"}


def _load_rent_schedule_cached(workbook_path: str) -> tuple[object | None, bool]:
    """
    Load the active rent schedule with a simple in-memory cache.

    Keyed by (workbook_path, mtime) so edits to the workbook invalidate the cache.
    Returns: (schedule|None, cache_hit)
    """
    if not workbook_path:
        return None, False
    try:
        mtime = float(os.path.getmtime(workbook_path))
    except Exception:
        return load_rent_schedule(workbook_path), False

    cached = _RENT_SCHEDULE_CACHE.get(workbook_path)
    if cached and float(cached[0]) == mtime:
        return cached[1], True

    schedule = load_rent_schedule(workbook_path)
    _RENT_SCHEDULE_CACHE[workbook_path] = (mtime, schedule)
    return schedule, False


def _build_program_config(
    base_config: dict,
    program: str,
    mih_option: str | None = None,
    mih_residential_sf: float | None = None,
    mih_max_band_percent: int | None = None,
) -> dict:
    config = copy.deepcopy(base_config)
    rules = config.get('optimization_rules', {}) or {}

    program_norm = (program or 'UAP').strip().upper()
    if program_norm not in {'UAP', 'MIH'}:
        raise ValueError("Invalid program. Expected 'UAP' or 'MIH'.")

    if program_norm == 'UAP':
        config['optimization_rules'] = rules
        return config

    # MIH mode: do not reuse UAP deep affordability defaults.
    option_norm = (mih_option or '').strip().upper().replace('_', ' ')
    if option_norm in {'1', 'OPTION1', 'OPTION 1'}:
        option_norm = 'OPTION 1'
    if option_norm in {'4', 'OPTION4', 'OPTION 4'}:
        option_norm = 'OPTION 4'
    if option_norm not in {'OPTION 1', 'OPTION 4'}:
        raise ValueError("MIH requires mih_option = 'Option 1' or 'Option 4'.")

    if mih_residential_sf is None or float(mih_residential_sf) <= 0:
        raise ValueError("MIH requires mih_residential_sf > 0 (from MIH!J21).")

    # Configure MIH constraints.
    # Note: the 40% AMI band is constrained to a NARROW WINDOW [10.0%, 12.5%]
    # of residential SF for both Option 1 and Option 4 (client rule, 2026-04).
    # Floor at 10% prevents below-spec scenarios; ceiling at 12.5% prevents the
    # optimizer from overshooting (e.g., piling SF into 40% to balance high
    # bands like 90/110, which produced 14-17% scenarios). The /api/optimize
    # MIH path SLIDES this window UP in 0.1% steps if no scenarios fit the
    # initial window — see the floor-walking loop near find_optimal_scenarios.
    if option_norm == 'OPTION 1':
        rules['waami_cap_percent'] = 60.0
        rules['max_bands_per_scenario'] = 3
        rules['share_thresholds'] = [
            # MIH Option 1: 40% band must be in [10.0%, 12.5%] of residential SF.
            {'band_threshold': 40, 'min_share': 0.10, 'max_share': 0.125, 'denominator': 'residential'},
        ]
    else:
        rules['waami_cap_percent'] = 115.0
        rules['max_bands_per_scenario'] = 4
        # MIH Option 4: same 40% window as Option 1, plus client-workbook
        # formulas requiring >=5% at <=70 and >=10% at <=90.
        rules['share_thresholds'] = [
            {'band_threshold': 40, 'min_share': 0.10, 'max_share': 0.125, 'denominator': 'residential'},
            {'band_threshold': 70, 'min_share': 0.05, 'denominator': 'residential'},
            {'band_threshold': 90, 'min_share': 0.10, 'denominator': 'residential'},
        ]

    # MIH uses the same WAAMI floor as UAP (59.1%).
    # Without a floor, infeasible band combos (e.g. 2-band [40,70] with MIH SF constraint)
    # produce scenarios with WAAMI below 59% that should not be shown.
    rules['waami_floor'] = 0.591

    # Provide the denominator for residential-share constraints.
    rules['residential_sf'] = float(mih_residential_sf)

    # Band cap. Client decision 2026-05-18: MIH must NEVER use bands above 100% AMI
    # (Option 1 AND Option 4). We hard-cap on the server regardless of what the
    # client workbook sends, so we don't have to re-edit every property workbook
    # in the field - existing workbooks may still set 130 or 135 in the Prog sheet
    # and we silently floor that to 100.
    # To intentionally allow a LOWER cap (e.g. 80) the workbook can still send a
    # smaller mih_max_band_percent; only the upper bound is enforced.
    MIH_HARD_BAND_CAP = 100
    if mih_max_band_percent is None:
        mih_max_band_percent = MIH_HARD_BAND_CAP
    max_band = min(int(mih_max_band_percent), MIH_HARD_BAND_CAP)
    candidate_bands = [40, 60, 70, 80, 90, 100, 110, 120, 130, 135]
    rules['potential_bands'] = [b for b in candidate_bands if b <= max_band and b != 50]

    # Disable UAP-specific deep affordability defaults.
    rules['deep_affordability_min_share'] = None
    rules['deep_affordability_max_share'] = None

    config['optimization_rules'] = rules
    return config


def _validate_api_key():
    """Validate API key from request header. Returns error response or None if valid."""
    if not API_KEY:
        # No API key configured - allow all requests (dev mode)
        return None

    provided_key = request.headers.get('X-API-Key', '')
    if provided_key != API_KEY:
        return jsonify({"error": "Invalid or missing API key"}), 401
    return None


def _validate_admin_key():
    """Validate admin key for rent calculator management. Returns error response or None if valid."""
    if not ADMIN_KEY:
        # No admin key configured - allow all requests (dev mode)
        return None

    provided_api_key = request.headers.get('X-API-Key', '')
    provided_admin_key = request.headers.get('X-Admin-Key', '')
    provided_key = provided_api_key or provided_admin_key

    if provided_key == ADMIN_KEY:
        return None

    # Compatibility mode for existing Excel add-ins that only have a single API key field.
    # This avoids breaking deployments where AMI_OPTIX_ADMIN_KEY differs from AMI_OPTIX_API_KEY.
    if ALLOW_API_KEY_FOR_ADMIN and API_KEY and provided_api_key == API_KEY:
        return None

    if provided_key != ADMIN_KEY:
        return jsonify({"error": "Invalid or missing admin key"}), 401
    return None


def _get_active_rent_calculator_path():
    """Get the path to the active rent calculator file."""
    # Check if there's an active selection
    if os.path.exists(ACTIVE_CALCULATOR_FILE):
        with open(ACTIVE_CALCULATOR_FILE, 'r') as f:
            active_name = f.read().strip()
        if active_name:
            active_path = os.path.join(RENT_CALCULATORS_DIR, active_name)
            if os.path.exists(active_path):
                return active_path

    # Fall back to default in repo root
    default_path = os.path.join(os.getcwd(), DEFAULT_RENT_CALCULATOR_FILENAME)
    if os.path.exists(default_path):
        return default_path

    return None


def _default_rent_calculator_path() -> str | None:
    default_path = os.path.join(os.getcwd(), DEFAULT_RENT_CALCULATOR_FILENAME)
    if os.path.exists(default_path):
        return default_path
    return None


def _normalize_rent_roll_year(value) -> int | None:
    if value is None:
        return None
    try:
        year = int(str(value).strip())
    except Exception:
        return None
    if year < 1900 or year > 2100:
        return None
    return year


def _infer_year_from_calculator_id(calculator_id: str | None) -> int | None:
    if not calculator_id:
        return None
    m = re.match(
        rf"^{re.escape(RENT_CALC_REMOTE_PREFIX)}(\d{{4}}){re.escape(RENT_CALC_REMOTE_SUFFIX)}$",
        str(calculator_id),
        flags=re.IGNORECASE,
    )
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None


def _resolve_rent_calculator_for_request(data: dict) -> tuple[str | None, dict]:
    """
    Resolve a rent calculator path deterministically for a single request.

    This enables stateless API calls (Fix-06d verify) by selecting the rent
    schedule based on request payload rather than server-global active state.

    Supported request fields:
      - rent_roll_year (int)
      - calculator_id (string; e.g., "AMI_Optix_Rent_Calculator_2024.xlsx" or "default")
    """
    requested_year = _normalize_rent_roll_year(data.get("rent_roll_year"))
    requested_calc = str(data.get("calculator_id") or "").strip()
    if requested_calc == "":
        requested_calc = None

    meta = {
        "rent_roll_year_requested": requested_year,
        "calculator_id_requested": requested_calc,
        "rent_roll_year_used": None,
        "calculator_id_used": None,
        "calculator_filename": None,
        "rent_schedule_source": None,
        "rent_schedule_warning": None,
    }

    default_path = _default_rent_calculator_path()

    def _use(path: str | None, year_used: int | None, calc_id_used: str | None, source: str, warning: str | None = None):
        meta["rent_roll_year_used"] = year_used
        meta["calculator_id_used"] = calc_id_used
        meta["calculator_filename"] = os.path.basename(path) if path else None
        meta["rent_schedule_source"] = source
        meta["rent_schedule_warning"] = warning
        return path, meta

    # 1) Explicit calculator_id takes precedence.
    if requested_calc:
        calc_norm = requested_calc.strip()
        if calc_norm.lower() in {"default", DEFAULT_RENT_CALCULATOR_FILENAME.lower()}:
            # Deterministic "default" selection (do not depend on ACTIVE_CALCULATOR_FILE).
            return _use(default_path, DEFAULT_RENT_ROLL_YEAR if default_path else None, "default", "request_calculator_id")

        safe_name = secure_filename(calc_norm)
        candidate = os.path.join(RENT_CALCULATORS_DIR, safe_name)
        if os.path.exists(candidate):
            year_used = _infer_year_from_calculator_id(safe_name)
            return _use(candidate, year_used, safe_name, "request_calculator_id")

        warning = f"calculator_id not found: {safe_name}; using default rent calculator"
        return _use(default_path, DEFAULT_RENT_ROLL_YEAR if default_path else None, "default", "request_calculator_id_fallback", warning)

    # 2) Explicit rent_roll_year selection.
    if requested_year:
        # 2025 default: prefer uploaded override if present.
        if requested_year == DEFAULT_RENT_ROLL_YEAR:
            override_name = f"{RENT_CALC_REMOTE_PREFIX}{requested_year}{RENT_CALC_REMOTE_SUFFIX}"
            override_path = os.path.join(RENT_CALCULATORS_DIR, override_name)
            if os.path.exists(override_path):
                return _use(override_path, requested_year, override_name, "request_year")
            return _use(default_path, requested_year if default_path else None, "default", "request_year")

        expected_name = f"{RENT_CALC_REMOTE_PREFIX}{requested_year}{RENT_CALC_REMOTE_SUFFIX}"
        expected_path = os.path.join(RENT_CALCULATORS_DIR, expected_name)
        if os.path.exists(expected_path):
            return _use(expected_path, requested_year, expected_name, "request_year")

        warning = f"rent_roll_year not found on server: {requested_year}; using default rent calculator ({DEFAULT_RENT_ROLL_YEAR})"
        return _use(default_path, DEFAULT_RENT_ROLL_YEAR if default_path else None, "default", "request_year_fallback", warning)

    # 3) Backward compatibility: use server-global active selection.
    active_path = _get_active_rent_calculator_path()
    if not active_path:
        return _use(None, None, None, "server_active")

    filename = os.path.basename(active_path)
    if filename.lower() == DEFAULT_RENT_CALCULATOR_FILENAME.lower():
        return _use(active_path, DEFAULT_RENT_ROLL_YEAR, "default", "server_active")

    if os.path.dirname(active_path).rstrip("\\/") == RENT_CALCULATORS_DIR.rstrip("\\/"):
        year_used = _infer_year_from_calculator_id(filename)
        return _use(active_path, year_used, filename, "server_active")

    return _use(active_path, None, filename, "server_active")


def _list_rent_calculators():
    """List all available rent calculator files."""
    calculators = []
    active_name = None

    # Get active calculator name
    if os.path.exists(ACTIVE_CALCULATOR_FILE):
        with open(ACTIVE_CALCULATOR_FILE, 'r') as f:
            active_name = f.read().strip()

    # List uploaded calculators
    if os.path.exists(RENT_CALCULATORS_DIR):
        for filename in os.listdir(RENT_CALCULATORS_DIR):
            if filename.startswith('.'):
                continue
            if filename.lower().endswith(('.xlsx', '.xlsm')):
                filepath = os.path.join(RENT_CALCULATORS_DIR, filename)
                stat = os.stat(filepath)
                calculators.append({
                    'name': filename,
                    'size': stat.st_size,
                    'modified': stat.st_mtime,
                    'is_active': filename == active_name,
                    'source': 'uploaded'
                })

    # Add default calculator if it exists
    default_path = os.path.join(os.getcwd(), "2025 AMI Rent Calculator Unlocked.xlsx")
    if os.path.exists(default_path):
        stat = os.stat(default_path)
        is_default_active = not active_name  # Default is active if no selection
        calculators.append({
            'name': '2025 AMI Rent Calculator Unlocked.xlsx',
            'size': stat.st_size,
            'modified': stat.st_mtime,
            'is_active': is_default_active,
            'source': 'default'
        })

    return calculators


def _sanitize_for_json(value):
    """Recursively convert numpy/pandas types and NaN values to JSON-safe primitives."""
    if isinstance(value, dict):
        return {key: _sanitize_for_json(val) for key, val in value.items()}
    if isinstance(value, list):
        return [_sanitize_for_json(item) for item in value]
    if isinstance(value, tuple):
        return [_sanitize_for_json(item) for item in value]
    if isinstance(value, np.ndarray):
        return [_sanitize_for_json(item) for item in value.tolist()]
    if isinstance(value, (float, np.floating)):
        if math.isnan(value) or math.isinf(value):
            return None
        return float(value)
    if isinstance(value, (np.integer,)):
        return int(value)
    try:
        return default_converter(value)
    except TypeError:
        return value


def _dashboard_file_exists(filename: str) -> bool:
    return os.path.exists(os.path.join(DASHBOARD_DIR, filename))


def _validate_assignment_payload(
    units: list[dict],
    config: dict,
) -> tuple[bool, list[str], dict]:
    """
    Validate an explicit unit -> assigned_ami mapping against the active optimization rules.

    Returns: (is_valid, errors, summary)
    """
    rules = (config or {}).get('optimization_rules', {}) or {}

    errors: list[str] = []
    potential_bands = sorted({int(b) for b in (rules.get('potential_bands') or []) if int(b) != 50})
    max_bands = int(rules.get('max_bands_per_scenario') or 3)

    # Normalize and validate assigned AMIs.
    total_sf = 0.0
    waami_num = 0.0
    used_bands: set[int] = set()
    low_band_stats: dict[int, float] = {}

    for i, unit in enumerate(units):
        unit_id = str(unit.get('unit_id', ''))
        try:
            net_sf = float(unit.get('net_sf') or 0.0)
        except Exception:
            net_sf = 0.0
        try:
            assigned = float(unit.get('assigned_ami'))
        except Exception:
            assigned = None

        if not unit_id:
            errors.append(f"Unit {i+1} missing unit_id.")
            continue
        if net_sf <= 0:
            errors.append(f"Unit '{unit_id}' has invalid net_sf.")
            continue
        if assigned is None:
            errors.append(f"Unit '{unit_id}' missing assigned_ami.")
            continue

        # Allow either 0.6 or 60 inputs; normalize to 0-2 range.
        if assigned > 2.0:
            assigned = assigned / 100.0
        if assigned <= 0 or assigned > 2.0:
            errors.append(f"Unit '{unit_id}' has invalid assigned_ami value.")
            continue

        band = int(round(assigned * 100))
        used_bands.add(band)
        if potential_bands and band not in potential_bands:
            errors.append(f"Unit '{unit_id}' uses {band}% which is not an allowed band for this program.")

        total_sf += net_sf
        waami_num += net_sf * assigned
        low_band_stats[band] = low_band_stats.get(band, 0.0) + net_sf

    if errors:
        return False, errors, {}

    if len(used_bands) > max_bands:
        errors.append(f"Too many bands used ({len(used_bands)}). Max allowed is {max_bands}.")

    waami = (waami_num / total_sf) if total_sf else 0.0
    waami_cap_percent = float(rules.get('waami_cap_percent') or 60.0)
    if waami > (waami_cap_percent / 100.0) + 1e-12:
        errors.append(f"WAAMI {waami*100:.2f}% exceeds cap {waami_cap_percent:.2f}%.")

    waami_floor = rules.get('waami_floor')
    if waami_floor is not None:
        try:
            waami_floor_f = float(waami_floor)
            if waami_floor_f > 0 and waami + 1e-12 < waami_floor_f:
                errors.append(f"WAAMI {waami*100:.2f}% is below floor {waami_floor_f*100:.2f}%.")
        except Exception:
            pass

    # Share threshold constraints (MIH or legacy UAP deep affordability).
    thresholds = rules.get('share_thresholds')
    if not thresholds:
        min_share = rules.get('deep_affordability_min_share')
        max_share = rules.get('deep_affordability_max_share')
        threshold_band = int(rules.get('low_band_band_threshold') or 40)
        if min_share is not None or max_share is not None:
            thresholds = [{
                'band_threshold': threshold_band,
                'min_share': min_share,
                'max_share': max_share,
                'denominator': 'affordable',
            }]

    if thresholds:
        residential_sf = rules.get('residential_sf')
        for t in thresholds:
            band_threshold = int(t.get('band_threshold') or 40)
            denom = str(t.get('denominator') or 'affordable').lower()
            denom_sf = total_sf if denom == 'affordable' else float(residential_sf or 0.0)
            if denom_sf <= 0:
                errors.append("Residential SF denominator is missing/invalid for share constraint validation.")
                continue
            num_sf = sum(sf for band, sf in low_band_stats.items() if int(band) <= band_threshold)
            share = num_sf / denom_sf
            if t.get('min_share') is not None and share + 1e-12 < float(t['min_share']):
                errors.append(f"Share at <= {band_threshold}% is {share*100:.2f}%, below required {float(t['min_share'])*100:.2f}%.")
            if t.get('max_share') is not None and share - 1e-12 > float(t['max_share']):
                errors.append(f"Share at <= {band_threshold}% is {share*100:.2f}%, above maximum {float(t['max_share'])*100:.2f}%.")

    summary = {
        'waami': waami,
        'waami_percent': waami * 100.0,
        'bands_used': sorted(list(used_bands)),
        'total_sf': total_sf,
        'total_units': len(units),
    }
    return len(errors) == 0, errors, summary


def _build_metrics_from_assignments(assignments: list[dict]) -> dict:
    """Compute lightweight scenario metrics for display in Excel (band mix, totals, etc.)."""
    total_sf = 0.0
    waami_num = 0.0
    band_stats: dict[int, dict] = {}

    for unit in assignments or []:
        try:
            net_sf = float(unit.get('net_sf') or 0.0)
        except Exception:
            net_sf = 0.0
        try:
            assigned = float(unit.get('assigned_ami') or 0.0)
        except Exception:
            assigned = 0.0

        total_sf += net_sf
        waami_num += net_sf * assigned

        band = int(round(assigned * 100))
        stats = band_stats.setdefault(band, {'band': band, 'units': 0, 'net_sf': 0.0})
        stats['units'] += 1
        stats['net_sf'] += net_sf

    waami = (waami_num / total_sf) if total_sf else 0.0
    band_mix = []
    for band in sorted(band_stats.keys()):
        stats = band_stats[band]
        share = (stats['net_sf'] / total_sf) if total_sf else 0.0
        band_mix.append({**stats, 'share_of_sf': share})

    return {
        'total_units': len(assignments or []),
        'total_sf': total_sf,
        'revenue_score': waami_num,
        'waami_percent': waami * 100.0,
        'band_mix': band_mix,
        'total_monthly_rent': 0.0,
        'total_annual_rent': 0.0,
    }


@app.route('/healthz')
def healthcheck():
    """Lightweight health endpoint for uptime checks."""
    return jsonify({"status": "ok"})


@app.route('/api/optimize', methods=['POST'])
def optimize_units():
    """
    JSON-based optimization endpoint for Excel VBA Add-in.

    Accepts unit data directly as JSON (no file upload needed).
    Requires API key authentication via X-API-Key header.

    Request body:
    {
        "units": [
            {"unit_id": "1A", "bedrooms": 2, "net_sf": 850, "floor": 1},
            {"unit_id": "1B", "bedrooms": 1, "net_sf": 650, "floor": 2}
        ],
        "utilities": {
            "electricity": "tenant_pays",
            "cooking": "gas",
            "heat": "gas",
            "hot_water": "gas"
        }
    }

    Returns:
    {
        "scenarios": { ... },
        "notes": [ ... ],
        "project_summary": { ... }
    }
    """
    timing_enabled = _timing_log_enabled()
    timing: dict[str, object] = {"endpoint": "optimize"}
    request_start = time.perf_counter()

    def _emit_timing(status: str, extra: dict | None = None) -> None:
        if not timing_enabled:
            return
        payload = dict(timing)
        payload["status"] = status
        payload["elapsed_ms"] = int(round((time.perf_counter() - request_start) * 1000))
        if extra:
            payload.update(extra)
        try:
            print(json.dumps(payload), flush=True)
        except Exception:
            pass

    # Validate API key
    auth_error = _validate_api_key()
    if auth_error:
        _emit_timing("auth_error")
        return auth_error

    # Parse JSON body
    try:
        data = request.get_json()
        if not data:
            _emit_timing("invalid_json", {"error": "Request body must be JSON"})
            return jsonify({"error": "Request body must be JSON"}), 400
    except Exception as e:
        _emit_timing("invalid_json", {"error": str(e)})
        return jsonify({"error": f"Invalid JSON: {str(e)}"}), 400

    # Extract units
    units = data.get('units', [])
    if not units or not isinstance(units, list):
        _emit_timing("invalid_units", {"error": "Missing or invalid 'units' array"})
        return jsonify({"error": "Missing or invalid 'units' array"}), 400

    # Validate required fields for each unit
    required_fields = ['unit_id', 'bedrooms', 'net_sf']
    for i, unit in enumerate(units):
        for field in required_fields:
            if field not in unit:
                _emit_timing("invalid_units", {"error": f"Unit {i+1} missing required field: {field}"})
                return jsonify({"error": f"Unit {i+1} missing required field: {field}"}), 400

    # Extract utilities (with defaults)
    utilities = data.get('utilities', {})
    utilities_clean = {
        'electricity': utilities.get('electricity', 'na'),
        'cooking': utilities.get('cooking', 'na'),
        'heat': utilities.get('heat', 'na'),
        'hot_water': utilities.get('hot_water', 'na'),
    }

    try:
        program = (data.get('program') or 'UAP')
        print(f"[MIH-DEBUG] payload top-level keys={list(data.keys())}", flush=True)
        mih_option = data.get('mih_option')
        mih_residential_sf = data.get('mih_residential_sf')
        mih_max_band_percent = data.get('mih_max_band_percent')
        raw_po = data.get('project_overrides')
        print(f"[MIH-DEBUG] raw project_overrides type={type(raw_po).__name__}, value={str(raw_po)[:200]}", flush=True)
        project_overrides = raw_po if isinstance(raw_po, dict) else None
        compare_baseline = bool(data.get('compare_baseline')) if data.get('compare_baseline') is not None else False

        # Convert units to DataFrame (same format parser produces)
        df_units = pd.DataFrame(units)

        # Ensure required columns exist with correct types
        df_units['unit_id'] = df_units['unit_id'].astype(str)
        df_units['bedrooms'] = pd.to_numeric(df_units['bedrooms'], errors='coerce')
        df_units['net_sf'] = pd.to_numeric(df_units['net_sf'], errors='coerce')

        # Optional columns
        if 'floor' in df_units.columns:
            df_units['floor'] = pd.to_numeric(df_units['floor'], errors='coerce')
        if 'balcony' in df_units.columns:
            df_units['balcony'] = df_units['balcony'].astype(bool)

        # Add client_ami column (required by solver, will be overwritten)
        # Set to a placeholder that indicates "needs assignment"
        if 'client_ami' not in df_units.columns:
            df_units['client_ami'] = 0.6  # Default placeholder

        # Load base config and specialize it per program.
        config = load_config()
        try:
            config = _build_program_config(
                config,
                program=program,
                mih_option=mih_option,
                mih_residential_sf=mih_residential_sf,
                mih_max_band_percent=mih_max_band_percent,
            )
        except ValueError as e:
            _emit_timing("config_error", {"error": str(e)})
            return jsonify({
                "success": False,
                "error": str(e),
                "notes": [],
            }), 200

        program_norm = str(program or 'UAP').strip().upper()

        timing["program"] = program_norm
        timing["unit_count"] = int(len(df_units))
        timing["parse_validation_ms"] = int(round((time.perf_counter() - request_start) * 1000))

        # MIH total_building_sf: store for project_summary display (Share of Building SF column).
        # Share thresholds are already set correctly by _build_program_config() — do NOT overwrite them.
        mih_constraint_injected = False
        total_building_sf = 0.0
        if program_norm == 'MIH' and mih_residential_sf is not None:
            try:
                total_building_sf = float(mih_residential_sf)
            except (TypeError, ValueError):
                total_building_sf = 0.0
            if total_building_sf > 0:
                rules = config.setdefault('optimization_rules', {})
                rules['total_building_sf'] = total_building_sf
                mih_constraint_injected = True
                print(f"[MIH-DEBUG] total_building_sf={total_building_sf:.2f}, share_thresholds (from _build_program_config)={rules.get('share_thresholds')}", flush=True)
        if not mih_constraint_injected:
            print(f"[MIH-DEBUG] total_building_sf NOT set: program={program_norm}, mih_residential_sf={mih_residential_sf}", flush=True)

        # Load rent schedule BEFORE the solver runs so we can build per-band
        # rent coefficients and pass them to find_optimal_scenarios. This makes
        # the solver maximize ACTUAL rent dollars per combo (using haircut-
        # adjusted rents from rent_components), instead of the WAAMI proxy.
        rent_load_start = time.perf_counter()
        rent_schedule = None
        rent_schedule_cache_hit = False
        rent_calc_path = _get_active_rent_calculator_path()
        if rent_calc_path:
            try:
                rent_schedule, rent_schedule_cache_hit = _load_rent_schedule_cached(rent_calc_path)
            except Exception as e:
                notes.append(f"Warning: Could not load rent calculator: {str(e)}")
        timing["rent_schedule_load_ms"] = int(round((time.perf_counter() - rent_load_start) * 1000))
        timing["rent_schedule_cache_hit"] = bool(rent_schedule_cache_hit)

        rent_by_band_cents = None
        if rent_schedule:
            try:
                bands_for_rent = sorted({int(b) for b in (config.get('optimization_rules', {}) or {}).get('potential_bands', [])})
                rent_by_band_cents = {}
                for band in bands_for_rent:
                    rents = []
                    for _, unit_row in df_units.iterrows():
                        components = rent_schedule.rent_components(band / 100.0, unit_row.get('bedrooms', 0), utilities_clean)
                        rents.append(int(round(float(components['gross']) * 100)))
                    rent_by_band_cents[int(band)] = rents
            except Exception as e:
                rent_by_band_cents = None
                notes.append(f"Warning: Could not build per-band rent coefficients: {str(e)}")

        solver_start = time.perf_counter()

        # Run the strict solver (optionally with project overrides for premium weights / unit rules).
        # MIH: SLIDE the 40 AMI band window [min_share, max_share] UP in 0.1% steps if no
        # scenarios fit the initial window [10.0%, 12.5%]. Window WIDTH stays constant (2.5%)
        # so each step is e.g. [10.1%, 12.6%], [10.2%, 12.7%]... Cap at min_share=15%. Per
        # client spec 2026-04: keep returning scenarios even when the strict window is
        # infeasible — slide the window up and rerun the optimizer. Mirrors the UAP widening
        # pattern below. After the walk settles, config is mutated so the same final window
        # flows into find_max_revenue_scenario later in the request.
        if program_norm == 'MIH':
            mih_floor_start = 0.100
            mih_floor_max = 0.150
            mih_floor_step = 0.001
            mih_window_width = 0.025  # max_share = min_share + 2.5%
            mih_walk_results = None
            mih_walk_last = None
            mih_walk_value = mih_floor_start

            while mih_walk_value <= mih_floor_max + 1e-9:
                for _t in (config.get('optimization_rules', {}) or {}).get('share_thresholds', []):
                    if int(_t.get('band_threshold', 0)) == 40:
                        _t['min_share'] = mih_walk_value
                        _t['max_share'] = round(mih_walk_value + mih_window_width, 5)
                _trial = find_optimal_scenarios(df_units, config, project_overrides=project_overrides, rent_by_band_cents=rent_by_band_cents)
                mih_walk_last = _trial
                if (_trial.get('scenarios') or {}).get('absolute_best'):
                    mih_walk_results = _trial
                    if mih_walk_value > mih_floor_start + 1e-9:
                        _max_pct = (mih_walk_value + mih_window_width) * 100
                        mih_walk_results.setdefault('notes', []).append(
                            f"40% AMI window slid up to [{mih_walk_value*100:.1f}%, {_max_pct:.1f}%] to find feasible scenarios for this building."
                        )
                    break
                mih_walk_value = round(mih_walk_value + mih_floor_step, 5)

            if mih_walk_results is None:
                solver_results = mih_walk_last or {'scenarios': {}, 'notes': []}
                _max_pct = (mih_floor_max + mih_window_width) * 100
                solver_results.setdefault('notes', []).append(
                    f"No feasible MIH scenarios found with 40% AMI window slid from [{mih_floor_start*100:.1f}%, {(mih_floor_start+mih_window_width)*100:.1f}%] up to [{mih_floor_max*100:.1f}%, {_max_pct:.1f}%]."
                )
            else:
                solver_results = mih_walk_results
        else:
            solver_results = find_optimal_scenarios(df_units, config, project_overrides=project_overrides, rent_by_band_cents=rent_by_band_cents)
        scenarios = solver_results.get('scenarios', {}) or {}
        notes = solver_results.get('notes', []) or []

        # Optional: baseline run for learning compare (runs strict rules without overrides).
        baseline_scenarios = None
        baseline_notes = None
        if compare_baseline and project_overrides:
            baseline_results = find_optimal_scenarios(df_units, config, project_overrides=None, rent_by_band_cents=rent_by_band_cents)
            baseline_scenarios = baseline_results.get('scenarios', {}) or {}
            baseline_notes = baseline_results.get('notes', []) or []

        # If strict UAP constraints yield no feasible solution, widen the deep affordability max-share
        # in small increments until a solution exists. This prevents "no results" on workbooks where the
        # 40%-band SF share cannot hit an extremely narrow target due to unit SF granularity.
        if (not scenarios or not scenarios.get('absolute_best')) and program_norm == 'UAP':
            try:
                rules = config.get('optimization_rules', {}) or {}
                min_share = rules.get('deep_affordability_min_share')
                max_share = rules.get('deep_affordability_max_share')
                widen_step = float(rules.get('deep_affordability_widen_step', 0.005) or 0.0)
                widen_cap = float(rules.get('deep_affordability_widen_cap', 0.4) or 0.0)

                if (min_share is not None) and (max_share is not None) and widen_step > 0 and widen_cap > 0:
                    start = float(max_share)
                    candidate = start + widen_step
                    while candidate <= widen_cap + 1e-12:
                        relaxed_config = copy.deepcopy(config)
                        relaxed_rules = relaxed_config.get('optimization_rules', {}) or {}
                        relaxed_rules['deep_affordability_max_share'] = float(candidate)
                        relaxed_config['optimization_rules'] = relaxed_rules

                        relaxed_results = find_optimal_scenarios(df_units, relaxed_config, project_overrides=project_overrides, rent_by_band_cents=rent_by_band_cents)
                        relaxed_scenarios = relaxed_results.get('scenarios', {}) or {}
                        if relaxed_scenarios.get('absolute_best'):
                            notes.append(
                                f"Strict deep affordability max share {start*100:.1f}% was infeasible; widened to {candidate*100:.1f}% to find a feasible solution."
                            )
                            notes.extend(relaxed_results.get('notes', []) or [])
                            scenarios = relaxed_scenarios
                            config = relaxed_config  # carry forward so edge validation uses the same strict baseline
                            break

                        candidate = candidate + widen_step
            except Exception:
                pass

        timing["find_optimal_scenarios_ms"] = int(round((time.perf_counter() - solver_start) * 1000))

        if not scenarios or not scenarios.get('absolute_best'):
            _emit_timing("no_solution", {"notes_count": int(len(notes or []))})
            return jsonify({
                "success": False,
                "error": "No optimal solution found",
                "notes": notes
            }), 200

        # Keep the full strict scenario set around as a fallback so we can top up to ~6 scenarios
        # even if edge/rent-max variants can't produce enough distinct mixes.
        strict_scenarios_full = copy.deepcopy(scenarios or {})

        # Keep strict output client-friendly: cap strict scenarios to 3 (best + 2 variants).
        # This prevents Excel from being overwhelmed while still showing meaningful alternatives.
        try:
            preferred_strict_keys = ["absolute_best", "best_3_band", "best_2_band", "alternative"]
            filtered = {}
            seen_canons = set()

            for k in preferred_strict_keys:
                s = scenarios.get(k)
                if not s:
                    continue

                canon = s.get("canonical_assignments")
                ck = None
                if canon:
                    try:
                        ck = tuple(tuple(pair) for pair in canon)
                    except Exception:
                        ck = None

                if ck and ck in seen_canons:
                    continue

                filtered[k] = s
                if ck:
                    seen_canons.add(ck)

                if len(filtered) >= 3:
                    break

            if filtered:
                scenarios = filtered
        except Exception:
            pass

        # rent_schedule was already loaded above (before solver) so we could
        # pass per-band rent coefficients into the solver.
        def _apply_rents_to_scenarios(scenarios_dict: dict) -> None:
            if not rent_schedule:
                return
            for scenario_key, scenario in scenarios_dict.items():
                if scenario and 'assignments' in scenario:
                    assignments, rent_totals = compute_rents_for_assignments(
                        rent_schedule,
                        scenario['assignments'],
                        utilities_clean
                    )
                    scenario['assignments'] = assignments
                    scenario['rent_totals'] = rent_totals

        # Apply rent calculations to each scenario
        if rent_schedule:
            rent_apply_start = time.perf_counter()
            _apply_rents_to_scenarios(scenarios)
            if strict_scenarios_full:
                _apply_rents_to_scenarios(strict_scenarios_full)
            if baseline_scenarios:
                _apply_rents_to_scenarios(baseline_scenarios)
            timing["rent_apply_ms"] = int(round((time.perf_counter() - rent_apply_start) * 1000))

            # Rent-first re-ranking: switch 'absolute_best' to the scenario with
            # the highest net monthly rent (which now reflects the 100% AMI
            # haircut applied in rent_components). Preserve the original
            # WAAMI-max winner under 'closest_to_60' so the client can compare.
            # The solver's revenue_score is a WAAMI proxy, not actual rent
            # dollars — actual rent is only known after _apply_rents_to_scenarios.
            original_best = scenarios.get('absolute_best')
            if original_best and original_best.get('rent_totals'):
                rented = [
                    (k, v) for k, v in scenarios.items()
                    if v and isinstance(v.get('rent_totals'), dict)
                       and v['rent_totals'].get('net_monthly') is not None
                ]
                if rented:
                    best_key, best_scenario = max(
                        rented,
                        key=lambda kv: float(kv[1]['rent_totals']['net_monthly'])
                    )
                    if best_scenario is not original_best:
                        scenarios['closest_to_60'] = original_best
                        scenarios['absolute_best'] = best_scenario
                        notes.append(
                            f"Recommended scenario switched from WAAMI-max to rent-max "
                            f"(higher net monthly rent by ${float(best_scenario['rent_totals']['net_monthly']) - float(original_best['rent_totals']['net_monthly']):.0f}). "
                            f"WAAMI-max scenario preserved as 'closest_to_60' for comparison."
                        )

            # 40%-share variants: give the client an explicit choice between
            # "minimal 40% allocation" (low_40_share, pinned near the legal
            # floor — matches what many developers prefer) and "maximal 40%
            # allocation" (max_40_share, pinned near the ceiling — typically
            # the rent-max landing zone). Both still use rent-max within the
            # narrowed window so the trade-off is purely about 40% exposure.
            if rent_schedule and rent_by_band_cents:
                def _solve_with_40_window(min_share, max_share):
                    v_config = copy.deepcopy(config)
                    v_rules = v_config.get('optimization_rules', {}) or {}
                    modified = False
                    thresholds = v_rules.get('share_thresholds')
                    if isinstance(thresholds, list):
                        for t in thresholds:
                            try:
                                if int(t.get('band_threshold', 0)) <= 40:
                                    if min_share is not None:
                                        t['min_share'] = float(min_share)
                                    if max_share is not None:
                                        t['max_share'] = float(max_share)
                                    modified = True
                            except (TypeError, ValueError):
                                pass
                    # UAP path
                    if v_rules.get('deep_affordability_min_share') is not None and v_rules.get('deep_affordability_max_share') is not None:
                        if min_share is not None:
                            v_rules['deep_affordability_min_share'] = float(min_share)
                        if max_share is not None:
                            v_rules['deep_affordability_max_share'] = float(max_share)
                        modified = True
                    if not modified:
                        return None
                    v_results = find_optimal_scenarios(
                        df_units, v_config,
                        project_overrides=project_overrides,
                        rent_by_band_cents=rent_by_band_cents,
                    )
                    v_scenarios = v_results.get('scenarios', {}) or {}
                    best_pick = None
                    best_pick_rent = -1.0
                    for k, s in v_scenarios.items():
                        if not s or 'assignments' not in s:
                            continue
                        a, t = compute_rents_for_assignments(rent_schedule, s['assignments'], utilities_clean)
                        s['assignments'] = a
                        s['rent_totals'] = t
                        nm = (t or {}).get('net_monthly')
                        if nm is not None and float(nm) > best_pick_rent:
                            best_pick = s
                            best_pick_rent = float(nm)
                    return best_pick

                # Read the EFFECTIVE 40% window from the (possibly-mutated) config.
                # MIH's floor-walk may have slid min_share above 10%; we respect that.
                eff_min, eff_max = None, None
                for t in (config.get('optimization_rules', {}) or {}).get('share_thresholds') or []:
                    try:
                        if int(t.get('band_threshold', 0)) <= 40:
                            eff_min = float(t.get('min_share') or 0.10)
                            eff_max = float(t.get('max_share') or 0.125)
                            break
                    except (TypeError, ValueError):
                        pass
                if eff_min is None:
                    eff_min = float((config.get('optimization_rules', {}) or {}).get('deep_affordability_min_share') or 0.10)
                    eff_max = float((config.get('optimization_rules', {}) or {}).get('deep_affordability_max_share') or 0.125)

                # low_40_share: narrow window at the floor (e.g., [10.0%, 10.5%])
                window_width = 0.005
                low_40 = _solve_with_40_window(eff_min, eff_min + window_width)
                if low_40 and low_40.get('rent_totals'):
                    scenarios['low_40_share'] = low_40

                # max_40_share: narrow window at the ceiling (e.g., [12.0%, 12.5%])
                if eff_max > eff_min + window_width + 1e-9:
                    max_40 = _solve_with_40_window(max(eff_min, eff_max - window_width), eff_max)
                    if max_40 and max_40.get('rent_totals'):
                        # Only add if it's distinct from absolute_best
                        ab_canon = (scenarios.get('absolute_best') or {}).get('canonical_assignments')
                        if max_40.get('canonical_assignments') != ab_canon:
                            scenarios['max_40_share'] = max_40

            # --- Edge / relaxed scenarios (UAP + MIH) ---
            # Generate up to N additional rent-maximizing scenarios to improve rent totals while still
            # respecting program rules. For UAP we may relax deep-affordability share bounds; for MIH
            # we keep share rules fixed and only relax the WAAMI floor (down to 58%).
            if program_norm in ('UAP', 'MIH'):
                edge_start = time.perf_counter()
                try:
                    strict_config = copy.deepcopy(config)
                    strict_rules = strict_config.get('optimization_rules', {}) or {}

                    edge_min_waami = 0.58  # never go below 58.00% in relaxed scenarios

                    def _canon_key(canon) -> tuple | None:
                        if not canon:
                            return None
                        try:
                            return tuple(tuple(pair) for pair in canon)
                        except Exception:
                            return None

                    existing_canons = set()
                    for s in (scenarios or {}).values():
                        ck = _canon_key((s or {}).get('canonical_assignments'))
                        if ck:
                            existing_canons.add(ck)

                    bands = sorted({int(b) for b in config.get('optimization_rules', {}).get('potential_bands', [])})
                    rent_by_band_cents = {}
                    for band in bands:
                        rents = []
                        for _, unit_row in df_units.iterrows():
                            components = rent_schedule.rent_components(band / 100.0, unit_row.get('bedrooms', 0), utilities_clean)
                            rents.append(int(round(float(components['gross']) * 100)))
                        rent_by_band_cents[int(band)] = rents

                    # Target total scenarios (strict + edge) returned to Excel.
                    # Excel can handle ~6 without becoming sluggish, and this matches the UI expectation.
                    target_total_count = max(1, _get_int_env("AMI_OPTIX_TARGET_TOTAL_SCENARIOS", 6))
                    target_edge_count = max(0, min(5, int(target_total_count - len(scenarios or {}))))
                    edge_keys_added: list[str] = []

                    def _maybe_add_edge(key: str, edge_config: dict, waami_floor: float, edge_settings: dict) -> bool:
                        nonlocal scenarios, notes, existing_canons
                        if len(edge_keys_added) >= target_edge_count:
                            return False

                        # Keep the rent search bounded and responsive.
                        # find_max_revenue_scenario applies its own conservative defaults if these are unset.
                        edge_config = copy.deepcopy(edge_config)

                        candidate = find_max_revenue_scenario(
                            df_units,
                            edge_config,
                            rent_by_band_cents=rent_by_band_cents,
                            waami_floor=float(waami_floor),
                            project_overrides=project_overrides,
                        )
                        if not candidate or candidate.get('status') != 'OPTIMAL':
                            return False

                        waami_val = float(candidate.get('waami') or 0.0)
                        if waami_val + 1e-12 < edge_min_waami:
                            return False

                        ck = _canon_key(candidate.get('canonical_assignments'))
                        if ck and ck in existing_canons:
                            return False

                        assignments, rent_totals = compute_rents_for_assignments(
                            rent_schedule,
                            candidate['assignments'],
                            utilities_clean,
                        )
                        candidate['assignments'] = assignments
                        candidate['rent_totals'] = rent_totals

                        # HARD WAAMI CAP: client requirement is the legal cap is
                        # never exceeded — not by a percent, not by a fraction.
                        # Tiny float overshoots (e.g. 60.00004% from rounding)
                        # still count as non-compliant and the scenario is dropped.
                        cap_fraction_strict = float(strict_rules.get('waami_cap_percent') or 60.0) / 100.0
                        total_sf_strict = sum(float(u.get('net_sf', 0) or 0) for u in assignments)
                        if total_sf_strict > 0:
                            actual_waami_strict = sum(
                                float(u.get('net_sf', 0) or 0) * float(u.get('assigned_ami', 0) or 0)
                                for u in assignments
                            ) / total_sf_strict
                            if actual_waami_strict > cap_fraction_strict:
                                notes.append(
                                    f"Dropped '{key}' edge scenario: WAAMI {actual_waami_strict*100:.6f}% exceeds the {cap_fraction_strict*100:.2f}% cap."
                                )
                                return False

                        # Tradeoffs: validate edge scenario under STRICT rules and keep only the failures.
                        is_valid, errors, _summary = _validate_assignment_payload(assignments, strict_config)
                        candidate['tradeoffs'] = [] if is_valid else errors[:8]
                        candidate['edge_settings'] = edge_settings
                        candidate['tier'] = 'edge'

                        scenarios[key] = candidate
                        if ck:
                            existing_canons.add(ck)
                        edge_keys_added.append(key)
                        return True

                    # Strict "Max Revenue" (no relaxation): rent-maximizing scenario under strict rules.
                    waami_cap_percent = float(strict_rules.get('waami_cap_percent') or 60.0)
                    if program_norm == 'MIH':
                        # MIH templates display compliance against ~60% Avg AMI when the cap is 60.
                        # Use 60% as the "strict" baseline so relaxed scenarios can show tradeoffs below 60.
                        if waami_cap_percent <= 60.0001:
                            strict_floor = 0.6
                        else:
                            strict_floor = float(strict_rules.get('waami_floor') or edge_min_waami)

                        strict_cfg_rules = strict_config.get('optimization_rules', {}) or {}
                        strict_cfg_rules['waami_floor'] = float(strict_floor)
                        strict_config['optimization_rules'] = strict_cfg_rules
                    else:
                        strict_floor = float(strict_rules.get('waami_floor') or 0.591)
                    _maybe_add_edge(
                        "max_revenue",
                        config,
                        waami_floor=strict_floor,
                        edge_settings={"mode": "rent_max", "relaxed": False, "waami_floor": float(strict_floor)},
                    )

                    if program_norm == 'UAP':
                        # Edge 1: relax max share at <=40% (try 22% -> 23%, kept tight to avoid
                        # obviously non-compliant scenarios like 47% at 40% AMI band).
                        strict_min_share = strict_rules.get('deep_affordability_min_share')
                        strict_max_share = strict_rules.get('deep_affordability_max_share')
                        if len(edge_keys_added) < target_edge_count:
                            for max_share in (0.22, 0.23):
                                if strict_max_share is not None and float(max_share) <= float(strict_max_share) + 1e-12:
                                    continue
                                edge_cfg = copy.deepcopy(config)
                                edge_cfg['optimization_rules']['deep_affordability_min_share'] = strict_min_share
                                edge_cfg['optimization_rules']['deep_affordability_max_share'] = float(max_share)
                                if _maybe_add_edge(
                                    f"edge_max_share_{int(round(max_share*100))}",
                                    edge_cfg,
                                    waami_floor=strict_floor,
                                    edge_settings={"mode": "rent_max", "relaxed": True, "deep_affordability_max_share": float(max_share)},
                                ):
                                    break

                        # Edge 2: relax min share at <=40% (try 19.9% down to 19%, kept tight
                        # to avoid scenarios like 16.96% at 40% AMI band).
                        if len(edge_keys_added) < target_edge_count:
                            for min_share in (0.199, 0.198, 0.195, 0.19):
                                if strict_min_share is not None and float(min_share) >= float(strict_min_share) - 1e-12:
                                    continue
                                edge_cfg = copy.deepcopy(config)
                                edge_cfg['optimization_rules']['deep_affordability_min_share'] = float(min_share)
                                edge_cfg['optimization_rules']['deep_affordability_max_share'] = strict_max_share
                                if _maybe_add_edge(
                                    f"edge_min_share_{int(round(min_share*1000))}",
                                    edge_cfg,
                                    waami_floor=strict_floor,
                                    edge_settings={"mode": "rent_max", "relaxed": True, "deep_affordability_min_share": float(min_share)},
                                ):
                                    break

                        # Edge 3: relax WAAMI floor only if we still don't have enough edge scenarios.
                        if len(edge_keys_added) < target_edge_count:
                            for floor in (0.589, 0.585, 0.58):
                                if float(floor) + 1e-12 < edge_min_waami:
                                    continue
                                if float(floor) >= strict_floor - 1e-12:
                                    continue
                                _maybe_add_edge(
                                    f"edge_waami_floor_{int(round(floor*1000))}",
                                    config,
                                    waami_floor=float(floor),
                                    edge_settings={"mode": "rent_max", "relaxed": True, "waami_floor": float(floor)},
                                )
                    elif program_norm == 'MIH':
                        # MIH relaxed scenarios: keep share rules fixed; only relax WAAMI floor down to 58%.
                        if len(edge_keys_added) < target_edge_count:
                            for floor in (0.59, 0.58):
                                if float(floor) + 1e-12 < edge_min_waami:
                                    continue
                                if float(floor) >= strict_floor - 1e-12:
                                    continue
                                _maybe_add_edge(
                                    f"edge_waami_floor_{int(round(floor*1000))}",
                                    config,
                                    waami_floor=float(floor),
                                    edge_settings={"mode": "rent_max", "relaxed": True, "waami_floor": float(floor)},
                                )
                                if len(edge_keys_added) >= target_edge_count:
                                    break

                    if target_edge_count > 0 and len(edge_keys_added) < target_edge_count:
                        notes.append(
                            f"Edge scenarios: generated {len(edge_keys_added)} of {target_edge_count}; remaining relaxations were infeasible or produced no distinct unit mix."
                        )

                    # Top-up: If we still have fewer than the target scenario count, append additional strict
                    # variants (e.g., client_oriented / alternative) so Excel users still see ~6 options.
                    if strict_scenarios_full and len(scenarios or {}) < target_total_count:
                        preferred_fill_keys = ["client_oriented", "alternative", "best_3_band", "best_2_band"]
                        for fill_key in preferred_fill_keys:
                            if len(scenarios) >= target_total_count:
                                break
                            if fill_key in scenarios:
                                continue
                            candidate = strict_scenarios_full.get(fill_key)
                            if not candidate:
                                continue
                            ck = _canon_key(candidate.get('canonical_assignments'))
                            if ck and ck in existing_canons:
                                continue
                            scenarios[fill_key] = candidate
                            if ck:
                                existing_canons.add(ck)
                except Exception as e:
                    notes.append(f"Warning: Could not compute edge scenarios: {str(e)}")
                timing["edge_scenarios_ms"] = int(round((time.perf_counter() - edge_start) * 1000))

        # --- Promote "Best Rent Roll" scenario (client request 2026-04) ---
        # Find the highest annual-rent scenario across ALL scenarios. If the
        # winner is currently keyed as edge_waami_floor_*, promote it to
        # best_rent_roll: clear its tradeoffs (the client wants this; not a
        # tradeoff), tag tier='rent_max', and delete the original edge key
        # to avoid duplicate display. If a strict scenario already has the
        # highest revenue, no promotion needed.
        try:
            best_rent_key = None
            best_rent_value = -1.0
            for _key, _scen in (scenarios or {}).items():
                if not _scen:
                    continue
                _rt = _scen.get("rent_totals") or {}
                _annual = float(_rt.get("net_annual") or _rt.get("total_annual_rent") or 0.0)
                if _annual > best_rent_value:
                    best_rent_value = _annual
                    best_rent_key = _key
            if best_rent_key and best_rent_key.startswith("edge_waami_floor_"):
                if "best_rent_roll" not in scenarios:
                    promoted = copy.deepcopy(scenarios[best_rent_key])
                    promoted["tradeoffs"] = []
                    promoted["tier"] = "rent_max"
                    promoted_settings = promoted.get("edge_settings") or {}
                    promoted_settings["promoted_to_best_rent_roll"] = True
                    promoted["edge_settings"] = promoted_settings
                    scenarios["best_rent_roll"] = promoted
                    del scenarios[best_rent_key]
                    waami_pct = float(promoted.get("waami") or 0.0) * 100
                    notes.append(
                        f"Promoted relaxed-floor scenario (WAAMI {waami_pct:.2f}%) to 'best_rent_roll' — annual rent ${best_rent_value:,.0f}."
                    )
        except Exception as e:
            notes.append(f"Warning: best_rent_roll promotion failed: {str(e)}")

        # --- Fix-02: De-dupe outcome-identical scenarios (post-processing) ---
        # Only remove scenarios when they are truly identical in outputs (band mix + rent totals),
        # then keep the best "placement" (40% units lower floors; higher AMI/higher rent higher floors).
        try:
            scenario_priority = [
                "absolute_best",
                "best_rent_roll",
                "best_3_band",
                "best_2_band",
                "alternative",
                "client_oriented",
                "max_revenue",
            ]
            priority_index = {k: i for i, k in enumerate(scenario_priority)}

            def _outcome_signature(scenario: dict) -> tuple | None:
                if not scenario:
                    return None

                metrics = scenario.get("metrics") or {}
                band_mix = metrics.get("band_mix") or []
                rent_totals = scenario.get("rent_totals") or {}

                # Be conservative: if we don't have rent totals, we can't safely claim outcomes match.
                if not band_mix:
                    return None
                if "net_monthly" not in rent_totals or "net_annual" not in rent_totals:
                    return None

                mix_items: list[tuple[int, int, float]] = []
                for item in band_mix:
                    if not isinstance(item, dict):
                        return None
                    try:
                        band = int(item.get("band"))
                    except Exception:
                        return None
                    try:
                        units = int(item.get("units") or 0)
                    except Exception:
                        units = 0
                    try:
                        net_sf = round(float(item.get("net_sf") or 0.0), 2)
                    except Exception:
                        net_sf = 0.0
                    mix_items.append((band, units, net_sf))
                mix_items.sort(key=lambda x: x[0])

                try:
                    net_monthly = round(float(rent_totals.get("net_monthly")), 2)
                    net_annual = round(float(rent_totals.get("net_annual")), 2)
                except Exception:
                    return None

                try:
                    total_sf = round(float(metrics.get("total_sf") or 0.0), 2)
                except Exception:
                    total_sf = 0.0
                try:
                    waami = round(float(scenario.get("waami") or 0.0), 10)
                except Exception:
                    waami = 0.0

                return (tuple(mix_items), net_monthly, net_annual, total_sf, waami)

            def _placement_key(scenario: dict) -> tuple:
                assignments = (scenario or {}).get("assignments") or []
                floors: list[float] = []
                bands: list[float] = []
                rents: list[float] = []

                for unit in assignments:
                    if not isinstance(unit, dict):
                        continue
                    try:
                        floor = float(unit.get("floor") or 0.0)
                    except Exception:
                        floor = 0.0
                    try:
                        band = float(unit.get("assigned_ami") or 0.0)
                    except Exception:
                        band = 0.0
                    try:
                        rent = float(unit.get("monthly_rent") or 0.0)
                    except Exception:
                        rent = 0.0
                    floors.append(floor)
                    bands.append(band)
                    rents.append(rent)

                # Count "inversions" where lower AMI is placed above higher AMI (or vice versa).
                inversions = 0
                n = len(floors)
                for i in range(n):
                    for j in range(i + 1, n):
                        if (bands[i] < bands[j] and floors[i] > floors[j]) or (bands[i] > bands[j] and floors[i] < floors[j]):
                            inversions += 1

                align_ami = sum(f * b for f, b in zip(floors, bands))
                align_rent = sum(f * r for f, r in zip(floors, rents))
                low40_floor_sum = sum(f for f, b in zip(floors, bands) if b <= 0.4000001)

                # Higher is better (fewer inversions; better alignment; 40% units lower floors).
                return (-inversions, align_ami, align_rent, -low40_floor_sum)

            groups: dict[tuple, list[str]] = {}
            for key, scenario in (scenarios or {}).items():
                sig = _outcome_signature(scenario)
                if sig is None:
                    continue
                groups.setdefault(sig, []).append(str(key))

            removed = 0
            if groups:
                for sig, keys in groups.items():
                    if len(keys) < 2:
                        continue

                    def _keeper_sort_key(k: str) -> tuple:
                        return (priority_index.get(k, 10_000), k)

                    keep_key = sorted(keys, key=_keeper_sort_key)[0]

                    # Prefer non-edge scenarios when available, then apply placement tie-break.
                    candidates = []
                    for k in keys:
                        sc = (scenarios or {}).get(k)
                        if not sc:
                            continue
                        tier = str(sc.get("tier") or "").strip().lower()
                        is_edge = (tier == "edge")
                        candidates.append((is_edge, _placement_key(sc), k, sc))
                    if not candidates:
                        continue

                    strict_candidates = [c for c in candidates if not c[0]]
                    pool = strict_candidates if strict_candidates else candidates
                    winner = max(
                        pool,
                        key=lambda c: (
                            c[1],
                            -priority_index.get(c[2], 10_000),
                            c[2],
                        ),
                    )
                    winner_key = winner[2]
                    winner_scenario = winner[3]

                    scenarios[keep_key] = winner_scenario
                    for k in keys:
                        if k == keep_key:
                            continue
                        if k in scenarios:
                            del scenarios[k]
                            removed += 1

            if removed > 0:
                notes.append(f"De-duped {removed} outcome-identical scenario(s) (Fix-02).")
        except Exception:
            pass

        learning_info = None
        if compare_baseline and baseline_scenarios and scenarios:
            def _scenario_snapshot(scenarios_dict: dict) -> dict:
                abs_best = scenarios_dict.get('absolute_best') or {}
                return {
                    "scenario_keys": sorted(list(scenarios_dict.keys())),
                    "absolute_best": {
                        "waami": float(abs_best.get('waami') or 0.0),
                        "waami_percent": float((abs_best.get('metrics') or {}).get('waami_percent') or 0.0),
                        "revenue_score": float(abs_best.get('revenue_score') or 0.0),
                        "rent_score": abs_best.get('rent_score'),
                        "canonical_assignments": abs_best.get('canonical_assignments'),
                    }
                }

            base_snap = _scenario_snapshot(baseline_scenarios)
            learned_snap = _scenario_snapshot(scenarios)

            base_canon = base_snap.get("absolute_best", {}).get("canonical_assignments") or []
            learned_canon = learned_snap.get("absolute_best", {}).get("canonical_assignments") or []
            base_map = {str(u): int(b) for u, b in base_canon if u is not None and b is not None}
            learned_map = {str(u): int(b) for u, b in learned_canon if u is not None and b is not None}
            changed_units = []
            for unit_id in sorted(set(base_map.keys()) | set(learned_map.keys())):
                if base_map.get(unit_id) != learned_map.get(unit_id):
                    changed_units.append({
                        "unit_id": unit_id,
                        "baseline_band": base_map.get(unit_id),
                        "learned_band": learned_map.get(unit_id),
                    })

            learning_info = {
                "compare_baseline": True,
                "project_overrides": project_overrides,
                "baseline": base_snap,
                "learned": learned_snap,
                "diff": {
                    "changed_unit_count": len(changed_units),
                    "changed_units": changed_units[:200],
                },
            }

        # FINAL HARD WAAMI CAP enforcement (belt + suspenders): scan every
        # scenario from every source and drop any whose WAAMI exceeds the
        # legal cap by even a fraction of a percent. Client requirement.
        try:
            cap_fraction_final = float((config.get('optimization_rules', {}) or {}).get('waami_cap_percent') or 60.0) / 100.0
            dropped_for_cap = []
            for _sk in list(scenarios.keys()):
                _sv = scenarios.get(_sk)
                if not _sv or 'assignments' not in _sv:
                    continue
                _a = _sv['assignments'] or []
                _tsf = sum(float(_u.get('net_sf', 0) or 0) for _u in _a)
                if _tsf <= 0:
                    continue
                _w = sum(float(_u.get('net_sf', 0) or 0) * float(_u.get('assigned_ami', 0) or 0) for _u in _a) / _tsf
                if _w > cap_fraction_final:
                    dropped_for_cap.append((_sk, _w))
                    del scenarios[_sk]
            for _k, _w in dropped_for_cap:
                notes.append(
                    f"Removed scenario '{_k}' from results: WAAMI {_w*100:.6f}% exceeds the {cap_fraction_final*100:.2f}% cap (client requires strict compliance)."
                )
        except Exception as _e:
            notes.append(f"Note: WAAMI cap enforcement pass failed: {_e}")

        # "Original Scenario": capture the client's AMI assignments as they were
        # in the workbook when this request fired, so the Scenarios page shows
        # one entry that's literally their input ("what I had before running
        # the program"). Not optimized by the solver — just computed for rent
        # so they can compare. Added AFTER the WAAMI cap pass so even a
        # non-compliant pre-saved scenario still appears for visibility.
        try:
            if rent_schedule:
                orig_assignments = []
                for _idx, _u in df_units.iterrows():
                    _ami = _u.get('client_ami')
                    if _ami is None:
                        continue
                    try:
                        _ami_f = float(_ami)
                    except (TypeError, ValueError):
                        continue
                    if _ami_f <= 0:
                        continue
                    orig_assignments.append({
                        'unit_id': str(_u.get('unit_id', '')),
                        'assigned_ami': _ami_f,
                        'bedrooms': _u.get('bedrooms'),
                        'net_sf': float(_u.get('net_sf') or 0.0),
                    })
                if orig_assignments:
                    _enriched, _totals = compute_rents_for_assignments(
                        rent_schedule, orig_assignments, utilities_clean
                    )
                    _orig_total_sf = sum(float(_a.get('net_sf', 0) or 0) for _a in _enriched)
                    _orig_waami = 0.0
                    if _orig_total_sf > 0:
                        _orig_waami = sum(
                            float(_a.get('net_sf', 0) or 0) * float(_a.get('assigned_ami', 0) or 0)
                            for _a in _enriched
                        ) / _orig_total_sf
                    _orig_bands = sorted(set(int(round(float(_a['assigned_ami']) * 100)) for _a in _enriched))
                    scenarios['original'] = {
                        'name': 'Original Scenario',
                        'description': "Your saved AMI assignments from the workbook (not optimized; shown for comparison).",
                        'assignments': _enriched,
                        'rent_totals': _totals,
                        'waami': _orig_waami,
                        'bands': _orig_bands,
                        'status': 'CLIENT_ORIGINAL',
                        'tier': 'reference',
                    }
        except Exception as _e:
            notes.append(f"Note: could not build Original Scenario: {_e}")

        # Build response
        response_start = time.perf_counter()
        safe_scenarios = _sanitize_for_json(scenarios)
        project_summary = {
                "total_units": len(df_units),
                "total_sf": float(df_units['net_sf'].sum()),
                "utility_selections": utilities_clean,
                "program": (program or 'UAP'),
                "mih_option": mih_option,
                "mih_residential_sf": mih_residential_sf,
        }
        if mih_constraint_injected:
            project_summary["total_building_sf"] = total_building_sf

        response = {
            "success": True,
            "scenarios": safe_scenarios,
            "notes": notes,
            "project_summary": project_summary,
        }

        if learning_info:
            response["learning"] = _sanitize_for_json(learning_info)

        timing["response_sanitize_ms"] = int(round((time.perf_counter() - response_start) * 1000))
        timing["total_elapsed_ms"] = int(round((time.perf_counter() - request_start) * 1000))
        if timing_enabled:
            response["timing"] = timing
        _emit_timing("ok", {"scenario_count": int(len(scenarios or {}))})
        return jsonify(response)

    except Exception as e:
        _emit_timing("exception", {"error": str(e)})
        app.logger.exception("optimize_units failed: %s", e)
        return jsonify({"error": f"Optimization failed: {str(e)}"}), 500


@app.route('/api/evaluate', methods=['POST'])
def evaluate_assignment():
    """
    Validate a user-specified unit assignment and compute rents/totals.

    Request body:
    {
      "rent_roll_year": 2025,
      "calculator_id": "AMI_Optix_Rent_Calculator_2025.xlsx" | "default",
      "program": "UAP" | "MIH",
      "mih_option": "Option 1" | "Option 4",
      "mih_residential_sf": 46197.57,
      "mih_max_band_percent": 135,
      "units": [
        {"unit_id": "207", "bedrooms": 2, "net_sf": 790.58, "assigned_ami": 0.6}
      ],
      "utilities": {...}
    }
    """
    auth_error = _validate_api_key()
    if auth_error:
        return auth_error

    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "Request body must be JSON"}), 400
    except Exception as e:
        return jsonify({"error": f"Invalid JSON: {str(e)}"}), 400

    units = data.get('units', [])
    if not units or not isinstance(units, list):
        return jsonify({"error": "Missing or invalid 'units' array"}), 400

    utilities = data.get('utilities', {})
    utilities_clean = {
        'electricity': utilities.get('electricity', 'na'),
        'cooking': utilities.get('cooking', 'na'),
        'heat': utilities.get('heat', 'na'),
        'hot_water': utilities.get('hot_water', 'na'),
    }

    program = (data.get('program') or 'UAP')
    mih_option = data.get('mih_option')
    mih_residential_sf = data.get('mih_residential_sf')
    mih_max_band_percent = data.get('mih_max_band_percent')

    rent_calc_path, rent_meta = _resolve_rent_calculator_for_request(data)
    # Default these fields so even early validation failures can report selection state.
    rent_meta.setdefault("rent_schedule_loaded", False)
    rent_meta.setdefault("rent_schedule_cache_hit", False)

    try:
        df_units = pd.DataFrame(units)
        if 'assigned_ami' not in df_units.columns:
            return jsonify({"error": "Each unit must include 'assigned_ami'."}), 400

        df_units['unit_id'] = df_units['unit_id'].astype(str)
        df_units['bedrooms'] = pd.to_numeric(df_units.get('bedrooms'), errors='coerce')
        df_units['net_sf'] = pd.to_numeric(df_units.get('net_sf'), errors='coerce')
        df_units['assigned_ami'] = pd.to_numeric(df_units.get('assigned_ami'), errors='coerce')

        config = load_config()
        try:
            config = _build_program_config(
                config,
                program=program,
                mih_option=mih_option,
                mih_residential_sf=mih_residential_sf,
                mih_max_band_percent=mih_max_band_percent,
            )
        except ValueError as e:
            resp = {
                "success": False,
                "errors": [str(e)],
                "summary": {},
            }
            resp.update(rent_meta)
            return jsonify(resp), 200

        # Validate constraints first.
        is_valid, errors, summary = _validate_assignment_payload(units, config)
        if not is_valid:
            resp = {
                "success": False,
                "errors": errors,
                "summary": summary,
            }
            resp.update(rent_meta)
            return jsonify(resp), 200

        rent_schedule = None
        cache_hit = False
        if rent_calc_path:
            rent_schedule, cache_hit = _load_rent_schedule_cached(rent_calc_path)
        rent_meta["rent_schedule_loaded"] = bool(rent_schedule)
        rent_meta["rent_schedule_cache_hit"] = bool(cache_hit)

        assignments = df_units.to_dict(orient='records')
        # Normalize assigned_ami to the 0-2 range (support 60/120 inputs as well as 0.6/1.2).
        for a in assignments:
            try:
                v = float(a.get('assigned_ami'))
                if v > 2.0:
                    a['assigned_ami'] = v / 100.0
            except Exception:
                pass

        metrics = _build_metrics_from_assignments(assignments)
        rent_totals = None
        if rent_schedule:
            assignments, rent_totals = compute_rents_for_assignments(rent_schedule, assignments, utilities_clean)
            try:
                metrics['total_monthly_rent'] = float((rent_totals or {}).get('net_monthly') or 0.0)
                metrics['total_annual_rent'] = float((rent_totals or {}).get('net_annual') or 0.0)
            except Exception:
                pass

        resp = {
            "success": True,
            "summary": summary,
            "metrics": _sanitize_for_json(metrics),
            "assignments": _sanitize_for_json(assignments),
            "rent_totals": _sanitize_for_json(rent_totals),
            "utility_selections": utilities_clean,
            "program": program,
            "mih_option": mih_option,
            "mih_residential_sf": mih_residential_sf,
        }
        resp.update(rent_meta)
        return jsonify(resp)

    except Exception as e:
        app.logger.exception("evaluate_assignment failed: %s", e)
        return jsonify({"error": f"Evaluation failed: {str(e)}"}), 500


@app.route('/api/manual_calculate', methods=['POST'])
def manual_calculate_assignment():
    """
    Compute rents/totals + diagnostics for a user-specified assignment, even if it violates program rules.

    This powers the Excel add-in "Live Sync OFF" workflow where the user types AMIs directly into the
    program sheet and clicks a button to compute the resulting rents/band mix without auto-reverting.
    """
    auth_error = _validate_api_key()
    if auth_error:
        return auth_error

    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "Request body must be JSON"}), 400
    except Exception as e:
        return jsonify({"error": f"Invalid JSON: {str(e)}"}), 400

    units = data.get('units', [])
    if not units or not isinstance(units, list):
        return jsonify({"error": "Missing or invalid 'units' array"}), 400

    utilities = data.get('utilities', {})
    utilities_clean = {
        'electricity': utilities.get('electricity', 'na'),
        'cooking': utilities.get('cooking', 'na'),
        'heat': utilities.get('heat', 'na'),
        'hot_water': utilities.get('hot_water', 'na'),
    }

    program = (data.get('program') or 'UAP')
    mih_option = data.get('mih_option')
    mih_residential_sf = data.get('mih_residential_sf')
    mih_max_band_percent = data.get('mih_max_band_percent')

    try:
        df_units = pd.DataFrame(units)
        if 'assigned_ami' not in df_units.columns:
            return jsonify({"error": "Each unit must include 'assigned_ami'."}), 400

        df_units['unit_id'] = df_units['unit_id'].astype(str)
        df_units['bedrooms'] = pd.to_numeric(df_units.get('bedrooms'), errors='coerce')
        df_units['net_sf'] = pd.to_numeric(df_units.get('net_sf'), errors='coerce')
        df_units['assigned_ami'] = pd.to_numeric(df_units.get('assigned_ami'), errors='coerce')

        # Normalize assigned_ami to 0-2 range (support 60/120 inputs).
        df_units.loc[df_units['assigned_ami'] > 2.0, 'assigned_ami'] = df_units['assigned_ami'] / 100.0

        config = load_config()
        try:
            config = _build_program_config(
                config,
                program=program,
                mih_option=mih_option,
                mih_residential_sf=mih_residential_sf,
                mih_max_band_percent=mih_max_band_percent,
            )
        except ValueError as e:
            return jsonify({
                "success": False,
                "tradeoffs": [str(e)],
                "summary": {},
                "metrics": {},
                "assignments": [],
                "rent_totals": None,
            }), 200

        is_valid, errors, summary = _validate_assignment_payload(units, config)

        assignments = df_units.to_dict(orient='records')
        metrics = _build_metrics_from_assignments(assignments)

        rent_calc_path, rent_meta = _resolve_rent_calculator_for_request(data)
        rent_schedule = None
        if rent_calc_path:
            rent_schedule, _cache_hit = _load_rent_schedule_cached(rent_calc_path)

        rent_totals = None
        if rent_schedule:
            assignments, rent_totals = compute_rents_for_assignments(rent_schedule, assignments, utilities_clean)
            try:
                metrics['total_monthly_rent'] = float((rent_totals or {}).get('net_monthly') or 0.0)
                metrics['total_annual_rent'] = float((rent_totals or {}).get('net_annual') or 0.0)
            except Exception:
                pass

        resp = {
            "success": True,
            "is_valid": bool(is_valid),
            "tradeoffs": [] if is_valid else errors[:12],
            "summary": summary,
            "metrics": _sanitize_for_json(metrics),
            "assignments": _sanitize_for_json(assignments),
            "rent_totals": _sanitize_for_json(rent_totals),
            "utility_selections": utilities_clean,
            "program": program,
            "mih_option": mih_option,
            "mih_residential_sf": mih_residential_sf,
        }
        resp.update(rent_meta)
        return jsonify(resp)

    except Exception as e:
        app.logger.exception("manual_calculate_assignment failed: %s", e)
        return jsonify({"error": f"Manual calculation failed: {str(e)}"}), 500


@app.route('/api/analyze', methods=['POST'])
def analyze_file():
    # Lock down file-upload endpoints if you are not using the public dashboard flow.
    auth_error = _validate_api_key()
    if auth_error:
        return auth_error

    # Prevent disk bloat from old zips (best-effort).
    _cleanup_uploads_dir()

    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected for uploading"}), 400

    if file:
        filename = secure_filename(file.filename)

        # Use a temporary directory for all processing
        with tempfile.TemporaryDirectory() as temp_dir:
            upload_filepath = os.path.join(temp_dir, filename)
            file.save(upload_filepath)

            utilities_payload = None
            overrides_payload = None
            rent_calculator_path = None

            utilities_raw = request.form.get('utilities')
            if utilities_raw:
                try:
                    utilities_payload = json.loads(utilities_raw)
                except json.JSONDecodeError:
                    return jsonify({"error": "Invalid utilities payload."}), 400

            overrides_raw = request.form.get('overrides')
            if overrides_raw:
                try:
                    overrides_payload = json.loads(overrides_raw)
                except json.JSONDecodeError:
                    return jsonify({"error": "Invalid overrides payload."}), 400

            rent_calculator_upload = request.files.get('rentCalculator')
            if rent_calculator_upload and rent_calculator_upload.filename:
                rent_filename = secure_filename(rent_calculator_upload.filename)
                rent_calculator_path = os.path.join(temp_dir, rent_filename)
                rent_calculator_upload.save(rent_calculator_path)

            try:
                # 1. Run the core analysis
                analysis_output = run_ami_optix_analysis(
                    upload_filepath,
                    utilities=utilities_payload,
                    overrides=overrides_payload,
                    rent_calculator_path=rent_calculator_path,
                )
                if "error" in analysis_output:
                    return jsonify(analysis_output), 400

                analysis_results = analysis_output['results']
                original_headers = analysis_output['original_headers']
                analysis_meta = analysis_results.get('analysis_meta', {})
                app.logger.info("analysis_id=%s combos=%s unique=%s duration=%.2fs truncated=%s", analysis_meta.get('analysis_id'), analysis_meta.get('solver_combination_count'), analysis_meta.get('solver_unique_scenarios'), analysis_meta.get('duration_sec'), analysis_meta.get('truncated'))

                # 2. Generate the internal summary (no LLM for now)
                narrative = generate_internal_summary(analysis_results)
                analysis_results['narrative_analysis'] = narrative

                # 3. Generate Excel reports, passing the original headers
                prefer_xlsb = filename.lower().endswith('.xlsb') or bool(request.form.get('preferXlsb'))
                utility_selections = analysis_results.get('project_summary', {}).get('utility_selections')
                rent_workbook_info = analysis_results.get('rent_workbook') or {}
                rent_workbook_source = rent_workbook_info.get('source_path')

                report_files = create_excel_reports(
                    analysis_results,
                    upload_filepath,
                    original_headers,
                    output_dir=temp_dir,
                    prefer_xlsb=prefer_xlsb,
                    utilities=utility_selections,
                    rent_workbook_path=rent_workbook_source,
                )

                # 4. Create a zip file containing all reports
                zip_filename = f"{os.path.splitext(filename)[0]}_reports.zip"
                zip_filepath = os.path.join(temp_dir, zip_filename)
                with zipfile.ZipFile(zip_filepath, 'w') as zipf:
                    for report_file in report_files:
                        zipf.write(report_file, os.path.basename(report_file))

                # 5. Add a download link for the zip file to the response
                analysis_results['download_link'] = f"/api/download/{zip_filename}"

                # Persist the zip in the uploads directory for the download endpoint
                os.makedirs(UPLOADS_DIR, exist_ok=True)
                shutil.move(zip_filepath, os.path.join(UPLOADS_DIR, zip_filename))

                safe_payload = _sanitize_for_json(analysis_results)

                return jsonify(safe_payload)

            except Exception as e:
                app.logger.exception("analysis_failed: %s", e)
                return jsonify({"error": f"An unexpected error occurred during analysis: {str(e)}"}), 500

    return jsonify({"error": "An unknown error occurred"}), 500


@app.route('/api/download/<filename>', methods=['GET'])
def download_report(filename):
    """Serves the generated zip file for download."""
    auth_error = _validate_api_key()
    if auth_error:
        return auth_error

    filename = secure_filename(filename)
    try:
        return send_from_directory(UPLOADS_DIR, filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found."}), 404


# =============================================================================
# RENT CALCULATOR MANAGEMENT ENDPOINTS
# =============================================================================

@app.route('/api/rent-calculators', methods=['GET'])
def list_rent_calculators():
    """List all available rent calculator files."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    calculators = _list_rent_calculators()
    active_path = _get_active_rent_calculator_path()

    return jsonify({
        "calculators": calculators,
        "active_path": active_path,
        "storage_dir": RENT_CALCULATORS_DIR
    })


@app.route('/api/rent-calculators/upload', methods=['POST'])
def upload_rent_calculator():
    """Upload a new rent calculator file."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    # Validate file extension
    filename = secure_filename(file.filename)
    if not filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({"error": "Only .xlsx and .xlsm files are supported"}), 400

    # Save the file
    filepath = os.path.join(RENT_CALCULATORS_DIR, filename)

    # Check if file already exists
    if os.path.exists(filepath):
        overwrite = request.form.get('overwrite', 'false').lower() == 'true'
        if not overwrite:
            return jsonify({"error": f"File '{filename}' already exists. Set overwrite=true to replace."}), 409

    try:
        file.save(filepath)

        # Validate the file is a valid rent calculator by trying to load it
        try:
            schedule = load_rent_schedule(filepath)
            # Basic validation - check it has rent data
            if not schedule.gross_rents:
                os.remove(filepath)
                return jsonify({"error": "File does not appear to be a valid rent calculator (no rent data found)"}), 400
        except Exception as e:
            os.remove(filepath)
            return jsonify({"error": f"Invalid rent calculator file: {str(e)}"}), 400

        return jsonify({
            "success": True,
            "message": f"Uploaded {filename}",
            "filename": filename
        })

    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500


@app.route('/api/rent-calculators/activate', methods=['POST'])
def activate_rent_calculator():
    """Set the active rent calculator."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    data = request.get_json()
    if not data or 'name' not in data:
        return jsonify({"error": "Missing 'name' in request body"}), 400

    name = data['name']

    # Check if it's the default (clear the active selection)
    if name == '2025 AMI Rent Calculator Unlocked.xlsx' or name == 'default':
        if os.path.exists(ACTIVE_CALCULATOR_FILE):
            os.remove(ACTIVE_CALCULATOR_FILE)
        return jsonify({
            "success": True,
            "message": "Activated default rent calculator",
            "active": "2025 AMI Rent Calculator Unlocked.xlsx"
        })

    # Verify the file exists
    filepath = os.path.join(RENT_CALCULATORS_DIR, name)
    if not os.path.exists(filepath):
        return jsonify({"error": f"Rent calculator '{name}' not found"}), 404

    # Set as active
    with open(ACTIVE_CALCULATOR_FILE, 'w') as f:
        f.write(name)

    return jsonify({
        "success": True,
        "message": f"Activated {name}",
        "active": name
    })


@app.route('/api/rent-calculators/<filename>', methods=['DELETE'])
def delete_rent_calculator(filename):
    """Delete a rent calculator file."""
    auth_error = _validate_admin_key()
    if auth_error:
        return auth_error

    # Cannot delete the default
    if filename == '2025 AMI Rent Calculator Unlocked.xlsx':
        return jsonify({"error": "Cannot delete the default rent calculator"}), 400

    filepath = os.path.join(RENT_CALCULATORS_DIR, secure_filename(filename))
    if not os.path.exists(filepath):
        return jsonify({"error": f"File '{filename}' not found"}), 404

    # If this was the active calculator, clear the selection
    if os.path.exists(ACTIVE_CALCULATOR_FILE):
        with open(ACTIVE_CALCULATOR_FILE, 'r') as f:
            active_name = f.read().strip()
        if active_name == filename:
            os.remove(ACTIVE_CALCULATOR_FILE)

    os.remove(filepath)

    return jsonify({
        "success": True,
        "message": f"Deleted {filename}"
    })


@app.route('/admin/rent-calculators')
def rent_calculator_admin():
    """Simple admin page for rent calculator management."""
    html = """
<!DOCTYPE html>
<html>
<head>
    <title>AMI Optix - Rent Calculator Admin</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 800px; margin: 40px auto; padding: 20px; }
        h1 { color: #2c3e50; }
        .card { background: #f8f9fa; border-radius: 8px; padding: 20px; margin: 20px 0; }
        .active { background: #d4edda; border-left: 4px solid #28a745; }
        .default { background: #e7f3ff; border-left: 4px solid #007bff; }
        button { background: #007bff; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; margin-right: 8px; }
        button:hover { background: #0056b3; }
        button.danger { background: #dc3545; }
        button.danger:hover { background: #c82333; }
        button.success { background: #28a745; }
        input[type="file"] { margin: 10px 0; }
        .status { padding: 10px; border-radius: 4px; margin: 10px 0; }
        .status.error { background: #f8d7da; color: #721c24; }
        .status.success { background: #d4edda; color: #155724; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; }
        .badge.active { background: #28a745; color: white; }
        .badge.default { background: #007bff; color: white; }
    </style>
</head>
<body>
    <h1>🏠 Rent Calculator Admin</h1>

    <div class="card">
        <h2>Upload New Calculator</h2>
        <p>Upload a new rent calculator Excel file (.xlsx or .xlsm)</p>
        <input type="file" id="fileInput" accept=".xlsx,.xlsm">
        <br>
        <label><input type="checkbox" id="overwrite"> Overwrite if exists</label>
        <br><br>
        <button onclick="uploadFile()">Upload</button>
        <div id="uploadStatus"></div>
    </div>

    <div class="card">
        <h2>Available Calculators</h2>
        <div id="calculatorList">Loading...</div>
    </div>

    <script>
        const apiKey = prompt('Enter API Key:');

        async function fetchCalculators() {
            try {
                const res = await fetch('/api/rent-calculators', {
                    headers: { 'X-API-Key': apiKey }
                });
                const data = await res.json();
                if (data.error) {
                    document.getElementById('calculatorList').innerHTML = '<p class="status error">' + data.error + '</p>';
                    return;
                }
                renderCalculators(data.calculators);
            } catch (e) {
                document.getElementById('calculatorList').innerHTML = '<p class="status error">Failed to load: ' + e.message + '</p>';
            }
        }

        function renderCalculators(calculators) {
            if (!calculators || calculators.length === 0) {
                document.getElementById('calculatorList').innerHTML = '<p>No calculators found.</p>';
                return;
            }

            let html = '<table><tr><th>Name</th><th>Source</th><th>Size</th><th>Actions</th></tr>';
            calculators.forEach(calc => {
                const badges = [];
                if (calc.is_active) badges.push('<span class="badge active">Active</span>');
                if (calc.source === 'default') badges.push('<span class="badge default">Default</span>');

                const actions = [];
                if (!calc.is_active) {
                    actions.push('<button onclick="activateCalc(\\'' + calc.name + '\\')">Activate</button>');
                }
                if (calc.source !== 'default') {
                    actions.push('<button class="danger" onclick="deleteCalc(\\'' + calc.name + '\\')">Delete</button>');
                }

                html += '<tr class="' + (calc.is_active ? 'active' : '') + ' ' + (calc.source === 'default' ? 'default' : '') + '">';
                html += '<td>' + calc.name + ' ' + badges.join(' ') + '</td>';
                html += '<td>' + calc.source + '</td>';
                html += '<td>' + Math.round(calc.size / 1024) + ' KB</td>';
                html += '<td>' + actions.join('') + '</td>';
                html += '</tr>';
            });
            html += '</table>';
            document.getElementById('calculatorList').innerHTML = html;
        }

        async function uploadFile() {
            const fileInput = document.getElementById('fileInput');
            const overwrite = document.getElementById('overwrite').checked;
            const statusDiv = document.getElementById('uploadStatus');

            if (!fileInput.files[0]) {
                statusDiv.innerHTML = '<p class="status error">Please select a file</p>';
                return;
            }

            const formData = new FormData();
            formData.append('file', fileInput.files[0]);
            formData.append('overwrite', overwrite);

            statusDiv.innerHTML = '<p>Uploading...</p>';

            try {
                const res = await fetch('/api/rent-calculators/upload', {
                    method: 'POST',
                    headers: { 'X-API-Key': apiKey },
                    body: formData
                });
                const data = await res.json();
                if (data.error) {
                    statusDiv.innerHTML = '<p class="status error">' + data.error + '</p>';
                } else {
                    statusDiv.innerHTML = '<p class="status success">' + data.message + '</p>';
                    fetchCalculators();
                }
            } catch (e) {
                statusDiv.innerHTML = '<p class="status error">Upload failed: ' + e.message + '</p>';
            }
        }

        async function activateCalc(name) {
            try {
                const res = await fetch('/api/rent-calculators/activate', {
                    method: 'POST',
                    headers: {
                        'X-API-Key': apiKey,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ name: name })
                });
                const data = await res.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    fetchCalculators();
                }
            } catch (e) {
                alert('Failed: ' + e.message);
            }
        }

        async function deleteCalc(name) {
            if (!confirm('Delete ' + name + '?')) return;

            try {
                const res = await fetch('/api/rent-calculators/' + encodeURIComponent(name), {
                    method: 'DELETE',
                    headers: { 'X-API-Key': apiKey }
                });
                const data = await res.json();
                if (data.error) {
                    alert('Error: ' + data.error);
                } else {
                    fetchCalculators();
                }
            } catch (e) {
                alert('Failed: ' + e.message);
            }
        }

        fetchCalculators();
    </script>
</body>
</html>
"""
    return html, 200, {'Content-Type': 'text/html'}


@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def serve_dashboard(path):
    """Serve the static dashboard that ships with the deployment."""
    # Guard API routes so they keep flowing to their own handlers
    if path.startswith('api/'):
        return jsonify({"error": "Not found"}), 404

    requested = path or 'index.html'
    requested_path = os.path.join(DASHBOARD_DIR, requested)

    if os.path.isdir(requested_path):
        requested = os.path.join(requested, 'index.html')

    if _dashboard_file_exists(requested):
        return send_from_directory(DASHBOARD_DIR, requested)

    if _dashboard_file_exists('index.html'):
        # SPA fallback - return index so client-side routing can take over
        return send_from_directory(DASHBOARD_DIR, 'index.html')

    return (
        "<html><head><title>NYC AMI Calculator API</title></head><body>"
        "<h1>NYC AMI Calculator API</h1>"
        "<p>The interactive dashboard has not been built. "
        "Deploy tip: run 'npm install' and 'npm run build && npm run export' inside the"
        " dashboard/ folder, then deploy the contents of the generated dashboard_static/"
        " directory alongside this service.</p>"
        "</body></html>",
        200,
        {"Content-Type": "text/html; charset=utf-8"}
    )


if __name__ == '__main__':
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    app.run(debug=True, port=5001)

