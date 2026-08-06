from __future__ import annotations
import os, sys, time, subprocess, re, math, platform, struct
import comtypes, comtypes.client
from dataclasses import dataclass
from typing import List, Literal, Optional, Tuple
from itertools import product, groupby

DO_BEAMS    = 1
DO_COLUMNS  = 0

# =========================
# Runtime paths and compatibility
# =========================
APP_VERSION = "0.1.0-alpha"
SUPPORTED_ETABS_MAJOR_VERSIONS = {20, 21}
ETABS_EXE_CANDIDATES = (
    r"C:\Program Files\Computers and Structures\ETABS 21\ETABS.exe",
    r"C:\Program Files\Computers and Structures\ETABS 20\ETABS.exe",
)

# Temporary working units; the model's original units are restored at the end.
# Options: kgf-cm-C, kN-mm-C, kN-m-C, kip-in-F, kip-ft-F
WORKING_UNITS = "kgf-cm-C"
REBAR_SYSTEM = "metric"  # Use "metric" or "us"; both input formats are detected when possible.
MAX_REBAR_DIAMETER_MISMATCH_MM = 0.80
VERIFY_COLUMN_REBAR_ROUNDTRIP = True

def log(*a): print(*a)

# =========================
# Shared utilities and caches
# =========================
_CACHE = {
    "rebar_list": None,          # Cached reinforcing-bar names
    "rebar_list_with_data": None,
    "materials_by_type": {},     # {type: [names]}
    "units_set": False,
    "original_units": None,
    "length_enum": 5,
}
_TR_F2E = str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
_RE_REBAR_MM = re.compile(r"(?:[dDtTφΦøØ]\s*)?(\d+(?:\.\d+)?)\s*[dD]?$")
_RE_US_BAR = re.compile(r"^#\s*(\d+)$")

# force, length, temperature, packed eUnits, mm per one ETABS length unit
UNIT_PRESETS = {
    "kgf-cm-C": (5, 5, 2, 14, 10.0),
    "kN-mm-C":  (4, 4, 2, 5, 1.0),
    "kN-m-C":   (4, 6, 2, 6, 1000.0),
    "kip-in-F": (2, 1, 1, 3, 25.4),
    "kip-ft-F": (2, 2, 1, 4, 304.8),
}
_MM_PER_LENGTH_ENUM = {1: 25.4, 2: 304.8, 3: 0.001, 4: 1.0, 5: 10.0, 6: 1000.0}
US_REBAR_DIAMETERS_MM = {
    3: 9.525, 4: 12.700, 5: 15.875, 6: 19.050, 7: 22.225,
    8: 25.400, 9: 28.650, 10: 32.260, 11: 35.810, 14: 43.000, 18: 57.330,
}

def _cache_reset():
    _CACHE["rebar_list"] = None
    _CACHE["rebar_list_with_data"] = None
    _CACHE["materials_by_type"].clear()
    # Preserve units_set.

# =========================
# ETABS connection with attach polling
# =========================
def _comtypes_cache_command():
    return f'"{sys.executable}" -m comtypes.clear_cache'


def _existing_etabs_exes():
    return [p for p in ETABS_EXE_CANDIDATES if os.path.isfile(p)]


def _query_helper_interface(helper, module_name):
    try:
        module = getattr(comtypes.gen, module_name)
        return helper.QueryInterface(module.cHelper)
    except Exception:
        # Some installations generate the wrapper in memory, where direct dispatch is sufficient.
        return helper


def link_etabs(launch_if_not_found: bool = True, verbose: bool = True):
    def vprint(msg):
        if verbose: log(msg)
    # 1) active
    try:
        vprint("[1] GetActiveObject …")
        etabs = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        sm = etabs.SapModel
        vprint("   → attached via ROT.")
        return sm
    except Exception as e:
        vprint(f"   × active failed: {e}")
    # 2) ETABS v19+ helper: primary path for ETABS 20 and 21
    try:
        vprint("[2] ETABSv1.Helper …")
        helper = comtypes.client.CreateObject("ETABSv1.Helper")
        helper = _query_helper_interface(helper, "ETABSv1")
        try:
            etabs = helper.GetObject("CSI.ETABS.API.ETABSObject")
            sm = etabs.SapModel
            vprint("   → helper attached.")
            return sm
        except Exception as e:
            vprint(f"   · helper attach failed: {e}")
            if launch_if_not_found:
                try:
                    vprint("   · helper CreateObjectProgID …")
                    etabs = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
                    ret = etabs.ApplicationStart()
                    if isinstance(ret, int) and ret != 0:
                        vprint(f"   · ApplicationStart ret={ret}")
                    start = time.time()
                    while time.time() - start < 8.0:
                        try:
                            sm = etabs.SapModel
                            vprint("   → helper launched via ProgID.")
                            return sm
                        except Exception:
                            time.sleep(0.25)
                except Exception as ee:
                    vprint(f"   · CreateObjectProgID failed: {ee}")
                # The ETABS helper CreateObject method expects the full executable path, not a ProgID.
                for exe in _existing_etabs_exes():
                    try:
                        vprint(f"   · helper CreateObject path → {exe}")
                        etabs = helper.CreateObject(exe)
                        ret = etabs.ApplicationStart()
                        if isinstance(ret, int) and ret != 0:
                            vprint(f"   · ApplicationStart ret={ret}")
                        start = time.time()
                        while time.time() - start < 8.0:
                            try:
                                sm = etabs.SapModel
                                vprint("   → helper launched via EXE path.")
                                return sm
                            except Exception:
                                time.sleep(0.25)
                    except Exception as ee:
                        vprint(f"   · helper path launch failed: {ee}")
    except Exception as e:
        vprint(f"   × ETABSv1.Helper path failed: {e}")
    # 3) Legacy ETABS 2016 helper retained as a fallback
    try:
        vprint("[3] ETABS2016.Helper …")
        helper = comtypes.client.CreateObject("ETABS2016.Helper")
        helper = _query_helper_interface(helper, "ETABS2016")
        try:
            etabs = helper.GetObject("CSI.ETABS.API.ETABSObject")
            sm = etabs.SapModel
            vprint("   → 2016 helper attached.")
            return sm
        except Exception as e:
            vprint(f"   · 2016 attach failed: {e}")
            if launch_if_not_found:
                try:
                    etabs = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
                    try: etabs.ApplicationStart()
                    except Exception as ee: vprint(f"   · ApplicationStart skipped: {ee}")
                    start = time.time()
                    while time.time() - start < 8.0:
                        try:
                            sm = etabs.SapModel
                            vprint("   → 2016 helper launched.")
                            return sm
                        except Exception:
                            time.sleep(0.25)
                except Exception as ee:
                    vprint(f"   · 2016 CreateObjectProgID failed: {ee}")
    except Exception as e:
        vprint(f"   × ETABS2016.Helper path failed: {e}")
    # 4) direct exe
    if launch_if_not_found:
        for exe in _existing_etabs_exes():
            vprint(f"[4] Direct EXE launch → {exe}")
            try:
                subprocess.Popen([exe])
                start = time.time()
                while time.time() - start < 10.0:
                    try:
                        etabs = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
                        sm = etabs.SapModel
                        vprint("   → attached after direct launch.")
                        return sm
                    except Exception:
                        time.sleep(0.25)
            except Exception as e:
                vprint(f"   × direct launch attach failed: {e}")
    vprint("!!! Connection failed after all attempts.")
    vprint(f"comtypes cache reset: {_comtypes_cache_command()}")
    return None

def _parse_units_result(res):
    if not isinstance(res, (tuple, list)) or len(res) < 4:
        return None
    vals = list(res)
    try:
        if int(vals[0]) == 0 and int(vals[1]) in range(1, 7) and int(vals[2]) in range(1, 7) and int(vals[3]) in (1, 2):
            return int(vals[0]), int(vals[1]), int(vals[2]), int(vals[3])
        if int(vals[-1]) == 0 and int(vals[0]) in range(1, 7) and int(vals[1]) in range(1, 7) and int(vals[2]) in (1, 2):
            return int(vals[-1]), int(vals[0]), int(vals[1]), int(vals[2])
    except Exception:
        pass
    return None


def get_present_units_safe(SapModel):
    try:
        parsed = _parse_units_result(SapModel.GetPresentUnits_2())
        if parsed:
            return parsed
    except Exception:
        pass
    try:
        packed = int(SapModel.GetPresentUnits())
        for _, (f, L, T, p, _) in UNIT_PRESETS.items():
            if p == packed:
                return 0, f, L, T
    except Exception:
        pass
    raise RuntimeError("خواندن واحدهای فعلی ETABS ناموفق بود.")


def _set_units_checked(SapModel, f, L, T):
    ret = SapModel.SetPresentUnits_2(int(f), int(L), int(T))
    if ret != 0:
        raise RuntimeError(f"SetPresentUnits_2 ret={ret}")
    actual = get_present_units_safe(SapModel)
    if actual[1:] != (f, L, T):
        raise RuntimeError(f"Unit verification failed: requested={(f,L,T)}, actual={actual[1:]}")


def ensure_units_kgf_cm(SapModel):
    """
    Preserve the legacy function name while selecting units from WORKING_UNITS.
    The original model units are stored once and restored after the main run.
    """
    _cache_reset()
    if _CACHE["units_set"]:
        return _CACHE["original_units"]
    preset = UNIT_PRESETS.get(WORKING_UNITS)
    if preset is None:
        raise ValueError(f"WORKING_UNITS نامعتبر است: {WORKING_UNITS}")
    original = get_present_units_safe(SapModel)
    f, L, T, _, _ = preset
    if original[1:] != (f, L, T):
        _set_units_checked(SapModel, f, L, T)
    _CACHE["original_units"] = original
    _CACHE["length_enum"] = L
    _CACHE["units_set"] = True
    return original


def restore_original_units(SapModel):
    original = _CACHE.get("original_units")
    if not original:
        return
    try:
        current = get_present_units_safe(SapModel)
        if current[1:] != original[1:]:
            _set_units_checked(SapModel, original[1], original[2], original[3])
        print("Present units restored:", get_present_units_safe(SapModel))
    finally:
        _CACHE["units_set"] = False
        _CACHE["original_units"] = None
        _CACHE["length_enum"] = original[2]
        _cache_reset()


def _cm_to_etabs(value_cm: float) -> float:
    mm_per_unit = _MM_PER_LENGTH_ENUM[int(_CACHE["length_enum"])]
    return float(value_cm) * 10.0 / mm_per_unit


def _etabs_length_to_mm(value: float) -> float:
    mm_per_unit = _MM_PER_LENGTH_ENUM[int(_CACHE["length_enum"])]
    return float(value) * mm_per_unit


def _etabs_area_to_mm2(value: float) -> float:
    mm_per_unit = _MM_PER_LENGTH_ENUM[int(_CACHE["length_enum"])]
    return float(value) * mm_per_unit * mm_per_unit

def get_model_filename_safe(SapModel):
    res = SapModel.GetModelFilename()
    if isinstance(res, (tuple, list)):
        ret = res[0] if len(res) > 0 else -1
        path = res[1] if len(res) > 1 else None
        is_saved = res[2] if len(res) > 2 else None
        return ret, path, is_saved
    if isinstance(res, str):
        return 0, res, None
    try:
        ret0 = res[0]; path1 = res[1] if len(res) > 1 else None
        return int(ret0), path1, (res[2] if len(res) > 2 else None)
    except Exception:
        return -1, None, None

def get_version_safe(SapModel):
    res = SapModel.GetVersion()
    if isinstance(res, (tuple, list)):
        a = res[0] if len(res) > 0 else None
        b = res[1] if len(res) > 1 else None
        if isinstance(a, int) and isinstance(b, str):
            return a, b
        if isinstance(a, str):
            return 0, a
        return 0, str(b) if b is not None else None
    if isinstance(res, str):
        return 0, res
    return -1, None

def print_units(SapModel):
    try:
        ret, f, L, T = get_present_units_safe(SapModel)
        print(f"Present units → ret={ret}, force={f}, length={L}, temp={T}")
    except Exception as e:
        print(f"Present units unavailable: {e}")

# =========================
# User settings
# =========================
MATERIALS = {"concrete": None, "rebar_long": None, "rebar_tie": None}

BEAM_WIDTHS_CM  = range(30, 50 + 1, 10)
BEAM_HEIGHTS_CM = range(50, 80 + 1, 10)

COVER_TOP_CM = 6.0
COVER_BOT_CM = 6.0
COVER_COL_CM = 6.0

# =========================
# Naming, text, and reinforcing-bar helpers
# =========================
def _digits_fa2en(s: str) -> str:
    return (s or "").translate(_TR_F2E)

def _norm(s: str) -> str:
    s = _digits_fa2en(s or "").lower().strip()
    return re.sub(r"[\s_]+", "", s)

def _short_name(prefix, b_cm, h_cm):
    return f"{prefix}{int(round(b_cm))}{int(round(h_cm))}"

def _column_name_with_counts_and_phi(prefix: str, b_cm: float, h_cm: float,
                                     nx: int, ny: int, phi_mm: float) -> str:
    base   = _short_name(prefix, b_cm, h_cm)                 # Example: C4060
    total  = 2*nx + 2*ny - 4                                 # Total number of longitudinal bars
    phi    = int(round(phi_mm))
    pair   = f"{nx}{ny}" if (nx < 10 and ny < 10) else f"{nx}x{ny}"
    return f"{base}-{total}T{phi}-{pair}"

def _extract_name_list_like(res):
    if isinstance(res, (tuple, list)):
        inner = None
        for item in res:
            if isinstance(item, (list, tuple)): inner = item
        if inner is not None: return [str(x) for x in inner]
        if len(res) >= 2 and isinstance(res[1], (list, tuple)): return [str(x) for x in res[1]]
        if len(res) >= 2 and isinstance(res[1], str): return [str(res[1])]
    if isinstance(res, str): return [res]
    return []

def _list_rebar_sizes(SapModel) -> list[str]:
    # Cache full bar data when available; otherwise cache names only.
    if _CACHE["rebar_list"] is None and _CACHE["rebar_list_with_data"] is None:
        try:
            res = SapModel.PropRebar.GetNameListWithData()
            if (isinstance(res, (tuple, list)) and len(res) >= 5 and
                    (not isinstance(res[0], int) or res[0] == 0) and isinstance(res[2], (list, tuple))):
                _CACHE["rebar_list_with_data"] = res
                return [str(x) for x in res[2]]
        except Exception:
            _CACHE["rebar_list_with_data"] = ()
        try:
            res = SapModel.PropRebar.GetNameList()
            _CACHE["rebar_list"] = _extract_name_list_like(res)
        except Exception:
            _CACHE["rebar_list"] = []
    if _CACHE["rebar_list_with_data"]:
        res = _CACHE["rebar_list_with_data"]
        return [str(x) for x in (res[2] if len(res) > 2 else [])]
    return list(_CACHE["rebar_list"] or [])


def _bar_size_to_mm(value) -> float:
    if isinstance(value, (int, float)):
        if float(value) <= 0:
            raise ValueError("قطر میلگرد باید مثبت باشد.")
        return float(value)
    token = _digits_fa2en(str(value)).strip()
    m_us = _RE_US_BAR.fullmatch(token)
    if m_us:
        number = int(m_us.group(1))
        if number not in US_REBAR_DIAMETERS_MM:
            raise ValueError(f"سایز میلگرد آمریکایی پشتیبانی نمی‌شود: {value}")
        return US_REBAR_DIAMETERS_MM[number]
    m = _RE_REBAR_MM.fullmatch(token.replace(" ", ""))
    if m and float(m.group(1)) > 0:
        return float(m.group(1))
    raise ValueError(f"سایز میلگرد قابل تشخیص نیست: {value}")


def _pick_rebar_size_by_mm(SapModel, phi_mm: float) -> str:
    names = _list_rebar_sizes(SapModel)
    # Try an exact metric name before matching by actual diameter.
    rounded_phi = int(round(float(phi_mm)))
    tag = f"{rounded_phi}d"
    if abs(float(phi_mm) - rounded_phi) < 1e-9 and tag in names:
        return tag
    try:
        res = _CACHE["rebar_list_with_data"]
        if not res:
            res = SapModel.PropRebar.GetNameListWithData()
            _CACHE["rebar_list_with_data"] = res
        if isinstance(res, (tuple, list)) and len(res) >= 5:
            nms = list(res[2]) if isinstance(res[2], (list, tuple)) else []
            dias = list(res[4]) if isinstance(res[4], (list, tuple)) else []
            best_i, best_err = None, 1e9
            for i, (nm, d) in enumerate(zip(nms, dias)):
                try:
                    err = abs(_etabs_length_to_mm(float(d)) - float(phi_mm))
                    if err < best_err:
                        best_i, best_err = i, err
                except Exception:
                    continue
            if best_i is not None:
                if best_err > MAX_REBAR_DIAMETER_MISMATCH_MM:
                    raise ValueError(
                        f"برای قطر {phi_mm:.3f} mm سایز منطبق پیدا نشد. "
                        f"نزدیک‌ترین سایز '{nms[best_i]}' با اختلاف {best_err:.3f} mm است."
                    )
                return str(nms[best_i])
    except ValueError:
        raise
    except Exception:
        pass
    raise ValueError(f"سایز میلگرد متناظر با قطر {phi_mm:.3f} mm پیدا نشد.")


def require_rebar_size(SapModel, size_name, label: str="Rebar size") -> str:
    names = _list_rebar_sizes(SapModel)
    token = str(size_name).strip()
    if token in names:
        return token
    normalized = {_norm(n): n for n in names}
    if _norm(token) in normalized:
        return normalized[_norm(token)]
    phi_mm = _bar_size_to_mm(size_name)
    mapped = _pick_rebar_size_by_mm(SapModel, phi_mm)
    if mapped in names:
        print(f"[MAP] {label}: '{size_name}' → '{mapped}' ({phi_mm:.3f} mm)")
        return mapped
    preview = ", ".join(names[:10]) + (" ..." if len(names) > 10 else "")
    raise ValueError(f"{label}: سایز '{size_name}' در مدل پیدا نشد. نمونه‌ها: {preview or '(خالی)'}")

# =========================
# Type-aware material lookup with caching
# =========================
E_STEEL = 1
E_CONC  = 2
E_REBAR = 6

def _get_material_type_safe(SapModel, name: str):
    try:
        res = SapModel.PropMaterial.GetMaterial(name)
        if isinstance(res, (tuple, list)) and len(res) >= 2:
            if int(res[0]) == 0 and int(res[1]) in range(1, 9):
                return int(res[1])
            if int(res[-1]) == 0 and int(res[0]) in range(1, 9):
                return int(res[0])
    except Exception:
        pass
    return None


def _list_materials_by_type(SapModel, mat_type: int):
    if mat_type in _CACHE["materials_by_type"]:
        return list(_CACHE["materials_by_type"][mat_type])
    names = []
    try:
        res = SapModel.PropMaterial.GetNameList(mat_type)
        if isinstance(res, (tuple, list)) and res and isinstance(res[0], int) and res[0] != 0:
            raise RuntimeError(f"GetNameList ret={res[0]}")
        names = _extract_name_list_like(res)
    except Exception:
        # Use the fallback only after verifying each material type; never accept an unfiltered list.
        try:
            all_names = _extract_name_list_like(SapModel.PropMaterial.GetNameList())
            names = [n for n in all_names if _get_material_type_safe(SapModel, n) == mat_type]
        except Exception as e:
            print(f"[WARN] GetNameList failed: {e}")
            names = []
    _CACHE["materials_by_type"][mat_type] = list(names)
    return names


def require_material(SapModel, name, wanted_type: int, label: str) -> str:
    names = _list_materials_by_type(SapModel, wanted_type)
    if not names:
        raise ValueError(f"{label}: هیچ متریالی از نوع {wanted_type} در مدل پیدا نشد.")
    if name is None or not str(name).strip():
        selected = names[0]
        print(f"[AUTO] {label}: '{selected}'")
        return selected
    if name in names: return name
    mm = { _norm(n): n for n in names }
    fix = mm.get(_norm(name))
    if fix:
        print(f"[WARN] {label}: '{name}' → '{fix}'")
        return fix
    preview = ", ".join(names[:10]) + (" ..." if len(names) > 10 else "")
    raise ValueError(f"{label}: متریال '{name}' با نوع موردنظر پیدا نشد. نمونه‌ها: {preview}")


def resolve_default_materials(SapModel):
    MATERIALS["concrete"] = require_material(SapModel, MATERIALS["concrete"], E_CONC, "Concrete")
    MATERIALS["rebar_long"] = require_material(SapModel, MATERIALS["rebar_long"], E_REBAR, "Rebar Longitudinal")
    MATERIALS["rebar_tie"] = require_material(SapModel, MATERIALS["rebar_tie"], E_REBAR, "Rebar Ties")
    print("Selected materials:")
    print("  Concrete:", MATERIALS["concrete"])
    print("  Longitudinal rebar:", MATERIALS["rebar_long"])
    print("  Tie rebar:", MATERIALS["rebar_tie"])

# =========================
# Rectangular frame sections
# =========================
def _existing_section_names(SapModel) -> set:
    names = set()
    try:
        res = SapModel.PropFrame.GetNameList()
        names.update(_extract_name_list_like(res))
    except Exception as e:
        print(f"[WARN] GetNameList(PropFrame) failed: {e}")
    return names

def _define_rect(SapModel, name: str, mat: str, b: float, h: float, overwrite: bool=True):
    ensure_units_kgf_cm(SapModel)
    final = name
    try:
        if not overwrite:
            existing = _existing_section_names(SapModel)
            if final in existing:
                k = 2
                while final in existing:
                    final = f"{name}_v{k}"; k += 1
        # ETABS convention: T3 = height (h), T2 = width (b).
        ret = SapModel.PropFrame.SetRectangle(final, mat, _cm_to_etabs(h), _cm_to_etabs(b))
        if ret != 0:
            raise RuntimeError(f"SetRectangle ret={ret}")
        return True, final
    except Exception as e:
        print(f"[ERR] SetRectangle({name},{mat},{b},{h}) → {e}")
        return False, name

# =========================
# Beam reinforcement
# =========================
def set_beam_rebar_for_section(SapModel, sec_name: str,
                               mat_long: str, mat_tie: str,
                               cover_top_cm: float, cover_bot_cm: float):
    try:
        ret = SapModel.PropFrame.SetRebarBeam(sec_name, mat_long, mat_tie,
                                        _cm_to_etabs(cover_top_cm), _cm_to_etabs(cover_bot_cm),
                                        0.0, 0.0, 0.0, 0.0)
        if ret != 0:
            raise RuntimeError(f"SetRebarBeam ret={ret}")
        return True
    except Exception as e:
        print(f"[WARN] SetRebarBeam({sec_name}) → {e}")
        return False

# =========================
# Reinforced-concrete column in CHECK mode
# =========================
def _verify_column_rebar_roundtrip(SapModel, sec_name: str, expected_r3: int, expected_r2: int):
    if not VERIFY_COLUMN_REBAR_ROUNDTRIP:
        return
    try:
        res = SapModel.PropFrame.GetRebarColumn(sec_name)
        if isinstance(res, (tuple, list)) and len(res) >= 15:
            if int(res[0]) == 0:
                actual = (int(res[7]), int(res[8]))
            elif int(res[-1]) == 0:
                actual = (int(res[6]), int(res[7]))
            else:
                actual = None
            if actual is not None and actual != (expected_r3, expected_r2):
                raise RuntimeError(
                    f"GetRebarColumn mismatch for {sec_name}: expected R3/R2="
                    f"{(expected_r3, expected_r2)}, actual={actual}"
                )
    except AttributeError:
        print(f"[WARN] GetRebarColumn unavailable for '{sec_name}'.")

def define_conc_column_section_check(
    SapModel,
    sec_name: str,
    b_cm: float,                 # T2 (width)
    h_cm: float,                 # T3 (depth)
    ConcreteMat: str,
    LongRebarMaterial: str,
    ConfinementRebarMaterial: str,
    cover_cols_cm: float,
    nx: int, ny: int,            # nx is along T2; ny is along T3.
    RebarSize_hint: str | None = None,
    TieSize_hint: str | None = None,
    phi_long_mm: float | None = None,
    tie_phi_mm: float = 8.0,
    TieSpacingLongit_cm: float = 10.0,
    nTieBarsDir2: int = 4,
    nTieBarsDir3: int = 4,
    overwrite_section: bool = True
):
    # 1) Resolve materials with type validation.
    conc = require_material(SapModel, ConcreteMat,               E_CONC,  "Concrete")
    rLong = require_material(SapModel, LongRebarMaterial,        E_REBAR, "Rebar Longitudinal")
    rTie  = require_material(SapModel, ConfinementRebarMaterial, E_REBAR, "Rebar Ties")

    # 2) Create or redefine the rectangular section (ETABS: SetRectangle(Name, Mat, T3=h, T2=b)).
    ok, final_name = _define_rect(SapModel, sec_name, conc, b_cm, h_cm, overwrite=overwrite_section)
    if not ok:
        raise RuntimeError(f"SetRectangle failed for '{sec_name}'.")

    # 3) Resolve reinforcing-bar size names from the model table.
    if RebarSize_hint:
        rebar_size = require_rebar_size(SapModel, RebarSize_hint, "Longitudinal size")
    elif phi_long_mm is not None:
        rebar_size = _pick_rebar_size_by_mm(SapModel, phi_long_mm)
    else:
        raise ValueError("یکی از RebarSize_hint یا phi_long_mm را برای میلگرد طولی بده.")

    if TieSize_hint:
        tie_size = require_rebar_size(SapModel, TieSize_hint, "Tie size")
    else:
        tie_size = _pick_rebar_size_by_mm(SapModel, tie_phi_mm)

    # 4) Map the resolved values to ETABS API parameters.
    Pattern = 1         # Rectangular
    ConfineType = 1     # Ties
    NumberCBars = 0
    ToBeDesigned = False

    # Project convention: nx is along T2=b and ny is along T3=h.
    # The ETABS API parameter order is NumberR3Bars followed by NumberR2Bars.
    R2 = int(nx)
    R3 = int(ny)

    ret = SapModel.PropFrame.SetRebarColumn(
        final_name, rLong, rTie,
        Pattern, ConfineType, _cm_to_etabs(cover_cols_cm),
        NumberCBars, R3, R2,
        rebar_size, tie_size, _cm_to_etabs(TieSpacingLongit_cm),
        int(nTieBarsDir2), int(nTieBarsDir3),
        ToBeDesigned
    )
    if ret != 0:
        names = _list_rebar_sizes(SapModel)
        preview = ", ".join(names[:10]) + (" ..." if len(names) > 10 else "")
        raise RuntimeError(
            f"SetRebarColumn ret={ret}. نام سایزها با جدول مدل سازگار نیست.\n"
            f"سایزهای موجود: {preview or '(لیست خالی)'}"
        )

    _verify_column_rebar_roundtrip(SapModel, final_name, R3, R2)

    # 5) Emit a clear diagnostic summary.
    print(f"[OK] Column '{final_name}' set as CHECK. T2(b)={b_cm} → nx={nx} | T3(h)={h_cm} → ny={ny}")
    return final_name

# =========================
# Column candidate and generation engine
# =========================
@dataclass
class DesignInputs:
    b_cm: float
    h_cm: float
    bar_diam_mm: float | List[float]
    rho_min_user: float
    rho_max_user: Optional[float] = None
    max_spacing_cm: float = 20.0
    min_clear_spacing_cm: float = 4.0
    min_bars_per_dim: int = 3
    equal_bars_both_dims: bool = False
    force_nx_le_ny: bool = True
    cover_cm: float = 6.0
    tie_phi_mm: float = 8.0
    # nx/ny ratio constraints
    narrow_width_threshold_cm: float = 35.0      # Narrow-width threshold
    ny_to_nx_ratio_max_wide: float = 2.0        # For b >= 35: ny < 2*nx
    ny_to_nx_ratio_max_narrow: float = 2.5      # For b < 35: ny <= 2.5*nx
    # ---------------------------------
    notes: bool = True

def _area_bar_cm2(phi_mm: float) -> float:
    d_cm = phi_mm / 10.0
    # Equivalent to pi*(d^2)/4 with a precomputed constant.
    return 0.7853981633974483 * d_cm * d_cm

def _rho_percent(As_cm2: float, Ag_cm2: float) -> float:
    return 100.0 * As_cm2 / Ag_cm2 if Ag_cm2 > 0 else 0.0

def _nx_min_by_spacing(b_cm: float, s_max: float) -> int:
    # s = b/(n-1) <= s_max  =>  n >= b/s_max + 1
    return int(math.ceil(b_cm / s_max)) + 1

def _nx_max_by_clear(b_cm: float, phi_cm: float, s_clear_min: float) -> int:
    # Clear spacing: s - phi >= s_clear_min, where s = b/(n-1).
    # Therefore b/(n-1) >= (s_clear_min + phi) => n-1 <= b/(s_clear_min + phi).
    return int(math.floor(b_cm / (s_clear_min + phi_cm))) + 1

def enumerate_column_candidates(di: DesignInputs):
    Ag = di.b_cm * di.h_cm
    # Reinforcement-ratio bounds
    rho_min_eff = max(1.0, di.rho_min_user)
    rho_max_eff = min(8.0, di.rho_max_user) if di.rho_max_user is not None else 8.0
    if rho_max_eff <= 0:
        return []

    phis = di.bar_diam_mm if isinstance(di.bar_diam_mm, list) else [di.bar_diam_mm]
    phis = sorted(float(p) for p in phis)

    cands = []
    for phi in phis:
        phi_cm = phi / 10.0
        As_bar = 0.7853981633974483 * phi_cm * phi_cm

        # Distance from the section edge to the longitudinal-bar centerline
        d_edge = di.cover_cm + (di.tie_phi_mm / 10.0) + (phi_cm / 2.0)
        b_eff = di.b_cm - 2.0 * d_edge
        h_eff = di.h_cm - 2.0 * d_edge
        if b_eff <= 0 or h_eff <= 0:
            continue

        # Lower bounds from maximum spacing and minimum bars per side
        nx_min = max(di.min_bars_per_dim, math.ceil(b_eff / di.max_spacing_cm) + 1)
        ny_min = max(di.min_bars_per_dim, math.ceil(h_eff / di.max_spacing_cm) + 1)

        # Upper bounds from minimum clear spacing
        nx_max_clear = math.floor(b_eff / (di.min_clear_spacing_cm + phi_cm)) + 1
        ny_max_clear = math.floor(h_eff / (di.min_clear_spacing_cm + phi_cm)) + 1
        if nx_max_clear < nx_min or ny_max_clear < ny_min:
            continue

        # Use rho_max to bound the total bar count N and consequently ny.
        n_tot_max_rho = math.floor((rho_max_eff/100.0) * Ag / As_bar)
        if n_tot_max_rho < (2*nx_min + 2*ny_min - 4):
            # Even the minimum arrangement exceeds rho_max.
            continue

        # Select the permitted ny/nx ratio based on section width.
        is_narrow = (di.b_cm < di.narrow_width_threshold_cm)
        ratio_max = di.ny_to_nx_ratio_max_narrow if is_narrow else di.ny_to_nx_ratio_max_wide

        # Prune nx early using the maximum permitted ratio.
        # Since ny >= ny_min and, when force_nx_le_ny is enabled, ny >= nx.
        # For wide sections, ny < ratio_max*nx requires nx >= ceil(ny_min/ratio_max).
        nx_lo_ratio = math.ceil(ny_min / ratio_max) if ratio_max > 0 else nx_min
        nx_min_eff = max(nx_min, nx_lo_ratio)

        for nx in range(nx_min_eff, nx_max_clear + 1):
            # Initial ny bounds
            ny_lo = ny_min
            ny_hi = ny_max_clear

            # Enforce nx <= ny.
            if di.force_nx_le_ny:
                ny_lo = max(ny_lo, nx)

            # Apply the analytical ny bound derived from rho_max.
            ny_hi_rho = math.floor((n_tot_max_rho + 4 - 2*nx) / 2.0)
            if ny_hi_rho < ny_lo:
                continue
            ny_hi = min(ny_hi, ny_hi_rho)

            # Apply the ny/nx ratio constraint.
            if is_narrow:
                # b < 35 => ny <= floor(2.5*nx)
                ny_hi_ratio = int(math.floor(ratio_max * nx))
            else:
                # b >= 35 => ny < 2*nx (strict)
                ny_hi_ratio = int(math.floor(ratio_max * nx)) - 1  # Enforce the strict inequality.
            if ny_lo > ny_hi_ratio:
                continue
            ny_hi = min(ny_hi, ny_hi_ratio)

            for ny in range(ny_lo, ny_hi + 1):
                # Require equal bar counts in both directions when requested.
                if di.equal_bars_both_dims and ny != nx:
                    continue

                # Actual center-to-center spacing
                sx = b_eff / (nx - 1)
                sy = h_eff / (ny - 1)

                # Check maximum spacing and minimum clear spacing.
                if sx > di.max_spacing_cm or sy > di.max_spacing_cm:
                    continue
                if (sx - phi_cm) < di.min_clear_spacing_cm:
                    continue
                if (sy - phi_cm) < di.min_clear_spacing_cm:
                    continue

                # Recheck the ratio as a defensive secondary filter.
                if is_narrow:
                    if ny > ratio_max * nx:
                        continue  # ny <= 2.5*nx
                else:
                    if ny >= ratio_max * nx:
                        continue  # ny < 2*nx

                # Reinforcement ratio
                n_total = 2*nx + 2*ny - 4
                As_tot  = n_total * As_bar
                rho     = 100.0 * As_tot / Ag
                if rho < rho_min_eff or rho > rho_max_eff:
                    continue

                cands.append((phi, nx, ny, sx, sy, As_tot, rho))
    return cands

def _total_bars(nx: int, ny: int) -> int:
    return 2*nx + 2*ny - 4

def build_all_feasible_columns_in_etabs(
    SapModel,
    di: DesignInputs,
    base_prefix: str,
    concrete_mat: str,
    long_rebar_mat: str,
    tie_rebar_mat: str,
    cover_cm: float,
    tie_size_hint: str = "8d",
    tie_spacing_cm: float = 10.0,
    overwrite_section: bool = True
):
    """
    Naming: C{b}{h}-{N}T{phi}-{nx}{ny}.
    Ordering: ascending total bar count N.
    Monotonicity rule: as N increases, nx and ny must not decrease;
    equal or larger values remain valid.
    """
    # 1) Generate candidates.
    cands = enumerate_column_candidates(di)  # [(phi, nx, ny, sx, sy, As_tot, rho), ...]

    # 2) Sort by N, then phi, nx, and ny without repeated rounding.
    def _key(c):
        phi, nx, ny, sx, sy, As_tot, rho = c
        N = 2*nx + 2*ny - 4
        return (N, int(phi), nx, ny)
    cands_sorted = sorted(cands, key=_key)

    # 3) Enforce nondecreasing nx and ny as N increases.
    kept = []
    nx_floor = 0   # Minimum permitted nx at this point
    ny_floor = 0   # Minimum permitted ny at this point

    for N, group_iter in groupby(cands_sorted, key=lambda c: 2*c[1] + 2*c[2] - 4):
        group = list(group_iter)

        # Keep only candidates with nx >= nx_floor and ny >= ny_floor.
        group_kept = [c for c in group if (c[1] >= nx_floor and c[2] >= ny_floor)]

        # Sort each group for deterministic output.
        group_kept.sort(key=lambda c: (int(c[0]), c[1], c[2]))  # (phi, nx, ny)

        kept.extend(group_kept)

        # Update the next floors using the maximum accepted values in this group.
        if group_kept:
            nx_floor = max(nx_floor, max(c[1] for c in group_kept))
            ny_floor = max(ny_floor, max(c[2] for c in group_kept))

    # 4) Create the sections in ETABS.
    made = []
    existing = _existing_section_names(SapModel)
    for (phi, nx, ny, sx, sy, As_tot, rho) in kept:
        name = _column_name_with_counts_and_phi(base_prefix, di.b_cm, di.h_cm, nx, ny, phi)

        final = name
        k = 2
        while final in existing:
            final = f"{name}_v{k}"; k += 1

        define_conc_column_section_check(
            SapModel,
            sec_name=final,
            b_cm=di.b_cm, h_cm=di.h_cm,
            ConcreteMat=concrete_mat,
            LongRebarMaterial=long_rebar_mat,
            ConfinementRebarMaterial=tie_rebar_mat,
            cover_cols_cm=cover_cm,
            nx=nx, ny=ny,
            RebarSize_hint=None,
            phi_long_mm=phi,
            TieSize_hint=tie_size_hint,
            TieSpacingLongit_cm=tie_spacing_cm,
            nTieBarsDir2=4, nTieBarsDir3=4,
            overwrite_section=overwrite_section
        )
        existing.add(final)
        N = 2*nx + 2*ny - 4
        made.append((final, phi, nx, ny, rho, N))

    # 5) Prepare the result list.
    made.sort(key=lambda r: r[5])  # Sort by N.
    return made

# =========================
# Beam generation with the existing flow and optional name versioning
# =========================
def run_beam_generation(SapModel, overwrite=True):
    conc = require_material(SapModel, MATERIALS["concrete"],   E_CONC,  "Concrete")
    rL   = require_material(SapModel, MATERIALS["rebar_long"], E_REBAR, "Rebar Longitudinal")
    rT   = require_material(SapModel, MATERIALS["rebar_tie"],  E_REBAR, "Rebar Ties")

    beams = []
    existing = _existing_section_names(SapModel) if not overwrite else set()

    for b, h in product(BEAM_WIDTHS_CM, BEAM_HEIGHTS_CM):
        name = _short_name("B", b, h)
        nm = name
        if not overwrite:
            k = 2
            while nm in existing:
                nm = f"{name}_v{k}"; k += 1
        ok, final = _define_rect(SapModel, nm, conc, b, h, overwrite=True)
        if not overwrite:
            existing.add(final)
        if ok and set_beam_rebar_for_section(SapModel, final, rL, rT, COVER_TOP_CM, COVER_BOT_CM):
            beams.append(final)

    print(f"[DONE] Beams: {len(beams)}")
    return beams

# =========================
# Example column parameters
# =========================
COL_B_CM = 60
COL_H_CM = 70
COL_BAR_DIAM_MM = [25]      # Metric: [20, 25] | US: ["#6", "#8"]
COL_TIE_SIZE = "8d"          # For US bar notation, use a value such as "#3".
COL_RHO_MIN = 1.0
COL_RHO_MAX = 8.0
COL_S_MAX   = 20.0
COL_S_CLEAR_MIN = 4.0
COL_MIN_PER_DIM = 3

def run_single_column_design_and_build(SapModel, overwrite=True):
    bar_diameters_mm = [_bar_size_to_mm(v) for v in COL_BAR_DIAM_MM]
    tie_phi_mm = _bar_size_to_mm(COL_TIE_SIZE)
    di = DesignInputs(
        b_cm=COL_B_CM, h_cm=COL_H_CM,
        bar_diam_mm=bar_diameters_mm,
        rho_min_user=COL_RHO_MIN, rho_max_user=COL_RHO_MAX,
        max_spacing_cm=COL_S_MAX,
        min_clear_spacing_cm=4.0,              # Minimum clear spacing
        min_bars_per_dim=COL_MIN_PER_DIM,
        equal_bars_both_dims=False,            # Not required to be equal; however:
        force_nx_le_ny=True,                   # n_x <= n_y
        cover_cm=COVER_COL_CM,
        tie_phi_mm=tie_phi_mm                  # Tie-bar diameter
    )

    conc = require_material(SapModel, MATERIALS["concrete"],   E_CONC,  "Concrete")
    rL   = require_material(SapModel, MATERIALS["rebar_long"], E_REBAR, "Rebar Longitudinal")
    rT   = require_material(SapModel, MATERIALS["rebar_tie"],  E_REBAR, "Rebar Ties")

    made = build_all_feasible_columns_in_etabs(
        SapModel, di,
        base_prefix="C",
        concrete_mat=conc, long_rebar_mat=rL, tie_rebar_mat=rT,
        cover_cm=COVER_COL_CM,
        tie_size_hint=COL_TIE_SIZE, tie_spacing_cm=10.0,
        overwrite_section=overwrite
    )

    print(f"[DONE] Built {len(made)} column sections for {int(di.b_cm)}x{int(di.h_cm)}.")
    for (nm, phi, nx, ny, rho, N) in made[:15]:
        print(f"  - {nm}: nx={nx}, ny={ny}, φ={int(phi)}mm, N={N}, ρ={rho:.2f}%")
    if len(made) > 15:
        print(f"  (+{len(made) - 15} more …)")

def debug_print_rebar_table(SapModel):
    try:
        res = SapModel.PropRebar.GetNameListWithData()
        names = list(res[2]) if (isinstance(res, (tuple, list)) and len(res) > 2) else []
        areas = list(res[3]) if (isinstance(res, (tuple, list)) and len(res) > 3) else []
        dias  = list(res[4]) if (isinstance(res, (tuple, list)) and len(res) > 4) else []
        if not names:
            print("Rebar table is empty.")
            return
        print("---- Rebar Table ----")
        for nm, ar, d in zip(names, areas, dias):
            print(f"Name='{nm}', Area={_etabs_area_to_mm2(ar):.3f} mm^2, Dia={_etabs_length_to_mm(d):.3f} mm")
    except Exception as e:
        print("PropRebar.GetNameListWithData failed:", e)

# =========================
# Program entry point
# =========================
if __name__ == "__main__":
    if not (DO_BEAMS or DO_COLUMNS):
        sys.exit("Both DO_BEAMS and DO_COLUMNS are 0 → nothing to do.")
    if REBAR_SYSTEM not in {"metric", "us"}:
        sys.exit("REBAR_SYSTEM must be 'metric' or 'us'.")

    print(f"ETABS RC Section Builder v{APP_VERSION}")
    print(f"Python         : {platform.python_version()} ({struct.calcsize('P')*8}-bit)")
    print(f"comtypes       : {getattr(comtypes, '__version__', 'unknown')}")
    print(f"Rebar system   : {REBAR_SYSTEM}")

    SapModel = link_etabs(launch_if_not_found=True, verbose=True)
    if not SapModel:
        sys.exit(
            "Connection failed. This code is written for ETABS 20/21. "
            f"If comtypes wrappers are corrupted run: {_comtypes_cache_command()}"
        )

    ret_ver, ver = get_version_safe(SapModel)
    print(f"ETABS version : {ver} (ret={ret_ver})")
    m = re.search(r"\b(\d{1,2})(?:\.\d+)?", ver or "")
    major = int(m.group(1)) if m else None
    if major not in SUPPORTED_ETABS_MAJOR_VERSIONS:
        print(f"[WARN] Supported ETABS versions are 20 and 21; detected: {ver}")

    ret_path, model_path, is_saved = get_model_filename_safe(SapModel)
    print(f"Model path    : {model_path} (ret={ret_path})")
    print(f"is_saved      : {is_saved}")

    original_units = None
    try:
        original_units = ensure_units_kgf_cm(SapModel)
        print_units(SapModel)
        resolve_default_materials(SapModel)

        if DO_BEAMS:
            run_beam_generation(SapModel, overwrite=True)
        if DO_COLUMNS:
            run_single_column_design_and_build(SapModel, overwrite=True)
        print("[DONE] Model was modified in memory; it was not saved automatically.")
    finally:
        if original_units is not None:
            restore_original_units(SapModel)
