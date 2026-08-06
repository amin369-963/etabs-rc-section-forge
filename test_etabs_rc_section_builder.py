import importlib.util
import math
import sys
import types
import unittest
from pathlib import Path

# Minimal comtypes stub so pure and fake-ETABS tests can run outside Windows.
comtypes = types.ModuleType("comtypes")
comtypes.__version__ = "test-stub"
comtypes.gen = types.SimpleNamespace()
client = types.ModuleType("comtypes.client")
comtypes.client = client
sys.modules["comtypes"] = comtypes
sys.modules["comtypes.client"] = client

MODULE_PATH = Path(__file__).with_name("etabs_rc_section_builder.py")
spec = importlib.util.spec_from_file_location("etabs_section_builder", MODULE_PATH)
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
assert spec.loader is not None
spec.loader.exec_module(module)


class FakePropMaterial:
    def __init__(self):
        self.by_type = {
            module.E_CONC: ["4000Psi", "C30"],
            module.E_REBAR: ["A615Gr60", "A706Gr60"],
        }

    def GetNameList(self, mat_type=None):
        if mat_type is None:
            names = [n for values in self.by_type.values() for n in values]
        else:
            names = self.by_type.get(mat_type, [])
        return 0, len(names), list(names)

    def GetMaterial(self, name):
        for mat_type, names in self.by_type.items():
            if name in names:
                return 0, mat_type, -1, "", ""
        return 1, 0, -1, "", ""


class FakePropRebar:
    def GetNameListWithData(self):
        names = ["8d", "25d", "#8"]
        diameters_cm = [0.8, 2.5, 2.54]
        areas_cm2 = [math.pi * d * d / 4 for d in diameters_cm]
        return 0, len(names), names, areas_cm2, diameters_cm, ["", "", ""]

    def GetNameList(self):
        return 0, 3, ["8d", "25d", "#8"]


class FakePropFrame:
    def __init__(self):
        self.sections = {}
        self.columns = {}
        self.beams = {}

    def GetNameList(self):
        names = list(self.sections)
        return 0, len(names), names

    def SetRectangle(self, name, material, t3, t2):
        self.sections[name] = (material, t3, t2)
        return 0

    def SetRebarBeam(self, name, long_mat, tie_mat, top, bot, *areas):
        self.beams[name] = (long_mat, tie_mat, top, bot, areas)
        return 0

    def SetRebarColumn(self, name, *data):
        self.columns[name] = data
        return 0

    def GetRebarColumn(self, name):
        return (0,) + self.columns[name]


class FakeSapModel:
    def __init__(self):
        self.units = (5, 5, 2)
        self.PropMaterial = FakePropMaterial()
        self.PropRebar = FakePropRebar()
        self.PropFrame = FakePropFrame()

    def GetPresentUnits_2(self):
        return (0,) + self.units

    def SetPresentUnits_2(self, force, length, temp):
        self.units = (force, length, temp)
        return 0

    def GetPresentUnits(self):
        for values in module.UNIT_PRESETS.values():
            if values[:3] == self.units:
                return values[3]
        return 0


class Tests(unittest.TestCase):
    def setUp(self):
        module._CACHE["units_set"] = False
        module._CACHE["original_units"] = None
        module._CACHE["length_enum"] = 5
        module._cache_reset()

    def test_metric_and_us_bar_parsing(self):
        self.assertEqual(module._bar_size_to_mm("25d"), 25.0)
        self.assertAlmostEqual(module._bar_size_to_mm("#8"), 25.4)

    def test_material_auto_selection(self):
        sap = FakeSapModel()
        previous = dict(module.MATERIALS)
        try:
            module.MATERIALS.update({"concrete": None, "rebar_long": None, "rebar_tie": None})
            module.resolve_default_materials(sap)
            self.assertEqual(module.MATERIALS["concrete"], "4000Psi")
            self.assertEqual(module.MATERIALS["rebar_long"], "A615Gr60")
            self.assertEqual(module.MATERIALS["rebar_tie"], "A615Gr60")
        finally:
            module.MATERIALS.clear()
            module.MATERIALS.update(previous)

    def test_units_convert_and_restore(self):
        sap = FakeSapModel()
        old = module.WORKING_UNITS
        try:
            module.WORKING_UNITS = "kip-in-F"
            original = module.ensure_units_kgf_cm(sap)
            self.assertEqual(sap.units, (2, 1, 1))
            self.assertAlmostEqual(module._cm_to_etabs(25.4), 10.0)
            module.restore_original_units(sap)
            self.assertEqual(sap.units, original[1:])
        finally:
            module.WORKING_UNITS = old

    def test_rebar_selection_by_actual_diameter(self):
        sap = FakeSapModel()
        self.assertEqual(module.require_rebar_size(sap, 25), "25d")
        self.assertEqual(module.require_rebar_size(sap, "#8"), "#8")
        self.assertEqual(module._pick_rebar_size_by_mm(sap, 25.4), "#8")

    def test_rectangle_units(self):
        sap = FakeSapModel()
        old = module.WORKING_UNITS
        try:
            module.WORKING_UNITS = "kip-in-F"
            ok, name = module._define_rect(sap, "B3050", "4000Psi", 30, 50)
            self.assertTrue(ok)
            mat, t3, t2 = sap.PropFrame.sections[name]
            self.assertEqual(mat, "4000Psi")
            self.assertAlmostEqual(t3, 50 / 2.54)
            self.assertAlmostEqual(t2, 30 / 2.54)
        finally:
            module.WORKING_UNITS = old

    def test_column_r3_r2_mapping(self):
        sap = FakeSapModel()
        created = module.define_conc_column_section_check(
            sap,
            "C6070-test",
            60,
            70,
            "4000Psi",
            "A615Gr60",
            "A615Gr60",
            6,
            nx=4,
            ny=6,
            RebarSize_hint="25d",
            TieSize_hint="8d",
        )
        self.assertEqual(created, "C6070-test")
        data = sap.PropFrame.columns[created]
        # indices after name: long, tie, pattern, confine, cover, circular, R3, R2...
        self.assertEqual(data[6], 6)
        self.assertEqual(data[7], 4)

    def test_candidate_engine_preserved(self):
        di = module.DesignInputs(60, 70, [25], 1.0, 8.0)
        cands = module.enumerate_column_candidates(di)
        self.assertTrue(cands)
        self.assertEqual(cands[0][0:3], (25.0, 4, 4))
        for phi, nx, ny, sx, sy, area, rho in cands:
            self.assertLessEqual(nx, ny)
            self.assertGreaterEqual(rho, 1.0)
            self.assertLessEqual(rho, 8.0)

    def test_python39_parse(self):
        import ast
        source = MODULE_PATH.read_text(encoding="utf-8")
        ast.parse(source, feature_version=(3, 9))


if __name__ == "__main__":
    unittest.main(verbosity=2)
