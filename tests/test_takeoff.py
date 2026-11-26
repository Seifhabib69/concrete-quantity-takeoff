"""Run with:  python -m pytest   or   python tests/test_takeoff.py"""
import sys, tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import concrete_takeoff as ct
from openpyxl import load_workbook


def approx(a, b, tol=1e-9):
    return abs(a - b) < tol


def test_member_volume():
    assert approx(ct.member_volume(4, 1.8, 1.8, 0.5), 6.48)


def test_plain_concrete_volume():
    # (1.8+0.2) x (1.8+0.2) x 0.1 x 4 = 1.6
    assert approx(ct.plain_concrete_volume(4, 1.8, 1.8, 0.1, 0.1), 1.6)


def test_calculate_with_waste():
    data = {s: [] for s in ct.MEMBER_SHEETS}
    data["Footings"] = [{"mark": "F1", "count": 4, "d1": 1.8, "d2": 1.8, "d3": 0.5, "notes": ""}]
    data["Columns"] = [{"mark": "C1", "count": 4, "d1": 0.3, "d2": 0.6, "d3": 3.0, "notes": ""}]
    s = dict(ct.DEFAULT_SETTINGS, waste_pct=0.05, plain_concrete="yes")
    r = ct.calculate(data, s)
    assert approx(r["Footings"]["net"], 6.48)
    assert approx(r["Columns"]["net"], 2.16)
    assert approx(r["Reinforced concrete total"]["net"], 8.64)
    assert approx(r["Reinforced concrete total"]["with_waste"], 8.64 * 1.05)
    assert approx(r["Plain concrete"]["net"], 1.6)


def test_template_roundtrip_and_report():
    with tempfile.TemporaryDirectory() as d:
        tpl, rep = Path(d) / "in.xlsx", Path(d) / "out.xlsx"
        ct.make_template(tpl)
        data, settings = ct.read_input(tpl)
        ct.write_report(data, settings, rep)
        wb = load_workbook(rep)
        assert wb.sheetnames[0] == "Summary"
        assert wb["Summary"]["B4"].value.startswith("=SUM(Footings!G2")


def test_bad_input_reports_every_problem():
    with tempfile.TemporaryDirectory() as d:
        tpl = Path(d) / "in.xlsx"
        ct.make_template(tpl)
        wb = load_workbook(tpl)
        wb["Footings"]["C2"] = -1
        wb["Columns"]["B2"] = "four"
        wb.save(tpl)
        try:
            ct.read_input(tpl)
        except ct.InputError as e:
            msg = str(e)
            assert "Footings row 2" in msg and "Columns row 2" in msg
        else:
            raise AssertionError("InputError not raised")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)
