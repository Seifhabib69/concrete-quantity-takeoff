# concrete-quantity-takeoff

A small, auditable **concrete quantity takeoff** tool for structural engineers.
You list footings, columns (necks), beams (tie/strap beams) and slabs in an Excel
file; the tool writes a report workbook whose volumes are **live Excel formulas**,
so every number can be checked by hand.

Built for the everyday case in design offices and on site: quick, transparent
concrete quantities without a closed spreadsheet.

## Install

```bash
pip install -r requirements.txt   # only openpyxl
```

## Use

```bash
# 1) create a blank input file (with a legend and one example row per sheet)
python concrete_takeoff.py template my_project.xlsx

# 2) fill in Footings / Columns / Beams / Slabs, optionally edit Settings

# 3) generate the report
python concrete_takeoff.py run my_project.xlsx report.xlsx
```

The console prints a summary and `report.xlsx` contains:

| Sheet | Content |
|---|---|
| Summary | net volume, waste %, volume incl. waste per element, RC total, plain concrete |
| Settings | waste %, plain concrete under footings (on/off, thickness, offset) |
| Footings / Columns / Beams / Slabs | your rows + unit volume and total volume formulas |

## Input format

All dimensions in **metres**. One row per group of identical members.

| Sheet | Columns |
|---|---|
| Footings | Mark, Count, Length, Width, Thickness |
| Columns | Mark, Count, b, t, Height (neck height) |
| Beams | Mark, Count, b, Depth, Length (clear length) |
| Slabs | Mark, Count, Length, Width, Thickness |

Bad input (negative or non-numeric values) is reported with sheet and row number.

## How quantities are calculated

- Volume of each row = `Count x dim1 x dim2 x dim3` (m3).
- Waste: `Volume incl. waste = Net x (1 + waste_pct)`; `waste_pct` is a fraction (3% = 0.03).
- Plain concrete (optional): `Count x (L + 2*offset) x (B + 2*offset) x PC thickness`.
- **Volumes are gross.** No deductions for overlaps between members, rebar or openings.
  Enter clear lengths/heights yourself where members intersect.

## Tests

```bash
python tests/test_takeoff.py        # or: python -m pytest
```

## Roadmap

- Reinforcement quantities (bar schedule / kg per m3 ratios)
- Formwork areas
- CSV input, stairs and walls
- Optional concrete mix materials (cement, sand, gravel)

## License

MIT

---

## بالعربي (ملخص)

أداة بسيطة لحصر كميات الخرسانة: بتملا ملف إكسيل بالقواعد والأعمدة (الرقاب) والكمرات والبلاطات
بالأبعاد بالمتر، وبتطلع تقرير إكسيل فيه معادلات حية (تقدر تراجع أي رقم بإيدك).
الكميات **إجمالية** (من غير خصم تداخلات)، والهالك بيتحدد من شيت Settings (مثلا 3% = 0.03).
