"""tools/xlsx_table.py: an Excel workbook read as tables with their formulas, number formats and merged cells."""
import io
import json
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

import xlsx_table  # noqa: E402

HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def workbook(sheets, styles="", shared=None) -> bytes:
    """A small .xlsx: `sheets` is [(name, sheetData xml, extra xml after sheetData, extra before)]."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", HEAD + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")
        z.writestr("xl/workbook.xml", HEAD + f"<workbook {MAIN}><sheets>" + "".join(
            f'<sheet name="{name}" sheetId="{i + 1}" r:id="rId{i + 1}"' + (' state="hidden"' if hidden else "") + "/>"
            for i, (name, _, _, _, hidden) in enumerate(sheets)) + "</sheets></workbook>")
        z.writestr("xl/_rels/workbook.xml.rels", HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(
            f'<Relationship Id="rId{i + 1}" Type="worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets))) + "</Relationships>")
        for i, (_, data, after, before, _) in enumerate(sheets):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", HEAD + f"<worksheet {MAIN}>{before}<sheetData>{data}</sheetData>{after}</worksheet>")
        if styles:
            z.writestr("xl/styles.xml", HEAD + f"<styleSheet {MAIN}>{styles}</styleSheet>")
        if shared:
            z.writestr("xl/sharedStrings.xml", HEAD + f"<sst {MAIN}>" + "".join(shared) + "</sst>")
    return buf.getvalue()


STYLES = (
    '<numFmts count="3"><numFmt numFmtId="164" formatCode="&quot;¥&quot;#,##0"/><numFmt numFmtId="165" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/>'
    '<numFmt numFmtId="166" formatCode="0.0%"/></numFmts>'
    '<cellXfs count="7"><xf numFmtId="0"/><xf numFmtId="3"/><xf numFmtId="164"/><xf numFmtId="14"/><xf numFmtId="165"/><xf numFmtId="166"/><xf numFmtId="2"/></cellXfs>'
)


class XlsxTableTests(unittest.TestCase):
    def test_values_formulas_and_number_formats(self):
        rows = (
            '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>'
            '<row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2" s="1"><v>1200</v></c><c r="C2" s="2"><v>3500.5</v></c></row>'
            '<row r="3"><c r="A3" t="inlineStr"><is><t>割合</t></is></c><c r="B3" s="5"><f>B2/C2</f><v>0.342665</v></c><c r="C3" s="6"><f>SUM(B2:C2)</f><v>4700.5</v></c></row>'
            '<row r="4"><c r="A4" t="b"><v>1</v></c><c r="B4" t="e"><v>#DIV/0!</v></c><c r="C4" t="str"><f>A3&amp;"!"</f><v>割合!</v></c></row>'
        )
        shared = ["<si><t>項目</t></si>", "<si><r><t>前期</t></r><r><t>（実績）</t></r></si>", "<si><t>今期</t><rPh><t>コンキ</t></rPh></si>", "<si><t>売上</t></si>"]
        sheets = xlsx_table.read_workbook(workbook([("売上", rows, "", "", False)], STYLES, shared))
        self.assertEqual(len(sheets), 1)
        sheet = sheets[0]
        self.assertEqual(sheet["name"], "売上")
        grid = sheet["rows"]
        self.assertEqual([c.get("text") for c in grid[0]], ["項目", "前期（実績）", "今期"], "runs joined, furigana left out")
        self.assertEqual(grid[1][1]["text"], "1,200", "#,##0 shows as the workbook shows it")
        self.assertEqual(grid[1][2]["text"], "¥3,501", "¥ and commas, half rounded away from zero")
        self.assertEqual(grid[2][1]["f"], "=B2/C2")
        self.assertEqual(grid[2][1]["text"], "34.3%")
        self.assertEqual((grid[2][1]["dec"], grid[2][1]["pct"]), (1, True), "a formula's number format is kept for when it is worked out again")
        self.assertEqual(grid[2][2]["text"], "4700.50")
        self.assertEqual(grid[2][2]["dec"], 2)
        self.assertEqual(grid[3][0]["text"], "TRUE")
        self.assertEqual(grid[3][1]["text"], "#DIV/0!")
        self.assertEqual((grid[3][2]["f"], grid[3][2]["text"]), ('=A3&"!"', "割合!"))
        self.assertAlmostEqual(sum(sheet["cols"]), 1, places=3)

    def test_dates_show_as_the_workbook_shows_them_and_lose_their_formula(self):
        rows = '<row r="1"><c r="A1" s="3"><v>45383</v></c><c r="B1" s="4"><v>45383.5</v></c><c r="C1" s="3"><f>A1+30</f><v>45413</v></c></row>'
        grid = xlsx_table.read_workbook(workbook([("日付", rows, "", "", False)], STYLES))[0]["rows"]
        self.assertEqual(grid[0][0]["text"], "2024/4/1")
        self.assertEqual(grid[0][1]["text"], "2024年4月1日")
        self.assertEqual(grid[0][2]["text"], "2024/5/1")
        self.assertNotIn("f", grid[0][2], "a date result is kept as it was (the studio does not do date arithmetic on results)")

    def test_shared_formulas_move_with_their_cell(self):
        rows = (
            '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2+$A$1</f><v>3</v></c></row>'
            '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>6</v></c></row>'
            '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>9</v></c></row>'
        )
        grid = xlsx_table.read_workbook(workbook([("s", rows, "", "", False)]))[0]["rows"]
        self.assertEqual([r[1]["f"] for r in grid], ["=A1*2+$A$1", "=A2*2+$A$1", "=A3*2+$A$1"])
        self.assertEqual(xlsx_table.shift_formula('=SUM(A1:B2)&"A1"&LOG10(A1)+1E5', 1, 1)[0:1], "=")

    def test_merged_cells_widths_hidden_sheets_and_the_used_range(self):
        rows = '<row r="2"><c r="B2" t="inlineStr"><is><t>見出し</t></is></c></row><row r="3"><c r="B3"><v>5</v></c><c r="C3"><v>6</v></c></row>'
        cols = '<cols><col min="1" max="1" width="4" customWidth="1"/><col min="2" max="3" width="12" customWidth="1"/></cols>'
        merges = '<mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells>'
        sheets = xlsx_table.read_workbook(workbook([("見出し", rows, merges, cols, False), ("隠し", rows, "", "", True), ("空", "", "", "", False)]))
        self.assertEqual([s["name"] for s in sheets], ["見出し", "空"], "hidden sheets are not offered")
        grid = sheets[0]["rows"]
        self.assertEqual((len(grid), len(grid[0])), (3, 3), "the used range, from A1 (so formulas keep their places)")
        self.assertEqual((grid[1][1]["rs"], grid[1][1]["cs"]), (1, 2))
        self.assertEqual(grid[1][2], {"merged": True})
        self.assertAlmostEqual(sheets[0]["cols"][0], 4 / 28, places=3)
        self.assertEqual(sheets[1]["rows"], [], "an empty sheet has no rows")

    def test_not_a_workbook(self):
        with self.assertRaises(ValueError):
            xlsx_table.read_workbook(b"not a zip")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("hello.txt", "x")
        with self.assertRaises(ValueError):
            xlsx_table.read_workbook(buf.getvalue())

    def test_command_line_prints_json(self):
        rows = '<row r="1"><c r="A1"><v>7</v></c></row>'
        done = subprocess.run([sys.executable, str(ROOT / "tools" / "xlsx_table.py"), "a.xlsx"], input=workbook([("s", rows, "", "", False)]), capture_output=True, check=False)
        self.assertEqual(done.returncode, 0, done.stderr.decode())
        self.assertEqual(json.loads(done.stdout)["sheets"][0]["rows"][0][0]["text"], "7")
        bad = subprocess.run([sys.executable, str(ROOT / "tools" / "xlsx_table.py"), "a.xlsx"], input=b"nope", capture_output=True, check=False)
        self.assertNotEqual(bad.returncode, 0)
        self.assertIn("Excel", bad.stderr.decode())


if __name__ == "__main__":
    unittest.main()
