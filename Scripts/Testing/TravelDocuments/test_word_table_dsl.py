"""Check the table DSL assets, expansion/grid contracts and Word scaffold without VBA execution."""

from copy import deepcopy
from pathlib import Path
import re
import unittest
from zipfile import ZipFile

from lxml import etree as ET

from prepare_table_example import CATALOG, SOURCE, TARGET, ROOT, NS as WORD_NS, W
from word_table_preview import grid, marker_bounds, styles, NS, P


def fragment(text):
    return ET.fromstring(f'<table xmlns="{P}"><columns><column name="A" width="40mm"/><column name="B" width="60mm"/></columns>{text}</table>')


class WordTableDslTests(unittest.TestCase):
    def test_vba_procedure_names_are_unique_per_module(self):
        folder = ROOT / "PROJECTS_EXCEL/vba"
        files = list((folder / "TravelDocumentsBuilder").glob("*.vba"))
        files.extend((folder / "common").rglob("*Word*.vba"))
        for file in files:
            text = file.read_text(encoding="utf-8-sig")
            names = re.findall(r"^\s*(?:Public|Private|Friend)\s+(?:Sub|Function)\s+(\w+)\s*\(", text, re.M | re.I)
            folded = [name.lower() for name in names]
            self.assertEqual(len(folded), len(set(folded)), file.name)

    def test_new_vba_declarations_avoid_conflicting_names(self):
        folder = ROOT / "PROJECTS_EXCEL/vba/common"
        for file in folder.rglob("*Word*.vba"):
            text = file.read_text(encoding="utf-8-sig")
            declarations = re.findall(r"\b(?:Dim|ByVal|ByRef)\s+(\w+)\s+As\b", text, re.I)
            forbidden = set("attribute shared binary namespace as boolean byte byref byval call case class const currency date declare dim do double each else elseif end enum error event exit false for friend function get global gosub goto if implements integer let lib long loop lset mod new next not nothing on option optional or paramarray preserve private property public raiseevent redim rem resume return rset select set single static stop string sub then to true type typeof until variant wend while with withevents xor".split())
            self.assertFalse(forbidden & {name.lower() for name in declarations}, file.name)

    def test_ticket_geometry_matches_existing_word_table(self):
        table = ET.parse(str(CATALOG)).find("p:template/p:table", NS)
        columns, rows, cells, occupied = grid(table, {})
        with ZipFile(SOURCE) as package:
            original = ET.fromstring(package.read("word/document.xml")).find(".//w:tbl", NS)
        self.assertEqual(len(rows), 30)
        self.assertEqual(len(columns), 11)
        expected = [int(node.get(f"{{{W}}}w")) for node in original.findall("w:tblGrid/w:gridCol", NS)]
        self.assertEqual([round(width * 20) for width in columns], expected)
        for index, (row, _) in enumerate(rows):
            original_cells = original.findall("w:tr", NS)[index].findall("w:tc", NS)
            expected_spans = [int(node.find("w:tcPr/w:gridSpan", NS).get(f"{{{W}}}val")) if node.find("w:tcPr/w:gridSpan", NS) is not None else 1 for node in original_cells]
            actual = []
            column = 0
            while column < len(columns):
                span = occupied[index, column][5]
                actual.append(span)
                column += span
            self.assertEqual(actual, expected_spans, index + 1)
        self.assertEqual(occupied[0, 0][4], 6)
        self.assertEqual(occupied[0, 4][4], 9)

    def test_scaffold_retains_everything_outside_first_table(self):
        with ZipFile(SOURCE) as source, ZipFile(TARGET) as target:
            original = ET.fromstring(source.read("word/document.xml"))
            scaffold = ET.fromstring(target.read("word/document.xml"))
            original_body = original.find("w:body", NS)
            scaffold_body = scaffold.find("w:body", NS)
            original_body.remove(original_body.find("w:tbl", NS))
            begin, end = marker_bounds(scaffold_body, "Ticket")
            self.assertEqual(end, begin + 1)
            for item in list(scaffold_body)[begin:end + 1]:
                scaffold_body.remove(item)
            self.assertEqual(ET.tostring(original), ET.tostring(scaffold))
            for name in source.namelist():
                if name != "word/document.xml":
                    self.assertEqual(source.read(name), target.read(name), name)

    def test_dynamic_rows_empty_single_many_and_condition(self):
        table = fragment('<row><cell columnSpan="2"><text>Header</text></cell></row><if test="[Enabled]"><forEach itemsSource="People" item="person"><row><cell><text>{person.[Name]}</text></cell><cell><text>{person.__index}</text></cell></row></forEach></if>')
        for count in (0, 1, 8):
            _, rows, _, _ = grid(table, {"Enabled": True, "People": [{"Name": str(i)} for i in range(count)]})
            self.assertEqual(len(rows), count + 1)
            if count:
                self.assertEqual(rows[-1][1]["person.__index"], count)
                self.assertTrue(rows[-1][1]["person.__last"])
        self.assertEqual(len(grid(table, {"Enabled": False})[1]), 1)

    def test_reject_holes_out_of_bounds_and_overlapping_spans(self):
        invalid = [
            '<row><cell/></row>',
            '<row><cell columnSpan="3"/></row>',
            '<row><cell rowSpan="2"/><cell/></row>',
            '<row><cell/><cell rowSpan="2"/></row><row><cell columnSpan="2"/></row>',
            '<row><cell columnSpan="0"/><cell/></row>',
            '<row header="true"><cell rowSpan="2"/><cell/></row><row><cell/></row>',
        ]
        for text in invalid:
            with self.subTest(text=text), self.assertRaises(ValueError):
                grid(fragment(text), {})

    def test_missing_collection_limit_and_empty_table(self):
        table = fragment('<forEach itemsSource="People" item="person"><row><cell columnSpan="2"/></row></forEach>')
        with self.assertRaises(KeyError):
            grid(table, {})
        with self.assertRaises(ValueError):
            grid(table, {"People": []})
        with self.assertRaises(ValueError):
            grid(table, {"People": [{}] * 513})

    def test_style_inheritance_missing_and_cycles(self):
        catalog = ET.parse(str(CATALOG)).getroot()
        resolved = styles(catalog)
        self.assertEqual(resolved["Underlined"]["horizontal"], "center")
        self.assertEqual(resolved["Logo"]["fontName"], "Times New Roman")
        for missing in (True, False):
            root = deepcopy(catalog)
            node = root.find("p:styles/p:cellStyle", NS)
            node.set("basedOn", "Missing" if missing else node.get("name"))
            with self.assertRaises(ValueError):
                styles(root)

    def test_marker_pairs_reject_missing_duplicate_inline_and_reversed(self):
        with ZipFile(TARGET) as package:
            body = ET.fromstring(package.read("word/document.xml")).find("w:body", NS)
        begin, end = marker_bounds(body, "Ticket")
        for mutation in ("missing", "duplicate", "inline", "reversed"):
            copy = deepcopy(body)
            if mutation == "missing":
                copy.remove(copy[end])
            elif mutation == "duplicate":
                copy.append(deepcopy(copy[begin]))
            elif mutation == "inline":
                copy[begin].find(".//w:t", NS).text += " extra"
            else:
                copy[begin], copy[end] = deepcopy(copy[end]), deepcopy(copy[begin])
            with self.subTest(mutation=mutation), self.assertRaises(ValueError):
                marker_bounds(copy, "Ticket")

    def test_vba_integration_and_portable_text_dsl(self):
        folder = ROOT / "PROJECTS_EXCEL/vba"
        source = (folder / "common/[2] classes/obj_WordTableTemplate.cls.vba").read_text(encoding="utf-8-sig")
        xml = (folder / "common/modules/ex_WordTableXml.vba").read_text(encoding="utf-8-sig")
        generator = (folder / "TravelDocumentsBuilder/obj_TDB_VacationGenerator.cls.vba").read_text(encoding="utf-8-sig")
        self.assertIn('plan.Add "xml", ex_WordTableXml.fn_Build(plan, m_images)', source)
        self.assertIn('first.InsertXML plan("xml")', source)
        self.assertNotIn(".Tables.Add", source)
        self.assertNotIn(".Merge ", source)
        self.assertLess(source.index('plan.Add "xml"'), source.index("Public Function TryInsert"))
        self.assertIn('m_document.setProperty "ProhibitDTD", True', source)
        self.assertIn('m_document.resolveExternals = False', source)
        self.assertNotIn(".hasAttribute(", source)
        self.assertIn('private_W(properties, "vMerge")', xml)
        self.assertIn('private_W(properties, "gridSpan")', xml)
        self.assertIn('If VBA.Len(paragraph("text")) = 0 Then\n        lines = VBA.Array(VBA.vbNullString)', xml)
        self.assertIn('For Each text In lines', xml)
        self.assertNotIn('For Each text In VBA.Split(paragraph("text")', xml)
        self.assertIn('A generated Word table cell has no paragraph.', xml)
        self.assertLess(generator.index("m_wordTableTemplate.TryPrepare"), generator.index('stage = "Session.Start"'))
        render = generator[generator.index("Private Sub private_GenerateWord"):]
        self.assertLess(render.index('document.SaveAs2 m_wordTemplatePath'), render.index("m_wordTableTemplate.TryInsert"))
        core = (folder / "common/[2] classes/obj_WordTableText.cls.vba").read_text(encoding="utf-8-sig")
        self.assertNotIn("ex_PEB_", core)
        original = (folder / "PersonalEventBuilder/obj_PEB_DslRenderer.cls.vba").read_text(encoding="utf-8-sig")
        for name in ("private_Render", "private_Condition", "private_Value", "private_Layout", "private_Date"):
            pattern = rf"Private Function {name}\([\s\S]*?^End Function"
            old = re.search(pattern, original, re.M)[0]
            new = re.search(pattern, core, re.M)[0]
            old = old.replace("ex_PEB_Export.fn_Text", "private_Text").replace("ex_PEB_Export.fn_Required", "private_Required").replace("ex_PEB_DslResources", "ex_WordDslResources")
            old = old.replace('            If collectionName = "BasisDocuments" Then\n                Set values(collectionName) = ex_PEB_Export.fn_BasisDocuments(values)\n            End If\n', "")
            old = re.sub(r"^\s*'.*\n", "", old, flags=re.M)
            new = re.sub(r"^\s*'.*\n", "", new, flags=re.M)
            simple_if = r"(?m)^( +)If ([^\n]+) Then\n\1    ([^\n]+)\n\1End If$"
            old = re.sub(simple_if, r"\1If \2 Then \3", old)
            new = re.sub(simple_if, r"\1If \2 Then \3", new)
            self.assertEqual(old, new, name)
        for file in (folder / "common").rglob("*.vba"):
            if file.stem.startswith(("obj_WordTable", "ex_Word")):
                for line in file.read_text(encoding="utf-8-sig").splitlines():
                    self.assertLessEqual(len(line), 1023, file.name)


if __name__ == "__main__":
    unittest.main(verbosity=2)