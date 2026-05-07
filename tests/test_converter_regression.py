import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

import converter


BASE_DIR = Path(__file__).resolve().parents[1]
DOCX_PATH = BASE_DIR / "files" / "doc" / "2020 P6 CL SA1 Revision 1.docx"
OPS_DOCX_PATH = BASE_DIR / "files" / "doc" / "2024 OPS P6 FMA CT1 1A.docx"
OPS_ANSWER_KEY_PATH = BASE_DIR / "files" / "doc" / "2024 OPS P6 FMA CT1 Answer Key.docx"
PHYSICS_P1_MS_PATH = BASE_DIR / "files" / "doc" / "2022 JC2 H1 Physics Prelim P1 MS.docx"
PHYSICS_P2_MS_PATH = BASE_DIR / "files" / "doc" / "2022 JC2 H1 Physics Prelim P2 MS.docx"
QTI_NS = {"q": "http://www.imsglobal.org/xsd/imsqti_v2p1"}


class ConverterRegressionTest(unittest.TestCase):
    def convert_docx(self, docx_path):
        temp_dir = tempfile.TemporaryDirectory(prefix="qti_regression_")
        self.addCleanup(temp_dir.cleanup)
        converter.convert_docx_to_qti(str(docx_path), temp_dir.name)
        item_paths = sorted((Path(temp_dir.name) / "items").glob("Q*.xml"))
        grouped = {}
        for path in item_paths:
            grouped.setdefault(path.name[:4], []).append(path)
        return item_paths, grouped

    def load_item(self, item_path):
        root = ET.parse(item_path).getroot()
        prompt = "".join(root.find(".//q:prompt", QTI_NS).itertext()).strip()
        choices = [
            "".join(choice.itertext()).strip()
            for choice in root.findall(".//q:simpleChoice", QTI_NS)
        ]
        return prompt, choices

    def assert_no_duplicate_prefixes(self, grouped):
        self.assertTrue(all(len(paths) == 1 for paths in grouped.values()))

    def test_mcq_mark_scheme_grid_still_parses(self):
        answers = converter.parse_mark_scheme(str(PHYSICS_P1_MS_PATH))

        self.assertEqual(30, len(answers))
        self.assertEqual("c", answers["1"])
        self.assertEqual("b", answers["2"])
        self.assertEqual("c", answers["30"])

    def test_qn_answer_column_pairs_are_normalized(self):
        answers = converter.parse_mark_scheme(str(OPS_ANSWER_KEY_PATH))

        self.assertEqual(20, len(answers))
        self.assertEqual("a", answers["1"])
        self.assertEqual("b", answers["9"])
        self.assertEqual("d", answers["11"])
        self.assertEqual("b", answers["17"])
        self.assertEqual("c", answers["20"])

    def test_structured_mark_scheme_still_parses(self):
        answers = converter.parse_mark_scheme(str(PHYSICS_P2_MS_PATH))

        self.assertEqual(8, len(answers))
        self.assertTrue(answers["1"].startswith("1a: rate of change of velocity"))
        self.assertIn("2a: Equilibrium is a state", answers["2"])

    def test_chinese_table_and_inline_mcq_variants(self):
        item_paths, grouped = self.convert_docx(DOCX_PATH)

        self.assertEqual(25, len(item_paths))
        self.assert_no_duplicate_prefixes(grouped)

        for qnum in ("Q001", "Q002", "Q003", "Q016", "Q021", "Q025"):
            _, choices = self.load_item(grouped[qnum][0])
            self.assertEqual(4, len(choices), qnum)

        prompt20, _ = self.load_item(grouped["Q020"][0])
        self.assertNotIn("\u9605\u8bfb\u7406\u89e3", prompt20)

        prompt21, _ = self.load_item(grouped["Q021"][0])
        self.assertIn("\u6709\u4e00\u4f4d\u753b\u5bb6", prompt21)

    def test_numbered_table_options_do_not_spawn_fake_question_numbers(self):
        item_paths, grouped = self.convert_docx(OPS_DOCX_PATH)

        self.assertEqual(20, len(item_paths))
        self.assertEqual(20, len(grouped))
        self.assert_no_duplicate_prefixes(grouped)

        prompt1, choices1 = self.load_item(grouped["Q001"][0])
        self.assertEqual(["4000", "400", "40", "4"], choices1)
        self.assertNotIn("(1)", prompt1)

        _, choices2 = self.load_item(grouped["Q002"][0])
        self.assertEqual(["7.60", "7.64", "7.65", "7.70"], choices2)

        _, choices4 = self.load_item(grouped["Q004"][0])
        self.assertEqual(["80", "120", "240", "400"], choices4)


if __name__ == "__main__":
    unittest.main()
