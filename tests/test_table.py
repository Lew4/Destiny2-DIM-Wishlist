import tempfile
import unittest
from pathlib import Path

from dim_wishlist.table import detect_columns, expand_rows, read_table


class TableTests(unittest.TestCase):
    def test_repeated_headers_and_perk_cartesian_product(self):
        content = (
            "自动步枪,,,,,,\n"
            "名字,Perk,Perk,Perk 1,Perk 2,注释,Tier\n"
            '测试枪,"枪管A\n枪管B",弹匣A,特性A,特性B,说明,S\n'
            "弓箭,,,,,,\n"
            "名字,Perk,Perk,Perk 1,Perk 2,注释,Tier\n"
            "测试弓,弓弦A,箭杆A,特性C,特性D,说明,A\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recommendations.csv"
            path.write_text(content, encoding="utf-8")
            headers, rows = read_table(path)

        weapon, perks, note, tier, rank = detect_columns(headers)
        self.assertEqual(perks, ["Perk", "Perk__2", "Perk 1", "Perk 2"])
        expanded = expand_rows(rows, weapon, perks, note, tier, rank)
        self.assertEqual(len(expanded), 3)
        self.assertEqual(expanded[0]["slot_order"], [
            "slot_1_barrel", "slot_2_magazine", "slot_3_trait", "slot_4_trait",
        ])
        self.assertEqual(expanded[0]["notes"], "Tier S | 说明")


if __name__ == "__main__":
    unittest.main()
