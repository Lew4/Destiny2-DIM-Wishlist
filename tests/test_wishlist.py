import unittest

from dim_wishlist.config import BuilderConfig
from dim_wishlist.manifest import ManifestIndex
from dim_wishlist.models import InventoryItem
from dim_wishlist.wishlist import build_wishlist


def item(item_hash, name, item_type, *, sockets=None, category="trait"):
    obj = {"hash": item_hash}
    if sockets is not None:
        obj["sockets"] = {"socketEntries": sockets}
    return InventoryItem(
        hash=item_hash,
        sql_id=item_hash,
        name=name,
        item_type=item_type,
        item_type_display="特性" if item_type == 19 else "武器",
        item_type_and_tier_display="普通特性" if item_type == 19 else "传说武器",
        tier_type_name="普通" if item_type == 19 else "传说",
        plug_category_identifier=category,
        has_plug=item_type == 19,
        json_obj=obj,
    )


class WishlistTests(unittest.TestCase):
    def test_all_versions_and_partial_version_fallback(self):
        barrel = item(101, "枪管A", 19, category="barrel")
        magazine = item(102, "弹匣A", 19, category="magazine")
        complete = item(201, "测试枪", 3, sockets=[
            {"randomizedPlugSetHash": 301},
            {"randomizedPlugSetHash": 302},
        ])
        old = item(202, "测试枪", 3, sockets=[
            {"randomizedPlugSetHash": 301},
        ])
        index = ManifestIndex(
            [barrel, magazine, complete, old],
            {301: [101], 302: [102]},
        )
        expanded = [{
            "source_row": 1,
            "weapon": "测试枪",
            "notes": "Tier S",
            "slot_order": ["slot_1_barrel", "slot_2_magazine"],
            "parsed_perk_columns": {
                "slot_1_barrel": ["枪管A"],
                "slot_2_magazine": ["弹匣A"],
            },
        }]

        lines, unresolved, audit = build_wishlist(BuilderConfig(), index, expanded)

        rules = [line for line in lines if line.startswith("dimwishlist:")]
        self.assertEqual(rules, [
            "dimwishlist:item=201&perks=101,102",
            "dimwishlist:item=202&perks=101",
        ])
        self.assertEqual([row["version_status"] for row in audit], ["full", "partial"])
        self.assertEqual(len(unresolved), 1)
        self.assertEqual(unresolved[0]["name"], "弹匣A")
        self.assertEqual(unresolved[0]["generated"], "yes")


if __name__ == "__main__":
    unittest.main()
