import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw

from dim_wishlist.icon_config import IconBuilderConfig
from dim_wishlist.icon_matching import image_similarity, signature_from_bytes
from dim_wishlist.icon_wishlist import resolve_named_perk_in_weapon_socket
from dim_wishlist.icon_xlsx import extract_icon_contexts
from dim_wishlist.manifest import ManifestIndex
from dim_wishlist.models import InventoryItem


def png_bytes(image):
    stream = BytesIO()
    image.save(stream, format="PNG")
    return stream.getvalue()


class IconWorkflowTests(unittest.TestCase):
    def test_sample_workbook_drawing_extraction(self):
        workbook = Path(__file__).parents[1] / "input" / "d2.xlsx"
        with tempfile.TemporaryDirectory() as directory:
            contexts, stats = extract_icon_contexts(
                workbook, Path(directory), IconBuilderConfig()
            )
        self.assertEqual(stats["drawing_count"], 4120)
        self.assertEqual(stats["perk_icon_position_count"], 3221)
        self.assertEqual(stats["unique_perk_icon_count"], 184)
        self.assertEqual(stats["missing_weapon_count"], 0)
        self.assertEqual(contexts[0].source_cell, "X6")

    def test_similarity_tolerates_small_translation(self):
        config = IconBuilderConfig()
        source = Image.new("RGBA", (96, 96), (0, 0, 0, 0))
        ImageDraw.Draw(source).ellipse((16, 12, 80, 84), fill=(255, 255, 255, 255))
        shifted = Image.new("RGBA", (96, 96), (0, 0, 0, 0))
        shifted.alpha_composite(source, (2, 1))
        score = image_similarity(
            signature_from_bytes(png_bytes(source), config),
            signature_from_bytes(png_bytes(shifted), config),
            config,
        )
        self.assertGreater(score, 0.99)

    def test_special_slot_override_prefers_normal_perk(self):
        def item(item_hash, display, tier):
            return InventoryItem(
                hash=item_hash,
                sql_id=item_hash,
                name="超频散热器",
                item_type=19,
                item_type_display=display,
                item_type_and_tier_display=f"{tier} {display}",
                tier_type_name=tier,
                plug_category_identifier="batteries",
                has_plug=True,
                json_obj={"hash": item_hash},
            )

        normal = item(1092016998, "电池", "普通")
        enhanced = item(226831738, "强化电池", "罕见")
        weapon = InventoryItem(
            hash=1229624538,
            sql_id=1229624538,
            name="维卡拉微冲4",
            item_type=3,
            item_type_display="武器",
            item_type_and_tier_display="传说武器",
            tier_type_name="传说",
            plug_category_identifier="",
            has_plug=False,
            json_obj={"sockets": {"socketEntries": [
                {}, {}, {"randomizedPlugSetHash": 9001},
            ]}},
        )
        index = ManifestIndex([weapon, normal, enhanced], {9001: [enhanced.hash, normal.hash]})
        selected, socket, matched = resolve_named_perk_in_weapon_socket(
            index, weapon.hash, "超频散热器"
        )
        self.assertEqual(selected.hash, normal.hash)
        self.assertEqual(socket["socket_index"], 2)
        self.assertEqual({candidate.hash for candidate in matched}, {normal.hash, enhanced.hash})


if __name__ == "__main__":
    unittest.main()
