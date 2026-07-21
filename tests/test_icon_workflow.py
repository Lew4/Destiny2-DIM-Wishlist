import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw

from dim_wishlist.icon_config import IconBuilderConfig
from dim_wishlist.icon_matching import image_similarity, signature_from_bytes
from dim_wishlist.icon_models import GlobalIconResolution
from dim_wishlist.icon_reports import (
    classify_version_results,
    resolve_manifest_weapon_name,
    select_weapon_versions,
)
from dim_wishlist.icon_wishlist import (
    is_recommendation_excluded,
    resolve_global_visual_in_actual_trait_socket,
    resolve_named_perk_in_weapon_socket,
)
from dim_wishlist.icon_xlsx import extract_icon_contexts
from dim_wishlist.manifest import ManifestIndex
from dim_wishlist.models import InventoryItem


def png_bytes(image):
    stream = BytesIO()
    image.save(stream, format="PNG")
    return stream.getvalue()


class IconWorkflowTests(unittest.TestCase):
    def test_source_row_selects_the_correct_same_name_versions(self):
        class Candidate:
            def __init__(self, item_hash):
                self.hash = item_hash

        candidates = [
            Candidate(1802315656), Candidate(1992309064),
            Candidate(3385326721), Candidate(4158265643),
        ]
        selected = select_weapon_versions(
            IconBuilderConfig(), "鲁莽神谕", candidates, 28
        )
        self.assertEqual(
            {candidate.hash for candidate in selected},
            {3385326721, 1992309064},
        )

    def test_explicit_recommendation_exclusion_is_source_scoped(self):
        config = IconBuilderConfig()
        self.assertTrue(is_recommendation_excluded(
            config, 28, "鲁莽神谕", "pve", ["集体爆破"]
        ))
        self.assertFalse(is_recommendation_excluded(
            config, 142, "鲁莽神谕", "pve", ["集体爆破"]
        ))

    def test_weapon_identity_override_uses_weapon_type(self):
        config = IconBuilderConfig()
        self.assertEqual(
            resolve_manifest_weapon_name(config, "信任", "烈日速射斥候"), "受托"
        )
        self.assertEqual(
            resolve_manifest_weapon_name(config, "信任", "烈日精密手炮"), "信任"
        )

    def test_sample_workbook_drawing_extraction(self):
        workbook = Path(__file__).parents[1] / "input" / "d2.xlsx"
        with tempfile.TemporaryDirectory() as directory:
            contexts, stats = extract_icon_contexts(
                workbook, Path(directory), IconBuilderConfig()
            )
            self.assertFalse((Path(directory) / "extracted_icons").exists())
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

    def test_wrong_xlsx_trait_column_uses_actual_weapon_socket(self):
        perk = InventoryItem(
            hash=501,
            sql_id=501,
            name="测试特性",
            item_type=19,
            item_type_display="特性",
            item_type_and_tier_display="普通 特性",
            tier_type_name="普通",
            plug_category_identifier="frames",
            has_plug=True,
            json_obj={"hash": 501},
        )
        result = GlobalIconResolution(
            icon_sha256="source", accepted=True, reason="accepted",
            best_visual_id="visual", best_score=1.0, second_score=0.0,
            margin=1.0, match_method="exact", candidate_summary=[],
            occurrence_count=1,
        )
        sockets = [
            {"socket_index": 3, "candidates": []},
            {"socket_index": 4, "candidates": [perk]},
        ]
        selected, socket, matched = resolve_global_visual_in_actual_trait_socket(
            result, sockets, {perk.hash: "visual"}
        )
        self.assertEqual(selected.hash, perk.hash)
        self.assertEqual(socket["socket_index"], 4)
        self.assertEqual([item.hash for item in matched], [perk.hash])

    def test_history_mismatch_is_not_a_source_error(self):
        base = {
            "excel_row": 1, "weapon_name": "测试枪", "manifest_weapon_name": "测试枪",
            "weapon_type": "测试类型", "usage": "pve", "source_cell": "A1",
            "icon_sha256": "icon", "recognized_names": "测试特性",
        }
        matches = [{**base, "weapon_hash": 101, "accepted": "yes"}]
        unresolved = [
            {**base, "weapon_hash": 102, "accepted": "no", "generated": "no"}
        ]
        real, history, groups = classify_version_results(matches, unresolved)
        self.assertEqual(real, [])
        self.assertEqual(len(history), 1)
        self.assertEqual(groups[0]["full_version_hashes"], "101")
        self.assertNotIn("reason", groups[0])


if __name__ == "__main__":
    unittest.main()
