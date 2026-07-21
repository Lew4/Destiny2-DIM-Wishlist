import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from PIL import Image, ImageDraw

from dim_wishlist.icon_config import IconBuilderConfig
from dim_wishlist.icon_matching import image_similarity, signature_from_bytes
from dim_wishlist.icon_xlsx import extract_icon_contexts


def png_bytes(image):
    stream = BytesIO()
    image.save(stream, format="PNG")
    return stream.getvalue()


class IconWorkflowTests(unittest.TestCase):
    def test_sample_workbook_drawing_extraction(self):
        workbook = Path(__file__).parents[1] / "examples" / "d2.xlsx"
        with tempfile.TemporaryDirectory() as directory:
            contexts, stats = extract_icon_contexts(
                workbook, Path(directory), IconBuilderConfig()
            )
        self.assertEqual(stats["drawing_count"], 4115)
        self.assertEqual(stats["perk_icon_position_count"], 3216)
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


if __name__ == "__main__":
    unittest.main()
