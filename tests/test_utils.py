import unittest

from dim_wishlist.utils import norm_name, split_options, to_dim_hash, to_sql_id


class UtilsTests(unittest.TestCase):
    def test_signed_and_unsigned_hash_conversion(self):
        self.assertEqual(to_dim_hash(-1001759469), 3293207827)
        self.assertEqual(to_sql_id(3293207827), -1001759469)

    def test_name_normalization_and_option_splitting(self):
        self.assertEqual(norm_name(" 鲁莽-神谕（万神殿） "), "鲁莽神谕万神殿")
        self.assertEqual(split_options("小口径\n箭头制退器/槽化枪管"), [
            "小口径", "箭头制退器", "槽化枪管",
        ])


if __name__ == "__main__":
    unittest.main()
