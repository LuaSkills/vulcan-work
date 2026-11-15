"""Verify release tags against the skill manifest version.
根据技能清单版本验证发布标签。
"""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from scripts.release_version import normalize_requested_version, resolve_release_tag


class ReleaseVersionTests(unittest.TestCase):
    """Verify that release tags follow the skill.yaml version.
    验证发布标签与 skill.yaml 版本一致。
    """

    def test_release_version_must_match_manifest(self) -> None:
        """Accept the manifest version and reject mismatched or invalid tags.
        接受清单版本，拒绝不同或无效的标签。
        """
        with TemporaryDirectory() as temporary_directory:
            # The isolated skill directory contains the manifest under test.
            # 独立的技能目录包含被测试的清单。
            root = Path(temporary_directory) / "vulcan-work"
            root.mkdir()
            (root / "skill.yaml").write_text(
                "name: Vulcan Work\nversion: 1.2.3\n",
                encoding="utf-8",
            )
            self.assertEqual("v1.2.3", resolve_release_tag(root, "1.2.3"))
            self.assertEqual("v1.2.3", resolve_release_tag(root, "v1.2.3"))
            with self.assertRaises(RuntimeError):
                resolve_release_tag(root, "1.2.4")
            with self.assertRaises(RuntimeError):
                normalize_requested_version("latest")
