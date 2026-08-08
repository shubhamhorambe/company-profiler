import re
import unittest
from pathlib import Path


class SecretHygieneTests(unittest.TestCase):
    def test_python_sources_do_not_contain_live_looking_openai_keys(self):
        pattern = re.compile(r"sk-[A-Za-z0-9_-]{20,}")
        offenders = []
        for path in Path(__file__).resolve().parents[1].glob("*.py"):
            if pattern.search(path.read_text(errors="ignore")):
                offenders.append(path.name)
        self.assertEqual(offenders, [], f"Potential hardcoded OpenAI keys: {offenders}")


if __name__ == "__main__":
    unittest.main()
