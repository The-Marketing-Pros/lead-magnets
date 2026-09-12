from pathlib import Path
import tempfile
import unittest

from validate_static import validate


class StaticAssetTests(unittest.TestCase):
    def test_missing_local_asset_fails_but_external_reference_needs_no_network(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'index.html').write_text('<a href="https://example.invalid">External</a><img src="missing.png">')
            self.assertEqual(len(validate(root)[1]), 1)
            (root / 'missing.png').write_bytes(b'fixture')
            self.assertEqual(validate(root)[1], [])

    def test_root_relative_directory_encoded_path_and_fragment(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'nested').mkdir()
            (root / 'index.html').write_text('<a href="/nested/">Next</a><a href="#local">Local</a>')
            (root / 'nested/index.html').write_text('<a href="../download%20file.csv?version=1">File</a>')
            (root / 'download file.csv').write_text('header')
            self.assertEqual(validate(root)[1], [])

    def test_missing_landing_page_and_escape_fail(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.assertTrue(validate(root)[1])
            (root / 'index.html').write_text('<a href="../outside.html">Outside</a>')
            self.assertIn('leaves the published root', validate(root)[1][0])

    def test_private_headers_required(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / 'index.html').write_text('Private fixture')
            self.assertTrue(validate(root, private=True)[1])
            (root / '_headers').write_text('/*\n X-Robots-Tag: noindex, nofollow, noarchive\n Cache-Control: private, no-store\n')
            self.assertEqual(validate(root, private=True)[1], [])
