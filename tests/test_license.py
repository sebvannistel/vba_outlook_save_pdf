import pathlib

LICENSE = pathlib.Path('LICENSE').read_text(encoding='utf-8')

def test_license_mentions_mit():
    assert 'MIT License' in LICENSE
