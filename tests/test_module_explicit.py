import pathlib


def test_module_has_option_explicit():
    lines = pathlib.Path('module.bas').read_text(encoding='utf-8').splitlines()
    assert 'Option Explicit' in lines
