import pathlib

def test_module_contains_functions():
    data = pathlib.Path('module.bas').read_text(encoding='utf-8')
    assert 'SaveAsPDFfile' in data
    assert 'SaveSelectedMails_AsPDF_NoPopups' in data
