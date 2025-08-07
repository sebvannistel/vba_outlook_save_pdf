import pathlib

README = pathlib.Path('readme.md').read_text(encoding='utf-8')

def test_readme_mentions_enhanced_guide():
    assert 'Enhanced Guide' in README

def test_readme_has_installation_section():
    assert '## Installation Guide' in README


def test_readme_describes_inspector_scenario():
    assert 'Manual Inspector Scenario' in README
