"""
Tests for factory.py — run with: pytest tests/
No AWS credentials required — tests build from the sample plan only.
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Mock boto3 before importing factory (not needed for build tests)
from unittest.mock import MagicMock

sys.modules["boto3"] = MagicMock()


from factory import BUILDERS, THEMES, build_pptx

SAMPLE_PLAN = {
    "title": "Test Deck",
    "executive_summary": "A test presentation.",
    "slides": [
        {"number":1,"layout":"title","title":"Test Title","subtitle":"Test Subtitle","speaker_line":"Name · Role","speaker_notes":"Notes."},
        {"number":2,"layout":"three_cards","title":"Three Cards","cards":[
            {"badge":"A","heading":"Heading One","body":"Body text for card one. More content here.","phase_color":"accent_1"},
            {"badge":"B","heading":"Heading Two","body":"Body text for card two. More content here.","phase_color":"accent_2"},
            {"badge":"C","heading":"Heading Three","body":"Body text for card three.","phase_color":"accent_3"}],
            "callout":"Bottom callout text.","speaker_notes":"Notes."},
        {"number":3,"layout":"two_column","title":"Two Column",
            "left_panel":{"heading":"Left","items":[{"label":"L1","description":"Desc one."},{"label":"L2","description":"Desc two."}]},
            "right_panel":{"heading":"Right","items":[{"label":"R1","description":"Desc A."},{"label":"R2","description":"Desc B."}]},
            "quote":"Pull quote here.","speaker_notes":"Notes."},
        {"number":4,"layout":"timeline","title":"Timeline","phases":[
            {"label":"Phase 1","timeframe":"0-90 Days","heading":"Foundation","items":["Item 1","Item 2","Item 3"],"outcome":"Outcome one.","phase_color":"accent_1"},
            {"label":"Phase 2","timeframe":"90-180 Days","heading":"Governance","items":["Item 4","Item 5","Item 6"],"outcome":"Outcome two.","phase_color":"accent_2"},
            {"label":"Phase 3","timeframe":"180-365 Days","heading":"Optimize","items":["Item 7","Item 8","Item 9"],"outcome":"Outcome three.","phase_color":"accent_3"}],
            "speaker_notes":"Notes."},
        {"number":5,"layout":"closing","title":"Closing Headline",
            "cards":[{"icon":"⚡","heading":"Outcome One","body":"Description of outcome one."},
                     {"icon":"🛡️","heading":"Outcome Two","body":"Description of outcome two."},
                     {"icon":"📋","heading":"Outcome Three","body":"Description of outcome three."}],
            "framework_title":"Key Takeaways",
            "framework_items":["Takeaway one.","Takeaway two.","Takeaway three."],
            "speaker_notes":"Notes."},
    ]
}


class TestThemes:
    def test_all_themes_present(self):
        assert set(THEMES.keys()) == {"meridian", "midnight", "teal", "corporate", "forest"}

    def test_theme_colors_are_valid_hex(self):
        color_keys = ["bg_dark","bg_light","bg_card","bg_card_dark","accent_1","accent_2","accent_3"]
        for name, t in THEMES.items():
            for k in color_keys:
                val = t[k]
                assert len(val) == 6, f"{name}.{k} invalid"
                int(val, 16)

    def test_accent_colors_distinct(self):
        for name, t in THEMES.items():
            assert len({t["accent_1"], t["accent_2"], t["accent_3"]}) == 3, f"{name} accents not distinct"


class TestBuilders:
    def test_all_layouts_registered(self):
        assert set(BUILDERS.keys()) == {"title","three_cards","two_column","timeline","closing"}

    def test_build_all_themes(self):
        for theme_name in THEMES:
            with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
                path = f.name
            try:
                build_pptx(SAMPLE_PLAN, theme_name, path)
                assert os.path.getsize(path) > 10_000, f"{theme_name} output too small"
            finally:
                os.unlink(path)
