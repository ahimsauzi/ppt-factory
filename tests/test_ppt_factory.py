"""
Tests for factory.py — run with: pytest tests/
No AWS credentials required — tests build from the sample plan only.
"""
import json
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


class TestThemeScaffold:
    def test_scaffold_keys_match_themes(self):
        import ast
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
        from new_theme import TEMPLATE

        entry = TEMPLATE.format(name_lower="x", name_title="X", description="d").split("#")[0].strip().rstrip(",")
        scaffold = ast.literal_eval("{" + entry + "}")["x"]
        assert set(scaffold) == set(THEMES["meridian"])


class TestPlanOnlyOutput:
    def test_progress_goes_to_stderr(self, capsys, monkeypatch):
        import factory

        monkeypatch.setattr(factory, "_call_bedrock", lambda system, user_msg: SAMPLE_PLAN)
        factory.generate_plan("Any topic", slide_count=5)
        out, err = capsys.readouterr()
        assert out == ""
        assert "[ppt-factory]" in err



def _plan(**slide_overrides):
    """A deep copy of SAMPLE_PLAN with one slide (by index) replaced."""
    import copy

    plan = copy.deepcopy(SAMPLE_PLAN)
    for idx, slide in slide_overrides.items():
        plan["slides"][int(idx.lstrip("s"))] = slide
    return plan


class TestValidatePlan:
    def test_sample_plan_is_valid(self):
        from factory import validate_plan

        assert validate_plan(SAMPLE_PLAN) == []

    def test_rejects_non_object(self):
        from factory import validate_plan

        assert validate_plan([]) == ["plan must be a JSON object"]

    def test_rejects_missing_slides(self):
        from factory import validate_plan

        assert "plan: missing 'slides'" in validate_plan({"title": "x"})

    def test_rejects_unknown_layout(self):
        from factory import validate_plan

        errors = validate_plan(_plan(s2={"layout": "pie_chart", "title": "x"}))
        assert any("unknown layout 'pie_chart'" in e for e in errors)

    def test_rejects_missing_title(self):
        from factory import validate_plan

        slide = dict(SAMPLE_PLAN["slides"][1])
        del slide["title"]
        assert any("missing 'title'" in e for e in validate_plan(_plan(s1=slide)))

    def test_rejects_content_the_builder_would_drop(self):
        from factory import validate_plan

        slide = dict(SAMPLE_PLAN["slides"][3])
        slide["phases"] = slide["phases"] * 2  # 6 phases, the timeline renders at most 3
        assert any("'phases' has 6 items, expected 1-3" in e for e in validate_plan(_plan(s3=slide)))

    def test_rejects_bad_phase_color(self):
        from factory import validate_plan

        slide = {**SAMPLE_PLAN["slides"][1], "cards": [{"heading": "h", "body": "b", "phase_color": "red"}]}
        assert any("phase_color 'red'" in e for e in validate_plan(_plan(s1=slide)))

    def test_rejects_non_string_text(self):
        from factory import validate_plan

        slide = {**SAMPLE_PLAN["slides"][1], "callout": 42}
        assert any("'callout' must be a non-empty string" in e for e in validate_plan(_plan(s1=slide)))

    def test_requires_title_first_and_closing_last(self):
        from factory import validate_plan

        plan = _plan()
        plan["slides"] = plan["slides"][1:-1]
        errors = validate_plan(plan)
        assert "slide 1: a deck must start with a 'title' slide" in errors
        assert any("must end with a 'closing' slide" in e for e in errors)


class TestBuildRefusesInvalid:
    def test_build_rejects_invalid_plan(self):
        import pytest

        from factory import PlanValidationError

        with pytest.raises(PlanValidationError):
            build_pptx({"title": "x", "slides": []}, "teal", os.devnull)

    def test_build_rejects_unknown_theme(self):
        import pytest

        with pytest.raises(ValueError, match="unknown theme"):
            build_pptx(SAMPLE_PLAN, "neon", os.devnull)


class TestRepairLoop:
    def _fake(self, monkeypatch, responses):
        import factory

        calls = []

        def fake_call(system, user_msg):
            calls.append(user_msg)
            result = responses[len(calls) - 1]
            if isinstance(result, Exception):
                raise result
            return result

        monkeypatch.setattr(factory, "_call_bedrock", fake_call)
        return calls

    def test_valid_first_response_makes_one_call(self, monkeypatch):
        import factory

        calls = self._fake(monkeypatch, [SAMPLE_PLAN])
        assert factory.generate_plan("t", slide_count=5) == SAMPLE_PLAN
        assert len(calls) == 1

    def test_invalid_response_is_repaired_with_the_errors(self, monkeypatch):
        import factory

        calls = self._fake(monkeypatch, [{"title": "x", "slides": []}, SAMPLE_PLAN])
        assert factory.generate_plan("t", slide_count=5) == SAMPLE_PLAN
        assert len(calls) == 2
        assert "was rejected" in calls[1] and "'slides' has 0 items" in calls[1]

    def test_invalid_json_is_repaired(self, monkeypatch):
        import factory

        calls = self._fake(monkeypatch, [json.JSONDecodeError("bad", "x", 0), SAMPLE_PLAN])
        assert factory.generate_plan("t", slide_count=5) == SAMPLE_PLAN
        assert "not valid JSON" in calls[1]

    def test_gives_up_after_one_repair(self, monkeypatch):
        import pytest

        import factory

        bad = {"title": "x", "slides": []}
        self._fake(monkeypatch, [bad, bad])
        with pytest.raises(factory.PlanValidationError):
            factory.generate_plan("t", slide_count=5)


class TestCli:
    def test_from_json_rejects_invalid_plan_with_exit_2(self, tmp_path):
        import subprocess

        bad = tmp_path / "bad.json"
        bad.write_text(json.dumps({"title": "x", "slides": [{"layout": "pie_chart", "title": "y"}]}))
        root = os.path.join(os.path.dirname(__file__), "..")
        result = subprocess.run(
            [sys.executable, os.path.join(root, "factory.py"), "--from-json", str(bad), "--output", str(tmp_path / "o.pptx")],
            capture_output=True, text=True, check=False,
        )
        assert result.returncode == 2
        assert "unknown layout 'pie_chart'" in result.stderr
        assert not (tmp_path / "o.pptx").exists()
