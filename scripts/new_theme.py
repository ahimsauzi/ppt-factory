#!/usr/bin/env python3
"""
scripts/new_theme.py — Scaffold a new theme entry for ppt_factory.

Usage:
    python scripts/new_theme.py --name obsidian

Prints a ready-to-paste Theme block.
"""
import argparse

TEMPLATE = '''
    # ── {name_title}: {description} ──────────────────────────────────────────
    "{name_lower}": Theme(
        name="{name_title}",
        bg_dark="",         bg_light="",
        bg_card="FFFFFF",   bg_card_dark="",
        accent_1="",        accent_2="",         accent_3="",
        text_on_dark="",    text_on_light="",    text_muted="6B7280",
        callout_bg="",      callout_text="",
        font_header="Calibri", font_body="Calibri",
    ),
'''

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scaffold a new ppt_factory theme")
    parser.add_argument("--name", required=True, help="Theme slug e.g. obsidian")
    parser.add_argument("--description", default="describe the style", help="One-line style description")
    args = parser.parse_args()

    print(TEMPLATE.format(
        name_lower=args.name.lower(),
        name_title=args.name.title(),
        description=args.description,
    ))
    print("→ Add the block above to src/ppt_factory/themes/__init__.py inside THEMES = {...}")
