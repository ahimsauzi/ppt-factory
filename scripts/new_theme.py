#!/usr/bin/env python3
"""
scripts/new_theme.py — Scaffold a new theme entry for ppt_factory.

Usage:
    python scripts/new_theme.py --name obsidian

Prints a ready-to-paste THEMES entry for factory.py.
"""
import argparse

TEMPLATE = (
    '    "{name_lower}": {{"bg_dark": "", "bg_light": "", "bg_card": "FFFFFF", "bg_card_dark": "", '
    '"accent_1": "", "accent_2": "", "accent_3": "", "text_dark": "", "text_light": "", '
    '"text_muted": "6B7280", "callout_bg": "", "callout_text": "", "font": "Calibri"}},  # {description}'
)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scaffold a new ppt-factory theme")
    parser.add_argument("--name", required=True, help="Theme slug e.g. obsidian")
    parser.add_argument("--description", default="describe the style", help="One-line style description")
    args = parser.parse_args()

    print(TEMPLATE.format(
        name_lower=args.name.lower(),
        name_title=args.name.title(),
        description=args.description,
    ))
    print("→ Paste the line above into THEMES in factory.py and fill in the hex colors (6 digits, no #).")
