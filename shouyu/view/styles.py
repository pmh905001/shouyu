"""Shared QSS / palette constants for the PySide6 windows."""
from __future__ import annotations

PENDING_COLOR_HEX = "#808080"
IN_PROGRESS_COLOR_HEX = "#C00000"
DONE_COLOR_HEX = "#107C10"
ACCENT_COLOR_HEX = "#0F62FE"
BG_COLOR_HEX = "#1F1F1F"
PANEL_COLOR_HEX = "#2B2B2B"
TEXT_COLOR_HEX = "#E6E6E6"
SUBTEXT_COLOR_HEX = "#9A9A9A"

GLOBAL_QSS = f"""
QWidget {{
    background-color: {BG_COLOR_HEX};
    color: {TEXT_COLOR_HEX};
    font-family: "Segoe UI", "Microsoft YaHei UI", sans-serif;
    font-size: 14px;
}}
QLabel#TitleLabel {{
    font-size: 22px;
    font-weight: 700;
}}
QLabel#SubtitleLabel {{
    color: {SUBTEXT_COLOR_HEX};
    font-size: 13px;
}}
QLabel#HintLabel {{
    color: {SUBTEXT_COLOR_HEX};
    font-size: 12px;
}}
QPushButton {{
    background-color: {PANEL_COLOR_HEX};
    color: {TEXT_COLOR_HEX};
    border: 1px solid #3A3A3A;
    border-radius: 6px;
    padding: 8px 18px;
}}
QPushButton:hover {{
    background-color: #353535;
}}
QPushButton#PrimaryButton {{
    background-color: {ACCENT_COLOR_HEX};
    color: white;
    border: none;
    font-weight: 600;
}}
QPushButton#PrimaryButton:hover {{
    background-color: #2D7BF1;
}}
QListWidget {{
    background-color: {PANEL_COLOR_HEX};
    border: 1px solid #3A3A3A;
    border-radius: 8px;
    padding: 6px;
    outline: 0;
}}
QListWidget::item {{
    padding: 10px 12px;
    border-radius: 6px;
}}
QListWidget::item:selected {{
    background-color: rgba(15, 98, 254, 0.25);
    color: {TEXT_COLOR_HEX};
}}
QLineEdit {{
    background-color: #181818;
    border: 1px solid #3A3A3A;
    border-radius: 6px;
    padding: 6px 8px;
    color: {TEXT_COLOR_HEX};
}}
QLineEdit:focus {{
    border: 1px solid {ACCENT_COLOR_HEX};
}}
"""
