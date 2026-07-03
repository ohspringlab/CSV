"""
Streamlit web demo for Gestionnaire de Debit.

This web app reuses the same parser, transformer, and exporter as the
desktop application so client feedback applies to the final Windows build.
"""

import html
import json
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from core.exporter import export_csv, export_csv_drawers, export_excel, split_piece_names
from core.parser import FILE_COLORS, parse_csv_file
from core.transformer import transform


APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "config" / "nomenclature.json"
USER_SETTINGS_PATH = APP_DIR / "config" / "user_settings.json"

_USER_SETTINGS_DEFAULTS = {
    "excel_font_size": 14,
    "excel_col_scale": 0.75,
    "excel_dim_width": 18,
    "excel_cab_width": 10,
    "excel_mod_width": 14,
    # Default file highlight colors (index matches FILE_COLORS order)
    "file_color_0": "#FFFF00",   # yellow (stronger than pale #FFFACD)
    "file_color_1": "#ADD8E6",   # light blue
    "file_color_2": "#E0FFE0",   # light green
    "file_color_3": "#FFE4E1",   # light rose
    "file_color_4": "#F0E6FF",   # light lavender
    "file_color_5": "#FFE8CC",   # light orange
    "file_color_6": "#E6FFFA",   # light teal
    "file_color_7": "#FFFACD",   # light yellow
}


def load_user_settings() -> dict:
    """Load persistent user settings from disk, falling back to defaults."""
    try:
        with USER_SETTINGS_PATH.open("r", encoding="utf-8") as f:
            saved = json.load(f)
        return {**_USER_SETTINGS_DEFAULTS, **saved}
    except Exception:
        return dict(_USER_SETTINGS_DEFAULTS)


def save_user_settings(settings: dict):
    """Persist user settings to disk."""
    try:
        with USER_SETTINGS_PATH.open("w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass

TRANSLATIONS = {
    "en": {
        "language": "Language",
        "current_language": "Current language: English",
        "app_title": "Cut List Manager",
        "app_subtitle": "CSV import · sorted preview · Excel/CSV export",
        "nomenclature": "Nomenclature",
        "part_names": "Part names / column order",
        "part_names_help": "One part name per line. Change these to test future naming changes.",
        "download_nomenclature": "Download nomenclature JSON",
        "file_uploader": "Import one or more CSV files",
        "empty_upload": "Upload the client's CSV files to preview the sorted cut list.",
        "loaded_files": "Loaded files",
        "sorted_preview": "Sorted preview",
        "plain_preview": "Plain data preview",
        "export": "Export",
        "download_excel": "⬇ Download Excel",
        "download_csv": "⬇ Download CSV",
        "cabinet": "Cabinet",
        "qty": "Qty",
        "dimensions": "Dimensions",
        "theme_settings": "Theme editor",
        "theme_primary": "Favorite color",
        "theme_secondary": "Accent color",
        "theme_button_style": "Button style",
        "theme_font_color": "Font color",
        "theme_table_font_color": "Table font color",
        "theme_font_family": "Font style",
        "theme_hover_effect": "Hover effect",
        "theme_radius": "Roundness",
        "theme_animation": "Animation effects",
        "theme_reset": "Reset theme",
        "theme_gradient": "Gradient",
        "theme_solid": "Solid",
        "theme_outline": "Outline",
        "theme_font_modern": "Modern",
        "theme_font_classic": "Classic",
        "theme_font_rounded": "Rounded",
        "theme_font_mono": "Mono",
        "theme_hover_lift": "Lift",
        "theme_hover_glow": "Glow",
        "theme_hover_zoom": "Zoom",
        "theme_hover_none": "None",
        "empty_upload_hint": "Drag & drop or click Browse files above.",
        "parsing": "Parsing files…",
        "files_loaded": "Files",
        "cabinets_word": "Cabinets",
        "parts_configured": "Part types",
        "pieces_word": "Non-empty cells",
        "export_tip": "Excel preserves colors and formatting. CSV is ideal for further processing.",
        "history_title": "Session history",
        "history_empty": "No previous sessions yet.",
        "history_session": "Session",
        "history_clear": "Clear history",
        "chart_title": "Pieces per cabinet",
        "chart_cabinet": "Cabinet",
        "chart_pieces": "Piece groups",
        "coverage_label": "Coverage",
        "search_placeholder": "Search cabinet or part…",
        "search_label": "Quick search",
        "warnings_title": "Warnings",
        "warn_empty_cabinet": "Cabinet with no pieces",
        "warn_high_qty": "Unusually high quantity",
        "warn_missing_dim": "Missing dimensions",
        "no_warnings": "No issues detected",
        "dim_chart_title": "Dimension spread (W × L)",
        "dim_chart_w": "Width",
        "dim_chart_l": "Length",
        "dim_chart_qty": "Qty",
        "copy_table": "Copy table as text",
        "copy_done": "Copied!",
        "donut_filled": "Filled",
        "donut_empty": "Empty",
        "analysis_title": "Analysis",
        "completeness": "Completeness",
        "sort_label": "Sort cut list by",
        "sort_cabinet": "Cabinet ID",
        "sort_completeness": "Completeness",
        "sort_alpha": "Alphabetical",
        "part_summary_title": "Part summary",
        "part_summary_part": "Part",
        "part_summary_total_qty": "Total qty",
        "part_summary_cabinets": "Cabinets",
        "part_summary_min_dim": "Min dim",
        "part_summary_max_dim": "Max dim",
        "material_title": "Material estimate",
        "material_part": "Part",
        "material_total_qty": "Total pcs",
        "material_sqft": "Total sq in",
        "material_sqm": "Total cm²",
        "duplicates_title": "Duplicate cabinets",
        "duplicates_found": "Cabinet in multiple files",
        "duplicates_none": "No duplicate cabinet IDs",
        "notes_label": "Job reference / notes",
        "notes_placeholder": "e.g. Client name, project #, date…",
        "drilldown_title": "Cabinet detail",
        "drilldown_select": "Select a cabinet",
        "drilldown_none": "Select a cabinet above to inspect it",
        "drawer_preview": "Drawer cut list",
        "no_drawer_parts": "No drawer parts in this selection.",
        "download_csv_drawers": "⬇ Download Drawer CSV",
        "file_colors": "File highlight colors",
        "excel_export_settings": "Excel export settings",
        "excel_font_size": "Font size",
        "excel_col_scale": "Column width scale",
        "excel_col_scale_help": "Scale all dimension columns. 0.75 = compact, 1.0 = standard.",
        "excel_dim_width": "Min dimension column width",
        "excel_cab_width": "Cabinet column width",
        "unknown_pieces": "unrecognized piece name(s) — assign them to a column below",
        "ignore": "ignore",
        "file_prefix_label": "Cabinet prefix (optional)",
        "file_prefix_help": "Add a prefix to bare N/alpha cabinet IDs in this file (e.g. R5:)",
    },
    "fr": {
        "language": "Langue",
        "current_language": "Langue actuelle : français",
        "app_title": "Gestionnaire de Débit",
        "app_subtitle": "Import CSV · aperçu trié · export Excel/CSV",
        "nomenclature": "Nomenclature",
        "part_names": "Noms des pièces / ordre des colonnes",
        "part_names_help": "Un nom de pièce par ligne. Modifiez cette liste pour tester les changements de nomenclature.",
        "download_nomenclature": "Télécharger la nomenclature JSON",
        "file_uploader": "Importer un ou plusieurs fichiers CSV",
        "empty_upload": "Importez les fichiers CSV du client pour voir la liste de débit triée.",
        "loaded_files": "Fichiers chargés",
        "sorted_preview": "Aperçu trié",
        "plain_preview": "Aperçu des données",
        "export": "Export",
        "download_excel": "⬇ Télécharger Excel",
        "download_csv": "⬇ Télécharger CSV",
        "cabinet": "Meuble",
        "qty": "Qté",
        "dimensions": "Dimensions",
        "theme_settings": "Éditeur de thème",
        "theme_primary": "Couleur favorite",
        "theme_secondary": "Couleur d'accent",
        "theme_button_style": "Style des boutons",
        "theme_font_color": "Couleur du texte",
        "theme_table_font_color": "Couleur du texte du tableau",
        "theme_font_family": "Style de police",
        "theme_hover_effect": "Effet au survol",
        "theme_radius": "Arrondi",
        "theme_animation": "Effets d'animation",
        "theme_reset": "Réinitialiser le thème",
        "theme_gradient": "Dégradé",
        "theme_solid": "Uni",
        "theme_outline": "Contour",
        "theme_font_modern": "Moderne",
        "theme_font_classic": "Classique",
        "theme_font_rounded": "Arrondie",
        "theme_font_mono": "Mono",
        "theme_hover_lift": "Lever",
        "theme_hover_glow": "Lueur",
        "theme_hover_zoom": "Zoom",
        "theme_hover_none": "Aucun",
        "empty_upload_hint": "Glissez-déposez ou cliquez sur Parcourir ci-dessus.",
        "parsing": "Analyse des fichiers en cours…",
        "files_loaded": "Fichiers",
        "cabinets_word": "Meubles",
        "parts_configured": "Types de pièces",
        "pieces_word": "Cellules non vides",
        "export_tip": "Excel conserve les couleurs et la mise en forme. CSV est idéal pour un traitement ultérieur.",
        "history_title": "Historique des sessions",
        "history_empty": "Aucune session précédente.",
        "history_session": "Session",
        "history_clear": "Effacer l'historique",
        "chart_title": "Pièces par meuble",
        "chart_cabinet": "Meuble",
        "chart_pieces": "Groupes de pièces",
        "coverage_label": "Couverture",
        "search_placeholder": "Rechercher meuble ou pièce…",
        "search_label": "Recherche rapide",
        "warnings_title": "Avertissements",
        "warn_empty_cabinet": "Meuble sans pièces",
        "warn_high_qty": "Quantité inhabituellement élevée",
        "warn_missing_dim": "Dimensions manquantes",
        "no_warnings": "Aucun problème détecté",
        "dim_chart_title": "Répartition des dimensions (L × l)",
        "dim_chart_w": "Largeur",
        "dim_chart_l": "Longueur",
        "dim_chart_qty": "Qté",
        "copy_table": "Copier le tableau en texte",
        "copy_done": "Copié !",
        "donut_filled": "Rempli",
        "donut_empty": "Vide",
        "analysis_title": "Analyse",
        "completeness": "Complétude",
        "sort_label": "Trier la liste par",
        "sort_cabinet": "N° de meuble",
        "sort_completeness": "Complétude",
        "sort_alpha": "Alphabétique",
        "part_summary_title": "Résumé des pièces",
        "part_summary_part": "Pièce",
        "part_summary_total_qty": "Qté totale",
        "part_summary_cabinets": "Meubles",
        "part_summary_min_dim": "Dim. min",
        "part_summary_max_dim": "Dim. max",
        "material_title": "Estimation matière",
        "material_part": "Pièce",
        "material_total_qty": "Nb pièces",
        "material_sqft": "Total po²",
        "material_sqm": "Total cm²",
        "duplicates_title": "Doublons",
        "duplicates_found": "Meuble présent dans plusieurs fichiers",
        "duplicates_none": "Aucun doublon détecté",
        "notes_label": "Référence / notes de travail",
        "notes_placeholder": "ex. Nom client, n° projet, date…",
        "drilldown_title": "Détail du meuble",
        "drilldown_select": "Choisir un meuble",
        "drilldown_none": "Sélectionnez un meuble ci-dessus pour l'inspecter",
        "drawer_preview": "Liste de débit tiroirs",
        "no_drawer_parts": "Aucune pièce de tiroir dans cette sélection.",
        "download_csv_drawers": "⬇ Télécharger CSV tiroirs",
        "file_colors": "Couleurs de surlignage par fichier",
        "excel_export_settings": "Paramètres export Excel",
        "excel_font_size": "Taille de police",
        "excel_col_scale": "Échelle largeur colonnes",
        "excel_col_scale_help": "0.75 = compact, 1.0 = standard, 1.2 = large.",
        "excel_dim_width": "Largeur min. col. dimensions",
        "excel_cab_width": "Largeur col. meuble",
        "unknown_pieces": "nom(s) de pièce non reconnu(s) — assignez-les à une colonne ci-dessous",
        "ignore": "ignorer",
        "file_prefix_label": "Préfixe meuble (optionnel)",
        "file_prefix_help": "Ajouter un préfixe aux ID de meuble N/alpha de ce fichier (ex: R5:)",
    },
}

DEFAULT_THEME = {
    "primary": "#38bdf8",
    "secondary": "#7c3aed",
    "font_color": "#e2e8f0",
    "table_font_color": "#1e293b",
    "font_family": "modern",
    "hover_effect": "lift",
    "button_style": "gradient",
    "radius": 16,
    "animation": True,
}

BUTTON_STYLE_OPTIONS = {
    "gradient": "theme_gradient",
    "solid": "theme_solid",
    "outline": "theme_outline",
}

FONT_FAMILY_OPTIONS = {
    "modern": ("theme_font_modern", "'Inter', 'Segoe UI', sans-serif"),
    "classic": ("theme_font_classic", "Georgia, 'Times New Roman', serif"),
    "rounded": ("theme_font_rounded", "'Trebuchet MS', 'Segoe UI', sans-serif"),
    "mono": ("theme_font_mono", "'Courier New', monospace"),
}

HOVER_EFFECT_OPTIONS = {
    "lift": "theme_hover_lift",
    "glow": "theme_hover_glow",
    "zoom": "theme_hover_zoom",
    "none": "theme_hover_none",
}


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    hex_color = hex_color.lstrip("#")
    return (
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


def get_theme() -> dict:
    for key, value in DEFAULT_THEME.items():
        st.session_state.setdefault(f"theme_{key}", value)

    return {
        "primary": st.session_state.theme_primary,
        "secondary": st.session_state.theme_secondary,
        "font_color": st.session_state.theme_font_color,
        "table_font_color": st.session_state.theme_table_font_color,
        "font_family": st.session_state.theme_font_family,
        "hover_effect": st.session_state.theme_hover_effect,
        "button_style": st.session_state.theme_button_style,
        "radius": st.session_state.theme_radius,
        "animation": st.session_state.theme_animation,
    }


def render_theme_editor(labels: dict) -> dict:
    st.markdown(
        f'<div style="margin:18px 0 8px;font-size:11px;font-weight:700;'
        f'letter-spacing:1.2px;text-transform:uppercase;color:var(--theme-primary);">'
        f'{html.escape(labels["theme_settings"])}</div>',
        unsafe_allow_html=True,
    )

    primary = st.color_picker(
        labels["theme_primary"],
        key="theme_primary",
    )
    secondary = st.color_picker(
        labels["theme_secondary"],
        key="theme_secondary",
    )
    font_color = st.color_picker(
        labels["theme_font_color"],
        key="theme_font_color",
    )
    table_font_color = st.color_picker(
        labels["theme_table_font_color"],
        key="theme_table_font_color",
    )

    option_keys = list(BUTTON_STYLE_OPTIONS.keys())
    selected_style = st.selectbox(
        labels["theme_button_style"],
        option_keys,
        format_func=lambda option: labels[BUTTON_STYLE_OPTIONS[option]],
        key="theme_button_style",
    )
    font_keys = list(FONT_FAMILY_OPTIONS.keys())
    selected_font = st.selectbox(
        labels["theme_font_family"],
        font_keys,
        format_func=lambda option: labels[FONT_FAMILY_OPTIONS[option][0]],
        key="theme_font_family",
    )
    hover_keys = list(HOVER_EFFECT_OPTIONS.keys())
    selected_hover = st.selectbox(
        labels["theme_hover_effect"],
        hover_keys,
        format_func=lambda option: labels[HOVER_EFFECT_OPTIONS[option]],
        key="theme_hover_effect",
    )
    radius = st.slider(
        labels["theme_radius"],
        min_value=4,
        max_value=28,
        key="theme_radius",
    )
    animation = st.toggle(
        labels["theme_animation"],
        key="theme_animation",
    )

    if st.button(labels["theme_reset"], use_container_width=True, key="theme_reset"):
        for key, value in DEFAULT_THEME.items():
            st.session_state[f"theme_{key}"] = value
        st.rerun()

    return {
        "primary": primary,
        "secondary": secondary,
        "font_color": font_color,
        "table_font_color": table_font_color,
        "font_family": selected_font,
        "hover_effect": selected_hover,
        "button_style": selected_style,
        "radius": radius,
        "animation": animation,
    }


def theme_css(theme: dict) -> str:
    pr, pg, pb = _hex_to_rgb(theme["primary"])
    sr, sg, sb = _hex_to_rgb(theme["secondary"])
    animation = "running" if theme["animation"] else "paused"
    transitions = "180ms ease" if theme["animation"] else "0ms linear"
    font_stack = FONT_FAMILY_OPTIONS[theme["font_family"]][1]

    if theme["button_style"] == "solid":
        button_background = theme["primary"]
        button_border = f"1px solid rgba({pr},{pg},{pb},0.40)"
        button_color = "#ffffff"
        button_shadow = f"0 2px 8px rgba({pr},{pg},{pb},0.25)"
    elif theme["button_style"] == "outline":
        button_background = "transparent"
        button_border = f"1.5px solid rgba({pr},{pg},{pb},0.70)"
        button_color = theme["primary"]
        button_shadow = "none"
    else:
        button_background = (
            f"linear-gradient(135deg, {theme['primary']} 0%, {theme['secondary']} 100%)"
        )
        button_border = "1px solid transparent"
        button_color = "#ffffff"
        button_shadow = f"0 2px 12px rgba({pr},{pg},{pb},0.22)"

    if not theme["animation"] or theme["hover_effect"] == "none":
        hover_transform = "none"
        hover_filter   = "none"
        hover_shadow   = button_shadow
        row_hover_bg   = "rgba(255,255,255,0.02)"
    elif theme["hover_effect"] == "glow":
        hover_transform = "translateY(-1px)"
        hover_filter    = "brightness(1.08)"
        hover_shadow    = f"0 0 20px rgba({pr},{pg},{pb},0.35)"
        row_hover_bg    = "rgba(255,255,255,0.03)"
    elif theme["hover_effect"] == "zoom":
        hover_transform = "scale(1.02)"
        hover_filter    = "brightness(1.06)"
        hover_shadow    = f"0 6px 24px rgba({pr},{pg},{pb},0.22)"
        row_hover_bg    = "rgba(255,255,255,0.03)"
    else:
        hover_transform = "translateY(-2px)"
        hover_filter    = "brightness(1.08)"
        hover_shadow    = f"0 8px 24px rgba({pr},{pg},{pb},0.20)"
        row_hover_bg    = "rgba(255,255,255,0.03)"

    return f"""
    <style>
    :root {{
        --p:    {theme["primary"]};
        --s:    {theme["secondary"]};
        --fc:   {theme["font_color"]};
        --tfc:  {theme["table_font_color"]};
        --ff:   {font_stack};
        --pr:   {pr}; --pg: {pg}; --pb: {pb};
        --sr:   {sr}; --sg: {sg}; --sb: {sb};
        --rad:  {theme["radius"]}px;
        --anim: {animation};
        --tr:   {transitions};
        --bb:   {button_background};
        --bbo:  {button_border};
        --bc:   {button_color};
        --bs:   {button_shadow};
        --ht:   {hover_transform};
        --hf:   {hover_filter};
        --hs:   {hover_shadow};
        --rhb:  {row_hover_bg};
        /* aliases kept for legacy HTML */
        --theme-primary:       {theme["primary"]};
        --theme-secondary:     {theme["secondary"]};
        --theme-font-color:    {theme["font_color"]};
        --theme-table-font-color: {theme["table_font_color"]};
        --theme-font-family:   {font_stack};
        --theme-primary-rgb:   {pr}, {pg}, {pb};
        --theme-secondary-rgb: {sr}, {sg}, {sb};
        --theme-radius:        {theme["radius"]}px;
        --theme-transition:    {transitions};
        --theme-button-shadow: {button_shadow};
        --theme-hover-transform: {hover_transform};
        --theme-hover-filter:  {hover_filter};
        --theme-hover-shadow:  {hover_shadow};
    }}

    /* ── App shell ── */
    .stApp {{
        color: var(--fc) !important;
        font-family: var(--ff) !important;
        background: #0c1220 !important;
        animation-play-state: var(--anim) !important;
    }}
    [data-testid="stAppViewContainer"] > .main {{ background: transparent; }}

    /* ── Title area ── */
    .app-title {{
        background: linear-gradient(120deg, #f1f5f9 20%, var(--p) 70%) !important;
        -webkit-background-clip: text !important;
        -webkit-text-fill-color: transparent !important;
        background-clip: text !important;
    }}
    .app-subtitle, h2, h3,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span,
    [data-testid="stFileUploader"] p,
    [data-testid="stFileUploader"] span,
    [data-testid="stFileUploader"] label,
    [data-testid="stExpander"] summary,
    [data-testid="stAlert"],
    .section-hdr span:last-child {{
        color: var(--fc) !important;
        font-family: var(--ff) !important;
    }}

    /* ── Hero accents ── */
    .app-hero {{ border-radius: var(--rad) !important; }}
    .app-hero::after {{
        background: linear-gradient(90deg, var(--p), var(--s)) !important;
    }}
    .hero-badge, .language-status {{
        color: var(--p) !important;
        border-color: rgba(var(--pr),var(--pg),var(--pb),0.28) !important;
        background: rgba(var(--pr),var(--pg),var(--pb),0.07) !important;
    }}
    .hero-badge-dot {{ background: var(--p) !important; }}

    /* ── Buttons ── */
    div.stButton > button, div.stDownloadButton > button {{
        font-family: var(--ff) !important;
        color: var(--bc) !important;
        background: var(--bb) !important;
        border: var(--bbo) !important;
        border-radius: calc(var(--rad) - 4px) !important;
        box-shadow: var(--bs) !important;
        transition: transform var(--tr), box-shadow var(--tr), filter var(--tr) !important;
        animation-play-state: var(--anim) !important;
    }}
    div.stButton > button:hover, div.stDownloadButton > button:hover {{
        transform: var(--ht) !important;
        filter: var(--hf) !important;
        box-shadow: var(--hs) !important;
    }}
    .theme-hover-card:hover {{
        transform: var(--ht) !important;
        filter: var(--hf) !important;
    }}

    /* ── Table theming ── */
    table.cutlist {{ font-family: var(--ff) !important; }}
    .cutlist .cab, .cutlist .cab-hdr {{
        color: var(--p) !important;
        border-right-color: rgba(var(--pr),var(--pg),var(--pb),0.20) !important;
    }}
    .cutlist td.data  {{ color: var(--tfc) !important; }}
    .cutlist td.data-light {{ color: var(--fc) !important; }}
    .cutlist tbody tr:hover td {{ background: var(--rhb) !important; }}
    .cutlist-wrap {{ border-radius: var(--rad) !important; }}

    /* ── Borders & focus states ── */
    [data-testid="stSidebar"] {{
        border-right-color: rgba(var(--pr),var(--pg),var(--pb),0.12) !important;
    }}
    [data-testid="stSidebar"] textarea:focus,
    [data-testid="stFileUploader"],
    [data-testid="stExpander"],
    [data-testid="stAlert"],
    [data-testid="stDataFrame"] {{
        border-color: rgba(var(--pr),var(--pg),var(--pb),0.22) !important;
    }}
    [data-testid="stFileUploader"] section {{
        border-color: rgba(var(--pr),var(--pg),var(--pb),0.30) !important;
    }}
    .section-hdr {{ border-bottom-color: rgba(var(--pr),var(--pg),var(--pb),0.18) !important; }}
    .empty-state  {{
        border-color: rgba(var(--pr),var(--pg),var(--pb),0.22) !important;
        background:   rgba(var(--pr),var(--pg),var(--pb),0.05) !important;
    }}
    ::-webkit-scrollbar-thumb {{
        background: rgba(var(--pr),var(--pg),var(--pb),0.35) !important;
    }}
    ::-webkit-scrollbar-thumb:hover {{
        background: rgba(var(--pr),var(--pg),var(--pb),0.55) !important;
    }}

    *, *::before, *::after {{ animation-play-state: var(--anim); }}
    </style>
    """


def inject_global_style(theme: dict):
    st.markdown(theme_css(theme), unsafe_allow_html=True)
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

        /* ── Keyframes ── */
        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(14px); }
            to   { opacity: 1; transform: translateY(0); }
        }
        @keyframes shimmer {
            0%   { transform: translateX(-120%); }
            100% { transform: translateX(120%); }
        }
        @keyframes pulseDot {
            0%, 100% { opacity: 1; }
            50%       { opacity: 0.35; }
        }
        @keyframes barFill {
            from { width: 0; }
        }
        @keyframes gradientShift {
            0%   { background-position: 0% 50%; }
            50%  { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        /* keep legacy name used by progress-bar */
        @keyframes floatUp {
            from { opacity: 0; transform: translateY(14px); }
            to   { opacity: 1; transform: translateY(0); }
        }

        /* ── Root ── */
        .stApp {
            font-family: 'Inter', sans-serif !important;
            background: #0c1220 !important;
            min-height: 100vh;
        }
        [data-testid="stAppViewContainer"] > .main { background: transparent; }

        /* ── Top bar ── */
        [data-testid="stHeader"] {
            background: rgba(12,18,32,0.95);
            backdrop-filter: blur(10px);
            border-bottom: 1px solid rgba(255,255,255,0.06);
        }

        /* ── Content width & entrance ── */
        .block-container {
            max-width: 1440px;
            padding-top: 1.8rem;
            padding-bottom: 5rem;
            animation: fadeUp 400ms ease both;
        }

        /* ── Hero banner ── */
        .app-hero {
            position: relative;
            overflow: hidden;
            padding: 32px 40px 28px;
            margin-bottom: 32px;
            border-radius: 14px;
            border: 1px solid rgba(255,255,255,0.07);
            background: #111827;
        }
        .app-hero::before {
            content: "";
            position: absolute;
            inset: 0;
            background: linear-gradient(110deg,
                transparent 0%, rgba(56,189,248,0.05) 50%, transparent 80%);
            transform: translateX(-120%);
            animation: shimmer 9s ease-in-out infinite;
            pointer-events: none;
        }
        .app-hero::after {
            content: "";
            position: absolute;
            bottom: 0; left: 0; right: 0; height: 2px;
            border-radius: 0 0 14px 14px;
            background: linear-gradient(90deg, #38bdf8 0%, #818cf8 100%);
            background-size: 200% 100%;
            animation: gradientShift 7s linear infinite;
        }
        .app-title {
            margin: 0 0 6px;
            font-size: clamp(24px, 3vw, 40px);
            font-weight: 800;
            line-height: 1.1;
            letter-spacing: -0.4px;
        }
        .app-subtitle {
            margin: 0;
            font-size: 13px;
            font-weight: 400;
            color: #4b5563;
            letter-spacing: 0.2px;
        }
        .hero-badge {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            margin-bottom: 12px;
            padding: 3px 10px;
            border-radius: 999px;
            font-size: 10px;
            font-weight: 600;
            letter-spacing: 0.8px;
            text-transform: uppercase;
        }
        .hero-badge-dot {
            width: 6px; height: 6px;
            border-radius: 50%;
            animation: pulseDot 2.2s ease-in-out infinite;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_component_style():
    """Component CSS — sidebar, buttons, table, widgets, cards."""
    st.markdown(
        """
        <style>
        /* ── Feedback banners ── */
        .success-message {
            padding: 12px 18px;
            border-radius: 10px;
            background: rgba(34,197,94,0.08);
            border: 1px solid rgba(34,197,94,0.25);
            color: #86efac;
            font-weight: 500;
            font-size: 13px;
            margin: 10px 0;
        }
        .progress-bar-container {
            width: 100%;
            height: 3px;
            background: rgba(255,255,255,0.06);
            border-radius: 999px;
            overflow: hidden;
            margin: 8px 0 0;
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #38bdf8, #818cf8);
            background-size: 200% 100%;
            animation: gradientShift 1.6s ease infinite;
            border-radius: 999px;
            width: 100%;
        }

        /* ── Sidebar ── */
        [data-testid="stSidebar"] {
            background: #0d1424 !important;
            border-right: 1px solid rgba(255,255,255,0.06) !important;
        }
        [data-testid="stSidebar"] > div:first-child { padding-top: 1.2rem; }
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] span { color: #94a3b8 !important; }
        [data-testid="stSidebar"] textarea {
            color: #e2e8f0 !important;
            background: #1a2236 !important;
            border: 1px solid rgba(255,255,255,0.10) !important;
            border-radius: 8px !important;
            font-family: 'Inter', monospace !important;
            font-size: 12px !important;
        }
        [data-testid="stSidebar"] textarea:focus {
            border-color: rgba(56,189,248,0.45) !important;
            box-shadow: 0 0 0 2px rgba(56,189,248,0.10) !important;
            outline: none !important;
        }
        .language-status {
            margin: 4px 0 16px;
            padding: 8px 12px;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: 0.4px;
            color: #38bdf8;
            border: 1px solid rgba(56,189,248,0.20);
            border-radius: 8px;
            background: rgba(56,189,248,0.05);
        }

        /* ── Buttons ── */
        div.stButton > button,
        div.stDownloadButton > button {
            border-radius: 8px;
            font-family: 'Inter', sans-serif;
            font-weight: 600;
            font-size: 13px;
            letter-spacing: 0.1px;
            padding: 0.5rem 1rem;
            transition: transform 150ms ease, box-shadow 150ms ease, filter 150ms ease;
        }
        div.stButton > button:active,
        div.stDownloadButton > button:active {
            transform: scale(0.98) !important;
        }

        /* ── File uploader ── */
        [data-testid="stFileUploader"] {
            padding: 16px 20px;
            border: 1.5px dashed rgba(255,255,255,0.14);
            border-radius: 12px;
            background: #111827;
        }
        [data-testid="stFileUploader"] section {
            border-color: rgba(255,255,255,0.10) !important;
            background: #1a2236 !important;
            border-radius: 8px !important;
        }
        [data-testid="stFileUploader"] p,
        [data-testid="stFileUploader"] span,
        [data-testid="stFileUploader"] label { color: #64748b !important; }

        /* ── Section headers (native) ── */
        h2, h3 {
            font-family: 'Inter', sans-serif !important;
            color: #e2e8f0 !important;
            font-weight: 700 !important;
            letter-spacing: -0.2px;
        }

        /* ── Expander ── */
        [data-testid="stExpander"] {
            border: 1px solid rgba(255,255,255,0.08) !important;
            border-radius: 10px !important;
            background: #111827 !important;
        }
        [data-testid="stExpander"] summary {
            color: #64748b !important;
            font-weight: 600;
            font-size: 13px;
        }

        /* ── Alert box ── */
        [data-testid="stAlert"] {
            background: #1a2236 !important;
            border: 1px solid rgba(255,255,255,0.08) !important;
            border-radius: 10px !important;
            color: #64748b !important;
        }

        /* ── Scrollbar ── */
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 999px; }
        ::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.22); }

        /* ── DataFrame ── */
        [data-testid="stDataFrame"] {
            border: 1px solid rgba(255,255,255,0.08) !important;
            border-radius: 10px !important;
            overflow: hidden;
        }

        /* ── Stat card base & tooltip ── */
        .stat-card-wrap { position: relative; cursor: default; }
        .stat-card {
            padding: 20px 18px 16px;
            border-radius: 12px;
            background: #111827;
            border: 1px solid rgba(255,255,255,0.07);
            text-align: center;
            transition: transform 180ms ease, border-color 180ms ease, box-shadow 180ms ease;
        }
        .stat-card-wrap:hover .stat-card {
            transform: translateY(-3px);
            border-color: rgba(56,189,248,0.30);
            box-shadow: 0 8px 24px rgba(0,0,0,0.35);
        }
        .stat-tooltip {
            display: none;
            position: absolute;
            top: calc(100% + 8px);
            left: 50%;
            transform: translateX(-50%);
            min-width: 220px;
            max-width: 260px;
            padding: 14px 16px;
            border-radius: 12px;
            background: #1e293b;
            border: 1px solid rgba(255,255,255,0.10);
            box-shadow: 0 16px 48px rgba(0,0,0,0.55);
            z-index: 9999;
            text-align: left;
            animation: fadeUp 180ms ease both;
            pointer-events: none;
        }
        .stat-tooltip::before {
            content: "";
            position: absolute;
            top: -6px;
            left: 50%;
            transform: translateX(-50%) rotate(45deg);
            width: 11px; height: 11px;
            background: #1e293b;
            border-left: 1px solid rgba(255,255,255,0.10);
            border-top: 1px solid rgba(255,255,255,0.10);
        }
        .stat-card-wrap:hover .stat-tooltip { display: block; }

        /* ── Section divider ── */
        .section-hdr {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 32px 0 16px;
            padding-bottom: 10px;
            border-bottom: 1px solid rgba(255,255,255,0.07);
        }
        .section-hdr-icon {
            width: 28px; height: 28px;
            border-radius: 7px;
            background: rgba(56,189,248,0.10);
            display: flex; align-items: center; justify-content: center;
            font-size: 14px;
        }
        .section-hdr-text {
            font-size: 15px;
            font-weight: 700;
            color: #e2e8f0;
            letter-spacing: -0.1px;
        }

        /* ── Empty state ── */
        .empty-state {
            margin-top: 20px;
            padding: 48px 32px;
            border: 1.5px dashed rgba(255,255,255,0.10);
            border-radius: 14px;
            background: #111827;
            text-align: center;
        }

        /* ── Export tip ── */
        .export-tip {
            margin: -4px 0 14px;
            padding: 10px 14px;
            border-radius: 8px;
            background: rgba(56,189,248,0.05);
            border: 1px solid rgba(56,189,248,0.12);
            color: #475569;
            font-size: 12px;
        }

        /* ── Footer ── */
        .app-footer {
            margin-top: 56px;
            padding: 18px 24px;
            border-radius: 10px;
            background: #111827;
            border: 1px solid rgba(255,255,255,0.06);
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 8px;
            font-size: 12px;
            color: #374151;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def load_piece_names() -> list[str]:
    """Return the combined list of main + drawer piece names from config."""
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        main = data.get("piece_names", [])
        drw  = data.get("drawer_piece_names", [])
        if main:
            return main + drw
    except Exception:
        pass

    # Hardcoded fallback — matches config defaults
    return [
        "PIGNON (L)", "PIGNON (R)", "DESSUS/DESSOUS", "DOS",
        "TAB AJUST", "PARTITION", "PORTES/FAÇADES",
        "Stile", "Poignées",
        "Drw L Side", "Drw R Side", "Drw Front", "Drw Back", "Drw Bottom",
        "Tray Bottom", "Tray Face", "Tray Side L", "Tray Side R", "Tray Back",
    ]


def render_language_buttons() -> str:
    if "language" not in st.session_state:
        st.session_state.language = "fr"

    language = st.session_state.language
    labels = TRANSLATIONS[language]

    st.markdown(
        """
        <div style="margin:0 0 6px;padding:8px 12px;border-radius:10px;
            background:rgba(56,189,248,0.07);border:1px solid rgba(56,189,248,0.18);
            font-size:11px;font-weight:700;letter-spacing:1px;
            text-transform:uppercase;color:#38bdf8;">
            🌐 &nbsp;Language / Langue
        </div>
        """,
        unsafe_allow_html=True,
    )
    col_en, col_fr = st.columns(2)
    with col_en:
        if st.button("🇬🇧 English", use_container_width=True, key="lang_en"):
            st.session_state.language = "en"
            st.rerun()
    with col_fr:
        if st.button("🇫🇷 Français", use_container_width=True, key="lang_fr"):
            st.session_state.language = "fr"
            st.rerun()

    st.markdown(
        f'<div class="language-status">✓ &nbsp;{TRANSLATIONS[language]["current_language"]}</div>',
        unsafe_allow_html=True,
    )
    return language


def parse_uploaded_files(uploaded_files, custom_colors: dict | None = None,
                         custom_prefixes: dict | None = None) -> tuple[list[dict], list[tuple[str, str]]]:
    all_rows = []
    imported_files = []

    for idx, uploaded in enumerate(uploaded_files):
        color = (custom_colors or {}).get(uploaded.name, FILE_COLORS[idx % len(FILE_COLORS)])
        suffix = Path(uploaded.name).suffix or ".csv"
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.getbuffer())
                tmp_path = tmp.name

            rows = parse_csv_file(tmp_path, source_color=color)

            # Apply manual cabinet-prefix override (e.g. add "R5:" to bare N-IDs)
            manual_prefix = (custom_prefixes or {}).get(uploaded.name, "").strip()
            if manual_prefix:
                if not manual_prefix.endswith(":"):
                    manual_prefix += ":"
                for row in rows:
                    cid = row.get("cabinet_id", "")
                    if cid and ":" not in cid and re.match(r'^[A-Za-z]', cid):
                        row["cabinet_id"] = manual_prefix + cid

            for row in rows:
                row["source_file"] = uploaded.name
            all_rows.extend(rows)
            imported_files.append((uploaded.name, color))
        finally:
            if tmp_path:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

    return all_rows, imported_files


def piece_entries(row_data: dict, piece_name: str) -> list[dict]:
    entries = row_data.get(f"{piece_name}_entries", [])
    if entries:
        return entries
    qty = row_data.get(f"{piece_name}_qty", "")
    if qty == "":
        return []
    return [{
        "qty": qty,
        "width": row_data.get(f"{piece_name}_w", ""),
        "length": row_data.get(f"{piece_name}_l", ""),
        "source_color": row_data.get(f"{piece_name}_color", "#FFFFFF"),
    }]


def entry_color(entries: list[dict]) -> str:
    colors = {
        entry.get("source_color", "#FFFFFF")
        for entry in entries
        if entry.get("source_color")
    }
    colors.discard("#FFFFFF")
    if not colors:
        return "transparent"
    if len(colors) == 1:
        return next(iter(colors))
    return "#FFE8CC"


def fmt_value(value) -> str:
    from core.parser import decimal_to_fraction_str
    if value == "" or value is None:
        return ""
    if isinstance(value, float):
        return decimal_to_fraction_str(value)
    return str(value)


def dimension_text(entry: dict) -> str:
    width = fmt_value(entry.get("width", ""))
    length = fmt_value(entry.get("length", ""))
    if not width or not length:
        return ""
    return f"{width} × {length}"


def _darken_hex(hex_color: str, amount: int = 30) -> str:
    """Slightly darken a hex color for borders."""
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r = max(0, r - amount)
    g = max(0, g - amount)
    b = max(0, b - amount)
    return f"#{r:02x}{g:02x}{b:02x}"


def preview_html(data: list[dict], piece_names: list[str], labels: dict) -> str:
    style = """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    .cutlist-outer { animation: fadeUp 350ms ease both; }
    .cutlist-wrap {
        overflow-x: auto;
        border: 1px solid rgba(255,255,255,0.07);
        border-radius: 12px;
        background: #111827;
    }
    .cutlist-wrap::-webkit-scrollbar { height: 5px; }
    .cutlist-wrap::-webkit-scrollbar-track { background: transparent; margin: 0 6px; }
    .cutlist-wrap::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.12); border-radius: 999px; }
    .cutlist-wrap::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.22); }

    table.cutlist {
        border-collapse: collapse;
        min-width: 1200px;
        width: max-content;
        font-family: 'Inter', sans-serif;
        font-size: 12px;
    }
    .cutlist th, .cutlist td {
        border-right: 1px solid rgba(255,255,255,0.05);
        border-bottom: 1px solid rgba(255,255,255,0.05);
        padding: 7px 10px;
        text-align: center;
        vertical-align: middle;
        white-space: nowrap;
    }
    .cutlist th:first-child, .cutlist td:first-child { border-left: none; }

    /* Top header row */
    .cutlist thead tr.top-row th {
        background: #1e293b;
        color: #94a3b8;
        font-weight: 600;
        font-size: 11px;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        border-bottom: 1px solid rgba(255,255,255,0.08);
        position: sticky;
        top: 0;
        z-index: 3;
        border-top: 2px solid #38bdf8;
    }
    .cutlist thead tr.top-row th.cab-hdr {
        border-top-color: #818cf8;
    }

    /* Sub-header row */
    .cutlist thead tr.sub-row th {
        background: #1a2236;
        color: #475569;
        font-size: 10px;
        font-weight: 600;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        border-bottom: 1px solid rgba(255,255,255,0.07);
        position: sticky;
        top: 33px;
        z-index: 3;
    }

    /* Cabinet column */
    .cutlist .cab-hdr {
        background: #1e293b !important;
        color: #818cf8 !important;
        font-weight: 700 !important;
        position: sticky !important;
        left: 0;
        z-index: 4 !important;
        border-right: 1px solid rgba(255,255,255,0.08) !important;
        min-width: 80px;
    }
    .cutlist .cab {
        background: #111827;
        color: #818cf8;
        font-weight: 700;
        font-size: 12px;
        position: sticky;
        left: 0;
        z-index: 2;
        border-right: 1px solid rgba(255,255,255,0.07) !important;
    }

    /* Data cells */
    .cutlist td.data {
        font-weight: 500;
        font-size: 12px;
        color: #1e293b;
    }
    .cutlist td.data-light { color: #cbd5e1 !important; }
    .cutlist td.empty { background: #0f172a; }

    /* Row hover */
    .cutlist tbody tr:hover td {
        background: rgba(255,255,255,0.025) !important;
    }
    .cutlist tbody tr:hover .cab {
        color: #a5b4fc !important;
        background: #131d2e !important;
    }
    </style>
    """

    ACCENT_CLASSES = ["grp-0"] * len(piece_names)  # single consistent accent

    out = [style, '<div class="cutlist-outer"><div class="cutlist-wrap">',
           '<table class="cutlist"><thead>']

    # Row 1: cabinet + piece group headers
    out.append('<tr class="top-row">')
    out.append(f'<th class="cab-hdr" rowspan="2">{html.escape(labels["cabinet"])}</th>')
    for name in piece_names:
        out.append(f'<th colspan="2">{html.escape(name)}</th>')
    out.append('</tr>')

    # Row 2: qty / dimensions
    out.append('<tr class="sub-row">')
    for _ in piece_names:
        out.append(
            f'<th>{html.escape(labels["qty"])}</th>'
            f'<th>{html.escape(labels["dimensions"])}</th>'
        )
    out.append('</tr></thead><tbody>')

    for row in data:
        # Find max number of entries across all pieces in this cabinet row
        all_entries = {name: piece_entries(row, name) for name in piece_names}

        # Skip cabinets that have no entries for any visible column
        if not any(all_entries.values()):
            continue

        max_rows = max((len(e) for e in all_entries.values()), default=1)
        max_rows = max(max_rows, 1)

        for sub_idx in range(max_rows):
            out.append('<tr>')

            # Cabinet ID cell — only on the first sub-row, spans all sub-rows
            if sub_idx == 0:
                out.append(
                    f'<td class="cab" rowspan="{max_rows}">'
                    f'{html.escape(str(row.get("cabinet_id", "")))}</td>'
                )

            for name in piece_names:
                entries = all_entries[name]

                if sub_idx >= len(entries):
                    # This piece has fewer entries than the max — empty continuation
                    out.append('<td class="empty"></td><td class="empty"></td>')
                    continue

                e = entries[sub_idx]
                src_color = e.get("source_color", "#FFFFFF")
                is_light  = src_color in ("#FFFFFF", "transparent", "")
                dc        = "data data-light" if is_light else "data"
                bg_style  = "" if is_light else f"background:{src_color};"

                qty  = html.escape(fmt_value(e.get("qty", "")))
                dims = html.escape(dimension_text(e))
                out.append(f'<td class="{dc}" style="{bg_style}">{qty}</td>')
                out.append(f'<td class="{dc}" style="{bg_style}">{dims}</td>')

            out.append('</tr>')

    out.append('</tbody></table></div></div>')
    return "".join(out)


def preview_dataframe(data: list[dict], piece_names: list[str], labels: dict) -> pd.DataFrame:
    records = []
    for row in data:
        record = {labels["cabinet"]: row.get("cabinet_id", "")}
        for name in piece_names:
            entries = piece_entries(row, name)
            record[f'{name} {labels["qty"]}'] = "\n".join(
                fmt_value(e.get("qty", "")) for e in entries
            )
            record[f'{name} {labels["dimensions"]}'] = "\n".join(
                dimension_text(e) for e in entries
            )
        records.append(record)
    return pd.DataFrame(records)


def export_bytes(export_func, data: list[dict], piece_names: list[str], suffix: str) -> bytes:
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp_path = tmp.name
        export_func(data, piece_names, tmp_path)
        with open(tmp_path, "rb") as f:
            return f.read()
    finally:
        if tmp_path:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass


def _legend_card(name: str, color: str) -> str:
    """Render a single file-legend chip as HTML."""
    is_colored = color not in ("#FFFFFF", "transparent")
    text_color = "#0f172a" if is_colored else "#cbd5e1"
    bg = color if is_colored else "#1e293b"
    border = _darken_hex(color, 20) if is_colored else "rgba(255,255,255,0.08)"
    return (
        f'<div class="theme-hover-card" style="display:flex;gap:8px;align-items:center;'
        f'padding:9px 14px;border:1px solid {border};'
        f'border-radius:9px;background:{bg};'
        f'transition:transform 150ms ease;">'
        f'<span style="width:10px;height:10px;background:{color};'
        f'border:1px solid rgba(255,255,255,0.20);border-radius:50%;'
        f'display:inline-block;flex-shrink:0;"></span>'
        f'<span style="color:{text_color};font-weight:600;font-size:12px;'
        f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px;">'
        f'{html.escape(name)}</span>'
        f'</div>'
    )


def _file_header_banners(imported_files: list[tuple[str, str]]) -> str:
    """
    Render one full-width banner per loaded file.
    Each banner shows the filename (without extension) highlighted in
    the color assigned to that file.
    White (#FFFFFF) files get a neutral dark card instead.
    """
    parts = []
    for fname, color in imported_files:
        stem = Path(fname).stem          # filename without extension
        is_colored = color not in ("#FFFFFF", "transparent")

        # Text readable on the file color
        r, g, b = _hex_to_rgb(color) if is_colored else (30, 41, 59)
        # luminance check → pick dark or light text
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        text_color = "#0f172a" if luminance > 0.55 else "#f1f5f9"

        bg      = color if is_colored else "#1e293b"
        border  = _darken_hex(color, 25) if is_colored else "rgba(255,255,255,0.08)"
        dot_shadow = f"0 0 8px {color}" if is_colored else "none"

        parts.append(
            f'<div style="'
            f'display:flex;align-items:center;gap:12px;'
            f'padding:14px 20px;'
            f'border-radius:10px;'
            f'background:{bg};'
            f'border:1px solid {border};'
            f'flex:1;min-width:0;'
            f'">'
            # colour dot
            f'<span style="'
            f'width:14px;height:14px;border-radius:50%;'
            f'background:{color};flex-shrink:0;'
            f'border:2px solid rgba(255,255,255,0.30);'
            f'box-shadow:{dot_shadow};'
            f'display:inline-block;"></span>'
            # filename stem
            f'<span style="'
            f'color:{text_color};'
            f'font-family:Inter,sans-serif;'
            f'font-size:15px;font-weight:700;'
            f'letter-spacing:-0.2px;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'
            f'">{html.escape(stem)}</span>'
            f'</div>'
        )

    # Wrap all banners side-by-side in a flex row
    inner = "".join(parts)
    return (
        f'<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px;">'
        + inner +
        f'</div>'
    )


def _section_header(icon: str, text: str) -> str:
    return (
        f'<div class="section-hdr">'
        f'<div class="section-hdr-icon">{icon}</div>'
        f'<span class="section-hdr-text">{html.escape(text)}</span>'
        f'</div>'
    )


# ── History helpers ────────────────────────────────────────────────────────────

MAX_HISTORY = 8

def record_session(files: int, cabinets: int, pieces: int, file_names: list[str]):
    """Append a snapshot of the current upload to session_state history."""
    if "_upload_history" not in st.session_state:
        st.session_state["_upload_history"] = []
    entry = {
        "time": datetime.now().strftime("%H:%M:%S"),
        "date": datetime.now().strftime("%d/%m/%Y"),
        "files": files,
        "cabinets": cabinets,
        "pieces": pieces,
        "names": file_names,
    }
    history: list = st.session_state["_upload_history"]
    # Avoid duplicate consecutive entries (same file list)
    if not history or history[-1]["names"] != file_names:
        history.append(entry)
        if len(history) > MAX_HISTORY:
            history.pop(0)


def render_history_sidebar(labels: dict):
    """Render upload history panel inside the sidebar."""
    history: list = st.session_state.get("_upload_history", [])

    st.markdown(
        f'<div style="margin:22px 0 8px;font-size:11px;font-weight:700;'
        f'letter-spacing:1.2px;text-transform:uppercase;color:var(--theme-primary,#38bdf8);">'
        f'🕑 &nbsp;{html.escape(labels.get("history_title","Session history"))}</div>',
        unsafe_allow_html=True,
    )

    if not history:
        st.markdown(
            f'<div style="padding:10px 14px;border-radius:10px;'
            f'background:rgba(56,189,248,0.05);border:1px solid rgba(56,189,248,0.14);'
            f'color:#475569;font-size:12px;">'
            f'{html.escape(labels.get("history_empty","No previous sessions yet."))}'
            f'</div>',
            unsafe_allow_html=True,
        )
        return

    for i, entry in enumerate(reversed(history)):
        label_sess = labels.get("history_session", "Session")
        st.markdown(
            f'<div style="margin-bottom:8px;padding:10px 14px;border-radius:10px;'
            f'background:rgba(15,31,61,0.70);border:1px solid rgba(56,189,248,0.16);'
            f'font-size:12px;">'
            f'<div style="display:flex;justify-content:space-between;'
            f'color:#38bdf8;font-weight:700;margin-bottom:4px;">'
            f'<span>{html.escape(label_sess)} {len(history)-i}</span>'
            f'<span style="color:#475569;">{html.escape(entry["time"])}</span>'
            f'</div>'
            f'<div style="color:#64748b;">'
            f'📄 {entry["files"]} &nbsp;·&nbsp; '
            f'🗂️ {entry["cabinets"]} &nbsp;·&nbsp; '
            f'✏️ {entry["pieces"]}'
            f'</div>'
            f'<div style="margin-top:4px;color:#334155;font-size:11px;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">'
            + " · ".join(html.escape(n) for n in entry["names"]) +
            f'</div>'
            f'</div>',
            unsafe_allow_html=True,
        )

    if st.button(
        "🗑 " + labels.get("history_clear", "Clear history"),
        use_container_width=True,
        key="clear_history",
    ):
        st.session_state["_upload_history"] = []
        st.rerun()


# ── Chart helpers ──────────────────────────────────────────────────────────────

def render_dimension_scatter(pivot_data: list[dict], piece_names: list[str], labels: dict):
    """Scatter: every cut piece plotted at W×L, bubble size = total qty."""
    points: list[dict] = []
    for row in pivot_data:
        cab = str(row.get("cabinet_id", "?"))
        for pname in piece_names:
            for entry in piece_entries(row, pname):
                w = entry.get("width") or 0
                l = entry.get("length") or 0
                q = entry.get("qty") or 0
                if w and l:
                    points.append({
                        "w": float(w), "l": float(l), "qty": int(q),
                        "part": pname, "cab": cab,
                        "label": f"{fmt_value(w)} × {fmt_value(l)}",
                    })
    if not points:
        return

    df = pd.DataFrame(points)
    agg = (
        df.groupby(["w", "l", "part", "label"])
        .agg(qty=("qty", "sum"), cabs=("cab", lambda x: ", ".join(sorted(set(x)))))
        .reset_index()
    )

    fig = go.Figure(go.Scatter(
        x=agg["w"], y=agg["l"],
        mode="markers",
        marker=dict(
            size=(agg["qty"].clip(upper=40) * 2 + 6).tolist(),
            color=agg["qty"].tolist(),
            colorscale=[[0, "#1e3a5f"], [0.5, "#38bdf8"], [1, "#818cf8"]],
            showscale=True,
            colorbar=dict(
                title=dict(text=labels.get("dim_chart_qty", "Qty"),
                           font=dict(color="#64748b", size=11)),
                tickfont=dict(color="#64748b", size=10),
                thickness=10, len=0.7,
            ),
            line=dict(width=0), opacity=0.82,
        ),
        text=agg.apply(
            lambda r: (
                f"<b>{r['part']}</b><br>{r['label']}<br>"
                f"{labels.get('dim_chart_qty','Qty')}: {r['qty']}<br>"
                f"{labels.get('cabinet','Cab')}: {r['cabs']}"
            ), axis=1,
        ),
        hovertemplate="%{text}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text=labels.get("dim_chart_title", "Dimension spread (W × L)"),
                   font=dict(color="#e2e8f0", size=14, family="Inter"), x=0),
        xaxis=dict(
            title=dict(text=labels.get("dim_chart_w", "Width"),
                       font=dict(color="#64748b", size=11)),
            tickfont=dict(color="#64748b", size=10),
            gridcolor="rgba(255,255,255,0.05)", zerolinecolor="rgba(255,255,255,0.08)",
        ),
        yaxis=dict(
            title=dict(text=labels.get("dim_chart_l", "Length"),
                       font=dict(color="#64748b", size=11)),
            tickfont=dict(color="#64748b", size=10),
            gridcolor="rgba(255,255,255,0.05)", zerolinecolor="rgba(255,255,255,0.08)",
        ),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=10, t=40, b=10),
        height=360, font=dict(family="Inter"),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False},
                    key="chart_dimension_scatter")


def render_warnings(pivot_data: list[dict], piece_names: list[str], labels: dict):
    """Auto-detect and display data quality warnings."""
    HIGH_QTY = 20
    warns: list[tuple[str, str, str]] = []

    for row in pivot_data:
        cab = str(row.get("cabinet_id", "?"))
        has_any = False
        for pname in piece_names:
            entries = piece_entries(row, pname)
            if not entries:
                continue
            has_any = True
            for e in entries:
                qty = e.get("qty", 0) or 0
                w   = e.get("width") or 0
                l   = e.get("length") or 0
                if qty > HIGH_QTY:
                    warns.append(("⚠️", cab,
                        f"{labels.get('warn_high_qty','High qty')}: {pname} × {qty}"))
                if qty and (not w or not l):
                    warns.append(("🔴", cab,
                        f"{labels.get('warn_missing_dim','Missing dims')}: {pname}"))
        if not has_any:
            warns.append(("🟡", cab,
                labels.get("warn_empty_cabinet", "Empty cabinet")))

    if not warns:
        st.markdown(
            f'<div style="padding:10px 14px;border-radius:8px;'
            f'background:rgba(52,211,153,0.07);border:1px solid rgba(52,211,153,0.20);'
            f'color:#34d399;font-size:12px;font-weight:500;">'
            f'✓ &nbsp;{html.escape(labels.get("no_warnings","No issues detected"))}'
            f'</div>',
            unsafe_allow_html=True,
        )
        return

    rows_html = ""
    for level, cab, msg in warns[:14]:
        rows_html += (
            f'<div style="display:flex;align-items:flex-start;gap:8px;padding:5px 0;'
            f'border-bottom:1px solid rgba(255,255,255,0.04);">'
            f'<span style="font-size:12px;flex-shrink:0;">{level}</span>'
            f'<div style="font-size:11px;">'
            f'<span style="color:#94a3b8;font-weight:600;">{html.escape(cab)}</span>'
            f' — <span style="color:#64748b;">{html.escape(msg)}</span>'
            f'</div></div>'
        )
    if len(warns) > 14:
        rows_html += (
            f'<div style="font-size:10px;color:#374151;margin-top:4px;">'
            f'+ {len(warns)-14} more…</div>'
        )
    st.markdown(
        f'<div style="padding:12px 14px;border-radius:10px;'
        f'background:#111827;border:1px solid rgba(255,255,255,0.07);">'
        + rows_html + '</div>',
        unsafe_allow_html=True,
    )


def render_completeness_donut(total: int, filled: int, labels: dict, color: str = "#38bdf8"):
    """Compact donut showing fill rate."""
    empty = max(0, total - filled)
    pct   = round(filled / total * 100) if total else 0
    fig = go.Figure(go.Pie(
        values=[filled, empty],
        labels=[labels.get("donut_filled", "Filled"), labels.get("donut_empty", "Empty")],
        hole=0.72,
        marker=dict(colors=[color, "rgba(255,255,255,0.05)"], line=dict(width=0)),
        textinfo="none",
        hovertemplate="%{label}: %{value}<extra></extra>",
        rotation=90,
    ))
    fig.add_annotation(
        text=f"<b>{pct}%</b>",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=18, color=color, family="Inter"),
    )
    fig.update_layout(
        showlegend=False,
        margin=dict(l=0, r=0, t=0, b=0),
        height=110,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False},
                    key="chart_completeness_donut")


# ── New feature helpers ────────────────────────────────────────────────────────

def render_part_summary(pivot_data: list[dict], piece_names: list[str], labels: dict):
    """Table: total qty + cabinet count + min/max dimensions per part type."""
    rows = []
    for pname in piece_names:
        total_qty = 0
        cab_count = 0
        dims = []
        for row in pivot_data:
            ents = piece_entries(row, pname)
            if not ents:
                continue
            cab_count += 1
            for e in ents:
                q = e.get("qty") or 0
                w = e.get("width") or 0
                l = e.get("length") or 0
                total_qty += q
                if w and l:
                    dims.append((float(w), float(l)))
        if total_qty == 0:
            continue
        if dims:
            min_w  = min(d[0] for d in dims)
            max_w  = max(d[0] for d in dims)
            min_l  = min(d[1] for d in dims)
            max_l  = max(d[1] for d in dims)
            min_dim = f"{fmt_value(min_w)} × {fmt_value(min_l)}"
            max_dim = f"{fmt_value(max_w)} × {fmt_value(max_l)}"
        else:
            min_dim = max_dim = "—"
        rows.append({
            labels.get("part_summary_part",    "Part"):       pname,
            labels.get("part_summary_total_qty","Total qty"): total_qty,
            labels.get("part_summary_cabinets", "Cabinets"):  cab_count,
            labels.get("part_summary_min_dim",  "Min dim"):   min_dim,
            labels.get("part_summary_max_dim",  "Max dim"):   max_dim,
        })
    if not rows:
        return
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_material_estimate(pivot_data: list[dict], piece_names: list[str], labels: dict):
    """Per-part material area estimate (total sq-in and cm²)."""
    rows = []
    for pname in piece_names:
        total_qty  = 0
        total_area = 0.0   # in square inches
        for row in pivot_data:
            for e in piece_entries(row, pname):
                q = e.get("qty") or 0
                w = float(e.get("width")  or 0)
                l = float(e.get("length") or 0)
                total_qty  += q
                total_area += q * w * l
        if total_qty == 0:
            continue
        sqcm = round(total_area * 6.4516, 1)   # 1 in² = 6.4516 cm²
        rows.append({
            labels.get("material_part",      "Part"):       pname,
            labels.get("material_total_qty", "Total pcs"): total_qty,
            labels.get("material_sqft",      "Total sq in"): round(total_area, 1),
            labels.get("material_sqm",       "Total cm²"):  sqcm,
        })
    if not rows:
        return
    df = pd.DataFrame(rows)
    # Highlight the top row by area
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_duplicate_detector(all_rows: list[dict], labels: dict):
    """Show cabinet IDs that appear in more than one source file."""
    from collections import defaultdict
    cab_files: dict[str, set] = defaultdict(set)
    for r in all_rows:
        cab  = str(r.get("cabinet_id", "?"))
        src  = r.get("source_file", "")
        if src:
            cab_files[cab].add(src)

    dupes = {cab: files for cab, files in cab_files.items() if len(files) > 1}
    if not dupes:
        st.markdown(
            f'<div style="padding:10px 14px;border-radius:8px;'
            f'background:rgba(52,211,153,0.07);border:1px solid rgba(52,211,153,0.20);'
            f'color:#34d399;font-size:12px;font-weight:500;">'
            f'✓ &nbsp;{html.escape(labels.get("duplicates_none","No duplicate cabinet IDs"))}'
            f'</div>',
            unsafe_allow_html=True,
        )
        return

    rows_html = ""
    for cab, files in sorted(dupes.items()):
        file_list = " · ".join(html.escape(Path(f).name) for f in sorted(files))
        rows_html += (
            f'<div style="display:flex;align-items:flex-start;gap:10px;'
            f'padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.05);">'
            f'<span style="color:#f87171;font-weight:700;font-size:12px;min-width:50px;">'
            f'{html.escape(cab)}</span>'
            f'<span style="color:#64748b;font-size:11px;">'
            f'{html.escape(labels.get("duplicates_found","In multiple files"))}: {file_list}'
            f'</span></div>'
        )
    st.markdown(
        f'<div style="padding:12px 14px;border-radius:10px;'
        f'background:#111827;border:1px solid rgba(248,113,113,0.20);">'
        + rows_html + '</div>',
        unsafe_allow_html=True,
    )


def render_cabinet_drilldown(
    pivot_data: list[dict],
    piece_names: list[str],
    labels:      dict,
):
    """Interactive per-cabinet detail view."""
    cab_ids = [str(row.get("cabinet_id", "?")) for row in pivot_data]
    if not cab_ids:
        return

    chosen = st.selectbox(
        labels.get("drilldown_select", "Select a cabinet"),
        options=["—"] + cab_ids,
        key="drilldown_select",
        label_visibility="collapsed",
    )
    if chosen == "—":
        st.caption(labels.get("drilldown_none", "Select a cabinet above to inspect it."))
        return

    row = next((r for r in pivot_data if str(r.get("cabinet_id")) == chosen), None)
    if not row:
        return

    # Build a detail card per part that has entries
    filled_parts = [p for p in piece_names if piece_entries(row, p)]
    if not filled_parts:
        st.caption("No parts for this cabinet.")
        return

    cols = st.columns(min(len(filled_parts), 4))
    for i, pname in enumerate(filled_parts):
        ents = piece_entries(row, pname)
        total = sum(e.get("qty") or 0 for e in ents)
        dim   = dimension_text(ents[0]) if ents else "—"
        color = entry_color(ents)
        dot_bg = color if color != "transparent" else "#334155"
        with cols[i % len(cols)]:
            st.markdown(
                f'<div style="padding:14px 12px;border-radius:10px;'
                f'background:#111827;border:1px solid rgba(255,255,255,0.07);'
                f'text-align:center;margin-bottom:8px;">'
                f'<div style="width:10px;height:10px;border-radius:50%;'
                f'background:{dot_bg};margin:0 auto 6px;"></div>'
                f'<div style="font-size:11px;font-weight:700;color:#e2e8f0;'
                f'margin-bottom:2px;">{html.escape(pname)}</div>'
                f'<div style="font-size:20px;font-weight:800;color:#38bdf8;">{total}</div>'
                f'<div style="font-size:10px;color:#475569;margin-top:3px;">'
                f'{html.escape(dim)}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )


def render_pieces_chart(pivot_data: list[dict], piece_names: list[str], labels: dict):
    """Horizontal bar chart: pieces per cabinet, using Plotly."""
    cabinet_ids = [str(row.get("cabinet_id", "?")) for row in pivot_data]
    counts = [
        sum(1 for name in piece_names if piece_entries(row, name))
        for row in pivot_data
    ]
    max_possible = len(piece_names) or 1

    fig = go.Figure()

    # Background reference bar (max capacity)
    fig.add_trace(go.Bar(
        y=cabinet_ids,
        x=[max_possible] * len(cabinet_ids),
        orientation="h",
        marker=dict(color="rgba(56,189,248,0.08)", line=dict(width=0)),
        showlegend=False,
        hoverinfo="skip",
    ))

    # Actual count bar
    bar_colors = [
        f"rgba(56,189,248,{0.55 + 0.40 * (c / max_possible):.2f})"
        for c in counts
    ]
    fig.add_trace(go.Bar(
        y=cabinet_ids,
        x=counts,
        orientation="h",
        marker=dict(
            color=bar_colors,
            line=dict(width=0),
        ),
        text=[f"{c}/{max_possible}" for c in counts],
        textposition="inside",
        textfont=dict(color="#0f172a", size=11, family="Inter"),
        hovertemplate=(
            f"<b>%{{y}}</b><br>"
            f"{html.escape(labels.get('chart_pieces','Pieces'))}: %{{x}}/{max_possible}<extra></extra>"
        ),
        showlegend=False,
    ))

    fig.update_layout(
        barmode="overlay",
        title=dict(
            text=labels.get("chart_title", "Pieces per cabinet"),
            font=dict(color="#e2e8f0", size=15, family="Inter"),
            x=0,
        ),
        xaxis=dict(
            title=labels.get("chart_pieces", "Piece groups"),
            range=[0, max_possible * 1.05],
            tickfont=dict(color="#64748b", size=11),
            gridcolor="rgba(255,255,255,0.06)",
            zerolinecolor="rgba(255,255,255,0.10)",
            title_font=dict(color="#64748b", size=11),
        ),
        yaxis=dict(
            title=labels.get("chart_cabinet", "Cabinet"),
            tickfont=dict(color="#94a3b8", size=11),
            title_font=dict(color="#64748b", size=11),
            autorange="reversed",
        ),
        paper_bgcolor="rgba(10,16,30,0.0)",
        plot_bgcolor="rgba(10,16,30,0.0)",
        margin=dict(l=10, r=10, t=40, b=10),
        height=max(220, min(40 * len(cabinet_ids) + 80, 520)),
        font=dict(family="Inter"),
    )

    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False},
                    key="chart_pieces_per_cabinet")


def main():
    st.set_page_config(
        page_title="Gestionnaire de Debit",
        page_icon="🪵",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    theme = get_theme()
    inject_global_style(theme)
    inject_component_style()
    st.markdown(theme_css(theme), unsafe_allow_html=True)

    with st.sidebar:
        language = render_language_buttons()

    labels = TRANSLATIONS[language]

    with st.sidebar:
        theme = render_theme_editor(labels)
        st.markdown(theme_css(theme), unsafe_allow_html=True)
        render_history_sidebar(labels)

    # ── Hero banner ────────────────────────────────────────────────────
    # Use session_state so the badge can reflect the previous upload count
    # without referencing the uploader widget before it exists.
    prev_file_count = st.session_state.get("_file_count", 0)
    badge_text = (
        f"{prev_file_count} {labels.get('files_loaded', 'file(s)')} &nbsp;✓"
        if prev_file_count > 0
        else "Woodshop Tool"
    )
    st.markdown(
        f"""
        <div class="app-hero">
            <div class="hero-badge">
                <span class="hero-badge-dot"></span>
                {badge_text}
            </div>
            <h1 class="app-title">{html.escape(labels["app_title"])}</h1>
            <p class="app-subtitle">{html.escape(labels["app_subtitle"])}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    default_piece_names = load_piece_names()

    with st.sidebar:
        st.markdown(
            f'<div style="margin:18px 0 8px;font-size:11px;font-weight:700;'
            f'letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;">'
            f'⚙ &nbsp;{html.escape(labels["nomenclature"])}</div>',
            unsafe_allow_html=True,
        )
        names_text = st.text_area(
            labels["part_names"],
            value="\n".join(default_piece_names),
            height=340,
            help=labels["part_names_help"],
        )
        piece_names = [line.strip() for line in names_text.splitlines() if line.strip()]

        st.download_button(
            "📥 " + labels["download_nomenclature"],
            data=json.dumps({"piece_names": piece_names}, ensure_ascii=False, indent=2),
            file_name="nomenclature.json",
            mime="application/json",
            use_container_width=True,
        )

        # ── Job reference / notes ──────────────────────────────────────
        st.markdown(
            f'<div style="margin:18px 0 6px;font-size:11px;font-weight:700;'
            f'letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;">'
            f'📝 &nbsp;{html.escape(labels.get("notes_label","Notes"))}</div>',
            unsafe_allow_html=True,
        )
        job_notes = st.text_area(
            labels.get("notes_label", "Notes"),
            placeholder=labels.get("notes_placeholder", "e.g. Client name, project #…"),
            height=80,
            key="job_notes",
            label_visibility="collapsed",
        )

        # ── Excel formatting controls ──────────────────────────────────
        st.markdown(
            f'<div style="margin:18px 0 6px;font-size:11px;font-weight:700;'
            f'letter-spacing:1.2px;text-transform:uppercase;color:#94a3b8;">'
            f'📐 &nbsp;{html.escape(labels.get("excel_export_settings","Excel export settings"))}</div>',
            unsafe_allow_html=True,
        )

        # Load saved defaults once per session
        if "_user_settings_loaded" not in st.session_state:
            saved = load_user_settings()
            for k, v in saved.items():
                st.session_state.setdefault(k, v)
            st.session_state["_user_settings_loaded"] = True

        excel_font_size = st.slider(
            labels.get("excel_font_size", "Font size"),
            min_value=8, max_value=20,
            key="excel_font_size",
        )
        excel_col_scale = st.slider(
            labels.get("excel_col_scale", "Column width scale"),
            min_value=0.4, max_value=1.5, step=0.05,
            key="excel_col_scale",
            help=labels.get("excel_col_scale_help",
                            "Scale all dimension columns. 0.75 = compact, 1.0 = standard, 1.2 = wide."),
        )
        excel_dim_width = st.slider(
            labels.get("excel_dim_width", "Min dimension column width"),
            min_value=8, max_value=40,
            key="excel_dim_width",
        )
        excel_cab_width = st.slider(
            labels.get("excel_cab_width", "Cabinet column width"),
            min_value=4, max_value=20,
            key="excel_cab_width",
        )

        # Save to disk whenever values change
        current_settings = {
            "excel_font_size": st.session_state.excel_font_size,
            "excel_col_scale": st.session_state.excel_col_scale,
            "excel_dim_width": st.session_state.excel_dim_width,
            "excel_cab_width": st.session_state.excel_cab_width,
        }
        save_user_settings(current_settings)

    uploaded_files = st.file_uploader(
        "📂 " + labels["file_uploader"],
        type=["csv"],
        accept_multiple_files=True,
    )

    # Keep count for the hero badge on the next render pass
    st.session_state["_file_count"] = len(uploaded_files) if uploaded_files else 0

    if not uploaded_files:
        st.markdown(
            f'<div class="empty-state">'
            f'<div style="font-size:40px;margin-bottom:14px;">🪵</div>'
            f'<div style="font-size:16px;font-weight:600;color:#e2e8f0;margin-bottom:6px;">'
            f'{html.escape(labels["empty_upload"])}</div>'
            f'<div style="font-size:13px;color:#4b5563;">'
            f'{html.escape(labels.get("empty_upload_hint","Drag & drop or click Browse files above."))}'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        return

    # ── Per-file color pickers + prefix override ──────────────────────
    custom_colors: dict[str, str] = {}
    custom_prefixes: dict[str, str] = {}
    if len(uploaded_files) > 0:
        st.markdown(
            f'<div style="font-size:11px;font-weight:600;color:#475569;'
            f'letter-spacing:0.8px;text-transform:uppercase;margin:8px 0 6px;">'
            f'🎨 &nbsp;{labels.get("file_colors","File highlight colors")}</div>',
            unsafe_allow_html=True,
        )
        color_changed = False
        for idx, uf in enumerate(uploaded_files):
            # Default: use saved persistent color, then fall back to FILE_COLORS
            saved_color_key = f"file_color_{idx}"
            saved_default = (
                st.session_state.get(f"_saved_{saved_color_key}")
                or load_user_settings().get(saved_color_key)
                or FILE_COLORS[idx % len(FILE_COLORS)]
            )
            key_color  = f"_file_color_{uf.name}"
            key_prefix = f"_file_prefix_{uf.name}"

            # Pre-fill with saved default on first appearance
            if key_color not in st.session_state:
                st.session_state[key_color] = saved_default

            col_a, col_b = st.columns([1, 2])
            with col_a:
                picked = st.color_picker(
                    Path(uf.name).stem[:20],
                    key=key_color,
                )
                custom_colors[uf.name] = picked
                # Detect if user changed from saved default
                if picked != saved_default:
                    color_changed = True
                    st.session_state[f"_saved_{saved_color_key}"] = picked
            with col_b:
                prefix_val = st.text_input(
                    labels.get("file_prefix_label", "Cabinet prefix (optional)"),
                    value=st.session_state.get(key_prefix, ""),
                    placeholder="e.g. R5:",
                    key=key_prefix,
                    label_visibility="collapsed",
                    help=labels.get("file_prefix_help",
                                    "Add a prefix to all bare N/alpha cabinet IDs in this file, e.g. 'R5:'"),
                )
                custom_prefixes[uf.name] = prefix_val.strip()

        # Persist updated colors to user_settings.json
        if color_changed:
            current = load_user_settings()
            for idx, uf in enumerate(uploaded_files):
                key_color = f"_file_color_{uf.name}"
                if key_color in st.session_state:
                    current[f"file_color_{idx}"] = st.session_state[key_color]
            save_user_settings(current)

    # ── Parse with progress feedback ──────────────────────────────────
    parse_placeholder = st.empty()
    parse_placeholder.markdown(
        '<div class="success-message">⏳ &nbsp;'
        + html.escape(labels.get("parsing", "Parsing files…"))
        + '<div class="progress-bar-container"><div class="progress-bar" style="width:100%;"></div></div>'
        + '</div>',
        unsafe_allow_html=True,
    )

    all_rows, imported_files = parse_uploaded_files(uploaded_files, custom_colors, custom_prefixes)
    pivot_data = transform(all_rows, piece_names)

    # ── Success banner ────────────────────────────────────────────────
    total_cabinets = len(pivot_data)
    total_pieces = sum(
        1
        for row in pivot_data
        for name in piece_names
        if piece_entries(row, name)
    )
    files_word = labels.get("files_loaded", "file(s) loaded")
    cabinets_word = labels.get("cabinets_word", "cabinet(s)")
    pieces_word = labels.get("pieces_word", "piece group(s)")

    parse_placeholder.markdown(
        f'<div class="success-message">'
        f'✓ &nbsp;{len(imported_files)} {html.escape(files_word)} &nbsp;·&nbsp; '
        f'<strong>{total_cabinets}</strong> {html.escape(cabinets_word)} &nbsp;·&nbsp; '
        f'<strong>{total_pieces}</strong> {html.escape(pieces_word)}'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Unknown piece name mapping ────────────────────────────────────
    known_names = set(piece_names)
    # Collect raw piece names from parsed rows that aren't in any column
    unknown_raw = sorted({
        r["piece_name"] for r in all_rows
        if r["piece_name"] not in known_names
        and r["piece_name"] not in (piece_aliases := {})  # placeholder
    })
    # Build using the actual transformer aliases
    from core.transformer import _ALIASES
    unknown_raw = sorted({
        r["piece_name"] for r in all_rows
        if r["piece_name"] not in known_names
        and r["piece_name"] not in _ALIASES
    })

    if unknown_raw:
        st.markdown(
            f'<div style="padding:10px 14px;border-radius:8px;margin-bottom:8px;'
            f'background:rgba(251,191,36,0.08);border:1px solid rgba(251,191,36,0.30);'
            f'color:#fbbf24;font-size:12px;font-weight:600;">'
            f'⚠️ &nbsp;{len(unknown_raw)} '
            f'{html.escape(labels.get("unknown_pieces","unrecognized piece name(s) — assign them to a column below"))}'
            f'</div>',
            unsafe_allow_html=True,
        )
        col_options = ["— " + labels.get("ignore","ignore") + " —"] + list(piece_names)
        mapping_changed = False
        for uname in unknown_raw:
            key = f"_piece_map_{uname}"
            current = st.session_state.get(key, col_options[0])
            # Bright visible label above the selectbox
            st.markdown(
                f'<div style="margin:10px 0 2px;font-size:13px;font-weight:700;'
                f'color:#f1f5f9;font-family:Inter,sans-serif;">'
                f'⚠️ &nbsp;<span style="color:#fbbf24;">{html.escape(uname)}</span>'
                f' → ?</div>',
                unsafe_allow_html=True,
            )
            chosen = st.selectbox(
                uname,
                col_options,
                index=col_options.index(current) if current in col_options else 0,
                key=key,
                label_visibility="collapsed",
            )
            if chosen != col_options[0]:
                # Apply mapping to all_rows in memory
                for r in all_rows:
                    if r["piece_name"] == uname:
                        r["piece_name"] = chosen
                mapping_changed = True

        if mapping_changed:
            # Re-transform with updated mappings
            pivot_data = transform(all_rows, piece_names)
            total_cabinets = len(pivot_data)
            total_pieces = sum(
                1 for row in pivot_data
                for name in piece_names
                if piece_entries(row, name)
            )

    # ── File header banners (filename + color) ───────────────────────
    st.markdown(_file_header_banners(imported_files), unsafe_allow_html=True)

    # ── Summary stat cards with progress bars + hover tooltips ──────────
    max_pieces = len(piece_names) * total_cabinets if total_cabinets else 1

    # Per-cabinet piece counts (used in sparklines)
    cab_counts = [
        sum(1 for name in piece_names if piece_entries(row, name))
        for row in pivot_data
    ]
    cab_ids = [str(row.get("cabinet_id", "?")) for row in pivot_data]

    # Most/least filled cabinet
    if cab_counts:
        max_cab_idx = cab_counts.index(max(cab_counts))
        min_cab_idx = cab_counts.index(min(cab_counts))
        most_filled  = f"{cab_ids[max_cab_idx]} ({cab_counts[max_cab_idx]})"
        least_filled = f"{cab_ids[min_cab_idx]} ({cab_counts[min_cab_idx]})"
        avg_fill = round(sum(cab_counts) / len(cab_counts), 1)
    else:
        most_filled = least_filled = "—"
        avg_fill = 0

    # History deltas
    history: list = st.session_state.get("_upload_history", [])
    prev = history[-2] if len(history) >= 2 else None

    def _delta(current: int, key: str) -> str:
        if prev is None:
            return ""
        diff = current - prev.get(key, current)
        if diff > 0:
            return f'<span style="color:#34d399;font-size:10px;">▲ +{diff}</span>'
        if diff < 0:
            return f'<span style="color:#f87171;font-size:10px;">▼ {diff}</span>'
        return '<span style="color:#475569;font-size:10px;">— same</span>'

    # Inline SVG bar sparkline for cabinet fill distribution
    def _sparkline_svg(values: list[int], color: str, width: int = 120, height: int = 28) -> str:
        if not values:
            return ""
        mx = max(values) or 1
        bar_w = max(2, (width - len(values)) // len(values))
        bars = []
        for i, v in enumerate(values):
            h = max(2, round(v / mx * height))
            x = i * (bar_w + 1)
            y = height - h
            bars.append(
                f'<rect x="{x}" y="{y}" width="{bar_w}" height="{h}" '
                f'rx="1" fill="{color}" opacity="0.85"/>'
            )
        return (
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" '
            f'style="display:block;margin:6px auto 0;">'
            + "".join(bars) +
            f'</svg>'
        )

    # File list rows for the "files" card tooltip
    file_rows = "".join(
        f'<div style="display:flex;align-items:center;gap:6px;padding:3px 0;'
        f'border-bottom:1px solid rgba(255,255,255,0.05);">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{c};'
        f'flex-shrink:0;display:inline-block;"></span>'
        f'<span style="color:#cbd5e1;font-size:11px;white-space:nowrap;'
        f'overflow:hidden;text-overflow:ellipsis;max-width:160px;">'
        f'{html.escape(n)}</span>'
        f'</div>'
        for n, c in imported_files
    )

    # Part coverage rows for "part types" tooltip
    part_rows_html = ""
    for pname in piece_names[:10]:   # cap at 10 to keep tooltip compact
        filled = sum(1 for row in pivot_data if piece_entries(row, pname))
        ppct   = round(filled / total_cabinets * 100) if total_cabinets else 0
        part_rows_html += (
            f'<div style="margin-bottom:5px;">'
            f'<div style="display:flex;justify-content:space-between;'
            f'font-size:10px;color:#94a3b8;margin-bottom:2px;">'
            f'<span>{html.escape(pname)}</span><span>{filled}/{total_cabinets}</span>'
            f'</div>'
            f'<div style="width:100%;height:3px;background:rgba(255,255,255,0.07);border-radius:999px;">'
            f'<div style="width:{ppct}%;height:100%;background:#34d399;border-radius:999px;"></div>'
            f'</div>'
            f'</div>'
        )
    if len(piece_names) > 10:
        part_rows_html += f'<div style="color:#475569;font-size:10px;margin-top:4px;">+ {len(piece_names)-10} more…</div>'

    sparkline = _sparkline_svg(cab_counts, "#38bdf8")

    coverage_lbl = labels.get("coverage_label", "Coverage")

    # Build each card: (icon, value, label, pct, color, tooltip_html)
    stat_data = [
        (
            "📄", str(len(imported_files)),
            labels.get("files_loaded", "Files"),
            100, "#38bdf8",
            f'<div style="font-weight:700;color:#38bdf8;font-size:12px;margin-bottom:8px;">'
            f'📄 {html.escape(labels.get("files_loaded","Files"))}</div>'
            + file_rows
            + f'<div style="margin-top:8px;padding-top:6px;border-top:1px solid rgba(255,255,255,0.07);'
            f'font-size:10px;color:#475569;">'
            f'{_delta(len(imported_files),"files")}'
            f'</div>',
        ),
        (
            "🗂️", str(total_cabinets),
            labels.get("cabinets_word", "Cabinets"),
            min(100, round(total_cabinets / max(total_cabinets, 1) * 100)), "#818cf8",
            f'<div style="font-weight:700;color:#818cf8;font-size:12px;margin-bottom:8px;">'
            f'🗂️ {html.escape(labels.get("cabinets_word","Cabinets"))}</div>'
            f'<div style="font-size:11px;color:#94a3b8;line-height:1.8;">'
            f'▲ Most filled: <strong style="color:#e2e8f0;">{html.escape(most_filled)}</strong><br>'
            f'▽ Least filled: <strong style="color:#e2e8f0;">{html.escape(least_filled)}</strong><br>'
            f'⌀ Avg: <strong style="color:#e2e8f0;">{avg_fill} / {len(piece_names)}</strong>'
            f'</div>'
            + sparkline
            + f'<div style="margin-top:8px;font-size:10px;color:#475569;">{_delta(total_cabinets,"cabinets")}</div>',
        ),
        (
            "🔩", str(len(piece_names)),
            labels.get("parts_configured", "Part types"),
            100, "#34d399",
            f'<div style="font-weight:700;color:#34d399;font-size:12px;margin-bottom:8px;">'
            f'🔩 {html.escape(labels.get("parts_configured","Part types"))}</div>'
            + part_rows_html,
        ),
        (
            "✏️", str(total_pieces),
            labels.get("pieces_word", "Non-empty cells"),
            min(100, round(total_pieces / max(max_pieces, 1) * 100)), "#f472b6",
            f'<div style="font-weight:700;color:#f472b6;font-size:12px;margin-bottom:8px;">'
            f'✏️ {html.escape(labels.get("pieces_word","Non-empty cells"))}</div>'
            f'<div style="font-size:11px;color:#94a3b8;line-height:1.8;">'
            f'Total possible: <strong style="color:#e2e8f0;">{max_pieces}</strong><br>'
            f'Filled: <strong style="color:#e2e8f0;">{total_pieces}</strong><br>'
            f'Empty: <strong style="color:#e2e8f0;">{max_pieces - total_pieces}</strong>'
            f'</div>'
            f'<div style="margin-top:10px;">'
            f'  <div style="display:flex;justify-content:space-between;font-size:10px;color:#475569;margin-bottom:3px;">'
            f'    <span>{coverage_lbl}</span>'
            f'    <span>{min(100,round(total_pieces/max(max_pieces,1)*100))}%</span>'
            f'  </div>'
            f'  <div style="width:100%;height:4px;background:rgba(255,255,255,0.07);border-radius:999px;">'
            f'    <div style="width:{min(100,round(total_pieces/max(max_pieces,1)*100))}%;'
            f'    height:100%;background:#f472b6;border-radius:999px;"></div>'
            f'  </div>'
            f'</div>'
            + f'<div style="margin-top:8px;font-size:10px;color:#475569;">{_delta(total_pieces,"pieces")}</div>',
        ),
    ]

    # ── Stat cards (CSS lives in inject_component_style) ─────────────
    stat_cols = st.columns(len(stat_data))
    for col, (icon, value, lbl, pct, color, tooltip_html) in zip(stat_cols, stat_data):
        with col:
            st.markdown(
                f'<div class="stat-card-wrap">'
                f'<div class="stat-card">'
                f'  <div style="font-size:24px;margin-bottom:6px;">{icon}</div>'
                f'  <div style="font-size:28px;font-weight:800;color:{color};line-height:1;">'
                f'    {html.escape(value)}'
                f'  </div>'
                f'  <div style="font-size:11px;font-weight:500;color:#4b5563;'
                f'  letter-spacing:0.6px;text-transform:uppercase;margin-top:5px;">'
                f'    {html.escape(lbl)}'
                f'  </div>'
                f'  <div style="margin-top:14px;">'
                f'    <div style="display:flex;justify-content:space-between;'
                f'    font-size:10px;color:#374151;margin-bottom:4px;">'
                f'      <span>{coverage_lbl}</span><span style="color:{color};">{pct}%</span>'
                f'    </div>'
                f'    <div style="width:100%;height:3px;background:rgba(255,255,255,0.06);'
                f'    border-radius:999px;overflow:hidden;">'
                f'      <div style="height:100%;width:{pct}%;background:{color};'
                f'      border-radius:999px;animation:barFill 700ms ease both;"></div>'
                f'    </div>'
                f'  </div>'
                f'</div>'
                f'<div class="stat-tooltip">'
                f'{tooltip_html}'
                f'</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

    # ── Record session history ────────────────────────────────────────
    record_session(
        files=len(imported_files),
        cabinets=total_cabinets,
        pieces=total_pieces,
        file_names=[n for n, _ in imported_files],
    )

    # ── Analysis section: donut + warnings + dimension scatter ────────
    st.markdown(_section_header("🔬", labels.get("analysis_title", "Analysis")),
                unsafe_allow_html=True)

    tab_chart, tab_dim, tab_warn = st.tabs([
        "📊 " + labels.get("chart_title", "Pieces / cabinet"),
        "📐 " + labels.get("dim_chart_title", "Dimension spread"),
        "⚠️ " + labels.get("warnings_title", "Warnings"),
    ])

    with tab_chart:
        col_donut, col_bar = st.columns([1, 3])
        with col_donut:
            st.markdown(
                f'<div style="font-size:11px;font-weight:600;color:#475569;'
                f'text-transform:uppercase;letter-spacing:0.6px;margin-bottom:4px;">'
                f'{html.escape(labels.get("completeness","Completeness"))}</div>',
                unsafe_allow_html=True,
            )
            render_completeness_donut(max_pieces, total_pieces, labels, "#38bdf8")
        with col_bar:
            render_pieces_chart(pivot_data, piece_names, labels)

    with tab_dim:
        render_dimension_scatter(pivot_data, piece_names, labels)

    with tab_warn:
        render_warnings(pivot_data, piece_names, labels)

    # ── Extra analysis tabs: part summary · material · duplicates ──────
    tab_parts, tab_mat, tab_dupe, tab_drill = st.tabs([
        "🔩 " + labels.get("part_summary_title", "Part summary"),
        "📦 " + labels.get("material_title", "Material estimate"),
        "🔁 " + labels.get("duplicates_title", "Duplicates"),
        "🔎 " + labels.get("drilldown_title", "Cabinet detail"),
    ])

    with tab_parts:
        render_part_summary(pivot_data, piece_names, labels)

    with tab_mat:
        render_material_estimate(pivot_data, piece_names, labels)

    with tab_dupe:
        render_duplicate_detector(all_rows, labels)

    with tab_drill:
        render_cabinet_drilldown(pivot_data, piece_names, labels)

    # ── Quick search + Loaded files legend ────────────────────────────
    st.markdown(_section_header("📁", labels["loaded_files"]), unsafe_allow_html=True)

    search_col, files_col = st.columns([2, 3])
    with search_col:
        search_term = st.text_input(
            labels.get("search_label", "Quick search"),
            placeholder=labels.get("search_placeholder", "Search cabinet or part…"),
            key="search_term",
            label_visibility="collapsed",
        )
        # Decorate with a subtle search icon hint
        st.markdown(
            '<div style="margin-top:-8px;font-size:10px;color:#374151;">'
            '🔍 filter by cabinet ID or part name</div>',
            unsafe_allow_html=True,
        )
    with files_col:
        legend_cols = st.columns(min(len(imported_files), 4) or 1)
        for idx, (name, color) in enumerate(imported_files):
            with legend_cols[idx % len(legend_cols)]:
                st.markdown(_legend_card(name, color), unsafe_allow_html=True)

    # Apply search filter
    if search_term.strip():
        term = search_term.strip().lower()
        filtered_data = [
            row for row in pivot_data
            if term in str(row.get("cabinet_id", "")).lower()
            or any(
                term in pname.lower()
                and piece_entries(row, pname)
                for pname in piece_names
            )
        ]
        filtered_names = [
            p for p in piece_names
            if term in p.lower()
            or any(
                piece_entries(row, p)
                and term in str(row.get("cabinet_id", "")).lower()
                for row in pivot_data
            )
        ] or piece_names
    else:
        filtered_data  = pivot_data
        filtered_names = piece_names

    # ── Sorted preview ────────────────────────────────────────────────
    st.markdown(_section_header("📋", labels["sorted_preview"]), unsafe_allow_html=True)

    sort_col, badge_col = st.columns([2, 3])
    with sort_col:
        sort_key = st.selectbox(
            labels.get("sort_label", "Sort by"),
            options=["cabinet", "completeness", "alpha"],
            format_func=lambda k: {
                "cabinet":      labels.get("sort_cabinet",      "Cabinet ID"),
                "completeness": labels.get("sort_completeness", "Completeness"),
                "alpha":        labels.get("sort_alpha",        "Alphabetical"),
            }[k],
            key="sort_key",
            label_visibility="collapsed",
        )
    with badge_col:
        st.markdown(
            f'<div style="padding-top:6px;font-size:11px;color:#374151;">'
            f'{"🔍 " + str(len(filtered_data)) + " / " + str(total_cabinets) if search_term.strip() else str(total_cabinets) + " " + html.escape(labels.get("cabinets_word","cabinets"))}'
            f'</div>',
            unsafe_allow_html=True,
        )

    # Apply sort
    if sort_key == "completeness":
        filtered_data = sorted(
            filtered_data,
            key=lambda r: sum(1 for p in piece_names if piece_entries(r, p)),
            reverse=True,
        )
    elif sort_key == "alpha":
        filtered_data = sorted(
            filtered_data,
            key=lambda r: str(r.get("cabinet_id", "")).lower(),
        )
    # "cabinet" keeps transformer's natural numeric sort

    # Split filtered names into main / drawer for two-tab preview
    main_filtered, drw_filtered = split_piece_names(filtered_names)

    tab_main, tab_drw = st.tabs([
        "🪵 " + labels.get("sorted_preview", "Sorted preview"),
        "🔲 " + labels.get("drawer_preview", "Drawer cut list"),
    ])
    with tab_main:
        st.markdown(preview_html(filtered_data, main_filtered, labels),
                    unsafe_allow_html=True)
    with tab_drw:
        if drw_filtered:
            st.markdown(preview_html(filtered_data, drw_filtered, labels),
                        unsafe_allow_html=True)
        else:
            st.caption(labels.get("no_drawer_parts", "No drawer parts in this selection."))

    # ── Plain data preview ────────────────────────────────────────────
    with st.expander("🔍 " + labels["plain_preview"]):
        st.dataframe(
            preview_dataframe(filtered_data, filtered_names, labels),
            use_container_width=True,
        )

    # ── Export ─────────────────────────────────────────────────────────
    st.markdown(_section_header("💾", labels["export"]), unsafe_allow_html=True)

    # Build a safe filename slug from job notes if provided
    job_notes = st.session_state.get("job_notes", "").strip()
    if job_notes:
        import re as _re
        safe_slug = _re.sub(r'[^\w\-]', '_', job_notes)[:40].strip('_')
        xlsx_name = f"liste_de_debit_{safe_slug}.xlsx"
        csv_name  = f"liste_de_debit_{safe_slug}.csv"
        tsv_name  = f"cut_list_{safe_slug}.tsv"
        st.markdown(
            f'<div style="margin-bottom:10px;padding:8px 12px;border-radius:8px;'
            f'background:rgba(129,140,248,0.07);border:1px solid rgba(129,140,248,0.18);'
            f'color:#818cf8;font-size:12px;">'
            f'📝 &nbsp;<strong>{html.escape(job_notes[:60])}</strong>'
            f'</div>',
            unsafe_allow_html=True,
        )
    else:
        xlsx_name = "liste_de_debit.xlsx"
        csv_name  = "liste_de_debit.csv"
        tsv_name  = "cut_list.tsv"

    # Export tip
    st.markdown(
        f'<div class="export-tip">'
        f'💡 &nbsp;{html.escape(labels.get("export_tip", "Excel preserves colors and formatting. CSV is ideal for further processing."))}'
        f'</div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        # Pass Excel formatting options
        excel_kw = dict(
            font_size     = st.session_state.get("excel_font_size", 14),
            col_scale     = st.session_state.get("excel_col_scale", 0.75),
            dim_col_width = st.session_state.get("excel_dim_width", 18),
            cab_col_width = st.session_state.get("excel_cab_width", 10),
        )
        def _export_excel_custom(data, pnames, path):
            export_excel(data, pnames, path, **excel_kw)

        st.download_button(
            labels["download_excel"],
            data=export_bytes(_export_excel_custom, filtered_data, filtered_names, ".xlsx"),
            file_name=xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Main + drawer sheets in one workbook",
        )
    with col2:
        st.download_button(
            labels["download_csv"],
            data=export_bytes(export_csv, filtered_data, filtered_names, ".csv"),
            file_name=csv_name,
            mime="text/csv",
            use_container_width=True,
            help="Main cabinet parts only",
        )
    with col3:
        drw_csv_name = (
            f"tiroirs_{safe_slug}.csv" if job_notes else "tiroirs.csv"
        )
        st.download_button(
            labels.get("download_csv_drawers", "⬇ Drawer CSV"),
            data=export_bytes(export_csv_drawers, filtered_data, filtered_names, ".csv"),
            file_name=drw_csv_name,
            mime="text/csv",
            use_container_width=True,
            help="Drawer box parts only (Drw*)",
        )
    with col4:
        def _table_as_text() -> str:
            main_n, _ = split_piece_names(filtered_names)
            lines = ["\t".join(
                [labels["cabinet"]] +
                [f"{p} {labels['qty']}\t{p} {labels['dimensions']}" for p in main_n]
            )]
            for row in filtered_data:
                cells = [str(row.get("cabinet_id", ""))]
                for pname in main_n:
                    ents = piece_entries(row, pname)
                    cells.append(", ".join(fmt_value(e.get("qty", "")) for e in ents))
                    cells.append(", ".join(dimension_text(e) for e in ents))
                lines.append("\t".join(cells))
            return "\n".join(lines)

        st.download_button(
            "📋 " + labels.get("copy_table", "Copy table"),
            data=_table_as_text().encode("utf-8"),
            file_name=tsv_name,
            mime="text/tab-separated-values",
            use_container_width=True,
            key="copy_tsv",
            help="Tab-separated — opens directly in Excel / Numbers",
        )

    # ── Footer ─────────────────────────────────────────────────────────
    notes_display = f" &nbsp;·&nbsp; 📝 {html.escape(job_notes[:50])}" if job_notes else ""
    st.markdown(
        f'<div class="app-footer">'
        f'<span>🪵 &nbsp;<strong>Gestionnaire de Débit</strong> &nbsp;·&nbsp; v1.0'
        f' &nbsp;·&nbsp; Built with ♥ for woodshop pros{notes_display}</span>'
        f'<span>CSV → Excel in seconds</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
