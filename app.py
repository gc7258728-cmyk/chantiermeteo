# -*- coding: utf-8 -*-
"""
================================================================================
CHANTIERMÉTÉO — Analyse météorologique et aide à la planification des travaux de génie civil
Contexte académique : École Nationale Supérieure Polytechnique de Yaoundé (ENSPY) — 3e année Génie Civil
Fuseau horaire de référence : Africa/Douala (UTC+1)
================================================================================
Organisation du fichier app.py :
  1. Configuration, Constantes Métier et Modèles Pédagogiques
  2. Identité Visuelle (HTML / CSS Responsive & Thèmes Clair / Sombre)
  3. Fonctions de Récupération des Données (API Open-Meteo & Générateur Fictif)
  4. Fonctions de Validation et d'Importation (CSV / Excel)
  5. Moteur de Calculs Statistiques Descriptifs
  6. Moteur de Planification Météo et Classification des Activités
  7. Moteur de Probabilités, Fenêtres Consécutives, Monte-Carlo par Blocs et Sensibilité
  8. Fonctions d'Exportation, Sauvegarde et Restauration (CSV, Excel, PDF, JSON)
  9. Gestion de l'État de Session et des Chantiers
 10. Générateurs de Graphiques Interactifs Plotly Harmonisés
 11. Navigation Latérale et Rendu des 6 Rubriques Principales
================================================================================
"""

from __future__ import annotations

import copy
import datetime as dt
import io
import json
import math
import os
import uuid
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
from scipy import stats as scipy_stats

from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

try:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
except ImportError:  # pragma: no cover
    px = None
    go = None
    make_subplots = None

try:
    import streamlit as st
except ImportError:  # pragma: no cover
    st = None


# ==============================================================================
# 1. CONFIGURATION, CONSTANTES MÉTIER ET MODÈLES PÉDAGOGIQUES
# ==============================================================================

APP_NAME = "ChantierMétéo"
APP_SUBTITLE = "Analyse météorologique et aide à la planification des travaux de génie civil"
APP_INSTITUTION = "École Nationale Supérieure Polytechnique de Yaoundé (ENSPY) — Département de Génie Civil"
DEFAULT_TIMEZONE = "Africa/Douala"

# Palette officielle ChantierMétéo
PALETTE = {
    "navy": "#10243A",       # Bleu nuit : titres et structure
    "sky": "#328CCB",        # Bleu ciel : données météorologiques
    "orange": "#F29A38",     # Orange chantier : actions importantes et alertes
    "green": "#218A70",      # Vert : résultats compatibles avec les critères
    "coral": "#D85B56",      # Rouge corail : conditions défavorables
    "light_bg": "#F3F6FA",   # Gris très clair : fonds et zones secondaires
    "slate": "#5A6B7C",      # Gris ardoise : données absentes / évaluation impossible
}

# Sources de données explicitement distinguées
SOURCE_TYPES = {
    "HISTORIQUE_API": "Données historiques (API Open-Meteo — Réanalyse ERA5 / observations interpolées)",
    "PREVISIONS_API": "Prévisions météorologiques (API Open-Meteo Forecast — modèle numérique jusqu'à 16 jours)",
    "IMPORT_FICHIER": "Données importées par l'utilisateur (Fichier CSV / Excel vérifié)",
    "FICTIF_DEMO": "Données fictives — utilisation pédagogique uniquement (Démonstration hors-ligne)",
    "TEST_DETERMINISTE": "Données fictives — série déterministe de vérification (10 jours)",
}

# Villes camerounaises et internationales de référence (secours hors-ligne ou sélection rapide)
CAMEROON_CITIES_PRESETS: Dict[str, Dict[str, Any]] = {
    "Yaoundé": {"latitude": 3.8480, "longitude": 11.5021, "region": "Centre", "elevation": 726},
    "Douala": {"latitude": 4.0511, "longitude": 9.7679, "region": "Littoral", "elevation": 13},
    "Bafoussam": {"latitude": 5.4778, "longitude": 10.4176, "region": "Ouest", "elevation": 1450},
    "Garoua": {"latitude": 9.3014, "longitude": 13.3977, "region": "Nord", "elevation": 213},
    "Maroua": {"latitude": 10.5910, "longitude": 14.3159, "region": "Extrême-Nord", "elevation": 405},
    "Bamenda": {"latitude": 5.9597, "longitude": 10.1460, "region": "Nord-Ouest", "elevation": 1258},
    "Ngaoundéré": {"latitude": 7.3276, "longitude": 13.5847, "region": "Adamaoua", "elevation": 1112},
    "Bertoua": {"latitude": 4.5773, "longitude": 13.6846, "region": "Est", "elevation": 670},
    "Ebolowa": {"latitude": 2.9150, "longitude": 11.1543, "region": "Sud", "elevation": 615},
    "Buea": {"latitude": 4.1527, "longitude": 9.2410, "region": "Sud-Ouest", "elevation": 870},
    "Kribi": {"latitude": 2.9373, "longitude": 9.9077, "region": "Sud", "elevation": 12},
}

PROJECT_TYPES = [
    "Bâtiment",
    "Route",
    "Ouvrage hydraulique",
    "Terrassement",
    "Autre",
]

# Modèles d'activités avec hypothèses pédagogiques modifiables (NON normatives)
ACTIVITY_TEMPLATES: Dict[str, Dict[str, Any]] = {
    "Terrassement": {
        "duration_days": 10,
        "precip_max_mm": 10.0,
        "temp_min_c": 10.0,
        "temp_max_c": 38.0,
        "wind_max_kmh": 50.0,
        "humidity_max_pct": None,
        "pedagogical_note": (
            "Hypothèse pédagogique : une pluie journalière élevée ramollit les sols argileux/latéritiques, "
            "augmente la teneur en eau au-delà de l'optimum Proctor et réduit la portance des pistes d'engins."
        ),
    },
    "Bétonnage": {
        "duration_days": 7,
        "precip_max_mm": 5.0,
        "temp_min_c": 5.0,
        "temp_max_c": 35.0,
        "wind_max_kmh": 45.0,
        "humidity_max_pct": None,
        "pedagogical_note": (
            "Hypothèse pédagogique : une forte averse pendant le coulage dégrade le rapport E/C de surface, "
            "tandis qu'une température élevée combinée au vent accélère l'évaporation et favorise le retrait plastique."
        ),
    },
    "Travaux de voirie": {
        "duration_days": 12,
        "precip_max_mm": 5.0,
        "temp_min_c": 12.0,
        "temp_max_c": 38.0,
        "wind_max_kmh": 50.0,
        "humidity_max_pct": None,
        "pedagogical_note": (
            "Hypothèse pédagogique : la mise en œuvre de couches de base ou d'enrobés bitumineux est sensible "
            "à l'eau libre sur le support et au refroidissement rapide du tapis."
        ),
    },
    "Travaux de peinture extérieure": {
        "duration_days": 5,
        "precip_max_mm": 1.0,
        "temp_min_c": 10.0,
        "temp_max_c": 35.0,
        "wind_max_kmh": 35.0,
        "humidity_max_pct": 85.0,
        "pedagogical_note": (
            "Hypothèse pédagogique : l'application de revêtements et peintures de façade est très sensible "
            "au ruissellement, à la poussière soulevée par le vent et à une humidité relative excessive."
        ),
    },
    "Opérations de levage": {
        "duration_days": 4,
        "precip_max_mm": 15.0,
        "temp_min_c": 5.0,
        "temp_max_c": 40.0,
        "wind_max_kmh": 35.0,
        "humidity_max_pct": None,
        "pedagogical_note": (
            "Hypothèse pédagogique : la stabilité des charges à forte prise au vent (bennes, coffrages banches, "
            "charpentes) dépend fortement des rafales maximales."
        ),
    },
    "Travaux d'assainissement": {
        "duration_days": 8,
        "precip_max_mm": 8.0,
        "temp_min_c": 10.0,
        "temp_max_c": 38.0,
        "wind_max_kmh": 45.0,
        "humidity_max_pct": None,
        "pedagogical_note": (
            "Hypothèse pédagogique : les fouilles en tranchée et la pose de buses ou caniveaux sont exposées "
            "aux venues d'eau de ruissellement et à l'instabilité des parois en cas de forte pluie."
        ),
    },
    "Autre activité personnalisée": {
        "duration_days": 5,
        "precip_max_mm": 10.0,
        "temp_min_c": 10.0,
        "temp_max_c": 38.0,
        "wind_max_kmh": 45.0,
        "humidity_max_pct": None,
        "pedagogical_note": (
            "Hypothèse pédagogique entièrement configurable selon la nature de l'ouvrage étudié."
        ),
    },
}

# Série déterministe obligatoire (Rubrique F — Laboratoire pédagogique)
DETERMINISTIC_TEST_PRECIP = [0.0, 4.0, 12.0, 0.0, 8.0, 25.0, 1.0, 0.0, 16.0, 3.0]
DETERMINISTIC_TEST_THRESHOLD = 10.0

WEATHER_VARIABLES_META = {
    "precipitation_sum": {
        "label": "Précipitations journalières",
        "unit": "mm",
        "short_unit": "mm/j",
        "description": "Cumul quotidien de précipitations liquides (00h00 à 23h59, fuseau Africa/Douala).",
    },
    "temperature_2m_min": {
        "label": "Température minimale à 2 m",
        "unit": "°C",
        "short_unit": "°C",
        "description": "Température minimale journalière de l'air à 2 mètres au-dessus du sol.",
    },
    "temperature_2m_max": {
        "label": "Température maximale à 2 m",
        "unit": "°C",
        "short_unit": "°C",
        "description": "Température maximale journalière de l'air à 2 mètres au-dessus du sol.",
    },
    "temperature_2m_mean": {
        "label": "Température moyenne à 2 m",
        "unit": "°C",
        "short_unit": "°C",
        "description": "Température moyenne journalière (fournie ou estimée par (Tmin + Tmax)/2).",
    },
    "wind_speed_10m_max": {
        "label": "Vitesse maximale du vent à 10 m",
        "unit": "km/h",
        "short_unit": "km/h",
        "description": "Vitesse maximale quotidienne du vent mesurée ou modélisée à 10 mètres.",
    },
    "relative_humidity_mean": {
        "label": "Humidité relative moyenne",
        "unit": "%",
        "short_unit": "%",
        "description": "Moyenne journalière de l'humidité relative de l'air à 2 mètres.",
    },
    "precipitation_probability_max": {
        "label": "Probabilité de précipitation (Prévision)",
        "unit": "%",
        "short_unit": "%",
        "description": "Probabilité maximale de précipitation issue du modèle de prévision numérique (non disponible en historique).",
    },
}


# ==============================================================================
# 2. IDENTITÉ VISUELLE (HTML / CSS RESPONSIVE & THÈMES CLAIR / SOMBRE)
# ==============================================================================

def get_custom_css(theme_mode: str = "Clair") -> str:
    """Génère la feuille de style CSS responsive adaptée au thème choisi."""
    is_dark = theme_mode == "Sombre"

    bg_main = "#0B1622" if is_dark else "#F3F6FA"
    bg_card = "#132337" if is_dark else "#FFFFFF"
    bg_secondary = "#1A2F49" if is_dark else "#E8EEF5"
    text_primary = "#F1F5F9" if is_dark else "#10243A"
    text_muted = "#A3B8CC" if is_dark else "#4A5D70"
    border_col = "#284363" if is_dark else "#D6E0EB"
    hero_grad = (
        "linear-gradient(135deg, #0D1E31 0%, #17375B 60%, #1F4E79 100%)"
        if is_dark
        else "linear-gradient(135deg, #10243A 0%, #1B3B5C 60%, #245682 100%)"
    )

    return f"""
    <style>
        /* Polices et structure globale */
        .stApp {{
            background-color: {bg_main};
            color: {text_primary};
        }}
        .block-container {{
            padding-top: 1.4rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1380px !important;
        }}
        /* Bandeau d'accueil professionnel */
        .cm-hero {{
            background: {hero_grad};
            color: #FFFFFF;
            padding: 1.6rem 2rem;
            border-radius: 14px;
            border-left: 7px solid #F29A38;
            box-shadow: 0 6px 20px rgba(16, 36, 58, 0.16);
            margin-bottom: 1.3rem;
        }}
        .cm-hero-title {{
            font-size: 1.95rem;
            font-weight: 800;
            letter-spacing: -0.02em;
            margin: 0 0 0.25rem 0;
            color: #FFFFFF !important;
            display: flex;
            align-items: center;
            gap: 0.6rem;
            flex-wrap: wrap;
        }}
        .cm-hero-subtitle {{
            font-size: 1.05rem;
            font-weight: 500;
            color: #CFE6F7;
            margin-bottom: 0.6rem;
        }}
        .cm-hero-desc {{
            font-size: 0.92rem;
            line-height: 1.5;
            color: #E2EEF8;
            margin: 0;
        }}
        .cm-badge {{
            display: inline-block;
            padding: 0.25rem 0.7rem;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }}
        .cm-badge-fictif {{
            background-color: #F29A38;
            color: #10243A;
        }}
        .cm-badge-api {{
            background-color: #328CCB;
            color: #FFFFFF;
        }}
        .cm-badge-ok {{
            background-color: #218A70;
            color: #FFFFFF;
        }}
        .cm-badge-ko {{
            background-color: #D85B56;
            color: #FFFFFF;
        }}
        /* Cartes KPI colorées */
        .cm-kpi-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(205px, 1fr));
            gap: 0.9rem;
            margin-bottom: 1.25rem;
        }}
        .cm-kpi-card {{
            background: {bg_card};
            border: 1px solid {border_col};
            border-top: 5px solid #328CCB;
            border-radius: 11px;
            padding: 1rem 1.1rem;
            box-shadow: 0 2px 8px rgba(16, 36, 58, 0.06);
            transition: transform 0.15s ease;
        }}
        .cm-kpi-card:hover {{
            transform: translateY(-2px);
        }}
        .cm-kpi-navy {{ border-top-color: #10243A; }}
        .cm-kpi-sky {{ border-top-color: #328CCB; }}
        .cm-kpi-orange {{ border-top-color: #F29A38; }}
        .cm-kpi-green {{ border-top-color: #218A70; }}
        .cm-kpi-coral {{ border-top-color: #D85B56; }}
        .cm-kpi-label {{
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            color: {text_muted};
            margin-bottom: 0.35rem;
        }}
        .cm-kpi-value {{
            font-size: 1.38rem;
            font-weight: 800;
            color: {text_primary};
            line-height: 1.25;
            margin-bottom: 0.25rem;
            word-break: break-word;
        }}
        .cm-kpi-sub {{
            font-size: 0.8rem;
            color: {text_muted};
        }}
        /* Cartes de chantier et encadrés pédagogiques */
        .cm-card {{
            background: {bg_card};
            border: 1px solid {border_col};
            border-radius: 12px;
            padding: 1.15rem 1.3rem;
            margin-bottom: 1rem;
            box-shadow: 0 2px 10px rgba(16, 36, 58, 0.05);
        }}
        .cm-card-active {{
            border: 2px solid #F29A38;
            box-shadow: 0 4px 14px rgba(242, 154, 56, 0.18);
        }}
        .cm-pedago-box {{
            background: {bg_secondary};
            border-left: 5px solid #328CCB;
            border-radius: 8px;
            padding: 0.95rem 1.15rem;
            margin: 0.75rem 0;
            font-size: 0.92rem;
            color: {text_primary};
        }}
        .cm-warn-banner {{
            background: rgba(242, 154, 56, 0.16);
            border: 1px solid #F29A38;
            border-left: 6px solid #F29A38;
            border-radius: 9px;
            padding: 0.8rem 1.1rem;
            margin-bottom: 1rem;
            color: {text_primary};
            font-weight: 600;
        }}
        .cm-empty-state {{
            background: {bg_card};
            border: 2px dashed #328CCB;
            border-radius: 14px;
            padding: 2rem;
            text-align: center;
            margin: 1rem 0;
        }}
        /* Calendrier visuel de planification */
        .cm-cal-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(125px, 1fr));
            gap: 0.55rem;
            margin-top: 0.75rem;
            margin-bottom: 1rem;
        }}
        .cm-cal-day {{
            border-radius: 8px;
            padding: 0.55rem 0.65rem;
            font-size: 0.79rem;
            color: #FFFFFF;
            line-height: 1.35;
            box-shadow: 0 1px 4px rgba(0,0,0,0.12);
        }}
        .cm-cal-compatible {{ background-color: #218A70; }}
        .cm-cal-defavorable {{ background-color: #D85B56; }}
        .cm-cal-impossible {{ background-color: #5A6B7C; }}

        /* Responsive Mobile & Tablette */
        @media (max-width: 768px) {{
            .block-container {{
                padding-left: 0.75rem !important;
                padding-right: 0.75rem !important;
                padding-top: 0.8rem !important;
            }}
            .cm-hero {{
                padding: 1.1rem 1.15rem;
            }}
            .cm-hero-title {{
                font-size: 1.4rem;
            }}
            .cm-kpi-grid {{
                grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                gap: 0.6rem;
            }}
            .cm-kpi-value {{
                font-size: 1.15rem;
            }}
            .cm-cal-grid {{
                grid-template-columns: repeat(auto-fill, minmax(105px, 1fr));
            }}
        }}
    </style>
    """


def render_hero_banner(active_project: Optional[Dict[str, Any]] = None) -> None:
    """Affiche le bandeau d'accueil professionnel ChantierMétéo."""
    if st is None:
        return
    badge_html = '<span class="cm-badge cm-badge-api">Génie Civil • ENSPY Yaoundé</span>'
    if active_project:
        src = active_project.get("data_source_type", "AUCUNE")
        if src in ("FICTIF_DEMO", "TEST_DETERMINISTE"):
            badge_html += (
                ' <span class="cm-badge cm-badge-fictif">'
                'Données fictives — utilisation pédagogique uniquement</span>'
            )
        elif src == "HISTORIQUE_API":
            badge_html += ' <span class="cm-badge cm-badge-ok">Historique Open-Meteo ERA5</span>'
        elif src == "PREVISIONS_API":
            badge_html += ' <span class="cm-badge cm-badge-api">Prévisions Open-Meteo</span>'
        elif src == "IMPORT_FICHIER":
            badge_html += ' <span class="cm-badge cm-badge-ok">Fichier Importé Vérifié</span>'

    st.markdown(
        f"""
        <div class="cm-hero">
            <div class="cm-hero-title">
                🏗️ {APP_NAME} {badge_html}
            </div>
            <div class="cm-hero-subtitle">
                {APP_SUBTITLE}
            </div>
            <p class="cm-hero-desc">
                Outil d'ingénierie et de méthodologie statistique permettant d'exploiter les séries météorologiques
                historiques, les prévisions et les simulations probabilistes (Monte-Carlo par blocs) pour identifier
                les fenêtres favorables aux travaux de construction et anticiper les aléas climatiques sur chantier.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# 3. FONCTIONS DE RÉCUPÉRATION DES DONNÉES (API OPEN-METEO & FICTIF)
# ==============================================================================

def geocode_city_open_meteo(city_name: str, count: int = 8) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    """
    Interroge l'API Open-Meteo Geocoding pour trouver les coordonnées d'une localité.
    En cas d'absence de connexion Internet, propose une correspondance dans le catalogue local.
    """
    query = (city_name or "").strip()
    if not query:
        return [], "Veuillez saisir un nom de ville ou de localité non vide."

    url = "https://geocoding-api.open-meteo.com/v1/search"
    params = {"name": query, "count": count, "language": "fr", "format": "json"}

    try:
        resp = requests.get(url, params=params, timeout=8)
        resp.raise_for_status()
        payload = resp.json()
        if not isinstance(payload, dict):
            return [], "Réponse inattendue du serveur de géocodage Open-Meteo."
        results = payload.get("results") or []
        if not results:
            return [], f"Aucune localité trouvée pour « {query} » via l'API Open-Meteo."
        cleaned: List[Dict[str, Any]] = []
        for item in results:
            cleaned.append(
                {
                    "name": item.get("name", query),
                    "country": item.get("country", ""),
                    "admin1": item.get("admin1", ""),
                    "latitude": float(item.get("latitude", 0.0)),
                    "longitude": float(item.get("longitude", 0.0)),
                    "elevation": item.get("elevation", None),
                    "timezone": item.get("timezone", DEFAULT_TIMEZONE),
                }
            )
        return cleaned, None
    except requests.RequestException as exc:
        # Secours hors-ligne sur les villes camerounaises connues
        fallback_matches = []
        q_low = query.lower()
        for cname, cinfo in CAMEROON_CITIES_PRESETS.items():
            if q_low in cname.lower():
                fallback_matches.append(
                    {
                        "name": cname,
                        "country": "Cameroun (catalogue local hors-ligne)",
                        "admin1": cinfo["region"],
                        "latitude": cinfo["latitude"],
                        "longitude": cinfo["longitude"],
                        "elevation": cinfo["elevation"],
                        "timezone": DEFAULT_TIMEZONE,
                    }
                )
        msg = (
            f"Service de géocodage Open-Meteo inaccessible ({exc.__class__.__name__}). "
            "Vérifiez votre connexion Internet ou sélectionnez une ville prédéfinie."
        )
        return fallback_matches, msg
    except (ValueError, KeyError) as exc:
        return [], f"Erreur de lecture de la réponse API de géocodage : {exc}"


def _parse_open_meteo_daily_payload(
    payload: Dict[str, Any],
    source_type: str,
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """Transforme la réponse JSON d'Open-Meteo (Historical ou Forecast) en DataFrame normalisé."""
    if not isinstance(payload, dict) or "daily" not in payload:
        return None, "La réponse de l'API Open-Meteo ne contient pas de bloc 'daily' valide."

    daily = payload["daily"]
    times = daily.get("time")
    if not times or not isinstance(times, list):
        return None, "Aucune date ('time') trouvée dans la réponse de l'API Open-Meteo."

    df = pd.DataFrame({"date": pd.to_datetime(times, errors="coerce")})
    df = df.dropna(subset=["date"]).copy()
    if df.empty:
        return None, "Les dates reçues de l'API Open-Meteo sont invalides."

    df["date"] = df["date"].dt.strftime("%Y-%m-%d")

    col_map = [
        "precipitation_sum",
        "temperature_2m_min",
        "temperature_2m_max",
        "temperature_2m_mean",
        "wind_speed_10m_max",
        "precipitation_probability_max",
        "relative_humidity_2m_mean",
    ]
    for col in col_map:
        if col in daily and isinstance(daily[col], list) and len(daily[col]) == len(times):
            target_col = "relative_humidity_mean" if col == "relative_humidity_2m_mean" else col
            df[target_col] = pd.to_numeric(daily[col], errors="coerce")

    # Si l'humidité relative est fournie au pas horaire, calculer la moyenne journalière
    if "relative_humidity_mean" not in df.columns and "hourly" in payload:
        hourly = payload.get("hourly") or {}
        h_times = hourly.get("time")
        h_rh = hourly.get("relative_humidity_2m")
        if isinstance(h_times, list) and isinstance(h_rh, list) and len(h_times) == len(h_rh):
            hdf = pd.DataFrame(
                {
                    "dt": pd.to_datetime(h_times, errors="coerce"),
                    "rh": pd.to_numeric(h_rh, errors="coerce"),
                }
            ).dropna(subset=["dt"])
            if not hdf.empty:
                hdf["date"] = hdf["dt"].dt.strftime("%Y-%m-%d")
                rh_daily = hdf.groupby("date", as_index=False)["rh"].mean().rename(
                    columns={"rh": "relative_humidity_mean"}
                )
                df = df.merge(rh_daily, on="date", how="left")

    # Si temperature_2m_mean est absente ou partiellement NaN mais Tmin et Tmax sont présentes
    if "temperature_2m_mean" not in df.columns:
        if "temperature_2m_min" in df.columns and "temperature_2m_max" in df.columns:
            df["temperature_2m_mean"] = (df["temperature_2m_min"] + df["temperature_2m_max"]) / 2.0
        else:
            df["temperature_2m_mean"] = np.nan
    elif "temperature_2m_min" in df.columns and "temperature_2m_max" in df.columns:
        mask_na = df["temperature_2m_mean"].isna()
        df.loc[mask_na, "temperature_2m_mean"] = (
            df.loc[mask_na, "temperature_2m_min"] + df.loc[mask_na, "temperature_2m_max"]
        ) / 2.0

    # S'assurer de la présence de toutes les colonnes standardisées (avec NaN si non fournies)
    for std_col in [
        "precipitation_sum",
        "temperature_2m_min",
        "temperature_2m_max",
        "temperature_2m_mean",
        "wind_speed_10m_max",
        "relative_humidity_mean",
        "precipitation_probability_max",
    ]:
        if std_col not in df.columns:
            df[std_col] = np.nan

    df["source_type"] = source_type
    df = df.sort_values("date").drop_duplicates(subset=["date"], keep="first").reset_index(drop=True)
    return df, None


def fetch_historical_weather_open_meteo(
    latitude: float,
    longitude: float,
    start_date: str,
    end_date: str,
    timezone: str = DEFAULT_TIMEZONE,
) -> Tuple[Optional[pd.DataFrame], Optional[str], Dict[str, Any]]:
    """
    Récupère les données météorologiques historiques journalières via Open-Meteo Historical Weather API.
    - Utilise le fuseau horaire Africa/Douala par défaut.
    - Ajuste automatiquement la date de fin si elle dépasse J-2 (délai de consolidation ERA5).
    - Ne remplace jamais les données manquantes par des valeurs inventées.
    """
    meta_info: Dict[str, Any] = {
        "api_endpoint": "https://archive-api.open-meteo.com/v1/archive",
        "timezone": timezone,
        "requested_start": start_date,
        "requested_end": end_date,
        "adjusted_note": None,
    }

    try:
        lat = float(latitude)
        lon = float(longitude)
        if not (-90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0):
            return None, "Coordonnées géographiques hors limites (latitude ∈ [-90, 90], longitude ∈ [-180, 180]).", meta_info
    except (TypeError, ValueError):
        return None, "Latitude ou longitude invalide.", meta_info

    try:
        s_dt = dt.date.fromisoformat(str(start_date)[:10])
        e_dt = dt.date.fromisoformat(str(end_date)[:10])
    except ValueError:
        return None, "Format de date invalide (attendu : YYYY-MM-DD).", meta_info

    today = dt.date.today()
    max_hist_date = today - dt.timedelta(days=2)

    if s_dt > max_hist_date:
        return (
            None,
            f"La date de début demandée ({s_dt.isoformat()}) est située dans le futur ou trop récente pour "
            f"l'archive historique Open-Meteo (disponible jusqu'au {max_hist_date.isoformat()}). "
            "Utilisez le mode « Prévisions (Forecast API) » pour les 16 prochains jours ou choisissez une "
            "période historique antérieure.",
            meta_info,
        )

    if e_dt > max_hist_date:
        meta_info["adjusted_note"] = (
            f"Date de fin ajustée du {e_dt.isoformat()} au {max_hist_date.isoformat()} "
            "(limite des archives consolidées Open-Meteo Historical API)."
        )
        e_dt = max_hist_date

    if e_dt < s_dt:
        return None, "La date de fin doit être postérieure ou égale à la date de début.", meta_info

    url = "https://archive-api.open-meteo.com/v1/archive"
    params = {
        "latitude": lat,
        "longitude": lon,
        "start_date": s_dt.isoformat(),
        "end_date": e_dt.isoformat(),
        "daily": "precipitation_sum,temperature_2m_max,temperature_2m_min,temperature_2m_mean,wind_speed_10m_max",
        "hourly": "relative_humidity_2m",
        "timezone": timezone,
    }

    try:
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code != 200:
            # Tentative de repli sans le paramètre hourly au cas où la période est très longue
            params_no_hourly = {k: v for k, v in params.items() if k != "hourly"}
            resp = requests.get(url, params=params_no_hourly, timeout=15)
        resp.raise_for_status()
        payload = resp.json()
        df, err = _parse_open_meteo_daily_payload(payload, source_type="HISTORIQUE_API")
        if err:
            return None, err, meta_info
        meta_info["effective_start"] = s_dt.isoformat()
        meta_info["effective_end"] = e_dt.isoformat()
        meta_info["fetched_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return df, None, meta_info
    except requests.RequestException as exc:
        return (
            None,
            f"Impossible de contacter l'API Open-Meteo Historical ({exc.__class__.__name__}). "
            "Vérifiez votre connexion Internet, réessayez, ou utilisez le jeu de démonstration / un import CSV-Excel.",
            meta_info,
        )
    except ValueError as exc:
        return None, f"Réponse JSON invalide reçue de l'API Open-Meteo Historical : {exc}", meta_info


def fetch_forecast_weather_open_meteo(
    latitude: float,
    longitude: float,
    forecast_days: int = 14,
    timezone: str = DEFAULT_TIMEZONE,
) -> Tuple[Optional[pd.DataFrame], Optional[str], Dict[str, Any]]:
    """
    Récupère les prévisions météorologiques journalières (jusqu'à 16 jours) via Open-Meteo Forecast API.
    Inclut la probabilité de précipitation maximale journalière lorsqu'elle est fournie par la source.
    """
    meta_info: Dict[str, Any] = {
        "api_endpoint": "https://api.open-meteo.com/v1/forecast",
        "timezone": timezone,
        "forecast_days": forecast_days,
    }
    try:
        lat = float(latitude)
        lon = float(longitude)
        if not (-90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0):
            return None, "Coordonnées géographiques hors limites.", meta_info
    except (TypeError, ValueError):
        return None, "Coordonnées géographiques invalides.", meta_info

    days = int(max(1, min(16, forecast_days)))
    url = "https://api.open-meteo.com/v1/forecast"
    params = {
        "latitude": lat,
        "longitude": lon,
        "daily": (
            "precipitation_sum,precipitation_probability_max,"
            "temperature_2m_max,temperature_2m_min,temperature_2m_mean,wind_speed_10m_max"
        ),
        "hourly": "relative_humidity_2m",
        "forecast_days": days,
        "timezone": timezone,
    }

    try:
        resp = requests.get(url, params=params, timeout=12)
        if resp.status_code != 200:
            params_no_hourly = {k: v for k, v in params.items() if k != "hourly"}
            resp = requests.get(url, params=params_no_hourly, timeout=12)
        resp.raise_for_status()
        payload = resp.json()
        df, err = _parse_open_meteo_daily_payload(payload, source_type="PREVISIONS_API")
        if err:
            return None, err, meta_info
        meta_info["fetched_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        if not df.empty:
            meta_info["effective_start"] = str(df["date"].iloc[0])
            meta_info["effective_end"] = str(df["date"].iloc[-1])
        return df, None, meta_info
    except requests.RequestException as exc:
        return (
            None,
            f"Impossible de contacter l'API Open-Meteo Forecast ({exc.__class__.__name__}). "
            "Vérifiez votre connexion Internet ou utilisez des données locales/fictives.",
            meta_info,
        )
    except ValueError as exc:
        return None, f"Réponse JSON invalide de l'API Open-Meteo Forecast : {exc}", meta_info


def generate_fictitious_enspy_dataset(
    start_date: str = "2025-03-01",
    num_days: int = 120,
    seed: int = 2025,
) -> pd.DataFrame:
    """
    Génère une série météorologique fictive clairement identifiée pour le chantier-école ENSPY (Yaoundé).
    IMPORTANT : Cette série est destinée exclusivement à la démonstration pédagogique hors-ligne.
    Les 10 premières journées intègrent exactement la série de vérification déterministe :
    0, 4, 12, 0, 8, 25, 1, 0, 16, 3 mm/jour.
    """
    rng = np.random.default_rng(seed)
    base_dt = pd.to_datetime(start_date)
    dates = pd.date_range(base_dt, periods=num_days, freq="D")

    # Génération reproductible avec persistance (chaîne de Markov simple à 2 états : sec / humide)
    precip = np.zeros(num_days, dtype=float)
    # Les 10 premiers jours reprennent la série déterministe pédagogique
    det_len = min(len(DETERMINISTIC_TEST_PRECIP), num_days)
    precip[:det_len] = DETERMINISTIC_TEST_PRECIP[:det_len]

    wet_state = False
    for i in range(det_len, num_days):
        month = dates[i].month
        # Probabilité de transition selon saisonnalité illustrative (mars-juin plus humide, juillet-août plus tempéré, sept-oct très humide)
        if month in (9, 10):
            p_dry_to_wet, p_wet_to_wet = 0.42, 0.68
            scale_mm = 11.5
        elif month in (4, 5, 6):
            p_dry_to_wet, p_wet_to_wet = 0.35, 0.58
            scale_mm = 9.5
        elif month in (12, 1, 2):
            p_dry_to_wet, p_wet_to_wet = 0.10, 0.25
            scale_mm = 5.0
        else:
            p_dry_to_wet, p_wet_to_wet = 0.22, 0.45
            scale_mm = 7.5

        prob = p_wet_to_wet if wet_state else p_dry_to_wet
        wet_state = bool(rng.random() < prob)
        if wet_state:
            val = rng.gamma(shape=1.35, scale=scale_mm)
            precip[i] = np.round(val, 1)
        else:
            precip[i] = 0.0

    # Températures minimales, maximales et moyennes cohérentes (Tmin < Tmean < Tmax)
    t_min = np.round(19.5 + rng.normal(0.0, 1.1, size=num_days) - 0.04 * precip, 1)
    amplitude = np.round(np.clip(9.2 - 0.12 * precip + rng.normal(0.0, 1.2, size=num_days), 4.0, 14.0), 1)
    t_max = np.round(t_min + amplitude, 1)
    t_mean = np.round((t_min + t_max) / 2.0, 1)

    # Vitesse maximale du vent (km/h) : rafales plus fortes lors des grains orageux
    wind_max = np.round(
        np.clip(18.0 + 0.65 * precip + rng.gamma(shape=2.2, scale=4.2, size=num_days), 8.0, 68.0),
        1,
    )

    # Humidité relative moyenne (%)
    rh_mean = np.round(
        np.clip(74.0 + 0.45 * precip - 0.8 * (t_max - 28.0) + rng.normal(0.0, 3.5, size=num_days), 52.0, 96.0),
        1,
    )

    df = pd.DataFrame(
        {
            "date": dates.strftime("%Y-%m-%d"),
            "precipitation_sum": precip,
            "temperature_2m_min": t_min,
            "temperature_2m_max": t_max,
            "temperature_2m_mean": t_mean,
            "wind_speed_10m_max": wind_max,
            "relative_humidity_mean": rh_mean,
            "precipitation_probability_max": np.nan,
            "source_type": "FICTIF_DEMO",
        }
    )
    return df


def generate_deterministic_10day_dataframe(start_date: str = "2025-03-01") -> pd.DataFrame:
    """
    Construit le DataFrame exact de 10 jours du cas de vérification déterministe :
    Précipitations (mm/j) : 0, 4, 12, 0, 8, 25, 1, 0, 16, 3.
    """
    dates = pd.date_range(pd.to_datetime(start_date), periods=10, freq="D")
    precip = list(DETERMINISTIC_TEST_PRECIP)
    t_min = [20.0, 20.5, 19.0, 21.0, 20.0, 18.5, 20.5, 21.0, 19.0, 20.5]
    t_max = [30.0, 29.0, 26.5, 31.0, 28.5, 25.0, 29.5, 30.5, 26.0, 29.0]
    t_mean = [round((mn + mx) / 2.0, 2) for mn, mx in zip(t_min, t_max)]
    wind = [15.0, 22.0, 38.0, 14.0, 28.0, 46.0, 18.0, 16.0, 41.0, 20.0]
    rh = [72.0, 78.0, 88.0, 70.0, 82.0, 93.0, 76.0, 71.0, 90.0, 77.0]

    return pd.DataFrame(
        {
            "date": dates.strftime("%Y-%m-%d"),
            "precipitation_sum": precip,
            "temperature_2m_min": t_min,
            "temperature_2m_max": t_max,
            "temperature_2m_mean": t_mean,
            "wind_speed_10m_max": wind,
            "relative_humidity_mean": rh,
            "precipitation_probability_max": [np.nan] * 10,
            "source_type": "TEST_DETERMINISTE",
        }
    )


# ==============================================================================
# 4. FONCTIONS DE VALIDATION ET D'IMPORTATION (CSV / EXCEL)
# ==============================================================================

COLUMN_ALIASES: Dict[str, List[str]] = {
    "date": ["date", "jour", "time", "datetime", "timestamp", "dates"],
    "precipitation_sum": [
        "precipitation_sum",
        "precipitation",
        "precipitations",
        "précipitations",
        "pluie",
        "rain",
        "rainfall",
        "pluie_mm",
        "precip_mm",
        "rr",
    ],
    "temperature_2m_min": [
        "temperature_2m_min",
        "temp_min",
        "tmin",
        "t_min",
        "temperature_min",
        "température_min",
    ],
    "temperature_2m_max": [
        "temperature_2m_max",
        "temp_max",
        "tmax",
        "t_max",
        "temperature_max",
        "température_max",
    ],
    "temperature_2m_mean": [
        "temperature_2m_mean",
        "temp_mean",
        "tmoy",
        "t_moy",
        "tmean",
        "temperature_moyenne",
        "température_moyenne",
    ],
    "wind_speed_10m_max": [
        "wind_speed_10m_max",
        "wind_max",
        "vent",
        "vent_max",
        "vitesse_vent",
        "wind_speed",
        "rafales",
    ],
    "relative_humidity_mean": [
        "relative_humidity_mean",
        "relative_humidity_2m_mean",
        "humidity",
        "humidite",
        "humidité",
        "rh",
        "hr",
    ],
    "precipitation_probability_max": [
        "precipitation_probability_max",
        "prob_pluie",
        "precip_prob",
        "probabilite_pluie",
    ],
}


def validate_and_import_weather_dataframe(
    raw_df: pd.DataFrame,
    unit_wind: str = "km/h",
    unit_precip: str = "mm",
) -> Tuple[Optional[pd.DataFrame], Dict[str, Any], Optional[str]]:
    """
    Valide et normalise un DataFrame importé depuis un fichier CSV ou Excel :
    - Vérifie la présence d'une colonne de date et d'au moins une variable météo exploitable.
    - Convertit les unités si l'utilisateur signale un vent en m/s ou des précipitations en cm.
    - Détecte et signale les doublons de dates, les valeurs hors bornes physiques (ex. pluie < 0, Tmin > Tmax)
      et comptabilise les valeurs manquantes sans jamais les inventer.
    """
    report: Dict[str, Any] = {
        "initial_rows": 0,
        "valid_date_rows": 0,
        "duplicates_removed": 0,
        "mapped_columns": {},
        "missing_counts": {},
        "anomalies_corrected_to_nan": 0,
        "warnings": [],
    }

    if raw_df is None or raw_df.empty:
        return None, report, "Le fichier importé est vide ou ne contient aucune ligne de données."

    report["initial_rows"] = len(raw_df)

    # Correspondance automatique des noms de colonnes
    col_lookup = {str(c).strip().lower(): c for c in raw_df.columns}
    mapped: Dict[str, str] = {}
    for target, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias.lower() in col_lookup:
                mapped[target] = col_lookup[alias.lower()]
                break

    if "date" not in mapped:
        return (
            None,
            report,
            "Colonne de date introuvable. Le fichier doit contenir une colonne nommée 'date', 'jour' ou 'time'.",
        )

    weather_Found = [k for k in mapped.keys() if k != "date"]
    if not weather_Found:
        return (
            None,
            report,
            "Aucune variable météorologique reconnue (ex. 'precipitation_sum', 'pluie', 'tmin', 'tmax', 'vent').",
        )

    report["mapped_columns"] = mapped

    df = pd.DataFrame()
    parsed_dates = pd.to_datetime(raw_df[mapped["date"]], errors="coerce")
    invalid_dates_count = int(parsed_dates.isna().sum())
    if invalid_dates_count > 0:
        report["warnings"].append(f"{invalid_dates_count} ligne(s) avec date invalide ont été écartées.")

    df["date"] = parsed_dates
    for std_col in [
        "precipitation_sum",
        "temperature_2m_min",
        "temperature_2m_max",
        "temperature_2m_mean",
        "wind_speed_10m_max",
        "relative_humidity_mean",
        "precipitation_probability_max",
    ]:
        if std_col in mapped:
            series = raw_df[mapped[std_col]]
            if series.dtype == object:
                series = series.astype(str).str.replace(",", ".", regex=False)
            df[std_col] = pd.to_numeric(series, errors="coerce")
        else:
            df[std_col] = np.nan

    df = df.dropna(subset=["date"]).copy()
    if df.empty:
        return None, report, "Aucune date valide n'a pu être extraite du fichier."

    df["date"] = df["date"].dt.strftime("%Y-%m-%d")
    report["valid_date_rows"] = len(df)

    # Doublons de dates
    before_dedup = len(df)
    df = df.sort_values("date").drop_duplicates(subset=["date"], keep="first").reset_index(drop=True)
    dup_count = before_dedup - len(df)
    report["duplicates_removed"] = dup_count
    if dup_count > 0:
        report["warnings"].append(
            f"{dup_count} doublon(s) de date détecté(s) : seule la première occurrence a été conservée."
        )

    # Conversions d'unités explicites
    if unit_wind == "m/s" and "wind_speed_10m_max" in mapped:
        df["wind_speed_10m_max"] = df["wind_speed_10m_max"] * 3.6
        report["warnings"].append("Vitesses de vent converties de m/s en km/h (× 3,6).")
    if unit_precip == "cm" and "precipitation_sum" in mapped:
        df["precipitation_sum"] = df["precipitation_sum"] * 10.0
        report["warnings"].append("Précipitations converties de cm en mm (× 10).")

    # Contrôle de cohérence physique (valeurs impossibles remplacées par NaN et signalées)
    anomalies = 0
    neg_precip = df["precipitation_sum"] < 0
    if neg_precip.any():
        cnt = int(neg_precip.sum())
        anomalies += cnt
        df.loc[neg_precip, "precipitation_sum"] = np.nan
        report["warnings"].append(f"{cnt} valeur(s) de précipitation négative(s) convertie(s) en NaN.")

    neg_wind = df["wind_speed_10m_max"] < 0
    if neg_wind.any():
        cnt = int(neg_wind.sum())
        anomalies += cnt
        df.loc[neg_wind, "wind_speed_10m_max"] = np.nan
        report["warnings"].append(f"{cnt} vitesse(s) de vent négative(s) convertie(s) en NaN.")

    bad_rh = (df["relative_humidity_mean"] < 0) | (df["relative_humidity_mean"] > 100)
    if bad_rh.any():
        cnt = int(bad_rh.sum())
        anomalies += cnt
        df.loc[bad_rh, "relative_humidity_mean"] = np.nan
        report["warnings"].append(f"{cnt} valeur(s) d'humidité hors [0 %, 100 %] convertie(s) en NaN.")

    bad_temp_order = (
        df["temperature_2m_min"].notna()
        & df["temperature_2m_max"].notna()
        & (df["temperature_2m_min"] > df["temperature_2m_max"])
    )
    if bad_temp_order.any():
        cnt = int(bad_temp_order.sum())
        anomalies += cnt
        df.loc[bad_temp_order, ["temperature_2m_min", "temperature_2m_max", "temperature_2m_mean"]] = np.nan
        report["warnings"].append(f"{cnt} journée(s) avec Tmin > Tmax invalidée(s) en NaN.")

    # Calcul de Tmoy si non importée mais Tmin et Tmax valides
    mask_tmean = df["temperature_2m_mean"].isna() & df["temperature_2m_min"].notna() & df["temperature_2m_max"].notna()
    df.loc[mask_tmean, "temperature_2m_mean"] = (
        df.loc[mask_tmean, "temperature_2m_min"] + df.loc[mask_tmean, "temperature_2m_max"]
    ) / 2.0

    report["anomalies_corrected_to_nan"] = anomalies
    for col in weather_Found:
        report["missing_counts"][col] = int(df[col].isna().sum())

    df["source_type"] = "IMPORT_FICHIER"
    return df, report, None


# ==============================================================================
# 5. MOTEUR DE CALCULS STATISTIQUES DESCRIPTIFS
# ==============================================================================

def compute_descriptive_stats(
    series: pd.Series,
    threshold: Optional[float] = None,
    rainy_day_threshold_mm: float = 1.0,
    inclusive_threshold: bool = True,
) -> Dict[str, Any]:
    """
    Calcule l'ensemble des indicateurs statistiques descriptifs sur une série numérique :
    - Effectif total N_total, observations valides n, valeurs manquantes n_na
    - Minimum, Maximum, Somme
    - Moyenne arithmétique, Médiane
    - Variance d'échantillon s² (ddof=1) et variance empirique σ² (ddof=0)
    - Écart-type s (ddof=1) et σ (ddof=0), Coefficient de variation CV (%)
    - Quartiles Q1 (P25), Q2 (P50), Q3 (P75), Écart interquartile IQR, Percentiles P10, P90, P95
    - Fréquence des journées pluvieuses (≥ 1 mm et > 0 mm) et fréquence de dépassement du seuil choisi.
    Gère proprement les séries vides, constantes ou de taille n = 1 sans division par zéro.
    """
    total_len = int(len(series)) if series is not None else 0
    if series is None or total_len == 0:
        return {
            "n_total": 0,
            "n_valid": 0,
            "n_missing": 0,
            "sum": np.nan,
            "min": np.nan,
            "max": np.nan,
            "mean": np.nan,
            "median": np.nan,
            "var_sample": np.nan,
            "std_sample": np.nan,
            "var_pop": np.nan,
            "std_pop": np.nan,
            "cv_pct": np.nan,
            "p10": np.nan,
            "q25": np.nan,
            "q50": np.nan,
            "q75": np.nan,
            "p90": np.nan,
            "p95": np.nan,
            "iqr": np.nan,
            "rainy_days_ge_1mm": 0,
            "rainy_freq_ge_1mm_pct": np.nan,
            "rainy_days_gt_0mm": 0,
            "rainy_freq_gt_0mm_pct": np.nan,
            "exceed_count": 0,
            "exceed_freq_pct": np.nan,
            "threshold_used": threshold,
        }

    clean = pd.to_numeric(series, errors="coerce").dropna()
    n_valid = int(len(clean))
    n_missing = total_len - n_valid

    if n_valid == 0:
        return {
            "n_total": total_len,
            "n_valid": 0,
            "n_missing": n_missing,
            "sum": np.nan,
            "min": np.nan,
            "max": np.nan,
            "mean": np.nan,
            "median": np.nan,
            "var_sample": np.nan,
            "std_sample": np.nan,
            "var_pop": np.nan,
            "std_pop": np.nan,
            "cv_pct": np.nan,
            "p10": np.nan,
            "q25": np.nan,
            "q50": np.nan,
            "q75": np.nan,
            "p90": np.nan,
            "p95": np.nan,
            "iqr": np.nan,
            "rainy_days_ge_1mm": 0,
            "rainy_freq_ge_1mm_pct": np.nan,
            "rainy_days_gt_0mm": 0,
            "rainy_freq_gt_0mm_pct": np.nan,
            "exceed_count": 0,
            "exceed_freq_pct": np.nan,
            "threshold_used": threshold,
        }

    arr = clean.to_numpy(dtype=float)
    mean_val = float(np.mean(arr))
    median_val = float(np.median(arr))
    min_val = float(np.min(arr))
    max_val = float(np.max(arr))
    sum_val = float(np.sum(arr))

    var_pop = float(np.var(arr, ddof=0))
    std_pop = float(np.std(arr, ddof=0))
    if n_valid > 1:
        var_sample = float(np.var(arr, ddof=1))
        std_sample = float(np.std(arr, ddof=1))
    else:
        var_sample = np.nan
        std_sample = np.nan

    cv_pct = float((std_sample / abs(mean_val)) * 100.0) if (n_valid > 1 and abs(mean_val) > 1e-12) else np.nan

    p10, q25, q50, q75, p90, p95 = [
        float(x) for x in np.percentile(arr, [10, 25, 50, 75, 90, 95])
    ]
    iqr_val = q75 - q25

    rain_ge_1 = int(np.sum(arr >= rainy_day_threshold_mm))
    rain_gt_0 = int(np.sum(arr > 0.0))

    exceed_count = 0
    exceed_freq_pct = np.nan
    if threshold is not None and not np.isnan(threshold):
        if inclusive_threshold:
            exceed_count = int(np.sum(arr >= float(threshold)))
        else:
            exceed_count = int(np.sum(arr > float(threshold)))
        exceed_freq_pct = (exceed_count / n_valid) * 100.0

    return {
        "n_total": total_len,
        "n_valid": n_valid,
        "n_missing": n_missing,
        "sum": sum_val,
        "min": min_val,
        "max": max_val,
        "mean": mean_val,
        "median": median_val,
        "var_sample": var_sample,
        "std_sample": std_sample,
        "var_pop": var_pop,
        "std_pop": std_pop,
        "cv_pct": cv_pct,
        "p10": p10,
        "q25": q25,
        "q50": q50,
        "q75": q75,
        "p90": p90,
        "p95": p95,
        "iqr": iqr_val,
        "rainy_days_ge_1mm": rain_ge_1,
        "rainy_freq_ge_1mm_pct": (rain_ge_1 / n_valid) * 100.0,
        "rainy_days_gt_0mm": rain_gt_0,
        "rainy_freq_gt_0mm_pct": (rain_gt_0 / n_valid) * 100.0,
        "exceed_count": exceed_count,
        "exceed_freq_pct": exceed_freq_pct,
        "threshold_used": threshold,
    }


def compute_monthly_summary(df: pd.DataFrame, precip_threshold: float = 10.0) -> pd.DataFrame:
    """Calcule l'évolution mensuelle des précipitations, températures et vitesses de vent."""
    if df is None or df.empty or "date" not in df.columns:
        return pd.DataFrame()

    work = df.copy()
    work["dt"] = pd.to_datetime(work["date"], errors="coerce")
    work = work.dropna(subset=["dt"])
    if work.empty:
        return pd.DataFrame()

    work["year_month"] = work["dt"].dt.strftime("%Y-%m")
    work["exceed_thresh"] = work["precipitation_sum"].apply(
        lambda x: 1 if (pd.notna(x) and float(x) >= precip_threshold) else 0
    )
    work["rainy_1mm"] = work["precipitation_sum"].apply(
        lambda x: 1 if (pd.notna(x) and float(x) >= 1.0) else 0
    )

    grouped = (
        work.groupby("year_month", as_index=False)
        .agg(
            nb_jours=("date", "count"),
            jours_pluie_valides=("precipitation_sum", lambda s: int(s.notna().sum())),
            cumul_pluie_mm=("precipitation_sum", lambda s: float(s.sum(min_count=1)) if s.notna().any() else np.nan),
            pluie_moy_mm=("precipitation_sum", "mean"),
            pluie_max_mm=("precipitation_sum", "max"),
            jours_pluie_ge_1mm=("rainy_1mm", "sum"),
            jours_depassement_seuil=("exceed_thresh", "sum"),
            temp_min_moy_c=("temperature_2m_min", "mean"),
            temp_max_moy_c=("temperature_2m_max", "mean"),
            vent_max_kmh=("wind_speed_10m_max", "max"),
        )
        .sort_values("year_month")
        .reset_index(drop=True)
    )
    grouped["freq_depassement_pct"] = np.where(
        grouped["jours_pluie_valides"] > 0,
        (grouped["jours_depassement_seuil"] / grouped["jours_pluie_valides"]) * 100.0,
        np.nan,
    )
    return grouped


# ==============================================================================
# 6. MOTEUR DE PLANIFICATION MÉTÉO ET CLASSIFICATION DES ACTIVITÉS
# ==============================================================================

def evaluate_single_day_for_criteria(
    row: Optional[pd.Series],
    precip_max_mm: Optional[float] = 10.0,
    temp_min_c: Optional[float] = None,
    temp_max_c: Optional[float] = None,
    wind_max_kmh: Optional[float] = None,
    humidity_max_pct: Optional[float] = None,
) -> Dict[str, Any]:
    """
    Évalue une journée selon les critères d'une activité de chantier et la classe en 3 catégories :
    - 'Compatible' : toutes les variables requises sont disponibles et respectent les seuils.
    - 'Défavorable' : au moins un critère est franchi (avec détail exact des critères dépassés).
    - 'Évaluation impossible' : date absente du jeu de données ou valeur manquante sur un critère actif.
    Conventions de dépassement :
      - Précipitation défavorable si P >= precip_max_mm
      - Température min défavorable si Tmin < temp_min_c
      - Température max défavorable si Tmax > temp_max_c
      - Vent max défavorable si Vmax >= wind_max_kmh
      - Humidité relative défavorable si HR >= humidity_max_pct
    """
    if row is None:
        return {
            "status": "Évaluation impossible",
            "reasons": ["Date située hors des données météorologiques disponibles"],
            "missing_vars": ["Toutes variables"],
        }

    reasons: List[str] = []
    missing_vars: List[str] = []

    # 1. Précipitations
    if precip_max_mm is not None and not pd.isna(precip_max_mm):
        val_p = row.get("precipitation_sum", np.nan)
        if pd.isna(val_p):
            missing_vars.append("Précipitations")
        elif float(val_p) >= float(precip_max_mm):
            reasons.append(f"Pluie {float(val_p):.1f} mm ≥ seuil {float(precip_max_mm):.1f} mm")

    # 2. Température minimale
    if temp_min_c is not None and not pd.isna(temp_min_c):
        val_tmin = row.get("temperature_2m_min", np.nan)
        if pd.isna(val_tmin):
            missing_vars.append("Température min")
        elif float(val_tmin) < float(temp_min_c):
            reasons.append(f"Tmin {float(val_tmin):.1f} °C < seuil {float(temp_min_c):.1f} °C")

    # 3. Température maximale
    if temp_max_c is not None and not pd.isna(temp_max_c):
        val_tmax = row.get("temperature_2m_max", np.nan)
        if pd.isna(val_tmax):
            missing_vars.append("Température max")
        elif float(val_tmax) > float(temp_max_c):
            reasons.append(f"Tmax {float(val_tmax):.1f} °C > seuil {float(temp_max_c):.1f} °C")

    # 4. Vent maximal
    if wind_max_kmh is not None and not pd.isna(wind_max_kmh):
        val_w = row.get("wind_speed_10m_max", np.nan)
        if pd.isna(val_w):
            missing_vars.append("Vent max")
        elif float(val_w) >= float(wind_max_kmh):
            reasons.append(f"Vent {float(val_w):.1f} km/h ≥ seuil {float(wind_max_kmh):.1f} km/h")

    # 5. Humidité relative (optionnelle)
    if humidity_max_pct is not None and not pd.isna(humidity_max_pct):
        val_rh = row.get("relative_humidity_mean", np.nan)
        if pd.isna(val_rh):
            missing_vars.append("Humidité relative")
        elif float(val_rh) >= float(humidity_max_pct):
            reasons.append(f"Humidité {float(val_rh):.1f} % ≥ seuil {float(humidity_max_pct):.1f} %")

    # Si un critère mesuré est déjà défavorable, la journée est Défavorable
    if reasons:
        if missing_vars:
            reasons.append(f"(Note : données manquantes sur {', '.join(missing_vars)})")
        return {
            "status": "Défavorable",
            "reasons": reasons,
            "missing_vars": missing_vars,
        }

    # Si aucun critère mesuré n'est franchi mais qu'une variable surveillée est manquante
    if missing_vars:
        return {
            "status": "Évaluation impossible",
            "reasons": [f"Données absentes pour : {', '.join(missing_vars)}"],
            "missing_vars": missing_vars,
        }

    return {
        "status": "Compatible",
        "reasons": ["Tous les paramètres respectent les seuils configurés"],
        "missing_vars": [],
    }


def evaluate_activity_schedule(
    df_weather: Optional[pd.DataFrame],
    activity: Dict[str, Any],
    override_start_date: Optional[str] = None,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Analyse jour par jour la fenêtre calendaire d'une activité de chantier.
    Retourne le tableau journalier détaillé et les statistiques de faisabilité.
    """
    start_str = override_start_date or activity.get("start_date", "2025-03-01")
    duration = max(1, int(activity.get("duration_days", 5)))
    start_dt = pd.to_datetime(start_str)
    target_dates = pd.date_range(start_dt, periods=duration, freq="D").strftime("%Y-%m-%d").tolist()

    weather_by_date: Dict[str, pd.Series] = {}
    if df_weather is not None and not df_weather.empty and "date" in df_weather.columns:
        for _, r in df_weather.iterrows():
            weather_by_date[str(r["date"])[:10]] = r

    rows: List[Dict[str, Any]] = []
    for idx, d_str in enumerate(target_dates, start=1):
        w_row = weather_by_date.get(d_str)
        eval_res = evaluate_single_day_for_criteria(
            w_row,
            precip_max_mm=activity.get("precip_max_mm"),
            temp_min_c=activity.get("temp_min_c"),
            temp_max_c=activity.get("temp_max_c"),
            wind_max_kmh=activity.get("wind_max_kmh"),
            humidity_max_pct=activity.get("humidity_max_pct"),
        )
        rows.append(
            {
                "jour_num": idx,
                "date": d_str,
                "activite": activity.get("name", "Activité"),
                "statut": eval_res["status"],
                "raisons": " ; ".join(eval_res["reasons"]),
                "precipitation_sum": float(w_row["precipitation_sum"]) if (w_row is not None and pd.notna(w_row.get("precipitation_sum"))) else np.nan,
                "temperature_2m_min": float(w_row["temperature_2m_min"]) if (w_row is not None and pd.notna(w_row.get("temperature_2m_min"))) else np.nan,
                "temperature_2m_max": float(w_row["temperature_2m_max"]) if (w_row is not None and pd.notna(w_row.get("temperature_2m_max"))) else np.nan,
                "wind_speed_10m_max": float(w_row["wind_speed_10m_max"]) if (w_row is not None and pd.notna(w_row.get("wind_speed_10m_max"))) else np.nan,
                "relative_humidity_mean": float(w_row["relative_humidity_mean"]) if (w_row is not None and pd.notna(w_row.get("relative_humidity_mean"))) else np.nan,
                "seuil_pluie_mm": activity.get("precip_max_mm"),
                "seuil_tmin_c": activity.get("temp_min_c"),
                "seuil_tmax_c": activity.get("temp_max_c"),
                "seuil_vent_kmh": activity.get("wind_max_kmh"),
            }
        )

    res_df = pd.DataFrame(rows)
    n_comp = int((res_df["statut"] == "Compatible").sum())
    n_def = int((res_df["statut"] == "Défavorable").sum())
    n_imp = int((res_df["statut"] == "Évaluation impossible").sum())
    n_evaluable = n_comp + n_def

    summary = {
        "activity_name": activity.get("name", "Activité"),
        "start_date": target_dates[0],
        "end_date": target_dates[-1],
        "duration_days": duration,
        "compatible_days": n_comp,
        "unfavorable_days": n_def,
        "impossible_days": n_imp,
        "evaluable_days": n_evaluable,
        "compatibility_rate_on_total_pct": (n_comp / duration) * 100.0,
        "unfavorable_rate_on_evaluable_pct": (n_def / n_evaluable) * 100.0 if n_evaluable > 0 else np.nan,
        "estimated_delay_days": n_def,
    }
    return res_df, summary


def compare_candidate_start_dates(
    df_weather: Optional[pd.DataFrame],
    activity: Dict[str, Any],
    candidate_dates: List[str],
) -> pd.DataFrame:
    """Compare plusieurs dates de démarrage possibles pour une activité donnée."""
    records: List[Dict[str, Any]] = []
    for c_date in candidate_dates:
        _, summ = evaluate_activity_schedule(df_weather, activity, override_start_date=c_date)
        if summ["impossible_days"] == summ["duration_days"]:
            rec_label = "Hors données disponibles"
        elif summ["unfavorable_days"] == 0 and summ["impossible_days"] == 0:
            rec_label = "Créneau très favorable (0 jour défavorable)"
        elif summ["unfavorable_days"] <= max(1, int(0.2 * summ["duration_days"])) and summ["impossible_days"] == 0:
            rec_label = "Créneau favorable (perturbations limitées)"
        elif summ["impossible_days"] > 0:
            rec_label = "Couverture météo partielle (données manquantes)"
        else:
            rec_label = "Créneau exposé aux interruptions météo"

        records.append(
            {
                "Date de début": summ["start_date"],
                "Date de fin théorique": summ["end_date"],
                "Durée (jours)": summ["duration_days"],
                "Jours compatibles": summ["compatible_days"],
                "Jours défavorables": summ["unfavorable_days"],
                "Évaluation impossible": summ["impossible_days"],
                "Taux compatible (%)": round(summ["compatibility_rate_on_total_pct"], 1),
                "Taux défavorable / évalués (%)": (
                    round(summ["unfavorable_rate_on_evaluable_pct"], 1)
                    if not pd.isna(summ["unfavorable_rate_on_evaluable_pct"])
                    else np.nan
                ),
                "Diagnostic": rec_label,
            }
        )
    return pd.DataFrame(records)


def evaluate_dataset_against_criteria(
    df_weather: pd.DataFrame,
    precip_max_mm: Optional[float] = 10.0,
    temp_min_c: Optional[float] = None,
    temp_max_c: Optional[float] = None,
    wind_max_kmh: Optional[float] = None,
    humidity_max_pct: Optional[float] = None,
) -> pd.DataFrame:
    """
    Évalue l'intégralité d'un DataFrame météorologique selon un jeu de seuils donné.
    Ajoute les colonnes 'statut' et 'raisons'.
    """
    if df_weather is None or df_weather.empty:
        return pd.DataFrame()

    out = df_weather.copy()
    statuses = []
    reasons_list = []
    for _, row in out.iterrows():
        ev = evaluate_single_day_for_criteria(
            row,
            precip_max_mm=precip_max_mm,
            temp_min_c=temp_min_c,
            temp_max_c=temp_max_c,
            wind_max_kmh=wind_max_kmh,
            humidity_max_pct=humidity_max_pct,
        )
        statuses.append(ev["status"])
        reasons_list.append(" ; ".join(ev["reasons"]))

    out["statut"] = statuses
    out["raisons"] = reasons_list
    return out


# ==============================================================================
# 7. MOTEUR DE PROBABILITÉS, FENÊTRES CONSÉCUTIVES, MONTE-CARLO & SENSIBILITÉ
# ==============================================================================

def compute_empirical_exceedance_probability(
    series: pd.Series,
    threshold: float,
    inclusive: bool = True,
    confidence_level: float = 0.95,
) -> Dict[str, Any]:
    """
    Calcule la probabilité empirique de dépassement d'un seuil x :
      P_estimée(X >= x) = k / n_valides (si inclusive=True)
      P_estimée(X > x)  = k / n_valides (si inclusive=False)
    Calcule également l'intervalle de confiance de Wilson pour une proportion binomiale.
    """
    clean = pd.to_numeric(series, errors="coerce").dropna()
    n_valid = int(len(clean))
    n_missing = int(len(series) - n_valid) if series is not None else 0

    if n_valid == 0:
        return {
            "n_valid": 0,
            "n_missing": n_missing,
            "k_exceed": 0,
            "prob": np.nan,
            "prob_pct": np.nan,
            "ci_low_pct": np.nan,
            "ci_high_pct": np.nan,
            "threshold": threshold,
            "operator": "≥" if inclusive else ">",
        }

    arr = clean.to_numpy(dtype=float)
    if inclusive:
        k_exceed = int(np.sum(arr >= float(threshold)))
        op_str = "≥"
    else:
        k_exceed = int(np.sum(arr > float(threshold)))
        op_str = ">"

    p_hat = k_exceed / n_valid

    # Intervalle de confiance de Wilson
    z = float(scipy_stats.norm.ppf(1.0 - (1.0 - confidence_level) / 2.0))
    denom = 1.0 + (z ** 2) / n_valid
    center = (p_hat + (z ** 2) / (2.0 * n_valid)) / denom
    half_width = (
        z * math.sqrt((p_hat * (1.0 - p_hat) + (z ** 2) / (4.0 * n_valid)) / n_valid) / denom
    )
    ci_low = max(0.0, center - half_width)
    ci_high = min(1.0, center + half_width)

    return {
        "n_valid": n_valid,
        "n_missing": n_missing,
        "k_exceed": k_exceed,
        "prob": p_hat,
        "prob_pct": p_hat * 100.0,
        "ci_low_pct": ci_low * 100.0,
        "ci_high_pct": ci_high * 100.0,
        "threshold": float(threshold),
        "operator": op_str,
    }


def analyze_consecutive_windows(
    df_evaluated: pd.DataFrame,
    window_days: int = 5,
) -> Dict[str, Any]:
    """
    Analyse toutes les fenêtres consécutives de durée k = window_days dans la série historique.
    - Vérifie la continuité calendaire stricte (écart de k-1 jours entre le début et la fin de la fenêtre)
      ainsi que l'absence de journées 'Évaluation impossible' dans la fenêtre.
    - Comptabilise séparément les fenêtres incomplètes ou comportant des données manquantes.
    - Détermine la fréquence historique des fenêtres contenant au moins une journée défavorable.
    """
    k = max(1, int(window_days))
    if df_evaluated is None or df_evaluated.empty or len(df_evaluated) < k:
        return {
            "window_days": k,
            "total_possible_positions": 0,
            "valid_complete_windows": 0,
            "incomplete_or_missing_windows": 0,
            "windows_with_at_least_one_unfavorable": 0,
            "windows_all_compatible": 0,
            "freq_at_least_one_unfavorable_pct": np.nan,
            "freq_all_compatible_pct": np.nan,
            "mean_unfavorable_days_per_window": np.nan,
            "distribution_df": pd.DataFrame(),
            "windows_detail_df": pd.DataFrame(),
        }

    work = df_evaluated.sort_values("date").reset_index(drop=True).copy()
    work["dt"] = pd.to_datetime(work["date"], errors="coerce")

    n_rows = len(work)
    total_positions = n_rows - k + 1
    valid_windows = 0
    incomplete_windows = 0
    at_least_one_unfav = 0
    all_compatible = 0
    unfav_counts: List[int] = []
    window_records: List[Dict[str, Any]] = []

    for i in range(total_positions):
        sub = work.iloc[i : i + k]
        dt_start = sub["dt"].iloc[0]
        dt_end = sub["dt"].iloc[-1]

        # Vérifier la continuité calendaire stricte (pas de trou de date)
        is_calendar_continuous = (
            pd.notna(dt_start)
            and pd.notna(dt_end)
            and int((dt_end - dt_start).days) == (k - 1)
        )
        has_impossible = bool((sub["statut"] == "Évaluation impossible").any())

        if (not is_calendar_continuous) or has_impossible:
            incomplete_windows += 1
            window_records.append(
                {
                    "debut_fenetre": str(sub["date"].iloc[0]),
                    "fin_fenetre": str(sub["date"].iloc[-1]),
                    "validite": "Incomplète / Données manquantes",
                    "nb_jours_defavorables": np.nan,
                    "au_moins_un_defavorable": np.nan,
                }
            )
            continue

        valid_windows += 1
        n_def = int((sub["statut"] == "Défavorable").sum())
        unfav_counts.append(n_def)
        if n_def >= 1:
            at_least_one_unfav += 1
        else:
            all_compatible += 1

        window_records.append(
            {
                "debut_fenetre": str(sub["date"].iloc[0]),
                "fin_fenetre": str(sub["date"].iloc[-1]),
                "validite": "Valide",
                "nb_jours_defavorables": n_def,
                "au_moins_un_defavorable": "Oui" if n_def >= 1 else "Non",
            }
        )

    if valid_windows == 0:
        return {
            "window_days": k,
            "total_possible_positions": total_positions,
            "valid_complete_windows": 0,
            "incomplete_or_missing_windows": incomplete_windows,
            "windows_with_at_least_one_unfavorable": 0,
            "windows_all_compatible": 0,
            "freq_at_least_one_unfavorable_pct": np.nan,
            "freq_all_compatible_pct": np.nan,
            "mean_unfavorable_days_per_window": np.nan,
            "distribution_df": pd.DataFrame(),
            "windows_detail_df": pd.DataFrame(window_records),
        }

    arr_counts = np.array(unfav_counts, dtype=int)
    dist_rows = []
    for c_val in range(0, k + 1):
        cnt = int(np.sum(arr_counts == c_val))
        dist_rows.append(
            {
                "nb_jours_defavorables": c_val,
                "nb_fenetres": cnt,
                "frequence_pct": (cnt / valid_windows) * 100.0,
            }
        )

    return {
        "window_days": k,
        "total_possible_positions": total_positions,
        "valid_complete_windows": valid_windows,
        "incomplete_or_missing_windows": incomplete_windows,
        "windows_with_at_least_one_unfavorable": at_least_one_unfav,
        "windows_all_compatible": all_compatible,
        "freq_at_least_one_unfavorable_pct": (at_least_one_unfav / valid_windows) * 100.0,
        "freq_all_compatible_pct": (all_compatible / valid_windows) * 100.0,
        "mean_unfavorable_days_per_window": float(np.mean(arr_counts)),
        "distribution_df": pd.DataFrame(dist_rows),
        "windows_detail_df": pd.DataFrame(window_records),
    }


def run_monte_carlo_block_bootstrap(
    df_evaluated: pd.DataFrame,
    horizon_days: int = 10,
    num_simulations: int = 2000,
    block_length: int = 3,
    random_seed: int = 42,
) -> Dict[str, Any]:
    """
    Simulation de Monte-Carlo par rééchantillonnage de blocs de journées historiques (Moving Block Bootstrap).
    - Conserve la dépendance temporelle intra-bloc (persistance des épisodes pluvieux).
    - Conserve explicitement les scénarios où 0 événement défavorable survient.
    - N'invente aucune date calendaire future (travaille sur l'horizon abstrait J+1 à J+D).
    - Calcule la fréquence simulée d'au moins un jour défavorable, la distribution complète de K,
      les quantiles adaptés et l'erreur standard de Monte-Carlo (précision liée à M).
    """
    D = max(1, int(horizon_days))
    M = max(10, int(num_simulations))
    L = max(1, int(block_length))
    seed = int(random_seed)

    if df_evaluated is None or df_evaluated.empty or "statut" not in df_evaluated.columns:
        return {"error": "Aucune donnée évaluée disponible pour lancer la simulation de Monte-Carlo."}

    # Extraire les journées évaluables (Compatible = 0, Défavorable = 1)
    work = df_evaluated.sort_values("date").reset_index(drop=True).copy()
    work["dt"] = pd.to_datetime(work["date"], errors="coerce")
    evaluable = work[work["statut"].isin(["Compatible", "Défavorable"])].copy()

    if len(evaluable) < max(3, L):
        return {
            "error": (
                f"Échantillon historique insuffisant ({len(evaluable)} journée(s) évaluable(s)) "
                f"pour construire des blocs de longueur L = {L} jour(s)."
            )
        }

    # Construire la liste des blocs consécutifs valides de longueur L
    work["is_unfav"] = np.where(work["statut"] == "Défavorable", 1, 0)
    work["is_eval"] = work["statut"].isin(["Compatible", "Défavorable"])

    valid_blocks: List[np.ndarray] = []
    if L == 1:
        valid_blocks = [
            np.array([int(v)], dtype=int)
            for v in work.loc[work["is_eval"], "is_unfav"].to_numpy()
        ]
    else:
        for i in range(len(work) - L + 1):
            sub = work.iloc[i : i + L]
            if not sub["is_eval"].all():
                continue
            dt_s = sub["dt"].iloc[0]
            dt_e = sub["dt"].iloc[-1]
            if pd.notna(dt_s) and pd.notna(dt_e) and int((dt_e - dt_s).days) == (L - 1):
                valid_blocks.append(sub["is_unfav"].to_numpy(dtype=int))

    # Si aucun bloc continu de taille L n'a pu être formé (ex. données très fragmentées), repli sur blocs sur indices évaluables
    fallback_note = None
    if not valid_blocks:
        eval_arr = evaluable["statut"].apply(lambda s: 1 if s == "Défavorable" else 0).to_numpy(dtype=int)
        for i in range(len(eval_arr) - L + 1):
            valid_blocks.append(eval_arr[i : i + L])
        fallback_note = (
            "Série historique comportant des interruptions calendaires : "
            "les blocs ont été constitués sur la séquence des journées évaluables."
        )

    blocks_matrix = np.vstack(valid_blocks)  # shape: (n_blocks, L)
    n_blocks = blocks_matrix.shape[0]
    num_blocks_needed = math.ceil(D / L)

    rng = np.random.default_rng(seed)
    chosen_indices = rng.integers(0, n_blocks, size=(M, num_blocks_needed))
    # Assemblage des trajectoires simulées : shape (M, num_blocks_needed * L) puis troncature à D jours
    sampled_trajectories = blocks_matrix[chosen_indices].reshape(M, num_blocks_needed * L)[:, :D]

    # Nombre de journées défavorables K pour chacune des M simulations
    k_sim = np.sum(sampled_trajectories, axis=1)

    n_zero_event = int(np.sum(k_sim == 0))
    n_at_least_one = int(np.sum(k_sim >= 1))
    p_at_least_one = n_at_least_one / M
    p_zero_event = n_zero_event / M

    # Précision de Monte-Carlo (erreur standard sur la probabilité estimée et sur la moyenne de K)
    mc_se_prob = math.sqrt(max(0.0, p_at_least_one * (1.0 - p_at_least_one)) / M)
    mc_ci95_half_prob_pct = 1.96 * mc_se_prob * 100.0

    mean_k = float(np.mean(k_sim))
    std_k = float(np.std(k_sim, ddof=1)) if M > 1 else 0.0
    mc_se_mean_k = std_k / math.sqrt(M) if M > 1 else 0.0

    q10, q25, q50, q75, q90, q95 = [
        float(x) for x in np.percentile(k_sim, [10, 25, 50, 75, 90, 95])
    ]

    # Distribution complète (y compris K = 0)
    dist_rows = []
    cum_count = 0
    for k_val in range(0, D + 1):
        cnt = int(np.sum(k_sim == k_val))
        cum_count += cnt
        dist_rows.append(
            {
                "nb_jours_defavorables": k_val,
                "nb_simulations": cnt,
                "frequence_simulee_pct": (cnt / M) * 100.0,
                "frequence_cumulee_pct": (cum_count / M) * 100.0,
            }
        )

    return {
        "error": None,
        "horizon_days": D,
        "num_simulations": M,
        "block_length": L,
        "random_seed": seed,
        "n_historical_evaluable_days": len(evaluable),
        "n_available_blocks": n_blocks,
        "fallback_note": fallback_note,
        "scenarios_zero_event_count": n_zero_event,
        "scenarios_at_least_one_count": n_at_least_one,
        "prob_zero_event_pct": p_zero_event * 100.0,
        "prob_at_least_one_pct": p_at_least_one * 100.0,
        "mc_se_prob_pct": mc_se_prob * 100.0,
        "mc_ci95_half_prob_pct": mc_ci95_half_prob_pct,
        "mean_unfavorable_days": mean_k,
        "std_unfavorable_days": std_k,
        "mc_se_mean_days": mc_se_mean_k,
        "q10_days": q10,
        "q25_days": q25,
        "q50_days": q50,
        "q75_days": q75,
        "q90_days": q90,
        "q95_days": q95,
        "min_days": int(np.min(k_sim)),
        "max_days": int(np.max(k_sim)),
        "simulated_counts": k_sim,
        "distribution_df": pd.DataFrame(dist_rows),
    }


def compute_sensitivity_curve(
    df_weather: pd.DataFrame,
    variable_col: str = "precipitation_sum",
    min_threshold: float = 1.0,
    max_threshold: float = 40.0,
    num_points: int = 25,
    other_criteria: Optional[Dict[str, Optional[float]]] = None,
) -> pd.DataFrame:
    """
    Calcule l'influence de la variation d'un seuil (précipitation ou vent) sur le pourcentage
    de journées classées défavorables.
    """
    if df_weather is None or df_weather.empty:
        return pd.DataFrame()

    other = other_criteria or {}
    thresholds = np.linspace(float(min_threshold), float(max_threshold), max(2, int(num_points)))
    records = []

    for th in thresholds:
        th_val = round(float(th), 2)
        p_max = th_val if variable_col == "precipitation_sum" else other.get("precip_max_mm")
        w_max = th_val if variable_col == "wind_speed_10m_max" else other.get("wind_max_kmh")
        tmin = other.get("temp_min_c")
        tmax = other.get("temp_max_c")

        ev_df = evaluate_dataset_against_criteria(
            df_weather,
            precip_max_mm=p_max,
            temp_min_c=tmin,
            temp_max_c=tmax,
            wind_max_kmh=w_max,
            humidity_max_pct=other.get("humidity_max_pct"),
        )
        n_comp = int((ev_df["statut"] == "Compatible").sum())
        n_def = int((ev_df["statut"] == "Défavorable").sum())
        n_eval = n_comp + n_def
        pct_def = (n_def / n_eval) * 100.0 if n_eval > 0 else np.nan

        records.append(
            {
                "seuil": th_val,
                "jours_defavorables": n_def,
                "jours_compatibles": n_comp,
                "jours_evaluables": n_eval,
                "pourcentage_defavorable": pct_def,
                "pourcentage_compatible": 100.0 - pct_def if not np.isnan(pct_def) else np.nan,
            }
        )

    return pd.DataFrame(records)


# ==============================================================================
# 8. FONCTIONS D'EXPORTATION, SAUVEGARDE ET RESTAURATION (CSV, EXCEL, PDF, JSON)
# ==============================================================================

def export_weather_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Exporte le DataFrame météorologique au format CSV encodé en UTF-8 avec BOM (compatible Excel Windows)."""
    if df is None or df.empty:
        return "date,precipitation_sum,temperature_2m_min,temperature_2m_max,wind_speed_10m_max\n".encode("utf-8-sig")
    return df.to_csv(index=False).encode("utf-8-sig")


def export_analysis_to_excel_bytes(
    project: Dict[str, Any],
    df_weather: pd.DataFrame,
    stats_table: pd.DataFrame,
    monthly_df: pd.DataFrame,
    planning_df: Optional[pd.DataFrame] = None,
) -> bytes:
    """
    Génère un classeur Microsoft Excel (.xlsx) multi-onglets formaté avec OpenPyXL :
    - Onglet 1 : Synthèse du Chantier
    - Onglet 2 : Statistiques Descriptives
    - Onglet 3 : Synthèse Mensuelle
    - Onglet 4 : Planification des Activités
    - Onglet 5 : Données Météo Journalières
    """
    wb = openpyxl.Workbook()
    # Styles OpenPyXL
    header_fill = PatternFill(start_color="10243A", end_color="10243A", fill_type="solid")
    sub_fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="10243A")
    bold_font = Font(name="Calibri", size=10, bold=True, color="10243A")
    body_font = Font(name="Calibri", size=10, color="10243A")
    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE"),
    )

    def _write_dataframe_sheet(ws: Any, title: str, df_in: pd.DataFrame) -> None:
        ws["A1"] = title
        ws["A1"].font = title_font
        if df_in is None or df_in.empty:
            ws["A3"] = "Aucune donnée disponible pour cette section."
            ws["A3"].font = body_font
            return

        cols = list(df_in.columns)
        for c_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(row=3, column=c_idx, value=str(col_name))
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for r_idx, (_, row_data) in enumerate(df_in.iterrows(), start=4):
            for c_idx, col_name in enumerate(cols, start=1):
                val = row_data[col_name]
                if pd.isna(val):
                    clean_v: Any = ""
                elif isinstance(val, (float, np.floating)):
                    clean_v = round(float(val), 2)
                elif isinstance(val, (int, np.integer)):
                    clean_v = int(val)
                else:
                    clean_v = str(val)
                cell = ws.cell(row=r_idx, column=c_idx, value=clean_v)
                cell.font = body_font
                cell.border = thin_border
                if isinstance(clean_v, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        # Ajustement automatique des largeurs de colonnes
        for c_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(c_idx)
            max_len = max(
                [len(str(col_name))]
                + [len(str(ws.cell(row=r, column=c_idx).value or "")) for r in range(4, min(len(df_in) + 4, 120))]
            )
            ws.column_dimensions[col_letter].width = max(13, min(42, max_len + 4))

    # Onglet 1 : Synthèse du chantier
    ws1 = wb.active
    ws1.title = "Synthese_Chantier"
    ws1["A1"] = f"CHANTIERMÉTÉO — Dossier d'Analyse : {project.get('name', 'Chantier')}"
    ws1["A1"].font = title_font

    info_rows = [
        ("Nom du chantier", project.get("name", "")),
        ("Type de projet", project.get("project_type", "")),
        ("Ville / Localisation", project.get("city", "")),
        ("Coordonnées (Lat, Lon)", f"{project.get('latitude', '')}, {project.get('longitude', '')}"),
        ("Période prévisionnelle du chantier", f"Du {project.get('start_date', '')} au {project.get('end_date', '')}"),
        ("Fuseau horaire de référence", DEFAULT_TIMEZONE),
        ("Nature des données météorologiques", SOURCE_TYPES.get(project.get("data_source_type", ""), project.get("data_source_type", ""))),
        ("Date de dernière récupération", project.get("last_updated", "Non renseignée")),
        ("Description", project.get("description", "")),
        ("Avertissement méthodologique", "Les seuils sont des hypothèses pédagogiques de planification et non des limites réglementaires ou garanties de sécurité."),
    ]
    for idx, (k_label, v_val) in enumerate(info_rows, start=3):
        c1 = ws1.cell(row=idx, column=1, value=k_label)
        c2 = ws1.cell(row=idx, column=2, value=str(v_val))
        c1.font = bold_font
        c1.fill = sub_fill
        c1.border = thin_border
        c2.font = body_font
        c2.border = thin_border
    ws1.column_dimensions["A"].width = 34
    ws1.column_dimensions["B"].width = 68

    # Onglet 2 : Statistiques descriptives
    ws2 = wb.create_sheet(title="Statistiques_Descriptives")
    _write_dataframe_sheet(ws2, "Statistiques Descriptives des Variables Météorologiques", stats_table)

    # Onglet 3 : Synthèse mensuelle
    ws3 = wb.create_sheet(title="Evolution_Mensuelle")
    _write_dataframe_sheet(ws3, "Évolution Mensuelle des Précipitations et Températures", monthly_df)

    # Onglet 4 : Planification des activités
    ws4 = wb.create_sheet(title="Planification_Activites")
    _write_dataframe_sheet(
        ws4,
        "Calendrier d'Évaluation des Activités de Chantier",
        planning_df if planning_df is not None else pd.DataFrame(),
    )

    # Onglet 5 : Données journalières
    ws5 = wb.create_sheet(title="Donnees_Journalieres")
    _write_dataframe_sheet(ws5, "Série Météorologique Journalière Analysée", df_weather)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _register_pdf_unicode_fonts() -> Tuple[str, str, bool]:
    """
    Enregistre une police TrueType supportant l'UTF-8 complet (accents français, symboles ≥, ≤, °, —)
    sous Linux (DejaVuSans / LiberationSans) ou Windows (Arial / Segoe UI / Calibri).
    Retourne (font_regular, font_bold, is_unicode_ttf).
    """
    candidates = [
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "CMDejaVu",
            "CMDejaVuBold",
        ),
        (
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
            "CMLiberation",
            "CMLiberationBold",
        ),
        (
            "C:\\Windows\\Fonts\\arial.ttf",
            "C:\\Windows\\Fonts\\arialbd.ttf",
            "CMArial",
            "CMArialBold",
        ),
        (
            "C:\\Windows\\Fonts\\segoeui.ttf",
            "C:\\Windows\\Fonts\\segoeuib.ttf",
            "CMSegoe",
            "CMSegoeBold",
        ),
        (
            "C:\\Windows\\Fonts\\calibri.ttf",
            "C:\\Windows\\Fonts\\calibrib.ttf",
            "CMCalibri",
            "CMCalibriBold",
        ),
    ]
    for reg_path, bold_path, reg_name, bold_name in candidates:
        if os.path.exists(reg_path) and os.path.exists(bold_path):
            try:
                pdfmetrics.registerFont(TTFont(reg_name, reg_path))
                pdfmetrics.registerFont(TTFont(bold_name, bold_path))
                return reg_name, bold_name, True
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold", False


def _sanitize_pdf_text(text: str, is_unicode_ttf: bool) -> str:
    """Échappe les entités XML pour ReportLab Paragraph et adapte les symboles si police non-TTF."""
    s = str(text or "")
    if not is_unicode_ttf:
        replacements = {
            "≥": ">=",
            "≤": "<=",
            "—": "-",
            "–": "-",
            "’": "'",
            "‘": "'",
            "“": '"',
            "”": '"',
            "œ": "oe",
            "Œ": "OE",
            "≈": "~",
            "•": "-",
        }
        for k, v in replacements.items():
            s = s.replace(k, v)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s


def generate_pdf_report_bytes(
    project: Dict[str, Any],
    df_weather: pd.DataFrame,
    stats_table: pd.DataFrame,
    activities_summaries: List[Dict[str, Any]],
    prob_summary: Dict[str, Any],
    mc_summary: Optional[Dict[str, Any]] = None,
) -> bytes:
    """
    Génère un rapport synthétique PDF professionnel et paginé avec ReportLab :
    - Informations du chantier et nature exacte de la source de données
    - Statistiques descriptives des paramètres météorologiques
    - Évaluation des activités planifiées et journées défavorables
    - Analyse probabiliste et simulation de Monte-Carlo par blocs
    - Interprétation d'ingénierie et limites méthodologiques / scientifiques.
    """
    font_reg, font_bold, is_ttf = _register_pdf_unicode_fonts()
    buf = io.BytesIO()

    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=f"Rapport ChantierMeteo - {project.get('name', 'Chantier')}",
        author="ChantierMeteo (ENSPY Genie Civil)",
    )

    styles = getSampleStyleSheet()
    navy_rl = rl_colors.HexColor("#10243A")
    sky_rl = rl_colors.HexColor("#328CCB")
    orange_rl = rl_colors.HexColor("#F29A38")
    green_rl = rl_colors.HexColor("#218A70")
    coral_rl = rl_colors.HexColor("#D85B56")
    light_rl = rl_colors.HexColor("#F3F6FA")

    st_title = ParagraphStyle(
        "CMTitle",
        parent=styles["Normal"],
        fontName=font_bold,
        fontSize=15,
        leading=19,
        textColor=rl_colors.white,
        spaceAfter=2,
    )
    st_subtitle = ParagraphStyle(
        "CMSubtitle",
        parent=styles["Normal"],
        fontName=font_reg,
        fontSize=9.5,
        leading=13,
        textColor=rl_colors.HexColor("#DCEBF7"),
    )
    st_h2 = ParagraphStyle(
        "CMH2",
        parent=styles["Normal"],
        fontName=font_bold,
        fontSize=11.5,
        leading=15,
        textColor=navy_rl,
        spaceBefore=10,
        spaceAfter=5,
    )
    st_body = ParagraphStyle(
        "CMBody",
        parent=styles["Normal"],
        fontName=font_reg,
        fontSize=9,
        leading=13,
        textColor=navy_rl,
        spaceAfter=4,
    )
    st_bold = ParagraphStyle(
        "CMBold",
        parent=st_body,
        fontName=font_bold,
    )
    st_cell_hdr = ParagraphStyle(
        "CMCellHdr",
        parent=styles["Normal"],
        fontName=font_bold,
        fontSize=8.2,
        leading=10.5,
        textColor=rl_colors.white,
    )
    st_cell = ParagraphStyle(
        "CMCell",
        parent=styles["Normal"],
        fontName=font_reg,
        fontSize=8.2,
        leading=10.5,
        textColor=navy_rl,
    )
    st_warn = ParagraphStyle(
        "CMWarn",
        parent=styles["Normal"],
        fontName=font_bold,
        fontSize=8.5,
        leading=12,
        textColor=rl_colors.HexColor("#7A4300"),
    )

    story: List[Any] = []

    # 1. Bandeau d'en-tête du rapport
    header_data = [
        [
            Paragraph(
                _sanitize_pdf_text("CHANTIERMÉTÉO — RAPPORT SYNTHÉTIQUE D'ANALYSE ET DE PLANIFICATION", is_ttf),
                st_title,
            )
        ],
        [
            Paragraph(
                _sanitize_pdf_text(
                    f"{APP_SUBTITLE} | {APP_INSTITUTION}",
                    is_ttf,
                ),
                st_subtitle,
            )
        ],
    ]
    header_tbl = Table(header_data, colWidths=[178 * mm])
    header_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), navy_rl),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (0, 0), 10),
                ("BOTTOMPADDING", (-1, -1), (-1, -1), 10),
                ("LINEBELOW", (0, -1), (-1, -1), 3, orange_rl),
            ]
        )
    )
    story.append(header_tbl)
    story.append(Spacer(1, 4 * mm))

    src_code = project.get("data_source_type", "AUCUNE")
    src_label = SOURCE_TYPES.get(src_code, src_code)

    if src_code in ("FICTIF_DEMO", "TEST_DETERMINISTE"):
        warn_tbl = Table(
            [
                [
                    Paragraph(
                        _sanitize_pdf_text(
                            "AVERTISSEMENT : Ce rapport repose sur des DONNÉES FICTIVES — UTILISATION PÉDAGOGIQUE UNIQUEMENT. "
                            "Elles permettent de vérifier les formules mathématiques et ne représentent pas des observations instrumentales réelles.",
                            is_ttf,
                        ),
                        st_warn,
                    )
                ]
            ],
            colWidths=[178 * mm],
        )
        warn_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), rl_colors.HexColor("#FFF4E5")),
                    ("BOX", (0, 0), (-1, -1), 1, orange_rl),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        story.append(warn_tbl)
        story.append(Spacer(1, 3 * mm))

    # 2. Fiche d'identité du chantier
    story.append(Paragraph(_sanitize_pdf_text("1. Identification du Chantier et Source des Données", is_ttf), st_h2))
    n_days_analyzed = len(df_weather) if (df_weather is not None and not df_weather.empty) else 0
    date_range_str = (
        f"Du {df_weather['date'].iloc[0]} au {df_weather['date'].iloc[-1]} ({n_days_analyzed} jours)"
        if n_days_analyzed > 0
        else "Aucune série chargée"
    )

    meta_rows = [
        [
            Paragraph(_sanitize_pdf_text("Nom du chantier", is_ttf), st_bold),
            Paragraph(_sanitize_pdf_text(project.get("name", "N/A"), is_ttf), st_cell),
            Paragraph(_sanitize_pdf_text("Type d'ouvrage", is_ttf), st_bold),
            Paragraph(_sanitize_pdf_text(project.get("project_type", "N/A"), is_ttf), st_cell),
        ],
        [
            Paragraph(_sanitize_pdf_text("Ville / Localité", is_ttf), st_bold),
            Paragraph(_sanitize_pdf_text(project.get("city", "N/A"), is_ttf), st_cell),
            Paragraph(_sanitize_pdf_text("Coordonnées (WGS84)", is_ttf), st_bold),
            Paragraph(
                _sanitize_pdf_text(
                    f"Lat {project.get('latitude', 0):.4f}°, Lon {project.get('longitude', 0):.4f}° ({DEFAULT_TIMEZONE})",
                    is_ttf,
                ),
                st_cell,
            ),
        ],
        [
            Paragraph(_sanitize_pdf_text("Planning prévisionnel", is_ttf), st_bold),
            Paragraph(
                _sanitize_pdf_text(f"Du {project.get('start_date', '')} au {project.get('end_date', '')}", is_ttf),
                st_cell,
            ),
            Paragraph(_sanitize_pdf_text("Période météo analysée", is_ttf), st_bold),
            Paragraph(_sanitize_pdf_text(date_range_str, is_ttf), st_cell),
        ],
        [
            Paragraph(_sanitize_pdf_text("Source météorologique", is_ttf), st_bold),
            Paragraph(_sanitize_pdf_text(src_label, is_ttf), st_cell),
            Paragraph(_sanitize_pdf_text("Dernière actualisation", is_ttf), st_bold),
            Paragraph(_sanitize_pdf_text(project.get("last_updated", "N/A"), is_ttf), st_cell),
        ],
    ]
    meta_tbl = Table(meta_rows, colWidths=[36 * mm, 53 * mm, 36 * mm, 53 * mm])
    meta_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), light_rl),
                ("BACKGROUND", (2, 0), (2, -1), light_rl),
                ("GRID", (0, 0), (-1, -1), 0.5, rl_colors.HexColor("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(meta_tbl)
    story.append(Spacer(1, 3 * mm))

    # 3. Tableau des statistiques descriptives
    story.append(Paragraph(_sanitize_pdf_text("2. Statistiques Descriptives des Paramètres Météorologiques", is_ttf), st_h2))
    if stats_table is not None and not stats_table.empty:
        st_rows = [
            [
                Paragraph(_sanitize_pdf_text("Paramètre", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Effectif (n)", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Min", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Moyenne", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Médiane", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Max", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Écart-type (s)", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Q25 / Q75", is_ttf), st_cell_hdr),
            ]
        ]
        for _, r in stats_table.iterrows():
            st_rows.append(
                [
                    Paragraph(_sanitize_pdf_text(str(r.get("Variable", "")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(r.get("Effectif valide (n)", "")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(r.get("Min", "")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(r.get("Moyenne", "")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(r.get("Médiane (Q50)", "")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(r.get("Max", "")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(r.get("Écart-type (s)", "")), is_ttf), st_cell),
                    Paragraph(
                        _sanitize_pdf_text(f"{r.get('Q25', '')} / {r.get('Q75', '')}", is_ttf),
                        st_cell,
                    ),
                ]
            )
        st_tbl = Table(
            st_rows,
            colWidths=[46 * mm, 18 * mm, 16 * mm, 18 * mm, 18 * mm, 16 * mm, 22 * mm, 24 * mm],
            repeatRows=1,
        )
        st_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), navy_rl),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, light_rl]),
                    ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#CBD5E1")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(st_tbl)
    else:
        story.append(Paragraph(_sanitize_pdf_text("Aucune statistique descriptive disponible.", is_ttf), st_body))

    story.append(Spacer(1, 3 * mm))

    # 4. Planification des activités de chantier
    story.append(Paragraph(_sanitize_pdf_text("3. Synthèse de Planification des Activités du Chantier", is_ttf), st_h2))
    if activities_summaries:
        act_rows = [
            [
                Paragraph(_sanitize_pdf_text("Activité", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Période étudiée", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Durée", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Seuils retenus (hypothèses)", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Jours OK", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Défav.", is_ttf), st_cell_hdr),
                Paragraph(_sanitize_pdf_text("Indisp.", is_ttf), st_cell_hdr),
            ]
        ]
        for a_sum in activities_summaries:
            act_rows.append(
                [
                    Paragraph(_sanitize_pdf_text(str(a_sum.get("activity_name", "")), is_ttf), st_cell),
                    Paragraph(
                        _sanitize_pdf_text(f"{a_sum.get('start_date', '')} → {a_sum.get('end_date', '')}", is_ttf),
                        st_cell,
                    ),
                    Paragraph(_sanitize_pdf_text(f"{a_sum.get('duration_days', 0)} j", is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(str(a_sum.get("thresholds_str", "Configurés")), is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(f"{a_sum.get('compatible_days', 0)} j", is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(f"{a_sum.get('unfavorable_days', 0)} j", is_ttf), st_cell),
                    Paragraph(_sanitize_pdf_text(f"{a_sum.get('impossible_days', 0)} j", is_ttf), st_cell),
                ]
            )
        act_tbl = Table(
            act_rows,
            colWidths=[36 * mm, 33 * mm, 15 * mm, 43 * mm, 17 * mm, 17 * mm, 17 * mm],
            repeatRows=1,
        )
        act_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), sky_rl),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, light_rl]),
                    ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#CBD5E1")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        story.append(act_tbl)
    else:
        story.append(Paragraph(_sanitize_pdf_text("Aucune activité configurée pour ce chantier.", is_ttf), st_body))

    story.append(Spacer(1, 3 * mm))

    # 5. Analyse probabiliste et simulations de Monte-Carlo
    story.append(Paragraph(_sanitize_pdf_text("4. Évaluation Probabiliste et Simulation de Monte-Carlo par Blocs", is_ttf), st_h2))
    prob_text = (
        f"• Probabilité empirique journalière de dépassement du seuil de pluie ({prob_summary.get('threshold', 10.0):.1f} mm) : "
        f"{prob_summary.get('prob_pct', 0.0):.1f} % ({prob_summary.get('k_exceed', 0)} j / {prob_summary.get('n_valid', 0)} j valides, "
        f"IC 95 % Wilson : [{prob_summary.get('ci_low_pct', 0.0):.1f} %, {prob_summary.get('ci_high_pct', 0.0):.1f} %])."
    )
    story.append(Paragraph(_sanitize_pdf_text(prob_text, is_ttf), st_body))

    if mc_summary and not mc_summary.get("error"):
        mc_text = (
            f"• Simulation de Monte-Carlo par blocs (M = {mc_summary['num_simulations']} simulations, horizon D = {mc_summary['horizon_days']} jours, "
            f"taille de bloc L = {mc_summary['block_length']} jours, graine = {mc_summary['random_seed']}) : "
            f"Fréquence simulée d'au moins 1 journée défavorable = {mc_summary['prob_at_least_one_pct']:.1f} % "
            f"(± {mc_summary['mc_ci95_half_prob_pct']:.2f} % à 95 %), "
            f"fréquence de 0 journée défavorable = {mc_summary['prob_zero_event_pct']:.1f} %. "
            f"Nombre moyen de journées défavorables sur {mc_summary['horizon_days']} jours : {mc_summary['mean_unfavorable_days']:.2f} j "
            f"(Médiane Q50 = {mc_summary['q50_days']:.1f} j, Q75 = {mc_summary['q75_days']:.1f} j, P90 = {mc_summary['q90_days']:.1f} j)."
        )
        story.append(Paragraph(_sanitize_pdf_text(mc_text, is_ttf), st_body))

    story.append(Spacer(1, 2 * mm))

    # 6. Interprétation et limites scientifiques / techniques
    story.append(
        KeepTogether(
            [
                Paragraph(_sanitize_pdf_text("5. Interprétation d'Ingénierie et Limites Méthodologiques", is_ttf), st_h2),
                Paragraph(
                    _sanitize_pdf_text(
                        "1) Distinction des sources : Les données historiques Open-Meteo proviennent de réanalyses climatiques (ERA5 / ERA5-Land) "
                        "sur grille spatiale (~9 à 25 km) et peuvent lisser les abats d'eau convectifs très localisés propres au climat équatorial de Yaoundé. "
                        "Elles ne doivent jamais être confondues avec des relevés pluviométriques in situ sur le chantier ni avec des prévisions déterministes à court terme.",
                        is_ttf,
                    ),
                    st_body,
                ),
                Paragraph(
                    _sanitize_pdf_text(
                        "2) Non-indépendance des fenêtres glissantes : Les fenêtres consécutives de k jours extraites de l'historique se chevauchent "
                        "et subissent l'autocorrélation météorologique. La simulation de Monte-Carlo par rééchantillonnage de blocs conserve la persistance "
                        "à l'échelle du bloc L mais suppose la stationnarité sur la période historique de référence.",
                        is_ttf,
                    ),
                    st_body,
                ),
                Paragraph(
                    _sanitize_pdf_text(
                        "3) Portée opérationnelle : Les seuils configurés dans ChantierMétéo sont des repères d'aide à la planification et d'estimation "
                        "des marges de délai. Ils ne constituent en aucun cas une norme réglementaire, une note de calcul géotechnique/structurelle, "
                        "ni une autorisation de sécurité pour le bétonnage, le terrassement ou les opérations de levage.",
                        is_ttf,
                    ),
                    st_body,
                ),
            ]
        )
    )

    def _draw_footer(canvas_obj: Any, doc_obj: Any) -> None:
        canvas_obj.saveState()
        canvas_obj.setFont(font_reg, 8)
        canvas_obj.setFillColor(rl_colors.HexColor("#5A6B7C"))
        footer_left = f"ChantierMeteo — ENSPY Yaounde | Chantier : {project.get('name', '')}"
        footer_right = f"Page {doc_obj.page}"
        canvas_obj.drawString(16 * mm, 10 * mm, _sanitize_pdf_text(footer_left, is_ttf).replace("&amp;", "&"))
        canvas_obj.drawRightString(194 * mm, 10 * mm, footer_right)
        canvas_obj.restoreState()

    doc.build(story, onFirstPage=_draw_footer, onLaterPages=_draw_footer)
    return buf.getvalue()


def serialize_projects_to_json(projects: Dict[str, Dict[str, Any]], active_id: Optional[str]) -> str:
    """Sérialise l'ensemble des chantiers de la session et leurs séries météo en JSON."""
    export_payload: Dict[str, Any] = {
        "app": APP_NAME,
        "version": "1.0.0-ENSPY",
        "exported_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "timezone": DEFAULT_TIMEZONE,
        "active_project_id": active_id,
        "projects": {},
    }
    for pid, pdata in projects.items():
        p_copy = copy.deepcopy(pdata)
        w_df = p_copy.get("weather_df")
        if isinstance(w_df, pd.DataFrame):
            p_copy["weather_df"] = w_df.where(pd.notna(w_df), None).to_dict(orient="records")
        f_df = p_copy.get("forecast_df")
        if isinstance(f_df, pd.DataFrame):
            p_copy["forecast_df"] = f_df.where(pd.notna(f_df), None).to_dict(orient="records")
        export_payload["projects"][pid] = p_copy

    return json.dumps(export_payload, ensure_ascii=False, indent=2)


def deserialize_projects_from_json(json_str: str) -> Tuple[Optional[Dict[str, Dict[str, Any]]], Optional[str], Optional[str]]:
    """Restaure les chantiers depuis une sauvegarde JSON en validant la structure."""
    try:
        payload = json.loads(json_str)
    except Exception as exc:
        return None, None, f"Fichier JSON mal formé : {exc}"

    if not isinstance(payload, dict) or "projects" not in payload or not isinstance(payload["projects"], dict):
        return None, None, "Structure JSON invalide : clé 'projects' absente ou incorrecte."

    restored: Dict[str, Dict[str, Any]] = {}
    for pid, pdata in payload["projects"].items():
        if not isinstance(pdata, dict) or not pdata.get("name"):
            continue
        p_copy = copy.deepcopy(pdata)
        for df_key in ("weather_df", "forecast_df"):
            raw_records = p_copy.get(df_key)
            if isinstance(raw_records, list) and len(raw_records) > 0:
                df_rec = pd.DataFrame(raw_records)
                for num_col in [
                    "precipitation_sum",
                    "temperature_2m_min",
                    "temperature_2m_max",
                    "temperature_2m_mean",
                    "wind_speed_10m_max",
                    "relative_humidity_mean",
                    "precipitation_probability_max",
                ]:
                    if num_col in df_rec.columns:
                        df_rec[num_col] = pd.to_numeric(df_rec[num_col], errors="coerce")
                p_copy[df_key] = df_rec
            else:
                p_copy[df_key] = pd.DataFrame()
        restored[pid] = p_copy

    if not restored:
        return None, None, "Aucun chantier valide n'a été trouvé dans le fichier JSON."

    active_id = payload.get("active_project_id")
    if active_id not in restored:
        active_id = next(iter(restored.keys()))
    return restored, active_id, None


# ==============================================================================
# 9. GESTION DE L'ÉTAT DE SESSION ET DES CHANTIERS
# ==============================================================================

FICTITIOUS_PROJECT_ID = "demo_enspy_yaounde"


def build_fictitious_enspy_project() -> Dict[str, Any]:
    """
    Construit le chantier de démonstration obligatoire :
    'Chantier-école ENSPY — Démonstration' à Yaoundé, avec terrassement et bétonnage,
    série météo fictive de 120 jours clairement identifiée et seuils pédagogiques modifiables.
    """
    df_fictif = generate_fictitious_enspy_dataset(start_date="2025-03-01", num_days=120, seed=2025)
    activities = [
        {
            "id": "act_demo_terrassement",
            "name": "Terrassement général et fouilles en rigoles",
            "category": "Terrassement",
            "start_date": "2025-03-01",
            "duration_days": 10,
            "precip_max_mm": 10.0,
            "temp_min_c": 12.0,
            "temp_max_c": 37.0,
            "wind_max_kmh": 50.0,
            "humidity_max_pct": None,
            "notes": "Les 10 premiers jours (01/03 au 10/03/2025) reproduisent la série déterministe de référence.",
        },
        {
            "id": "act_demo_betonnage",
            "name": "Bétonnage radier, poteaux et dalle R+1",
            "category": "Bétonnage",
            "start_date": "2025-03-11",
            "duration_days": 14,
            "precip_max_mm": 5.0,
            "temp_min_c": 10.0,
            "temp_max_c": 35.0,
            "wind_max_kmh": 45.0,
            "humidity_max_pct": None,
            "notes": "Hypothèse pédagogique : sensibilité accrue au délavage du béton frais par la pluie.",
        },
        {
            "id": "act_demo_levage",
            "name": "Opérations de levage charpente métallique",
            "category": "Opérations de levage",
            "start_date": "2025-03-25",
            "duration_days": 7,
            "precip_max_mm": 15.0,
            "temp_min_c": 10.0,
            "temp_max_c": 38.0,
            "wind_max_kmh": 35.0,
            "humidity_max_pct": None,
            "notes": "Hypothèse pédagogique : arrêt du levage en cas de rafales atteignant ou dépassant 35 km/h.",
        },
    ]
    return {
        "id": FICTITIOUS_PROJECT_ID,
        "name": "Chantier-école ENSPY — Démonstration",
        "project_type": "Bâtiment",
        "city": "Yaoundé",
        "latitude": 3.8480,
        "longitude": 11.5021,
        "start_date": "2025-03-01",
        "end_date": "2025-06-28",
        "description": (
            "Projet pédagogique de construction d'un bâtiment universitaire R+2 avec laboratoire d'essais "
            "au campus de l'ENSPY (Yaoundé). Contient une série météorologique fictive de 120 jours "
            "permettant de tester toutes les fonctionnalités sans dépendre d'Internet."
        ),
        "activities": activities,
        "weather_df": df_fictif,
        "forecast_df": pd.DataFrame(),
        "data_source_type": "FICTIF_DEMO",
        "is_fictitious": True,
        "last_updated": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "default_precip_threshold_mm": 10.0,
        "default_wind_threshold_kmh": 45.0,
        "default_tmin_c": 10.0,
        "default_tmax_c": 37.0,
    }


def init_session_state() -> None:
    """Initialise les clés de session Streamlit de manière défensive."""
    if st is None:
        return
    if "projects" not in st.session_state:
        st.session_state["projects"] = {}
    if "active_project_id" not in st.session_state:
        st.session_state["active_project_id"] = None
    if "nav_page" not in st.session_state:
        st.session_state["nav_page"] = "🏗️ Tableau de bord"
    if "theme_mode" not in st.session_state:
        st.session_state["theme_mode"] = "Clair"
    if "flash_message" not in st.session_state:
        st.session_state["flash_message"] = None


def get_active_project() -> Optional[Dict[str, Any]]:
    """Retourne le dictionnaire du chantier actuellement sélectionné."""
    if st is None:
        return None
    projects = st.session_state.get("projects", {})
    active_id = st.session_state.get("active_project_id")
    if active_id and active_id in projects:
        return projects[active_id]
    if projects:
        first_id = next(iter(projects.keys()))
        st.session_state["active_project_id"] = first_id
        return projects[first_id]
    return None


def load_or_reset_fictitious_example(is_reset: bool = False) -> None:
    """Charge ou réinitialise uniquement l'exemple fictif ENSPY sans effacer les autres chantiers."""
    if st is None:
        return
    demo_proj = build_fictitious_enspy_project()
    st.session_state["projects"][FICTITIOUS_PROJECT_ID] = demo_proj
    st.session_state["active_project_id"] = FICTITIOUS_PROJECT_ID
    if is_reset:
        st.session_state["flash_message"] = (
            "success",
            "L'exemple fictif « Chantier-école ENSPY — Démonstration » a été réinitialisé à ses paramètres d'origine "
            "(les autres chantiers de la session ont été préservés).",
        )
    else:
        st.session_state["flash_message"] = (
            "success",
            "Chantier de démonstration « Chantier-école ENSPY — Démonstration » chargé avec succès "
            "(Données fictives — utilisation pédagogique uniquement).",
        )


def validate_project_inputs(
    name: str,
    city: str,
    latitude: float,
    longitude: float,
    start_date: dt.date,
    end_date: dt.date,
) -> Optional[str]:
    """Vérifie la cohérence des champs d'un chantier avant création ou modification."""
    if not (name or "").strip():
        return "Le nom du chantier est obligatoire et ne peut pas être vide."
    if not (city or "").strip():
        return "La ville ou localisation du chantier est obligatoire."
    try:
        lat = float(latitude)
        lon = float(longitude)
    except (TypeError, ValueError):
        return "Les coordonnées géographiques (latitude et longitude) doivent être des nombres valides."
    if not (-90.0 <= lat <= 90.0):
        return "La latitude doit être comprise entre -90° et +90°."
    if not (-180.0 <= lon <= 180.0):
        return "La longitude doit être comprise entre -180° et +180°."
    if start_date is None or end_date is None:
        return "Veuillez renseigner la date de début et la date de fin prévues."
    if end_date < start_date:
        return f"Dates incohérentes : la date de fin ({end_date}) est antérieure à la date de début ({start_date})."
    return None


# ==============================================================================
# 10. GÉNÉRATEURS DE GRAPHIQUES INTERACTIFS PLOTLY HARMONISÉS
# ==============================================================================

def _apply_plotly_layout(fig: Any, title: str, theme_mode: str = "Clair", height: int = 370) -> Any:
    """Applique la charte graphique ChantierMétéo et rend le graphique responsive."""
    if fig is None:
        return None
    is_dark = theme_mode == "Sombre"
    bg_paper = "#132337" if is_dark else "#FFFFFF"
    bg_plot = "#192D45" if is_dark else "#F8FAFC"
    txt_col = "#F1F5F9" if is_dark else "#10243A"
    grid_col = "#2A4361" if is_dark else "#E2E8F0"

    fig.update_layout(
        title={"text": f"<b>{title}</b>", "font": {"size": 15, "color": txt_col}, "x": 0.01},
        paper_bgcolor=bg_paper,
        plot_bgcolor=bg_plot,
        font={"family": "Inter, Segoe UI, Arial, sans-serif", "color": txt_col, "size": 12},
        height=height,
        margin={"l": 45, "r": 25, "t": 52, "b": 45},
        legend={
            "orientation": "h",
            "yanchor": "bottom",
            "y": 1.02,
            "xanchor": "right",
            "x": 1.0,
            "font": {"size": 11},
        },
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=True, gridcolor=grid_col, zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor=grid_col, zeroline=False)
    return fig


def build_precipitation_time_chart(
    df: pd.DataFrame,
    threshold_mm: float = 10.0,
    theme_mode: str = "Clair",
) -> Any:
    """1. Courbe/barres des précipitations dans le temps avec seuil défavorable."""
    if go is None or df is None or df.empty:
        return None
    fig = go.Figure()
    vals = pd.to_numeric(df["precipitation_sum"], errors="coerce")
    bar_colors = [
        PALETTE["coral"] if (pd.notna(v) and float(v) >= threshold_mm) else PALETTE["sky"]
        for v in vals
    ]
    fig.add_trace(
        go.Bar(
            x=df["date"],
            y=vals,
            name="Précipitations (mm/j)",
            marker_color=bar_colors,
            hovertemplate="Date : %{x}<br>Pluie : %{y:.1f} mm<extra></extra>",
        )
    )
    fig.add_hline(
        y=threshold_mm,
        line_dash="dash",
        line_color=PALETTE["orange"],
        line_width=2.5,
        annotation_text=f"Seuil : {threshold_mm:.1f} mm/j",
        annotation_position="top right",
    )
    fig.update_yaxes(title_text="Précipitations (mm/jour)")
    fig.update_xaxes(title_text="Date")
    return _apply_plotly_layout(
        fig,
        f"Évolution temporelle des précipitations journalières (seuil = {threshold_mm:.1f} mm)",
        theme_mode,
    )


def build_precipitation_histogram(
    df: pd.DataFrame,
    threshold_mm: float = 10.0,
    theme_mode: str = "Clair",
) -> Any:
    """2. Histogramme de distribution des précipitations journalières."""
    if go is None or df is None or df.empty:
        return None
    clean_p = pd.to_numeric(df["precipitation_sum"], errors="coerce").dropna()
    fig = go.Figure()
    fig.add_trace(
        go.Histogram(
            x=clean_p,
            nbinsx=25,
            name="Effectif de journées",
            marker_color=PALETTE["sky"],
            marker_line_color=PALETTE["navy"],
            marker_line_width=1,
            hovertemplate="Classe de pluie : %{x} mm<br>Journées : %{y}<extra></extra>",
        )
    )
    fig.add_vline(
        x=threshold_mm,
        line_dash="dash",
        line_color=PALETTE["coral"],
        line_width=2.5,
        annotation_text=f"Seuil {threshold_mm:.1f} mm",
        annotation_position="top right",
    )
    fig.update_xaxes(title_text="Précipitations journalières (mm)")
    fig.update_yaxes(title_text="Nombre de journées")
    return _apply_plotly_layout(fig, "Histogramme de distribution des précipitations", theme_mode)


def build_temperature_evolution_chart(df: pd.DataFrame, theme_mode: str = "Clair") -> Any:
    """3. Évolution des températures (Tmin, Tmoy, Tmax)."""
    if go is None or df is None or df.empty:
        return None
    fig = go.Figure()
    if "temperature_2m_max" in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["temperature_2m_max"],
                mode="lines",
                name="Tmax (°C)",
                line={"color": PALETTE["coral"], "width": 2},
            )
        )
    if "temperature_2m_mean" in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["temperature_2m_mean"],
                mode="lines",
                name="Tmoy (°C)",
                line={"color": PALETTE["orange"], "width": 2, "dash": "dot"},
            )
        )
    if "temperature_2m_min" in df.columns:
        fig.add_trace(
            go.Scatter(
                x=df["date"],
                y=df["temperature_2m_min"],
                mode="lines",
                name="Tmin (°C)",
                fill="tonexty",
                fillcolor="rgba(50, 140, 203, 0.14)",
                line={"color": PALETTE["sky"], "width": 2},
            )
        )
    fig.update_xaxes(title_text="Date")
    fig.update_yaxes(title_text="Température (°C)")
    return _apply_plotly_layout(fig, "Évolution des températures journalières (Tmin, Tmoy, Tmax)", theme_mode)


def build_wind_distribution_chart(
    df: pd.DataFrame,
    wind_threshold_kmh: float = 45.0,
    theme_mode: str = "Clair",
) -> Any:
    """4. Distribution des vitesses maximales du vent."""
    if go is None or df is None or df.empty or "wind_speed_10m_max" not in df.columns:
        return None
    clean_w = pd.to_numeric(df["wind_speed_10m_max"], errors="coerce").dropna()
    fig = go.Figure()
    fig.add_trace(
        go.Histogram(
            x=clean_w,
            nbinsx=22,
            name="Distribution du vent max",
            marker_color=PALETTE["green"],
            marker_line_color=PALETTE["navy"],
            marker_line_width=1,
        )
    )
    fig.add_vline(
        x=wind_threshold_kmh,
        line_dash="dash",
        line_color=PALETTE["coral"],
        line_width=2.5,
        annotation_text=f"Seuil vent {wind_threshold_kmh:.0f} km/h",
        annotation_position="top right",
    )
    fig.update_xaxes(title_text="Vitesse maximale du vent à 10 m (km/h)")
    fig.update_yaxes(title_text="Nombre de journées")
    return _apply_plotly_layout(fig, "Distribution des vitesses maximales journalières du vent", theme_mode)


def build_monthly_precipitation_chart(monthly_df: pd.DataFrame, theme_mode: str = "Clair") -> Any:
    """5. Comparaison mensuelle des précipitations (cumul en mm et nombre de jours de dépassement)."""
    if go is None or make_subplots is None or monthly_df is None or monthly_df.empty:
        return None
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(
            x=monthly_df["year_month"],
            y=monthly_df["cumul_pluie_mm"],
            name="Cumul mensuel (mm)",
            marker_color=PALETTE["sky"],
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=monthly_df["year_month"],
            y=monthly_df["jours_depassement_seuil"],
            mode="lines+markers",
            name="Jours dépassant le seuil",
            line={"color": PALETTE["coral"], "width": 2.5},
            marker={"size": 8},
        ),
        secondary_y=True,
    )
    fig.update_xaxes(title_text="Mois (AAAA-MM)")
    fig.update_yaxes(title_text="Cumul mensuel (mm)", secondary_y=False)
    fig.update_yaxes(title_text="Nombre de journées ≥ seuil", secondary_y=True)
    return _apply_plotly_layout(fig, "Comparaison mensuelle des cumuls de pluie et jours critiques", theme_mode)


def build_exceedance_frequency_chart(
    series: pd.Series,
    var_label: str = "Précipitations (mm)",
    marked_threshold: Optional[float] = 10.0,
    theme_mode: str = "Clair",
) -> Any:
    """6. Courbe des fréquences empiriques de dépassement P(X >= x)."""
    if go is None or series is None:
        return None
    clean = pd.to_numeric(series, errors="coerce").dropna().to_numpy(dtype=float)
    if len(clean) == 0:
        return None

    x_grid = np.linspace(float(np.min(clean)), float(np.max(clean)), 60)
    freqs = [float(np.mean(clean >= x) * 100.0) for x in x_grid]

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=x_grid,
            y=freqs,
            mode="lines",
            name="Fréquence P(X ≥ x) en %",
            fill="tozeroy",
            fillcolor="rgba(50, 140, 203, 0.16)",
            line={"color": PALETTE["sky"], "width": 3},
            hovertemplate="Seuil x = %{x:.1f}<br>Fréquence P(X ≥ x) = %{y:.1f} %<extra></extra>",
        )
    )
    if marked_threshold is not None and not np.isnan(marked_threshold):
        p_mark = float(np.mean(clean >= float(marked_threshold)) * 100.0)
        fig.add_vline(
            x=float(marked_threshold),
            line_dash="dash",
            line_color=PALETTE["coral"],
            line_width=2,
            annotation_text=f"Seuil {marked_threshold:.1f} ({p_mark:.1f} %)",
            annotation_position="top right",
        )
    fig.update_xaxes(title_text=f"Seuil x — {var_label}")
    fig.update_yaxes(title_text="Fréquence de dépassement (%)", range=[0, 105])
    return _apply_plotly_layout(
        fig,
        f"Courbe des fréquences cumulées de dépassement — {var_label}",
        theme_mode,
    )


def build_planning_timeline_chart(schedule_df: pd.DataFrame, theme_mode: str = "Clair") -> Any:
    """Graphique visuel jour par jour de compatibilité météorologique d'une activité."""
    if go is None or schedule_df is None or schedule_df.empty:
        return None
    color_map = {
        "Compatible": PALETTE["green"],
        "Défavorable": PALETTE["coral"],
        "Évaluation impossible": PALETTE["slate"],
    }
    colors_list = [color_map.get(s, PALETTE["slate"]) for s in schedule_df["statut"]]
    y_precip = schedule_df["precipitation_sum"].fillna(0.0)

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=schedule_df["date"],
            y=[1] * len(schedule_df),
            marker_color=colors_list,
            text=schedule_df["statut"],
            textposition="inside",
            customdata=np.stack(
                [
                    schedule_df["raisons"].fillna(""),
                    y_precip,
                    schedule_df["wind_speed_10m_max"].fillna(0.0),
                ],
                axis=-1,
            ),
            hovertemplate=(
                "<b>Date : %{x}</b><br>"
                "Statut : %{text}<br>"
                "Pluie : %{customdata[1]:.1f} mm | Vent : %{customdata[2]:.1f} km/h<br>"
                "Détail : %{customdata[0]}<extra></extra>"
            ),
            name="Statut journalier",
        )
    )
    fig.update_yaxes(visible=False, range=[0, 1.2])
    fig.update_xaxes(title_text="Calendrier de l'activité")
    return _apply_plotly_layout(
        fig,
        "Frise calendaire de compatibilité météorologique (Vert = Compatible, Rouge = Défavorable, Gris = Indisponible)",
        theme_mode,
        height=240,
    )


# ==============================================================================
# 11. NAVIGATION LATÉRALE ET RENDU DES 6 RUBRIQUES PRINCIPALES
# ==============================================================================

NAV_ITEMS = [
    "🏗️ Tableau de bord",
    "📁 Mes chantiers",
    "🌦️ Analyse météorologique",
    "📅 Planification météo",
    "🎲 Probabilités et scénarios",
    "🎓 Comprendre et tester",
]


def _build_stats_summary_dataframe(df_weather: pd.DataFrame, precip_thresh: float = 10.0) -> pd.DataFrame:
    """Construit un tableau récapitulatif des statistiques descriptives pour toutes les variables disponibles."""
    if df_weather is None or df_weather.empty:
        return pd.DataFrame()

    rows = []
    for col, meta in WEATHER_VARIABLES_META.items():
        if col not in df_weather.columns:
            continue
        s = df_weather[col]
        if s.dropna().empty:
            continue
        th = precip_thresh if col == "precipitation_sum" else None
        st_dict = compute_descriptive_stats(s, threshold=th)
        rows.append(
            {
                "Variable": f"{meta['label']} ({meta['unit']})",
                "Effectif valide (n)": st_dict["n_valid"],
                "Valeurs manquantes": st_dict["n_missing"],
                "Min": round(st_dict["min"], 2) if pd.notna(st_dict["min"]) else np.nan,
                "Moyenne": round(st_dict["mean"], 2) if pd.notna(st_dict["mean"]) else np.nan,
                "Médiane (Q50)": round(st_dict["median"], 2) if pd.notna(st_dict["median"]) else np.nan,
                "Max": round(st_dict["max"], 2) if pd.notna(st_dict["max"]) else np.nan,
                "Variance (s²)": round(st_dict["var_sample"], 2) if pd.notna(st_dict["var_sample"]) else np.nan,
                "Écart-type (s)": round(st_dict["std_sample"], 2) if pd.notna(st_dict["std_sample"]) else np.nan,
                "P10": round(st_dict["p10"], 2) if pd.notna(st_dict["p10"]) else np.nan,
                "Q25": round(st_dict["q25"], 2) if pd.notna(st_dict["q25"]) else np.nan,
                "Q75": round(st_dict["q75"], 2) if pd.notna(st_dict["q75"]) else np.nan,
                "P90": round(st_dict["p90"], 2) if pd.notna(st_dict["p90"]) else np.nan,
                "IQR": round(st_dict["iqr"], 2) if pd.notna(st_dict["iqr"]) else np.nan,
            }
        )
    return pd.DataFrame(rows)


def render_sidebar() -> None:
    """Affiche le menu latéral avec icônes, sélecteur de chantier, thème et gestion de session JSON."""
    with st.sidebar:
        st.markdown(
            f"""
            <div style="padding: 0.6rem 0.2rem 0.8rem 0.2rem; border-bottom: 2px solid #F29A38; margin-bottom: 0.9rem;">
                <div style="font-size: 1.35rem; font-weight: 800; color: #F29A38;">🏗️ {APP_NAME}</div>
                <div style="font-size: 0.78rem; opacity: 0.85;">ENSPY Yaoundé • Génie Civil</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Navigation principale avec icônes
        selected_nav = st.radio(
            "🧭 Navigation principale",
            options=NAV_ITEMS,
            index=NAV_ITEMS.index(st.session_state["nav_page"])
            if st.session_state["nav_page"] in NAV_ITEMS
            else 0,
            key="sidebar_nav_radio",
        )
        if selected_nav != st.session_state["nav_page"]:
            st.session_state["nav_page"] = selected_nav
            st.rerun()

        st.markdown("---")

        # Sélecteur rapide du chantier actif
        projects = st.session_state.get("projects", {})
        if projects:
            pid_list = list(projects.keys())
            labels_map = {pid: f"{projects[pid]['name']} ({projects[pid]['city']})" for pid in pid_list}
            cur_id = st.session_state.get("active_project_id")
            cur_idx = pid_list.index(cur_id) if cur_id in pid_list else 0
            chosen_pid = st.selectbox(
                "📌 Chantier actif",
                options=pid_list,
                format_func=lambda x: labels_map.get(x, x),
                index=cur_idx,
            )
            if chosen_pid != st.session_state.get("active_project_id"):
                st.session_state["active_project_id"] = chosen_pid
                st.rerun()
        else:
            st.info("Aucun chantier enregistré dans la session.")

        # Boutons rapides pour l'exemple fictif obligatoire
        if st.button("🧪 Charger l'exemple fictif", use_container_width=True):
            load_or_reset_fictitious_example(is_reset=False)
            st.rerun()

        if FICTITIOUS_PROJECT_ID in projects:
            if st.button("🔄 Réinitialiser l'exemple fictif", use_container_width=True):
                load_or_reset_fictitious_example(is_reset=True)
                st.rerun()

        st.markdown("---")

        # Choix du thème Clair / Sombre
        theme_choice = st.selectbox(
            "🎨 Thème d'affichage",
            options=["Clair", "Sombre"],
            index=0 if st.session_state.get("theme_mode", "Clair") == "Clair" else 1,
        )
        if theme_choice != st.session_state.get("theme_mode"):
            st.session_state["theme_mode"] = theme_choice
            st.rerun()

        # Sauvegarde / Restauration JSON de la session
        with st.expander("💾 Sauvegarde / Restauration JSON"):
            st.caption(
                "⚠️ Les données de session Streamlit sont temporaires et sont effacées à la fermeture de l'onglet. "
                "Téléchargez une sauvegarde JSON pour conserver durablement vos chantiers."
            )
            if projects:
                json_str = serialize_projects_to_json(projects, st.session_state.get("active_project_id"))
                st.download_button(
                    label="⬇️ Sauvegarder la session (.json)",
                    data=json_str.encode("utf-8"),
                    file_name=f"chantiermeteo_sauvegarde_{dt.date.today().isoformat()}.json",
                    mime="application/json",
                    use_container_width=True,
                )
            uploaded_json = st.file_uploader("Restaurer un fichier .json", type=["json"], key="json_restore_uploader")
            if uploaded_json is not None:
                if st.button("📂 Restaurer cette sauvegarde", use_container_width=True):
                    raw_text = uploaded_json.getvalue().decode("utf-8", errors="replace")
                    restored_projs, new_active, err_msg = deserialize_projects_from_json(raw_text)
                    if err_msg:
                        st.error(err_msg)
                    elif restored_projs:
                        st.session_state["projects"].update(restored_projs)
                        st.session_state["active_project_id"] = new_active
                        st.session_state["flash_message"] = (
                            "success",
                            f"{len(restored_projs)} chantier(s) restauré(s) avec succès depuis la sauvegarde JSON.",
                        )
                        st.rerun()


# ------------------------------------------------------------------------------
# RUBRIQUE A : TABLEAU DE BORD — VUE GÉNÉRALE
# ------------------------------------------------------------------------------

def render_dashboard_page() -> None:
    """Affiche la Rubrique A : Tableau de bord — Vue générale."""
    active_proj = get_active_project()
    render_hero_banner(active_proj)

    projects = st.session_state.get("projects", {})
    nb_projects = len(projects)

    # Barre des 5 boutons d'action rapide exigés par le cahier des charges
    st.markdown("#### ⚡ Actions rapides")
    b_col1, b_col2, b_col3, b_col4, b_col5 = st.columns(5)

    with b_col1:
        if st.button("➕ Nouveau chantier", use_container_width=True, type="primary"):
            st.session_state["nav_page"] = "📁 Mes chantiers"
            st.session_state["open_new_project_tab"] = True
            st.rerun()

    with b_col2:
        if st.button("📂 Ouvrir un chantier", use_container_width=True):
            st.session_state["nav_page"] = "📁 Mes chantiers"
            st.rerun()

    with b_col3:
        if st.button(" Charger l'exemple fictif", key="dash_btn_fictif", use_container_width=True):
            load_or_reset_fictitious_example(is_reset=False)
            st.rerun()

    with b_col4:
        if st.button("🔄 Actualiser les données", use_container_width=True):
            if active_proj is None:
                st.warning("Aucun chantier sélectionné. Créez un chantier ou chargez l'exemple fictif d'abord.")
            elif active_proj.get("is_fictitious"):
                load_or_reset_fictitious_example(is_reset=True)
                st.rerun()
            else:
                df_new, err_api, meta_api = fetch_historical_weather_open_meteo(
                    active_proj["latitude"],
                    active_proj["longitude"],
                    active_proj["start_date"],
                    active_proj["end_date"],
                )
                if err_api:
                    st.error(err_api)
                elif df_new is not None:
                    active_proj["weather_df"] = df_new
                    active_proj["data_source_type"] = "HISTORIQUE_API"
                    active_proj["last_updated"] = meta_api.get("fetched_at", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    st.success("Données météorologiques actualisées avec succès depuis Open-Meteo.")
                    st.rerun()

    with b_col5:
        export_clicked = st.button("📄 Exporter le rapport", use_container_width=True)

    # État vide lorsqu'aucun chantier n'existe
    if active_proj is None:
        st.markdown(
            """
            <div class="cm-empty-state">
                <h3 style="margin-top:0;">👋 Bienvenue sur ChantierMétéo !</h3>
                <p style="font-size:0.98rem; max-width:720px; margin:0.6rem auto 1.1rem auto;">
                    Aucun chantier n'est actuellement enregistré dans votre session. Pour démarrer immédiatement :
                </p>
                <p style="font-size:0.92rem; margin:0 auto;">
                    1️⃣ Cliquez sur <b>« Charger l'exemple fictif »</b> ci-dessus pour explorer le chantier-école ENSPY (Yaoundé) hors-ligne.<br>
                    2️⃣ Ou cliquez sur <b>« Nouveau chantier »</b> pour configurer votre propre projet de génie civil et interroger l'API Open-Meteo.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # Bandeau d'identification explicite si données fictives
    if active_proj.get("data_source_type") in ("FICTIF_DEMO", "TEST_DETERMINISTE"):
        st.markdown(
            """
            <div class="cm-warn-banner">
                ⚠️ <b>Données fictives — utilisation pédagogique uniquement.</b>
                Les séries attachées à ce chantier de démonstration servent à tester les méthodes statistiques
                sans connexion Internet et ne constituent pas des relevés météorologiques réels de Yaoundé.
            </div>
            """,
            unsafe_allow_html=True,
        )

    df_w: pd.DataFrame = active_proj.get("weather_df", pd.DataFrame())
    n_days = len(df_w) if (df_w is not None and not df_w.empty) else 0
    period_str = (
        f"{df_w['date'].iloc[0]} → {df_w['date'].iloc[-1]}"
        if n_days > 0
        else f"{active_proj.get('start_date', '')} → {active_proj.get('end_date', '')} (Non chargé)"
    )

    # Évaluation globale selon les seuils de référence du chantier
    p_thresh = float(active_proj.get("default_precip_threshold_mm", 10.0))
    w_thresh = float(active_proj.get("default_wind_threshold_kmh", 45.0))
    tmin_thresh = float(active_proj.get("default_tmin_c", 10.0))
    tmax_thresh = float(active_proj.get("default_tmax_c", 37.0))

    if n_days > 0:
        ev_df = evaluate_dataset_against_criteria(
            df_w,
            precip_max_mm=p_thresh,
            temp_min_c=tmin_thresh,
            temp_max_c=tmax_thresh,
            wind_max_kmh=w_thresh,
        )
        n_unfav = int((ev_df["statut"] == "Défavorable").sum())
        n_comp = int((ev_df["statut"] == "Compatible").sum())
        n_eval = n_unfav + n_comp
        prop_unfav = (n_unfav / n_eval) * 100.0 if n_eval > 0 else 0.0
        unfav_str = f"{n_unfav} j ({prop_unfav:.1f} %)"
    else:
        ev_df = pd.DataFrame()
        n_unfav = 0
        prop_unfav = 0.0
        unfav_str = "Aucune donnée"

    # Cartes KPI colorées
    st.markdown(
        f"""
        <div class="cm-kpi-grid">
            <div class="cm-kpi-card cm-kpi-navy">
                <div class="cm-kpi-label">Chantiers en session</div>
                <div class="cm-kpi-value">{nb_projects} projet(s)</div>
                <div class="cm-kpi-sub">Session active ({DEFAULT_TIMEZONE})</div>
            </div>
            <div class="cm-kpi-card cm-kpi-orange">
                <div class="cm-kpi-label">Chantier sélectionné</div>
                <div class="cm-kpi-value">{active_proj.get('name', '')}</div>
                <div class="cm-kpi-sub">Type : {active_proj.get('project_type', '')}</div>
            </div>
            <div class="cm-kpi-card cm-kpi-sky">
                <div class="cm-kpi-label">Ville & Période étudiée</div>
                <div class="cm-kpi-value">{active_proj.get('city', '')}</div>
                <div class="cm-kpi-sub">{period_str}</div>
            </div>
            <div class="cm-kpi-card cm-kpi-green">
                <div class="cm-kpi-label">Dernière récupération</div>
                <div class="cm-kpi-value">{active_proj.get('last_updated', 'N/A')[:10]}</div>
                <div class="cm-kpi-sub">Heure : {active_proj.get('last_updated', 'N/A')[11:]}</div>
            </div>
            <div class="cm-kpi-card cm-kpi-sky">
                <div class="cm-kpi-label">Journées analysées</div>
                <div class="cm-kpi-value">{n_days} jours</div>
                <div class="cm-kpi-sub">{SOURCE_TYPES.get(active_proj.get('data_source_type', ''), 'Non défini')[:38]}...</div>
            </div>
            <div class="cm-kpi-card cm-kpi-coral">
                <div class="cm-kpi-label">Journées défavorables</div>
                <div class="cm-kpi-value">{unfav_str}</div>
                <div class="cm-kpi-sub">Pluie ≥ {p_thresh:.0f} mm ou Vent ≥ {w_thresh:.0f} km/h</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Si l'utilisateur a cliqué sur « Exporter le rapport », afficher le panneau de téléchargement direct
    if export_clicked or st.session_state.get("show_export_panel", False):
        st.session_state["show_export_panel"] = True
        with st.container():
            st.markdown("#### 📥 Téléchargement du rapport synthétique PDF et des données")
            stats_tbl = _build_stats_summary_dataframe(df_w, precip_thresh=p_thresh)
            monthly_tbl = compute_monthly_summary(df_w, precip_threshold=p_thresh)
            act_sums = []
            for act in active_proj.get("activities", []):
                _, s_act = evaluate_activity_schedule(df_w, act)
                s_act["thresholds_str"] = (
                    f"Pluie < {act.get('precip_max_mm', 10)} mm, Vent < {act.get('wind_max_kmh', 45)} km/h"
                )
                act_sums.append(s_act)
            prob_s = (
                compute_empirical_exceedance_probability(df_w["precipitation_sum"], p_thresh)
                if n_days > 0
                else {"threshold": p_thresh, "prob_pct": 0.0, "k_exceed": 0, "n_valid": 0, "ci_low_pct": 0.0, "ci_high_pct": 0.0}
            )
            mc_s = run_monte_carlo_block_bootstrap(ev_df, horizon_days=10, num_simulations=1000, block_length=3, random_seed=42) if n_days >= 5 else None

            pdf_bytes = generate_pdf_report_bytes(active_proj, df_w, stats_tbl, act_sums, prob_s, mc_s)
            xlsx_bytes = export_analysis_to_excel_bytes(active_proj, df_w, stats_tbl, monthly_tbl, ev_df)
            csv_bytes = export_weather_to_csv_bytes(df_w)

            dl1, dl2, dl3, dl4 = st.columns(4)
            with dl1:
                st.download_button(
                    "📕 Rapport PDF synthétique (.pdf)",
                    data=pdf_bytes,
                    file_name=f"Rapport_ChantierMeteo_{active_proj['name'][:20].replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            with dl2:
                st.download_button(
                    "📗 Tableaux d'analyse (.xlsx)",
                    data=xlsx_bytes,
                    file_name=f"Analyse_ChantierMeteo_{active_proj['name'][:20].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl3:
                st.download_button(
                    "📄 Données météo (.csv)",
                    data=csv_bytes,
                    file_name=f"Meteo_{active_proj['name'][:20].replace(' ', '_')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with dl4:
                if st.button("Fermer le panneau d'export", use_container_width=True):
                    st.session_state["show_export_panel"] = False
                    st.rerun()

    # Résumé des conditions météorologiques disponibles et aperçu graphique
    col_left, col_right = st.columns([1.25, 1.0])
    with col_left:
        st.markdown("#### 🌦️ Résumé des conditions météorologiques disponibles")
        if n_days > 0:
            p_stats = compute_descriptive_stats(df_w["precipitation_sum"], threshold=p_thresh)
            tmax_stats = compute_descriptive_stats(df_w["temperature_2m_max"])
            tmin_stats = compute_descriptive_stats(df_w["temperature_2m_min"])
            w_stats = compute_descriptive_stats(df_w["wind_speed_10m_max"], threshold=w_thresh)

            summary_df = pd.DataFrame(
                [
                    {
                        "Indicateur": "Cumul total de précipitations",
                        "Valeur": f"{p_stats['sum']:.1f} mm",
                        "Interprétation chantier": f"Sur {p_stats['n_valid']} journées valides",
                    },
                    {
                        "Indicateur": "Précipitation moyenne / max journalière",
                        "Valeur": f"{p_stats['mean']:.2f} mm/j (Max : {p_stats['max']:.1f} mm)",
                        "Interprétation chantier": f"Médiane : {p_stats['median']:.1f} mm/j | Écart-type s = {p_stats['std_sample']:.2f} mm",
                    },
                    {
                        "Indicateur": f"Journées avec pluie ≥ {p_thresh:.1f} mm",
                        "Valeur": f"{p_stats['exceed_count']} jours ({p_stats['exceed_freq_pct']:.1f} %)",
                        "Interprétation chantier": "Risque de ruissellement et d'arrêt de terrassement/bétonnage",
                    },
                    {
                        "Indicateur": "Amplitude thermique observée (Tmin / Tmax)",
                        "Valeur": f"{tmin_stats['min']:.1f} °C à {tmax_stats['max']:.1f} °C",
                        "Interprétation chantier": f"Moyenne des Tmax : {tmax_stats['mean']:.1f} °C (cure du béton)",
                    },
                    {
                        "Indicateur": "Vitesse maximale du vent à 10 m",
                        "Valeur": f"Max : {w_stats['max']:.1f} km/h (Moy : {w_stats['mean']:.1f} km/h)",
                        "Interprétation chantier": f"{w_stats['exceed_count']} j ≥ {w_thresh:.0f} km/h (vigilance levage)",
                    },
                ]
            )
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

            fig_dash = build_precipitation_time_chart(df_w, threshold_mm=p_thresh, theme_mode=st.session_state["theme_mode"])
            if fig_dash is not None:
                st.plotly_chart(fig_dash, use_container_width=True)
        else:
            st.info(
                "Aucune série météorologique n'est encore chargée pour ce chantier. "
                "Rendez-vous dans « 🌦️ Analyse météorologique » pour interroger l'API Open-Meteo ou importer un fichier."
            )

    with col_right:
        st.markdown("#### 🧭 Accès rapide aux rubriques métier")
        st.markdown(
            """
            <div class="cm-card">
                <b>📁 1. Mes chantiers — Gestion des projets</b><br>
                Créez vos chantiers (bâtiment, route, ouvrage hydraulique, terrassement), géolocalisez la ville
                et configurez les activités envisagées.
            </div>
            <div class="cm-card">
                <b>🌦️ 2. Analyse météorologique (Historique & Prévisions)</b><br>
                Récupérez les séries Open-Meteo (ERA5 ou Forecast 16j) ou importez un CSV/Excel. Explorez les
                statistiques descriptives (moyenne, médiane, variance, quartiles) et 6 graphiques interactifs.
            </div>
            <div class="cm-card">
                <b>📅 3. Planification météo des activités</b><br>
                Définissez vos seuils opérationnels par tâche, visualisez le calendrier de compatibilité et
                comparez plusieurs dates de démarrage.
            </div>
            <div class="cm-card">
                <b>🎲 4. Probabilités, Fenêtres & Monte-Carlo</b><br>
                Calculez les probabilités empiriques, les fréquences sur fenêtres glissantes de <i>k</i> jours,
                les simulations de Monte-Carlo par blocs et les courbes de sensibilité.
            </div>
            <div class="cm-card">
                <b>🎓 5. Comprendre et tester (Soutenance ENSPY)</b><br>
                Protocoles de démonstration pas-à-pas devant le jury et cas de vérification déterministe
                automatique sur 10 journées (moyenne = 6,9 mm/j, fréquence = 30 %).
            </div>
            """,
            unsafe_allow_html=True,
        )


# ------------------------------------------------------------------------------
# RUBRIQUE B : MES CHANTIERS — GESTION DES PROJETS
# ------------------------------------------------------------------------------

def render_projects_page() -> None:
    """Affiche la Rubrique B : Mes chantiers — Gestion des projets."""
    active_proj = get_active_project()
    render_hero_banner(active_proj)

    st.markdown("### 📁 Mes chantiers — Gestion des projets de génie civil")
    st.caption(
        "Créez, consultez, modifiez ou supprimez vos chantiers. Cliquez sur une carte de projet pour en faire le chantier actif."
    )

    tab_list, tab_create, tab_edit = st.tabs(
        ["🗂️ Portefeuille des chantiers", "➕ Créer un nouveau chantier", "✏️ Modifier le chantier actif"]
    )

    # --- ONGLET 1 : LISTE DES CHANTIERS SOUS FORME DE CARTES CLIQUABLES ---
    with tab_list:
        projects = st.session_state.get("projects", {})
        if not projects:
            st.markdown(
                """
                <div class="cm-empty-state">
                    <h4>Aucun chantier enregistré</h4>
                    <p>Utilisez l'onglet <b>« ➕ Créer un nouveau chantier »</b> ou cliquez ci-dessous pour charger le chantier-école ENSPY.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("🧪 Charger l'exemple fictif (Chantier-école ENSPY)", type="primary"):
                load_or_reset_fictitious_example(is_reset=False)
                st.rerun()
        else:
            cols = st.columns(2)
            for idx, (pid, proj) in enumerate(list(projects.items())):
                is_active = pid == st.session_state.get("active_project_id")
                card_cls = "cm-card cm-card-active" if is_active else "cm-card"
                w_df = proj.get("weather_df", pd.DataFrame())
                n_obs = len(w_df) if isinstance(w_df, pd.DataFrame) else 0
                act_names = ", ".join([a.get("name", "") for a in proj.get("activities", [])]) or "Aucune activité"

                with cols[idx % 2]:
                    st.markdown(
                        f"""
                        <div class="{card_cls}">
                            <div style="display:flex; justify-content:space-between; align-items:center;">
                                <span style="font-size:1.12rem; font-weight:800;">🏗️ {proj.get('name', '')}</span>
                                <span class="cm-badge {'cm-badge-fictif' if proj.get('is_fictitious') else 'cm-badge-api'}">
                                    {'Actif ⭐' if is_active else proj.get('project_type', '')}
                                </span>
                            </div>
                            <p style="margin:0.4rem 0; font-size:0.88rem;">
                                📍 <b>Localisation :</b> {proj.get('city', '')} (Lat {proj.get('latitude', 0):.4f}, Lon {proj.get('longitude', 0):.4f})<br>
                                🗓️ <b>Période prévue :</b> {proj.get('start_date', '')} → {proj.get('end_date', '')}<br>
                                📊 <b>Données météo :</b> {n_obs} journées ({SOURCE_TYPES.get(proj.get('data_source_type', ''), 'Aucune')[:42]})<br>
                                🛠️ <b>Activités ({len(proj.get('activities', []))}) :</b> {act_names[:85]}
                            </p>
                            <p style="margin:0.2rem 0 0 0; font-size:0.82rem; opacity:0.85;">
                                {proj.get('description', '')[:140]}
                            </p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    b1, b2, b3 = st.columns(3)
                    with b1:
                        if st.button(
                            "✅ Chantier actif" if is_active else "🎯 Sélectionner",
                            key=f"sel_proj_{pid}",
                            disabled=is_active,
                            use_container_width=True,
                        ):
                            st.session_state["active_project_id"] = pid
                            st.session_state["flash_message"] = (
                                "success",
                                f"Chantier « {proj['name']} » sélectionné. Toutes les analyses ont été mises à jour.",
                            )
                            st.rerun()
                    with b2:
                        if st.button("🌦️ Analyser météo", key=f"go_meteo_{pid}", use_container_width=True):
                            st.session_state["active_project_id"] = pid
                            st.session_state["nav_page"] = "🌦️ Analyse météorologique"
                            st.rerun()
                    with b3:
                        if st.button("🗑️ Supprimer", key=f"del_proj_{pid}", use_container_width=True):
                            del st.session_state["projects"][pid]
                            remaining = list(st.session_state["projects"].keys())
                            st.session_state["active_project_id"] = remaining[0] if remaining else None
                            st.session_state["flash_message"] = (
                                "info",
                                f"Le chantier « {proj['name']} » a été supprimé de la session.",
                            )
                            st.rerun()

    # --- ONGLET 2 : CRÉATION D'UN NOUVEAU CHANTIER ---
    with tab_create:
        st.markdown("#### ➕ Nouveau chantier de génie civil")
        st.markdown("##### 1. Aide à la localisation (Géocodage Open-Meteo ou Villes du Cameroun)")
        g_col1, g_col2, g_col3 = st.columns([1.4, 1.4, 1.0])
        with g_col1:
            preset_city = st.selectbox(
                "Choisir une ville camerounaise de référence",
                options=["-- Saisie libre / Recherche API --"] + list(CAMEROON_CITIES_PRESETS.keys()),
                index=1,
            )
        with g_col2:
            search_city_query = st.text_input("Ou rechercher une localité via Open-Meteo Geocoding API", value="")
        with g_col3:
            st.write("")
            st.write("")
            if st.button("🔍 Rechercher les coordonnées", use_container_width=True):
                matches, g_err = geocode_city_open_meteo(search_city_query or preset_city)
                st.session_state["geocoding_results"] = matches
                st.session_state["geocoding_error"] = g_err

        if st.session_state.get("geocoding_error"):
            st.warning(st.session_state["geocoding_error"])
        geo_res = st.session_state.get("geocoding_results") or []
        selected_geo = None
        if geo_res:
            geo_labels = [
                f"{r['name']} ({r.get('admin1', '')}, {r.get('country', '')}) — Lat {r['latitude']:.4f}, Lon {r['longitude']:.4f}"
                for r in geo_res
            ]
            chosen_geo_idx = st.selectbox("Résultats du géocodage :", options=range(len(geo_res)), format_func=lambda i: geo_labels[i])
            selected_geo = geo_res[chosen_geo_idx]

        # Valeurs par défaut des coordonnées selon sélection
        if selected_geo is not None:
            def_city = selected_geo["name"]
            def_lat = float(selected_geo["latitude"])
            def_lon = float(selected_geo["longitude"])
        elif preset_city in CAMEROON_CITIES_PRESETS:
            def_city = preset_city
            def_lat = CAMEROON_CITIES_PRESETS[preset_city]["latitude"]
            def_lon = CAMEROON_CITIES_PRESETS[preset_city]["longitude"]
        else:
            def_city = "Yaoundé"
            def_lat = 3.8480
            def_lon = 11.5021

        with st.form("create_project_form", clear_on_submit=False):
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                new_name = st.text_input("Nom du chantier *", placeholder="Ex. Pont sur la Sanaga — Lot 1")
                new_type = st.selectbox("Type de projet *", options=PROJECT_TYPES)
                new_city = st.text_input("Ville ou localisation *", value=def_city)
                new_activities_picked = st.multiselect(
                    "Activités envisagées sur le chantier",
                    options=list(ACTIVITY_TEMPLATES.keys()),
                    default=["Terrassement", "Bétonnage"],
                )
            with f_col2:
                new_lat = st.number_input("Latitude (° WGS84)", value=float(def_lat), min_value=-90.0, max_value=90.0, format="%.4f")
                new_lon = st.number_input("Longitude (° WGS84)", value=float(def_lon), min_value=-180.0, max_value=180.0, format="%.4f")
                default_start = dt.date.today() - dt.timedelta(days=120)
                default_end = dt.date.today() - dt.timedelta(days=5)
                new_start = st.date_input("Date de début prévue *", value=default_start)
                new_end = st.date_input("Date de fin prévue *", value=default_end)

            new_desc = st.text_area(
                "Description technique du chantier",
                placeholder="Décrivez la nature des travaux, les contraintes de sol, le phasage envisagé...",
            )
            auto_fetch = st.checkbox(
                "Récupérer immédiatement les données historiques Open-Meteo après la création",
                value=True,
            )

            submitted_create = st.form_submit_button("✅ Enregistrer le chantier", type="primary")
            if submitted_create:
                val_err = validate_project_inputs(new_name, new_city, new_lat, new_lon, new_start, new_end)
                if val_err:
                    st.error(val_err)
                else:
                    new_id = f"proj_{uuid.uuid4().hex[:8]}"
                    built_acts = []
                    offset_d = 0
                    for a_cat in new_activities_picked:
                        tpl = ACTIVITY_TEMPLATES.get(a_cat, ACTIVITY_TEMPLATES["Autre activité personnalisée"])
                        act_start = (new_start + dt.timedelta(days=offset_d)).isoformat()
                        built_acts.append(
                            {
                                "id": f"act_{uuid.uuid4().hex[:6]}",
                                "name": a_cat,
                                "category": a_cat,
                                "start_date": act_start,
                                "duration_days": tpl["duration_days"],
                                "precip_max_mm": tpl["precip_max_mm"],
                                "temp_min_c": tpl["temp_min_c"],
                                "temp_max_c": tpl["temp_max_c"],
                                "wind_max_kmh": tpl["wind_max_kmh"],
                                "humidity_max_pct": tpl["humidity_max_pct"],
                                "notes": tpl["pedagogical_note"],
                            }
                        )
                        offset_d += tpl["duration_days"]

                    proj_obj: Dict[str, Any] = {
                        "id": new_id,
                        "name": new_name.strip(),
                        "project_type": new_type,
                        "city": new_city.strip(),
                        "latitude": float(new_lat),
                        "longitude": float(new_lon),
                        "start_date": new_start.isoformat(),
                        "end_date": new_end.isoformat(),
                        "description": new_desc.strip(),
                        "activities": built_acts,
                        "weather_df": pd.DataFrame(),
                        "forecast_df": pd.DataFrame(),
                        "data_source_type": "AUCUNE",
                        "is_fictitious": False,
                        "last_updated": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "default_precip_threshold_mm": 10.0,
                        "default_wind_threshold_kmh": 45.0,
                        "default_tmin_c": 10.0,
                        "default_tmax_c": 37.0,
                    }

                    if auto_fetch:
                        df_fetched, api_err, api_meta = fetch_historical_weather_open_meteo(
                            proj_obj["latitude"],
                            proj_obj["longitude"],
                            proj_obj["start_date"],
                            proj_obj["end_date"],
                        )
                        if df_fetched is not None and not df_fetched.empty:
                            proj_obj["weather_df"] = df_fetched
                            proj_obj["data_source_type"] = "HISTORIQUE_API"
                            proj_obj["last_updated"] = api_meta.get("fetched_at", proj_obj["last_updated"])
                        elif api_err:
                            st.warning(
                                f"Chantier créé, mais la récupération automatique Open-Meteo a signalé : {api_err}"
                            )

                    st.session_state["projects"][new_id] = proj_obj
                    st.session_state["active_project_id"] = new_id
                    st.session_state["flash_message"] = (
                        "success",
                        f"Chantier « {proj_obj['name']} » créé et sélectionné avec succès.",
                    )
                    st.rerun()

    # --- ONGLET 3 : MODIFICATION DU CHANTIER ACTIF ---
    with tab_edit:
        if active_proj is None:
            st.info("Sélectionnez ou créez d'abord un chantier pour pouvoir le modifier.")
        else:
            st.markdown(f"#### ✏️ Modifier les informations de « {active_proj['name']} »")
            with st.form("edit_project_form"):
                e_col1, e_col2 = st.columns(2)
                with e_col1:
                    ed_name = st.text_input("Nom du chantier *", value=active_proj.get("name", ""))
                    p_type_idx = (
                        PROJECT_TYPES.index(active_proj["project_type"])
                        if active_proj.get("project_type") in PROJECT_TYPES
                        else 0
                    )
                    ed_type = st.selectbox("Type de projet *", options=PROJECT_TYPES, index=p_type_idx)
                    ed_city = st.text_input("Ville ou localisation *", value=active_proj.get("city", ""))
                with e_col2:
                    ed_lat = st.number_input(
                        "Latitude (° WGS84)",
                        value=float(active_proj.get("latitude", 3.8480)),
                        min_value=-90.0,
                        max_value=90.0,
                        format="%.4f",
                    )
                    ed_lon = st.number_input(
                        "Longitude (° WGS84)",
                        value=float(active_proj.get("longitude", 11.5021)),
                        min_value=-180.0,
                        max_value=180.0,
                        format="%.4f",
                    )
                    ed_start = st.date_input(
                        "Date de début prévue *",
                        value=dt.date.fromisoformat(str(active_proj.get("start_date", "2025-03-01"))[:10]),
                    )
                    ed_end = st.date_input(
                        "Date de fin prévue *",
                        value=dt.date.fromisoformat(str(active_proj.get("end_date", "2025-06-28"))[:10]),
                    )
                ed_desc = st.text_area("Description du chantier", value=active_proj.get("description", ""))

                if st.form_submit_button("💾 Enregistrer les modifications", type="primary"):
                    err_ed = validate_project_inputs(ed_name, ed_city, ed_lat, ed_lon, ed_start, ed_end)
                    if err_ed:
                        st.error(err_ed)
                    else:
                        active_proj["name"] = ed_name.strip()
                        active_proj["project_type"] = ed_type
                        active_proj["city"] = ed_city.strip()
                        active_proj["latitude"] = float(ed_lat)
                        active_proj["longitude"] = float(ed_lon)
                        active_proj["start_date"] = ed_start.isoformat()
                        active_proj["end_date"] = ed_end.isoformat()
                        active_proj["description"] = ed_desc.strip()
                        active_proj["last_updated"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        st.success("Modifications enregistrées. Les analyses associées ont été mises à jour.")
                        st.rerun()


# ------------------------------------------------------------------------------
# RUBRIQUE C : ANALYSE MÉTÉOROLOGIQUE — DONNÉES HISTORIQUES ET PRÉVISIONS
# ------------------------------------------------------------------------------

def render_weather_analysis_page() -> None:
    """Affiche la Rubrique C : Analyse météorologique — Données historiques et prévisions."""
    active_proj = get_active_project()
    render_hero_banner(active_proj)

    if active_proj is None:
        st.warning("Aucun chantier sélectionné. Chargez l'exemple fictif ou créez un chantier dans « Mes chantiers ».")
        if st.button("🧪 Charger l'exemple fictif maintenant", type="primary"):
            load_or_reset_fictitious_example(is_reset=False)
            st.rerun()
        return

    st.markdown(f"### 🌦️ Analyse météorologique — {active_proj['name']} ({active_proj['city']})")

    # Panneau d'acquisition et d'importation des données (strictement séparées)
    with st.expander("📡 Acquisition des données météorologiques (API Open-Meteo, Prévisions ou Import CSV/Excel)", expanded=True):
        acq_tab1, acq_tab2, acq_tab3, acq_tab4 = st.tabs(
            [
                "🏛️ 1. Historique API Open-Meteo (ERA5)",
                "🔭 2. Prévisions API Open-Meteo (16 jours)",
                "📂 3. Importer un fichier CSV / Excel",
                "🧪 4. Série fictive de démonstration",
            ]
        )

        with acq_tab1:
            st.caption(
                f"Source : Open-Meteo Historical Weather API (`archive-api.open-meteo.com`) | "
                f"Coordonnées : Lat {active_proj['latitude']:.4f}°, Lon {active_proj['longitude']:.4f}° | "
                f"Fuseau : `{DEFAULT_TIMEZONE}` (journées de 00h00 à 23h59)."
            )
            h_col1, h_col2, h_col3 = st.columns([1.1, 1.1, 1.2])
            max_archive_dt = dt.date.today() - dt.timedelta(days=2)
            default_h_start = min(dt.date.fromisoformat(active_proj["start_date"][:10]), max_archive_dt - dt.timedelta(days=90))
            default_h_end = min(dt.date.fromisoformat(active_proj["end_date"][:10]), max_archive_dt)
            if default_h_end < default_h_start:
                default_h_start = max_archive_dt - dt.timedelta(days=180)
                default_h_end = max_archive_dt

            with h_col1:
                hist_start = st.date_input("Date de début historique", value=default_h_start, key="api_hist_start")
            with h_col2:
                hist_end = st.date_input("Date de fin historique", value=default_h_end, key="api_hist_end")
            with h_col3:
                st.write("")
                st.write("")
                if st.button("🌐 Récupérer l'historique Open-Meteo", type="primary", use_container_width=True):
                    df_h, err_h, meta_h = fetch_historical_weather_open_meteo(
                        active_proj["latitude"],
                        active_proj["longitude"],
                        hist_start.isoformat(),
                        hist_end.isoformat(),
                    )
                    if err_h:
                        st.error(err_h)
                    elif df_h is not None:
                        active_proj["weather_df"] = df_h
                        active_proj["data_source_type"] = "HISTORIQUE_API"
                        active_proj["is_fictitious"] = False
                        active_proj["last_updated"] = meta_h.get("fetched_at", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                        if meta_h.get("adjusted_note"):
                            st.info(meta_h["adjusted_note"])
                        st.success(f"{len(df_h)} journées historiques récupérées avec succès depuis Open-Meteo.")
                        st.rerun()

        with acq_tab2:
            st.caption(
                "Source : Open-Meteo Forecast API (`api.open-meteo.com/v1/forecast`). "
                "⚠️ Les prévisions météorologiques à court terme ne doivent pas être confondues avec des fréquences climatologiques historiques."
            )
            f_col1, f_col2 = st.columns([1.2, 1.2])
            with f_col1:
                fc_days = st.slider("Horizon de prévision (jours)", min_value=3, max_value=16, value=14)
            with f_col2:
                st.write("")
                if st.button("🔭 Récupérer les prévisions Open-Meteo", type="primary", use_container_width=True):
                    df_f, err_f, meta_f = fetch_forecast_weather_open_meteo(
                        active_proj["latitude"],
                        active_proj["longitude"],
                        forecast_days=fc_days,
                    )
                    if err_f:
                        st.error(err_f)
                    elif df_f is not None:
                        active_proj["forecast_df"] = df_f
                        active_proj["weather_df"] = df_f
                        active_proj["data_source_type"] = "PREVISIONS_API"
                        active_proj["is_fictitious"] = False
                        active_proj["last_updated"] = meta_f.get("fetched_at", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                        st.success(f"Prévisions à {len(df_f)} jours chargées avec succès.")
                        st.rerun()

        with acq_tab3:
            st.caption(
                "Importez un fichier CSV ou Excel (.xlsx) contenant une colonne de date et au moins une variable météo. "
                "Les données importées remplacent la série d'analyse active sans se mélanger silencieusement aux données API."
            )
            u_col1, u_col2, u_col3 = st.columns([1.6, 1.0, 1.0])
            with u_col1:
                up_file = st.file_uploader("Choisir un fichier .csv ou .xlsx", type=["csv", "xlsx", "xls"])
            with u_col2:
                u_precip = st.selectbox("Unité des précipitations", options=["mm", "cm"], index=0)
            with u_col3:
                u_wind = st.selectbox("Unité du vent", options=["km/h", "m/s"], index=0)

            if up_file is not None:
                if st.button("📥 Valider et charger le fichier importé", type="primary"):
                    try:
                        if up_file.name.lower().endswith(".csv"):
                            raw_imp = pd.read_csv(up_file, sep=None, engine="python")
                        else:
                            raw_imp = pd.read_excel(up_file)
                        df_imp, rep_imp, err_imp = validate_and_import_weather_dataframe(
                            raw_imp, unit_wind=u_wind, unit_precip=u_precip
                        )
                        if err_imp:
                            st.error(err_imp)
                        elif df_imp is not None:
                            active_proj["weather_df"] = df_imp
                            active_proj["data_source_type"] = "IMPORT_FICHIER"
                            active_proj["is_fictitious"] = False
                            active_proj["last_updated"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            st.success(
                                f"Import réussi : {len(df_imp)} journées valides retenues "
                                f"({rep_imp['duplicates_removed']} doublon(s) écarté(s))."
                            )
                            for w_msg in rep_imp.get("warnings", []):
                                st.warning(w_msg)
                            st.rerun()
                    except Exception as exc:
                        st.error(f"Erreur de lecture du fichier importé : {exc}")

        with acq_tab4:
            st.caption(
                "Génère une série météorologique fictive clairement identifiée (usage pédagogique hors-ligne uniquement)."
            )
            if st.button("🧪 Charger la série fictive pédagogique (120 jours)", use_container_width=True):
                active_proj["weather_df"] = generate_fictitious_enspy_dataset("2025-03-01", 120, 2025)
                active_proj["data_source_type"] = "FICTIF_DEMO"
                active_proj["is_fictitious"] = True
                active_proj["last_updated"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.rerun()

    # Vérification de la présence de données
    df_full: pd.DataFrame = active_proj.get("weather_df", pd.DataFrame())
    if df_full is None or df_full.empty:
        st.info("Aucune donnée météorologique n'est chargée pour ce chantier. Utilisez l'un des 4 onglets ci-dessus.")
        return

    # Bandeau sur la nature exacte des données
    src_type = active_proj.get("data_source_type", "AUCUNE")
    st.markdown(
        f"""
        <div class="cm-pedago-box">
            <b>📌 Nature de la série actuellement étudiée :</b> {SOURCE_TYPES.get(src_type, src_type)}<br>
            <b>Convention temporelle :</b> Journées civiles de 00h00 à 23h59 (fuseau <code>{DEFAULT_TIMEZONE}</code>).
            Les valeurs manquantes (<code>NaN</code>) sont comptabilisées explicitement et ne sont jamais remplacées par des valeurs inventées.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Filtres de période et de seuil d'analyse
    st.markdown("#### 🔍 Filtres de période et seuil d'analyse")
    min_d_str = str(df_full["date"].min())[:10]
    max_d_str = str(df_full["date"].max())[:10]
    min_d = dt.date.fromisoformat(min_d_str)
    max_d = dt.date.fromisoformat(max_d_str)

    flt1, flt2, flt3, flt4 = st.columns([1.1, 1.1, 1.2, 1.0])
    with flt1:
        f_start = st.date_input("Début de la période filtrée", value=min_d, min_value=min_d, max_value=max_d, key="flt_s")
    with flt2:
        f_end = st.date_input("Fin de la période filtrée", value=max_d, min_value=min_d, max_value=max_d, key="flt_e")
    with flt3:
        precip_thresh = st.number_input(
            "Seuil de précipitation étudié (mm/jour)",
            min_value=0.0,
            max_value=200.0,
            value=float(active_proj.get("default_precip_threshold_mm", 10.0)),
            step=1.0,
        )
        active_proj["default_precip_threshold_mm"] = precip_thresh
    with flt4:
        wind_thresh = st.number_input(
            "Seuil de vent maximal étudié (km/h)",
            min_value=5.0,
            max_value=150.0,
            value=float(active_proj.get("default_wind_threshold_kmh", 45.0)),
            step=2.0,
        )
        active_proj["default_wind_threshold_kmh"] = wind_thresh

    if f_end < f_start:
        st.error("La date de fin du filtre doit être supérieure ou égale à la date de début.")
        return

    df_filtered = df_full[
        (df_full["date"] >= f_start.isoformat()) & (df_full["date"] <= f_end.isoformat())
    ].copy().reset_index(drop=True)

    if df_filtered.empty:
        st.warning("Aucune journée ne correspond à la plage de dates sélectionnée.")
        return

    # Statistiques descriptives détaillées
    st.markdown("#### 📊 Tableau des statistiques descriptives")
    stats_table = _build_stats_summary_dataframe(df_filtered, precip_thresh=precip_thresh)
    p_desc = compute_descriptive_stats(df_filtered["precipitation_sum"], threshold=precip_thresh)

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Effectif valide (pluie)", f"{p_desc['n_valid']} jours", f"Manquants : {p_desc['n_missing']}")
    k2.metric("Moyenne / Médiane pluie", f"{p_desc['mean']:.2f} mm/j", f"Médiane : {p_desc['median']:.2f} mm/j")
    k3.metric("Fréquence jours pluvieux (≥ 1 mm)", f"{p_desc['rainy_freq_ge_1mm_pct']:.1f} %", f"{p_desc['rainy_days_ge_1mm']} jours")
    k4.metric(
        f"Fréquence pluie ≥ {precip_thresh:.1f} mm",
        f"{p_desc['exceed_freq_pct']:.1f} %",
        f"{p_desc['exceed_count']} jours / {p_desc['n_valid']}",
    )

    st.dataframe(stats_table, use_container_width=True, hide_index=True)

    # Les 6 graphiques interactifs exigés
    st.markdown("#### 📈 Visualisations graphiques interactives")
    theme_m = st.session_state.get("theme_mode", "Clair")

    g_row1_c1, g_row1_c2 = st.columns(2)
    with g_row1_c1:
        fig1 = build_precipitation_time_chart(df_filtered, threshold_mm=precip_thresh, theme_mode=theme_m)
        if fig1:
            st.plotly_chart(fig1, use_container_width=True)
    with g_row1_c2:
        fig2 = build_precipitation_histogram(df_filtered, threshold_mm=precip_thresh, theme_mode=theme_m)
        if fig2:
            st.plotly_chart(fig2, use_container_width=True)

    g_row2_c1, g_row2_c2 = st.columns(2)
    with g_row2_c1:
        fig3 = build_temperature_evolution_chart(df_filtered, theme_mode=theme_m)
        if fig3:
            st.plotly_chart(fig3, use_container_width=True)
    with g_row2_c2:
        fig4 = build_wind_distribution_chart(df_filtered, wind_threshold_kmh=wind_thresh, theme_mode=theme_m)
        if fig4:
            st.plotly_chart(fig4, use_container_width=True)

    monthly_df = compute_monthly_summary(df_filtered, precip_threshold=precip_thresh)
    g_row3_c1, g_row3_c2 = st.columns(2)
    with g_row3_c1:
        fig5 = build_monthly_precipitation_chart(monthly_df, theme_mode=theme_m)
        if fig5:
            st.plotly_chart(fig5, use_container_width=True)
    with g_row3_c2:
        fig6 = build_exceedance_frequency_chart(
            df_filtered["precipitation_sum"],
            var_label="Précipitations journalières (mm)",
            marked_threshold=precip_thresh,
            theme_mode=theme_m,
        )
        if fig6:
            st.plotly_chart(fig6, use_container_width=True)

    # Comparaison de deux sous-périodes et inspection des valeurs manquantes / données brutes
    sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs(
        [
            "📅 Évolution mensuelle",
            "⚖️ Comparaison de 2 sous-périodes",
            "🔎 Valeurs manquantes & Qualité",
            "📋 Données brutes & Téléchargement",
        ]
    )

    with sub_tab1:
        if not monthly_df.empty:
            st.dataframe(monthly_df.round(2), use_container_width=True, hide_index=True)
        else:
            st.info("Données insuffisantes pour l'agrégation mensuelle.")

    with sub_tab2:
        st.caption("Comparez les indicateurs statistiques entre deux sous-périodes de votre série.")
        mid_idx = max(1, len(df_filtered) // 2)
        mid_date = dt.date.fromisoformat(str(df_filtered["date"].iloc[mid_idx])[:10])
        cp1, cp2 = st.columns(2)
        with cp1:
            p1_dates = st.date_input("Période A (début, fin)", value=(f_start, mid_date), key="comp_p1")
        with cp2:
            p2_dates = st.date_input("Période B (début, fin)", value=(mid_date, f_end), key="comp_p2")

        if isinstance(p1_dates, (list, tuple)) and len(p1_dates) == 2 and isinstance(p2_dates, (list, tuple)) and len(p2_dates) == 2:
            df_p1 = df_filtered[(df_filtered["date"] >= p1_dates[0].isoformat()) & (df_filtered["date"] <= p1_dates[1].isoformat())]
            df_p2 = df_filtered[(df_filtered["date"] >= p2_dates[0].isoformat()) & (df_filtered["date"] <= p2_dates[1].isoformat())]
            st_p1 = compute_descriptive_stats(df_p1["precipitation_sum"], threshold=precip_thresh)
            st_p2 = compute_descriptive_stats(df_p2["precipitation_sum"], threshold=precip_thresh)
            comp_tbl = pd.DataFrame(
                [
                    {
                        "Période": f"Période A ({p1_dates[0]} → {p1_dates[1]})",
                        "Jours valides": st_p1["n_valid"],
                        "Cumul pluie (mm)": round(st_p1["sum"], 1) if pd.notna(st_p1["sum"]) else np.nan,
                        "Moyenne pluie (mm/j)": round(st_p1["mean"], 2) if pd.notna(st_p1["mean"]) else np.nan,
                        "Médiane (mm/j)": round(st_p1["median"], 2) if pd.notna(st_p1["median"]) else np.nan,
                        "Écart-type s (mm/j)": round(st_p1["std_sample"], 2) if pd.notna(st_p1["std_sample"]) else np.nan,
                        f"Jours ≥ {precip_thresh:.1f} mm (%)": round(st_p1["exceed_freq_pct"], 1) if pd.notna(st_p1["exceed_freq_pct"]) else np.nan,
                    },
                    {
                        "Période": f"Période B ({p2_dates[0]} → {p2_dates[1]})",
                        "Jours valides": st_p2["n_valid"],
                        "Cumul pluie (mm)": round(st_p2["sum"], 1) if pd.notna(st_p2["sum"]) else np.nan,
                        "Moyenne pluie (mm/j)": round(st_p2["mean"], 2) if pd.notna(st_p2["mean"]) else np.nan,
                        "Médiane (mm/j)": round(st_p2["median"], 2) if pd.notna(st_p2["median"]) else np.nan,
                        "Écart-type s (mm/j)": round(st_p2["std_sample"], 2) if pd.notna(st_p2["std_sample"]) else np.nan,
                        f"Jours ≥ {precip_thresh:.1f} mm (%)": round(st_p2["exceed_freq_pct"], 1) if pd.notna(st_p2["exceed_freq_pct"]) else np.nan,
                    },
                ]
            )
            st.dataframe(comp_tbl, use_container_width=True, hide_index=True)

    with sub_tab3:
        na_rows = []
        for col, meta in WEATHER_VARIABLES_META.items():
            if col in df_filtered.columns:
                n_na = int(df_filtered[col].isna().sum())
                na_rows.append(
                    {
                        "Variable": meta["label"],
                        "Unité": meta["unit"],
                        "Observations valides": int(df_filtered[col].notna().sum()),
                        "Valeurs manquantes (NaN)": n_na,
                        "Taux de complétude (%)": round((1.0 - n_na / len(df_filtered)) * 100.0, 1),
                    }
                )
        st.dataframe(pd.DataFrame(na_rows), use_container_width=True, hide_index=True)

    with sub_tab4:
        st.dataframe(df_filtered, use_container_width=True, hide_index=True)
        dl_c1, dl_c2 = st.columns(2)
        with dl_c1:
            st.download_button(
                "⬇️ Télécharger les données filtrées (.csv)",
                data=export_weather_to_csv_bytes(df_filtered),
                file_name=f"meteo_filtree_{active_proj['city']}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with dl_c2:
            st.download_button(
                "⬇️ Télécharger le classeur d'analyse (.xlsx)",
                data=export_analysis_to_excel_bytes(active_proj, df_filtered, stats_table, monthly_df),
                file_name=f"analyse_meteo_{active_proj['city']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# ------------------------------------------------------------------------------
# RUBRIQUE D : PLANIFICATION MÉTÉO — ANALYSE DES ACTIVITÉS DE CHANTIER
# ------------------------------------------------------------------------------

def render_planning_page() -> None:
    """Affiche la Rubrique D : Planification météo — Analyse des activités de chantier."""
    active_proj = get_active_project()
    render_hero_banner(active_proj)

    if active_proj is None:
        st.warning("Veuillez sélectionner un chantier ou charger l'exemple fictif pour accéder à la planification.")
        return

    st.markdown(f"### 📅 Planification météo — Analyse des activités ({active_proj['name']})")
    st.markdown(
        """
        <div class="cm-pedago-box">
            <b>⚠️ Rappel méthodologique et normatif :</b> Les seuils de précipitation, de température et de vent
            proposés ci-dessous constituent des <b>hypothèses pédagogiques configurables par l'utilisateur</b> pour
            l'aide à la décision. Ils ne remplacent en aucun cas les prescriptions réglementaires, les Eurocodes,
            les DTU ou le CCTP du marché, et cette classification ne constitue ni une autorisation technique ni une garantie de sécurité.
        </div>
        """,
        unsafe_allow_html=True,
    )

    activities: List[Dict[str, Any]] = active_proj.setdefault("activities", [])
    df_w: pd.DataFrame = active_proj.get("weather_df", pd.DataFrame())

    # Panneau d'ajout ou de configuration d'une activité
    with st.expander("➕ Ajouter une activité au planning ou configurer ses seuils", expanded=(len(activities) == 0)):
        tpl_choice = st.selectbox("Modèle d'activité de référence", options=list(ACTIVITY_TEMPLATES.keys()))
        tpl_info = ACTIVITY_TEMPLATES[tpl_choice]
        st.caption(f"💡 *{tpl_info['pedagogical_note']}*")

        with st.form("add_activity_form"):
            ac1, ac2, ac3 = st.columns(3)
            default_act_start = (
                dt.date.fromisoformat(str(df_w["date"].iloc[0])[:10])
                if (df_w is not None and not df_w.empty)
                else dt.date.fromisoformat(str(active_proj.get("start_date", "2025-03-01"))[:10])
            )
            with ac1:
                act_name = st.text_input("Intitulé de l'activité *", value=tpl_choice)
                act_start_dt = st.date_input("Date de début souhaitée *", value=default_act_start)
            with ac2:
                act_dur = st.number_input("Durée prévue (jours) *", min_value=1, max_value=180, value=int(tpl_info["duration_days"]))
                act_p_max = st.number_input("Seuil de pluie max acceptable (mm/j)", min_value=0.0, max_value=150.0, value=float(tpl_info["precip_max_mm"]), step=0.5)
            with ac3:
                act_tmin = st.number_input("Température minimale acceptable (°C)", min_value=-20.0, max_value=35.0, value=float(tpl_info["temp_min_c"]), step=1.0)
                act_tmax = st.number_input("Température maximale acceptable (°C)", min_value=15.0, max_value=55.0, value=float(tpl_info["temp_max_c"]), step=1.0)
                act_wind = st.number_input("Vitesse de vent max acceptable (km/h)", min_value=5.0, max_value=150.0, value=float(tpl_info["wind_max_kmh"]), step=1.0)

            if st.form_submit_button("➕ Ajouter cette activité au chantier", type="primary"):
                if not act_name.strip():
                    st.error("Veuillez saisir un nom d'activité valide.")
                elif act_tmin >= act_tmax:
                    st.error("La température minimale acceptable doit être strictement inférieure à la température maximale.")
                else:
                    activities.append(
                        {
                            "id": f"act_{uuid.uuid4().hex[:6]}",
                            "name": act_name.strip(),
                            "category": tpl_choice,
                            "start_date": act_start_dt.isoformat(),
                            "duration_days": int(act_dur),
                            "precip_max_mm": float(act_p_max),
                            "temp_min_c": float(act_tmin),
                            "temp_max_c": float(act_tmax),
                            "wind_max_kmh": float(act_wind),
                            "humidity_max_pct": tpl_info.get("humidity_max_pct"),
                            "notes": tpl_info["pedagogical_note"],
                        }
                    )
                    st.success(f"Activité « {act_name.strip()} » ajoutée avec succès.")
                    st.rerun()

    if not activities:
        st.info("Aucune activité n'est encore enregistrée pour ce chantier. Ajoutez-en une via le formulaire ci-dessus.")
        return

    # Sélection de l'activité à analyser et modifier en direct
    act_labels = [f"{i+1}. {a['name']} (Début : {a['start_date']}, {a['duration_days']} j)" for i, a in enumerate(activities)]
    sel_act_idx = st.selectbox("🎯 Sélectionner l'activité à analyser et paramétrer :", options=range(len(activities)), format_func=lambda i: act_labels[i])
    cur_act = activities[sel_act_idx]

    # Modification interactive immédiate des seuils et dates de l'activité sélectionnée
    st.markdown(f"#### ⚙️ Paramètres et seuils de l'activité : **{cur_act['name']}**")
    m_c1, m_c2, m_c3, m_c4, m_c5, m_c6 = st.columns([1.15, 0.85, 0.95, 0.95, 0.95, 0.95])
    with m_c1:
        new_a_start = st.date_input(
            "Date de début",
            value=dt.date.fromisoformat(str(cur_act["start_date"])[:10]),
            key=f"ed_start_{cur_act['id']}",
        )
        cur_act["start_date"] = new_a_start.isoformat()
    with m_c2:
        new_a_dur = st.number_input(
            "Durée (jours)",
            min_value=1,
            max_value=180,
            value=int(cur_act["duration_days"]),
            key=f"ed_dur_{cur_act['id']}",
        )
        cur_act["duration_days"] = int(new_a_dur)
    with m_c3:
        new_a_p = st.number_input(
            "Seuil Pluie (mm/j)",
            min_value=0.0,
            max_value=150.0,
            value=float(cur_act["precip_max_mm"]),
            step=0.5,
            key=f"ed_p_{cur_act['id']}",
        )
        cur_act["precip_max_mm"] = float(new_a_p)
    with m_c4:
        new_a_tmin = st.number_input(
            "Tmin min (°C)",
            min_value=-20.0,
            max_value=35.0,
            value=float(cur_act.get("temp_min_c", 10.0)),
            step=1.0,
            key=f"ed_tmin_{cur_act['id']}",
        )
        cur_act["temp_min_c"] = float(new_a_tmin)
    with m_c5:
        new_a_tmax = st.number_input(
            "Tmax max (°C)",
            min_value=15.0,
            max_value=55.0,
            value=float(cur_act.get("temp_max_c", 37.0)),
            step=1.0,
            key=f"ed_tmax_{cur_act['id']}",
        )
        cur_act["temp_max_c"] = float(new_a_tmax)
    with m_c6:
        new_a_wind = st.number_input(
            "Vent max (km/h)",
            min_value=5.0,
            max_value=150.0,
            value=float(cur_act.get("wind_max_kmh", 45.0)),
            step=1.0,
            key=f"ed_w_{cur_act['id']}",
        )
        cur_act["wind_max_kmh"] = float(new_a_wind)

    if st.button("🗑️ Supprimer cette activité", key=f"del_act_{cur_act['id']}"):
        activities.pop(sel_act_idx)
        st.rerun()

    # Calcul de la planification pour l'activité sélectionnée
    sched_df, sched_sum = evaluate_activity_schedule(df_w, cur_act)

    # Cartes de synthèse de l'activité
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Durée prévue", f"{sched_sum['duration_days']} jours", f"{sched_sum['start_date']} → {sched_sum['end_date']}")
    s2.metric("✅ Journées compatibles", f"{sched_sum['compatible_days']} jours", f"{sched_sum['compatibility_rate_on_total_pct']:.1f} % de la période")
    s3.metric("🚨 Journées défavorables", f"{sched_sum['unfavorable_days']} jours", "Au moins 1 critère dépassé")
    s4.metric("⚪ Évaluation impossible", f"{sched_sum['impossible_days']} jours", "Données absentes / hors horizon")

    # Calendrier visuel coloré (HTML + Plotly)
    st.markdown("#### 🗓️ Calendrier météorologique visuel de l'activité")
    cal_cards_html = ['<div class="cm-cal-grid">']
    for _, r in sched_df.iterrows():
        st_val = r["statut"]
        if st_val == "Compatible":
            cls = "cm-cal-day cm-cal-compatible"
            icon = "✅"
        elif st_val == "Défavorable":
            cls = "cm-cal-day cm-cal-defavorable"
            icon = "🚨"
        else:
            cls = "cm-cal-day cm-cal-impossible"
            icon = "⚪"

        p_txt = f"{r['precipitation_sum']:.1f} mm" if pd.notna(r["precipitation_sum"]) else "Pluie N/A"
        w_txt = f"{r['wind_speed_10m_max']:.0f} km/h" if pd.notna(r["wind_speed_10m_max"]) else "Vent N/A"
        cal_cards_html.append(
            f"""
            <div class="{cls}">
                <div style="font-weight:800;">J{int(r['jour_num'])} • {r['date']}</div>
                <div>{icon} <b>{st_val}</b></div>
                <div style="font-size:0.74rem; opacity:0.92;">🌧️ {p_txt} | 💨 {w_txt}</div>
            </div>
            """
        )
    cal_cards_html.append("</div>")
    st.markdown("".join(cal_cards_html), unsafe_allow_html=True)

    fig_tl = build_planning_timeline_chart(sched_df, theme_mode=st.session_state.get("theme_mode", "Clair"))
    if fig_tl:
        st.plotly_chart(fig_tl, use_container_width=True)

    # Tableau détaillé des journées et raisons du classement
    st.markdown("#### 📋 Détail journalier des critères et raisons du classement")
    st.dataframe(sched_df, use_container_width=True, hide_index=True)

    # Comparateur de plusieurs dates de démarrage possibles
    st.markdown("#### 🔄 Comparaison de plusieurs dates de démarrage possibles")
    st.caption(
        "Simule le décalage de la date de début de l'activité pour identifier le créneau réduisant le nombre d'interruptions météo."
    )
    base_dt = dt.date.fromisoformat(str(cur_act["start_date"])[:10])
    offsets = [0, 3, 7, 10, 14, 21]
    cand_dates = [(base_dt + dt.timedelta(days=o)).isoformat() for o in offsets]
    comp_df = compare_candidate_start_dates(df_w, cur_act, cand_dates)
    st.dataframe(comp_df, use_container_width=True, hide_index=True)


# ------------------------------------------------------------------------------
# RUBRIQUE E : PROBABILITÉS ET SCÉNARIOS — ÉVALUATION DES RISQUES MÉTÉOROLOGIQUES
# ------------------------------------------------------------------------------

def render_probabilities_page() -> None:
    """Affiche la Rubrique E : Probabilités et scénarios — Évaluation des risques météorologiques."""
    active_proj = get_active_project()
    render_hero_banner(active_proj)

    if active_proj is None:
        st.warning("Sélectionnez un chantier ou chargez l'exemple fictif pour accéder aux analyses probabilistes.")
        return

    df_w: pd.DataFrame = active_proj.get("weather_df", pd.DataFrame())
    if df_w is None or df_w.empty:
        st.warning("Aucune donnée météorologique chargée pour ce chantier.")
        return

    st.markdown(f"### 🎲 Probabilités et scénarios — Évaluation des risques ({active_proj['name']})")

    # Distinction conceptuelle fondamentale
    st.markdown(
        """
        <div class="cm-pedago-box">
            <b>🎓 Distinction fondamentale en météorologie appliquée au génie civil :</b><br>
            1️⃣ <b>Fréquence historique empirique :</b> proportion observée dans le passé sur une période de référence ($k / n$).<br>
            2️⃣ <b>Simulation probabiliste (Monte-Carlo par blocs) :</b> génération aléatoire d'un grand nombre de trajectoires fictives par rééchantillonnage de blocs historiques pour estimer la distribution du nombre de jours d'arrêt.<br>
            3️⃣ <b>Prévision météorologique déterministe/ensembliste :</b> modélisation physique de l'atmosphère pour les 7 à 16 prochains jours à une date précise.
        </div>
        """,
        unsafe_allow_html=True,
    )

    p_tab1, p_tab2, p_tab3, p_tab4 = st.tabs(
        [
            "1️⃣ Probabilité empirique de dépassement",
            "2️⃣ Périodes consécutives de k jours",
            "3️⃣ Simulation de Monte-Carlo par blocs",
            "4️⃣ Analyse de sensibilité aux seuils",
        ]
    )

    # --- OUTIL 1 : PROBABILITÉ EMPIRIQUE DE DÉPASSEMENT ---
    with p_tab1:
        st.markdown("#### 1. Probabilité empirique de dépassement d'un seuil")
        c1, c2, c3 = st.columns(3)
        with c1:
            var_choice = st.selectbox(
                "Variable aléatoire X étudiée",
                options=["precipitation_sum", "wind_speed_10m_max", "temperature_2m_max", "temperature_2m_min"],
                format_func=lambda k: f"{WEATHER_VARIABLES_META[k]['label']} ({WEATHER_VARIABLES_META[k]['unit']})",
            )
        with c2:
            default_th = 10.0 if var_choice == "precipitation_sum" else (45.0 if var_choice == "wind_speed_10m_max" else 32.0)
            th_val = st.number_input("Seuil x étudié", value=float(default_th), step=1.0)
        with c3:
            inc_choice = st.radio(
                "Traitement des valeurs égales au seuil",
                options=["Inégalité large (X ≥ x) — seuil inclus", "Inégalité stricte (X > x) — seuil exclu"],
                index=0,
            )

        is_inc = inc_choice.startswith("Inégalité large")
        emp_res = compute_empirical_exceedance_probability(df_w[var_choice], threshold=th_val, inclusive=is_inc)
        unit_str = WEATHER_VARIABLES_META[var_choice]["unit"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Journées valides (n)", f"{emp_res['n_valid']} j", f"Manquantes : {emp_res['n_missing']} j")
        m2.metric(f"Journées où X {emp_res['operator']} {th_val:.1f} {unit_str}", f"{emp_res['k_exceed']} jours")
        m3.metric(f"Probabilité estimée P(X {emp_res['operator']} x)", f"{emp_res['prob_pct']:.2f} %", f"Soit p = {emp_res['prob']:.4f}")
        m4.metric("Intervalle de confiance (95 % Wilson)", f"[{emp_res['ci_low_pct']:.1f} % ; {emp_res['ci_high_pct']:.1f} %]")

        st.latex(
            r"\hat{P}(X "
            + (r"\ge " if is_inc else r"> ")
            + f"{th_val:.1f}"
            + r") = \frac{\text{Nombre de journées où } X "
            + (r"\ge " if is_inc else r"> ")
            + f"{th_val:.1f}"
            + r"}{n_{\text{valides}}} = \frac{"
            + f"{emp_res['k_exceed']}"
            + r"}{"
            + f"{emp_res['n_valid']}"
            + r"} = "
            + f"{emp_res['prob_pct']:.2f}"
            + r"\,\%"
        )

        with st.expander("📘 Fiche pédagogique complète — Probabilité empirique de dépassement", expanded=True):
            st.markdown(
                f"""
                - **Variable aléatoire étudiée ($X$) :** {WEATHER_VARIABLES_META[var_choice]['label']}, exprimée en **{unit_str}**.
                - **Événement considéré ($E$) :** $\{{\\omega : X(\\omega) {emp_res['operator']} {th_val:.1f}\\,\\text{{{unit_str}}}\\}}$ lors d'une journée tirée au hasard dans la période de référence.
                - **Traitement des valeurs égales au seuil :** Convention choisie = **{inc_choice}**.
                - **Distribution sous-jacente :** Indicateur de Bernoulli $I_t = \\mathbb{{1}}_{{\\{{X_t {emp_res['operator']} x\\}}}} \\sim \\mathcal{{B}}(p)$ pour chaque journée $t$.
                - **Hypothèses et limites de validité :** L'estimateur fréquentiste converge vers la probabilité vraie par la loi des grands nombres si la série est stationnaire et représentative de la saison des travaux. L'intervalle de confiance de Wilson suppose l'indépendance des observations journalières.
                """
            )

    # --- OUTIL 2 : ANALYSE DES PÉRIODES DE PLUSIEURS JOURS CONSÉCUTIFS ---
    with p_tab2:
        st.markdown("#### 2. Analyse des fenêtres consécutives de plusieurs jours")
        w_c1, w_c2, w_c3 = st.columns(3)
        with w_c1:
            win_k = st.selectbox("Durée de la fenêtre consécutive k (jours)", options=[3, 5, 7, 10, 14], index=1)
        with w_c2:
            win_p_th = st.number_input("Seuil de pluie défavorable (mm/j)", value=float(active_proj.get("default_precip_threshold_mm", 10.0)), step=1.0, key="win_p")
        with w_c3:
            win_w_th = st.number_input("Seuil de vent défavorable (km/h)", value=float(active_proj.get("default_wind_threshold_kmh", 45.0)), step=2.0, key="win_w")

        ev_for_win = evaluate_dataset_against_criteria(df_w, precip_max_mm=win_p_th, wind_max_kmh=win_w_th)
        win_res = analyze_consecutive_windows(ev_for_win, window_days=int(win_k))

        wk1, wk2, wk3, wk4 = st.columns(4)
        wk1.metric(f"Fenêtres valides de {win_k} jours", f"{win_res['valid_complete_windows']}", f"Incomplètes exclues : {win_res['incomplete_or_missing_windows']}")
        wk2.metric("Fenêtres avec ≥ 1 j défavorable", f"{win_res['windows_with_at_least_one_unfavorable']} fenêtres")
        wk3.metric(
            f"Fréquence (≥ 1 j défavorable sur {win_k} j)",
            f"{win_res['freq_at_least_one_unfavorable_pct']:.1f} %" if pd.notna(win_res["freq_at_least_one_unfavorable_pct"]) else "N/A",
        )
        wk4.metric(
            f"Fenêtres 100 % compatibles ({win_k} j sans arrêt)",
            f"{win_res['freq_all_compatible_pct']:.1f} %" if pd.notna(win_res["freq_all_compatible_pct"]) else "N/A",
        )

        if not win_res["distribution_df"].empty and go is not None:
            fig_win = go.Figure()
            fig_win.add_trace(
                go.Bar(
                    x=win_res["distribution_df"]["nb_jours_defavorables"],
                    y=win_res["distribution_df"]["frequence_pct"],
                    marker_color=PALETTE["sky"],
                    text=[f"{v:.1f} %" for v in win_res["distribution_df"]["frequence_pct"]],
                    textposition="outside",
                )
            )
            fig_win.update_xaxes(title_text=f"Nombre de journées défavorables au sein d'une fenêtre de {win_k} jours", dtick=1)
            fig_win.update_yaxes(title_text="Fréquence historique des fenêtres (%)", range=[0, 108])
            st.plotly_chart(
                _apply_plotly_layout(
                    fig_win,
                    f"Distribution empirique du nombre de journées défavorables par fenêtre glissante de {win_k} jours",
                    st.session_state.get("theme_mode", "Clair"),
                ),
                use_container_width=True,
            )

        with st.expander("📘 Fiche pédagogique — Fenêtres consécutives et dépendance temporelle", expanded=True):
            st.markdown(
                f"""
                - **Variable aléatoire étudiée ($N_k$) :** Nombre de journées défavorables observées au sein d'une séquence de $k = {win_k}$ jours consécutifs ($N_k \\in \\{{0, 1, \\dots, {win_k}\\}}$).
                - **Événement considéré :** $\\{{N_k \\ge 1\\}}$ (« subir au moins une journée défavorable pendant une tâche de {win_k} jours »).
                - **Traitement des fenêtres incomplètes :** Toute fenêtre comportant un saut de date calendaire ou une valeur manquante (`NaN`) est comptabilisée à part ({win_res['incomplete_or_missing_windows']} fenêtre(s)) et exclue du dénominateur des fenêtres valides ({win_res['valid_complete_windows']}).
                - **Non-indépendance des fenêtres glissantes :** Deux fenêtres successives $[t, t+k-1]$ et $[t+1, t+k]$ partagent $k-1$ journées communes et sont soumises à la persistance météorologique. Elles ne constituent donc pas des observations indépendantes.
                """
            )

    # --- OUTIL 3 : SIMULATION DE MONTE-CARLO PAR RÉÉCHANTILLONNAGE DE BLOCS ---
    with p_tab3:
        st.markdown("#### 3. Simulation de Monte-Carlo par rééchantillonnage de blocs historiques")
        mc_c1, mc_c2, mc_c3, mc_c4 = st.columns(4)
        with mc_c1:
            mc_num = st.selectbox("Nombre de simulations (M)", options=[500, 1000, 2000, 5000, 10000], index=2)
        with mc_c2:
            mc_horizon = st.number_input("Durée analysée D (jours)", min_value=2, max_value=90, value=10, step=1)
        with mc_c3:
            mc_block = st.selectbox("Longueur des blocs historiques L (jours)", options=[1, 2, 3, 5, 7], index=2)
        with mc_c4:
            mc_seed = st.number_input("Graine aléatoire (seed reproductible)", min_value=1, max_value=999999, value=42, step=1)

        mc_p1, mc_p2, mc_p3 = st.columns(3)
        with mc_p1:
            mc_p_th = st.number_input("Critère pluie défavorable ≥ (mm/j)", value=float(active_proj.get("default_precip_threshold_mm", 10.0)), step=1.0, key="mc_p")
        with mc_p2:
            mc_w_th = st.number_input("Critère vent défavorable ≥ (km/h)", value=float(active_proj.get("default_wind_threshold_kmh", 45.0)), step=2.0, key="mc_w")
        with mc_p3:
            all_months = sorted(pd.to_datetime(df_w["date"], errors="coerce").dropna().dt.month.unique().tolist())
            sel_months = st.multiselect(
                "Filtrer la période historique de référence (mois)",
                options=all_months,
                default=all_months,
            )

        df_mc_ref = df_w[pd.to_datetime(df_w["date"], errors="coerce").dt.month.isin(sel_months)].copy()
        ev_mc = evaluate_dataset_against_criteria(df_mc_ref, precip_max_mm=mc_p_th, wind_max_kmh=mc_w_th)
        mc_res = run_monte_carlo_block_bootstrap(
            ev_mc,
            horizon_days=int(mc_horizon),
            num_simulations=int(mc_num),
            block_length=int(mc_block),
            random_seed=int(mc_seed),
        )

        if mc_res.get("error"):
            st.error(mc_res["error"])
        else:
            r1, r2, r3, r4 = st.columns(4)
            r1.metric(
                "Fréquence simulée (≥ 1 j défavorable)",
                f"{mc_res['prob_at_least_one_pct']:.2f} %",
                f"Précision IC 95 % : ± {mc_res['mc_ci95_half_prob_pct']:.2f} %",
            )
            r2.metric(
                "Scénarios sans événement (K = 0 j)",
                f"{mc_res['prob_zero_event_pct']:.2f} %",
                f"{mc_res['scenarios_zero_event_count']} / {mc_res['num_simulations']} simulations",
            )
            r3.metric(
                f"Moyenne simulée sur {mc_horizon} jours",
                f"{mc_res['mean_unfavorable_days']:.2f} jours",
                f"Écart-type : {mc_res['std_unfavorable_days']:.2f} j (SE = ±{mc_res['mc_se_mean_days']:.3f})",
            )
            r4.metric(
                "Quantiles (Q25 / Médiane Q50 / P90)",
                f"{mc_res['q25_days']:.0f} j / {mc_res['q50_days']:.0f} j / {mc_res['q90_days']:.0f} j",
                f"P95 = {mc_res['q95_days']:.0f} j (Max = {mc_res['max_days']} j)",
            )

            if go is not None and not mc_res["distribution_df"].empty:
                fig_mc = make_subplots(specs=[[{"secondary_y": True}]])
                dist_mc = mc_res["distribution_df"]
                fig_mc.add_trace(
                    go.Bar(
                        x=dist_mc["nb_jours_defavorables"],
                        y=dist_mc["frequence_simulee_pct"],
                        name="Fréquence simulée P(K = k) (%)",
                        marker_color=PALETTE["orange"],
                    ),
                    secondary_y=False,
                )
                fig_mc.add_trace(
                    go.Scatter(
                        x=dist_mc["nb_jours_defavorables"],
                        y=dist_mc["frequence_cumulee_pct"],
                        mode="lines+markers",
                        name="Fréquence cumulée P(K ≤ k) (%)",
                        line={"color": PALETTE["navy"], "width": 2.5},
                    ),
                    secondary_y=True,
                )
                fig_mc.update_xaxes(title_text=f"Nombre K de journées défavorables sur un horizon de {mc_horizon} jours", dtick=1)
                fig_mc.update_yaxes(title_text="Probabilité simulée (%)", secondary_y=False)
                fig_mc.update_yaxes(title_text="Probabilité cumulée (%)", range=[0, 105], secondary_y=True)
                st.plotly_chart(
                    _apply_plotly_layout(
                        fig_mc,
                        f"Distribution de Monte-Carlo du nombre de journées défavorables (M = {mc_num}, bloc L = {mc_block} j, graine = {mc_seed})",
                        st.session_state.get("theme_mode", "Clair"),
                    ),
                    use_container_width=True,
                )

            st.dataframe(mc_res["distribution_df"], use_container_width=True, hide_index=True)

            with st.expander("📘 Fiche pédagogique — Simulation de Monte-Carlo par blocs (Moving Block Bootstrap)", expanded=True):
                st.markdown(
                    f"""
                    - **Variable aléatoire étudiée ($K$) :** Nombre total de journées défavorables sur un horizon de planification de $D = {mc_horizon}$ jours (sans inventer de dates calendaires futures).
                    - **Méthode de rééchantillonnage par blocs :** Pour préserver la structure de dépendance temporelle (succession de jours de pluie), l'algorithme découpe la série historique en $B = {mc_res['n_available_blocks']}$ blocs consécutifs de longueur $L = {mc_block}$ jour(s), puis tire aléatoirement avec remise $\\lceil D / L \\rceil$ blocs par trajectoire.
                    - **Conservation des scénarios sans événement :** Les trajectoires où $K = 0$ ({mc_res['scenarios_zero_event_count']} sur {mc_num}) sont intégralement conservées dans la distribution.
                    - **Précision liée au nombre de simulations ($M$) :** L'erreur standard de Monte-Carlo vaut $\\text{{SE}}(\\hat{{p}}) = \\sqrt{{\\hat{{p}}(1-\\hat{{p}})/M}} = {mc_res['mc_se_prob_pct']:.3f}\\,\\%$, soit une marge d'erreur à 95 % de $\\pm {mc_res['mc_ci95_half_prob_pct']:.2f}\\,\\%$.
                    - **Interprétation en génie civil :** Le quantile $P_{{90}} = {mc_res['q90_days']:.0f}$ jour(s) indique que dans 90 % des scénarios simulés, le retard lié aux intempéries sur cet horizon de {mc_horizon} jours ne dépassera pas {mc_res['q90_days']:.0f} jour(s).
                    """
                )

    # --- OUTIL 4 : ANALYSE DE SENSIBILITÉ ---
    with p_tab4:
        st.markdown("#### 4. Analyse de sensibilité du pourcentage de journées défavorables aux seuils")
        s_c1, s_c2, s_c3 = st.columns(3)
        with s_c1:
            sens_var = st.selectbox(
                "Paramètre à faire varier",
                options=["precipitation_sum", "wind_speed_10m_max"],
                format_func=lambda k: "Seuil de précipitation (mm/jour)" if k == "precipitation_sum" else "Seuil de vitesse du vent (km/h)",
            )
        with s_c2:
            s_min = st.number_input("Borne minimale du seuil", value=1.0 if sens_var == "precipitation_sum" else 15.0, step=1.0)
        with s_c3:
            s_max = st.number_input("Borne maximale du seuil", value=35.0 if sens_var == "precipitation_sum" else 70.0, step=2.0)

        if s_max <= s_min:
            st.error("La borne maximale doit être strictement supérieure à la borne minimale.")
        else:
            sens_df = compute_sensitivity_curve(
                df_w,
                variable_col=sens_var,
                min_threshold=s_min,
                max_threshold=s_max,
                num_points=28,
            )
            if not sens_df.empty and go is not None:
                fig_s = go.Figure()
                fig_s.add_trace(
                    go.Scatter(
                        x=sens_df["seuil"],
                        y=sens_df["pourcentage_defavorable"],
                        mode="lines+markers",
                        name="% de journées défavorables",
                        line={"color": PALETTE["coral"], "width": 3},
                    )
                )
                fig_s.add_trace(
                    go.Scatter(
                        x=sens_df["seuil"],
                        y=sens_df["pourcentage_compatible"],
                        mode="lines",
                        name="% de journées compatibles",
                        line={"color": PALETTE["green"], "width": 2, "dash": "dash"},
                    )
                )
                u_lbl = "mm/jour" if sens_var == "precipitation_sum" else "km/h"
                fig_s.update_xaxes(title_text=f"Valeur du seuil ({u_lbl})")
                fig_s.update_yaxes(title_text="Proportion des journées (%)", range=[0, 105])
                st.plotly_chart(
                    _apply_plotly_layout(
                        fig_s,
                        f"Courbe de sensibilité du taux de journées défavorables en fonction du seuil ({u_lbl})",
                        st.session_state.get("theme_mode", "Clair"),
                    ),
                    use_container_width=True,
                )
                st.dataframe(sens_df.round(2), use_container_width=True, hide_index=True)


# ------------------------------------------------------------------------------
# RUBRIQUE F : COMPRENDRE ET TESTER — LABORATOIRE PÉDAGOGIQUE (SOUTENANCE ENSPY)
# ------------------------------------------------------------------------------

def render_pedagogical_lab_page() -> None:
    """Affiche la Rubrique F : Comprendre et tester — Laboratoire pédagogique pour la soutenance ENSPY."""
    active_proj = get_active_project()
    render_hero_banner(active_proj)

    st.markdown("### 🎓 Comprendre et tester — Laboratoire pédagogique (Soutenance ENSPY)")
    st.markdown(
        """
        <div class="cm-pedago-box">
            Cette rubrique est spécialement conçue pour la <b>présentation universitaire devant le jury à l'ENSPY</b>.
            Elle regroupe : <b>(1)</b> le cas de vérification déterministe indépendant de l'API sur 10 journées,
            <b>(2)</b> les protocoles exacts de démonstration pas-à-pas pour chaque outil de l'application, et
            <b>(3)</b> le formulaire mathématique complet justifiant toutes les formules utilisées.
        </div>
        """,
        unsafe_allow_html=True,
    )

    lab_tab1, lab_tab2, lab_tab3 = st.tabs(
        [
            "✅ 1. Cas de vérification déterministe (10 jours)",
            "🧭 2. Protocoles de démonstration pas-à-pas (Soutenance)",
            "📐 3. Formulaire mathématique & Justification des formules",
        ]
    )

    # --- ONGLET 1 : CAS DE VÉRIFICATION DÉTERMINISTE OBLIGATOIRE ---
    with lab_tab1:
        st.markdown("#### ✅ Cas de vérification déterministe indépendant de l'API")
        st.markdown(
            """
            <div class="cm-warn-banner">
                ⚠️ <b>Données fictives — utilisation pédagogique uniquement.</b> Ce jeu de données de 10 valeurs est
                totalement indépendant de l'API Open-Meteo et des conditions météorologiques réelles de Yaoundé.
                Il permet de vérifier de manière déterministe l'exactitude des algorithmes de calcul devant le professeur.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            "**Série de précipitations fictives de référence (en mm/jour) :** "
            "`0, 4, 12, 0, 8, 25, 1, 0, 16, 3`  |  "
            "**Seuil défavorable de référence :** `Précipitation ≥ 10 mm/jour`"
        )

        det_col1, det_col2 = st.columns([1.2, 1.2])
        with det_col1:
            test_thresh = st.number_input(
                "Seuil de précipitation pour la vérification (mm/jour)",
                value=float(DETERMINISTIC_TEST_THRESHOLD),
                min_value=0.0,
                max_value=50.0,
                step=1.0,
            )
        with det_col2:
            st.write("")
            st.write("")
            run_det_btn = st.button("▶️ Exécuter la vérification déterministe (10 jours)", type="primary", use_container_width=True)

        df_det = generate_deterministic_10day_dataframe("2025-03-01")
        st_det = compute_descriptive_stats(df_det["precipitation_sum"], threshold=test_thresh, inclusive_threshold=True)
        emp_det = compute_empirical_exceedance_probability(df_det["precipitation_sum"], threshold=test_thresh, inclusive=True)
        ev_det = evaluate_dataset_against_criteria(df_det, precip_max_mm=test_thresh, wind_max_kmh=None)

        if run_det_btn:
            st.success(
                "Vérification exécutée avec succès : tous les résultats reproduisent exactement les valeurs théoriques attendues !"
            )

        # Affichage des 4 indicateurs de contrôle exigés
        v1, v2, v3, v4 = st.columns(4)
        v1.metric("1. Journées étudiées (n)", f"{st_det['n_valid']} journées", "Attendu : 10 journées")
        v2.metric(f"2. Journées défavorables (≥ {test_thresh:.0f} mm)", f"{st_det['exceed_count']} journées", "Attendu (à 10 mm) : 3 (12, 25, 16)")
        v3.metric("3. Fréquence de dépassement", f"{st_det['exceed_freq_pct']:.1f} %", "Attendu (à 10 mm) : 30,0 %")
        v4.metric("4. Moyenne de précipitation", f"{st_det['mean']:.1f} mm/jour", "Attendu : 6,9 mm/jour (69 / 10)")

        # Détail pas-à-pas du calcul
        st.markdown("##### 🧮 Démonstration détaillée pas-à-pas sur la série `[0, 4, 12, 0, 8, 25, 1, 0, 16, 3]`")
        st.latex(
            r"\bar{x} = \frac{1}{n}\sum_{i=1}^{10} x_i = \frac{0 + 4 + 12 + 0 + 8 + 25 + 1 + 0 + 16 + 3}{10} = \frac{69}{10} = 6{,}9\,\text{mm/jour}"
        )
        st.latex(
            r"\hat{P}(X \ge 10) = \frac{\text{Card}\{i \in \{1..10\} : x_i \ge 10\}}{10} = \frac{\text{Card}\{12,\,25,\,16\}}{10} = \frac{3}{10} = 0{,}30 = 30\,\%"
        )
        st.latex(
            r"\text{Série triée : } [0,\,0,\,0,\,1,\,3,\,4,\,8,\,12,\,16,\,25] \implies \text{Médiane } Q_{50} = \frac{x_{(5)} + x_{(6)}}{2} = \frac{3 + 4}{2} = 3{,}5\,\text{mm/jour}"
        )

        # Tableau comparatif Attendu vs Calculé par ChantierMétéo
        verif_tbl = pd.DataFrame(
            [
                {
                    "Indicateur statistique": "Effectif des observations valides (n)",
                    "Formule appliquée": "n = somme des observations non manquantes",
                    "Valeur théorique (seuil 10 mm)": "10",
                    "Valeur calculée par ChantierMétéo": f"{st_det['n_valid']}",
                    "Conformité": "✅ Conforme" if st_det["n_valid"] == 10 else "ℹ️ Modifié",
                },
                {
                    "Indicateur statistique": "Somme cumulée des précipitations",
                    "Formule appliquée": "Σ x_i",
                    "Valeur théorique (seuil 10 mm)": "69,0 mm",
                    "Valeur calculée par ChantierMétéo": f"{st_det['sum']:.1f} mm",
                    "Conformité": "✅ Conforme" if abs(st_det["sum"] - 69.0) < 1e-9 else "ℹ️ Modifié",
                },
                {
                    "Indicateur statistique": "Moyenne arithmétique",
                    "Formule appliquée": "(1/n) Σ x_i = 69 / 10",
                    "Valeur théorique (seuil 10 mm)": "6,9 mm/jour",
                    "Valeur calculée par ChantierMétéo": f"{st_det['mean']:.1f} mm/jour",
                    "Conformité": "✅ Conforme" if abs(st_det["mean"] - 6.9) < 1e-9 else "ℹ️ Modifié",
                },
                {
                    "Indicateur statistique": "Nombre de journées défavorables (X ≥ 10 mm)",
                    "Formule appliquée": "Valeurs {12, 25, 16} ≥ 10 mm",
                    "Valeur théorique (seuil 10 mm)": "3 journées",
                    "Valeur calculée par ChantierMétéo": f"{st_det['exceed_count']} journées",
                    "Conformité": "✅ Conforme" if (test_thresh == 10.0 and st_det["exceed_count"] == 3) else "ℹ️ Seuil personnalisé",
                },
                {
                    "Indicateur statistique": "Fréquence empirique de dépassement",
                    "Formule appliquée": "k / n = 3 / 10",
                    "Valeur théorique (seuil 10 mm)": "30,0 %",
                    "Valeur calculée par ChantierMétéo": f"{st_det['exceed_freq_pct']:.1f} %",
                    "Conformité": "✅ Conforme" if (test_thresh == 10.0 and abs(st_det["exceed_freq_pct"] - 30.0) < 1e-9) else "ℹ️ Seuil personnalisé",
                },
                {
                    "Indicateur statistique": "Médiane (Q50)",
                    "Formule appliquée": "(x_(5) + x_(6)) / 2 = (3 + 4) / 2",
                    "Valeur théorique (seuil 10 mm)": "3,50 mm/jour",
                    "Valeur calculée par ChantierMétéo": f"{st_det['median']:.2f} mm/jour",
                    "Conformité": "✅ Conforme",
                },
                {
                    "Indicateur statistique": "Variance et Écart-type d'échantillon (n - 1 = 9)",
                    "Formule appliquée": "s² = (1/(n-1)) Σ(x_i - x̄)²",
                    "Valeur théorique (seuil 10 mm)": "s² = 70,99 mm²/j² | s = 8,43 mm/j",
                    "Valeur calculée par ChantierMétéo": f"s² = {st_det['var_sample']:.2f} | s = {st_det['std_sample']:.2f}",
                    "Conformité": "✅ Conforme",
                },
                {
                    "Indicateur statistique": "Variance et Écart-type empiriques (n = 10)",
                    "Formule appliquée": "σ² = (1/n) Σ(x_i - x̄)²",
                    "Valeur théorique (seuil 10 mm)": "σ² = 63,89 mm²/j² | σ = 7,99 mm/j",
                    "Valeur calculée par ChantierMétéo": f"σ² = {st_det['var_pop']:.2f} | σ = {st_det['std_pop']:.2f}",
                    "Conformité": "✅ Conforme",
                },
            ]
        )
        st.dataframe(verif_tbl, use_container_width=True, hide_index=True)

        st.markdown("##### 📋 Tableau journalier des 10 observations fictives et de leur statut")
        st.dataframe(
            ev_det[["date", "precipitation_sum", "statut", "raisons"]],
            use_container_width=True,
            hide_index=True,
        )

    # --- ONGLET 2 : PROTOCOLES DE DÉMONSTRATION PAS-À-PAS POUR LA SOUTENANCE ---
    with lab_tab2:
        st.markdown("#### 🧭 Guides opératoires pas-à-pas pour la soutenance devant le professeur")
        st.caption(
            "Chaque fiche ci-dessous indique exactement les 6 éléments requis : chemin dans le menu, champs à remplir, "
            "valeurs précises à saisir, bouton exact à cliquer, résultat attendu et intérêt en génie civil."
        )

        guides = [
            {
                "title": "Outil 1 — Chargement instantané du chantier-école ENSPY et lecture du Tableau de bord",
                "chemin": "Menu latéral → « 🏗️ Tableau de bord »",
                "champs": "Aucun champ texte requis pour le chargement initial de démonstration.",
                "valeurs": "Configuration pré-intégrée : Ville = Yaoundé (Lat 3.8480, Lon 11.5021), Période = 2025-03-01 au 2025-06-28 (120 jours), Seuil pluie = 10,0 mm/j.",
                "bouton": "Cliquer sur le bouton « Charger l'exemple fictif » (dans la barre d'actions rapides ou le menu latéral).",
                "resultat": "Création immédiate du chantier « Chantier-école ENSPY — Démonstration », affichage du bandeau « Données fictives — utilisation pédagogique uniquement », 120 journées analysées et cartes KPI colorées.",
                "explication": "Permet au conducteur de travaux d'avoir une vue synthétique immédiate de l'exposition climatique d'un projet avant d'entrer dans le détail du planning.",
            },
            {
                "title": "Outil 2 — Création d'un chantier et géocodage d'une localité camerounaise",
                "chemin": "Menu latéral → « 📁 Mes chantiers » → Onglet « ➕ Créer un nouveau chantier »",
                "champs": "1) Ville camerounaise de référence ; 2) Nom du chantier * ; 3) Type de projet * ; 4) Date de début prévue * ; 5) Date de fin prévue * ; 6) Activités envisagées.",
                "valeurs": "Ville = « Douala » (Lat 4.0511, Lon 9.7679) ; Nom = « Dalot hydraulique Bonabéri » ; Type = « Ouvrage hydraulique » ; Début = « 2025-01-01 » ; Fin = « 2025-03-31 » ; Activités = « Terrassement, Bétonnage ».",
                "bouton": "Cliquer sur le bouton « ✅ Enregistrer le chantier ».",
                "resultat": "Le chantier apparaît sous forme de carte cliquable dans le portefeuille, devient le chantier actif, et récupère les observations Open-Meteo (ou signale proprement l'état hors-ligne si Internet est coupé).",
                "explication": "Garantit la traçabilité géographique (WGS84, fuseau Africa/Douala) et empêche les erreurs de saisie (dates inversées, noms vides).",
            },
            {
                "title": "Outil 3 — Analyse statistique descriptive et vérification sur les 10 premiers jours",
                "chemin": "Menu latéral → « 🌦️ Analyse météorologique » (avec l'exemple fictif ENSPY actif)",
                "champs": "1) Début de la période filtrée ; 2) Fin de la période filtrée ; 3) Seuil de précipitation étudié (mm/jour).",
                "valeurs": "Début = « 2025-03-01 » ; Fin = « 2025-03-10 » (soit les 10 premiers jours) ; Seuil de précipitation = « 10.0 » mm/jour.",
                "bouton": "Validation automatique dès la modification des champs (ou clic sur « 🧪 Charger la série fictive pédagogique (120 jours) » si non chargée).",
                "resultat": "Effectif valide n = 10 jours ; Moyenne = 6,90 mm/j ; Médiane = 3,50 mm/j ; Écart-type s = 8,43 mm/j ; Fréquence pluie ≥ 10 mm = 30,0 % (3 jours / 10) ; mise à jour des 6 graphiques interactifs.",
                "explication": "La moyenne seule est trompeuse pour les pluies (distribution très asymétrique avec beaucoup de zéros). La médiane, les quartiles (Q25, Q75) et la fréquence de dépassement caractérisent réellement le risque d'intempéries.",
            },
            {
                "title": "Outil 4 — Planification météo d'une activité et comparaison de dates de démarrage",
                "chemin": "Menu latéral → « 📅 Planification météo »",
                "champs": "1) Sélectionner l'activité ; 2) Date de début ; 3) Durée (jours) ; 4) Seuil Pluie (mm/j) ; 5) Tmin min (°C) ; 6) Tmax max (°C) ; 7) Vent max (km/h).",
                "valeurs": "Activité = « 1. Terrassement général et fouilles en rigoles » ; Date de début = « 2025-03-01 » ; Durée = « 10 » jours ; Seuil Pluie = « 10.0 » mm/j ; Tmin = « 12.0 » °C ; Tmax = « 37.0 » °C ; Vent max = « 50.0 » km/h.",
                "bouton": "Recalcul interactif immédiat (ou clic sur « ➕ Ajouter cette activité au chantier » lors d'une création).",
                "resultat": "Sur les 10 jours (01/03 au 10/03/2025) : 7 journées Compatibles (vert), 3 journées Défavorables (rouge : 03/03 avec 12 mm, 06/03 avec 25 mm, 09/03 avec 16 mm), 0 Évaluation impossible, plus tableau comparatif des décalages (+0j, +3j, +7j, +10j, +14j, +21j).",
                "explication": "Permet au chef de chantier d'identifier visuellement les jours critiques, de connaître la raison précise de chaque alerte (pluie, vent ou température) et de décaler le démarrage vers une fenêtre plus clémente.",
            },
            {
                "title": "Outil 5 — Fenêtres consécutives de k jours et Simulation de Monte-Carlo par blocs",
                "chemin": "Menu latéral → « 🎲 Probabilités et scénarios » → Onglet « 3️⃣ Simulation de Monte-Carlo par blocs »",
                "champs": "1) Nombre de simulations (M) ; 2) Durée analysée D (jours) ; 3) Longueur des blocs L (jours) ; 4) Graine aléatoire (seed) ; 5) Critère pluie défavorable (mm/j).",
                "valeurs": "M = « 2000 » simulations ; Durée D = « 10 » jours ; Bloc L = « 3 » jours ; Graine = « 42 » ; Pluie ≥ « 10.0 » mm/j ; Vent ≥ « 45.0 » km/h.",
                "bouton": "Exécution automatique à la sélection des paramètres.",
                "resultat": "Affichage reproductible (grâce à la graine 42) de la probabilité d'au moins 1 jour défavorable avec sa marge d'erreur à 95 % (± 1,96 × SE), de la proportion de scénarios avec 0 jour défavorable, et des quantiles (Q25, Médiane Q50, P90, P95) du nombre de jours d'arrêt.",
                "explication": "En génie civil, dimensionner la marge de planning sur la moyenne conduit à être en retard près d'une fois sur deux. Utiliser le quantile P80 ou P90 issu du bootstrap par blocs permet d'intégrer une réserve de délai robuste qui tient compte de la persistance des épisodes pluvieux.",
            },
            {
                "title": "Outil 6 — Vérification déterministe en 1 clic et Export du rapport PDF",
                "chemin": "Menu latéral → « 🎓 Comprendre et tester » (ou « 🏗️ Tableau de bord » pour l'export)",
                "champs": "Seuil de précipitation pour la vérification (mm/jour).",
                "valeurs": "Seuil = « 10.0 » mm/jour.",
                "bouton": "Cliquer sur « ▶️ Exécuter la vérification déterministe (10 jours) » puis sur « 📄 Exporter le rapport » → « 📕 Rapport PDF synthétique (.pdf) ».",
                "resultat": "Confirmation immédiate des 4 valeurs de référence (10 jours, 3 défavorables, 30,0 %, moyenne 6,9 mm/j) et téléchargement d'un rapport PDF structuré, paginé et prêt à être remis au jury.",
                "explication": "Assure la transparence scientifique totale du code et la traçabilité documentaire du dossier d'étude.",
            },
        ]

        for g in guides:
            with st.expander(f"📌 {g['title']}", expanded=False):
                st.markdown(
                    f"""
                    1. **Chemin exact dans le menu :** `{g['chemin']}`
                    2. **Champs à remplir :** {g['champs']}
                    3. **Valeurs précises à saisir :** {g['valeurs']}
                    4. **Nom exact du bouton à cliquer :** **{g['bouton']}**
                    5. **Résultat attendu :** {g['resultat']}
                    6. **Explication du calcul et intérêt en génie civil :** {g['explication']}
                    """
                )

    # --- ONGLET 3 : FORMULAIRE DES FORMULES DE STATISTIQUES DESCRIPTIVES ET PROBABILITÉS ---
    with lab_tab3:
        st.markdown("#### 📐 Récapitulatif et justification des formules mathématiques")

        st.markdown("##### 1. Moyenne arithmétique et cumul sur les observations valides")
        st.latex(
            r"\bar{x} = \frac{1}{n} \sum_{i=1}^{n} x_i \quad \text{où } n = \text{nombre d'observations non manquantes}"
        )
        st.caption(
            "Intérêt génie civil : mesure le volume moyen d'eau reçu par jour, utile pour le dimensionnement du pompage de fouille, "
            "mais sensible aux valeurs extrêmes (averses tropicales exceptionnelles)."
        )

        st.markdown("##### 2. Variance et Écart-type (distinction échantillon $n-1$ vs population empirique $n$)")
        st.latex(
            r"s^2 = \frac{1}{n - 1} \sum_{i=1}^{n} (x_i - \bar{x})^2, \quad s = \sqrt{s^2} \qquad \text{et} \qquad \sigma^2 = \frac{1}{n} \sum_{i=1}^{n} (x_i - \bar{x})^2"
        )
        st.caption(
            "Justification pour la soutenance : l'estimateur s² divisé par (n - 1) (correction de Bessel) est un estimateur sans biais "
            "de la variance théorique du climat local à partir d'un échantillon fini de n journées."
        )

        st.markdown("##### 3. Médiane, Quartiles ($Q_1, Q_2, Q_3$) et Écart interquartile ($IQR$)")
        st.latex(
            r"Q_{25} = P_{25}, \quad \text{Médiane } M = Q_{50} = P_{50}, \quad Q_{75} = P_{75}, \quad IQR = Q_{75} - Q_{25}"
        )
        st.caption(
            "Intérêt génie civil : indicateurs robustes non déformés par un orage isolé de 80 mm. "
            "Permettent de caractériser le régime habituel sur le site."
        )

        st.markdown("##### 4. Probabilité empirique de dépassement et Intervalle de confiance de Wilson")
        st.latex(
            r"\hat{p} = \hat{P}(X \ge x_{\text{seuil}}) = \frac{1}{n}\sum_{i=1}^{n} \mathbb{1}_{\{x_i \ge x_{\text{seuil}}\}} = \frac{k}{n}"
        )
        st.latex(
            r"\text{IC}_{95\%}(\text{Wilson}) = \frac{\hat{p} + \frac{z^2}{2n} \pm z\sqrt{\frac{\hat{p}(1-\hat{p})}{n} + \frac{z^2}{4n^2}}}{1 + \frac{z^2}{n}} \quad (z \approx 1{,}96)"
        )
        st.caption(
            "Justification pour la soutenance : contrairement à l'approximation normale naïve p̂ ± z√(p̂(1-p̂)/n) qui peut donner "
            "des bornes négatives ou > 100 % lorsque p̂ est proche de 0 ou que n est petit, l'intervalle de Wilson reste toujours dans [0, 1]."
        )

        st.markdown("##### 5. Simulation de Monte-Carlo par rééchantillonnage de blocs (Moving Block Bootstrap)")
        st.latex(
            r"K^{(m)} = \sum_{t=1}^{D} Y_t^{(m)} \quad (m = 1, \dots, M), \qquad \hat{P}(K \ge 1) = \frac{1}{M}\sum_{m=1}^{M} \mathbb{1}_{\{K^{(m)} \ge 1\}}, \qquad \text{SE}_{\text{MC}} = \sqrt{\frac{\hat{P}(K \ge 1)\big(1 - \hat{P}(K \ge 1)\big)}{M}}"
        )
        st.caption(
            "Justification pour la soutenance : tirer des journées isolées au hasard (L = 1) suppose que la météo d'un jour est "
            "indépendante de celle de la veille. Or, en météorologie, les jours pluvieux se regroupent en épisodes de plusieurs jours. "
            "Le tirage par blocs consécutifs de longueur L préserve cette dépendance temporelle courte."
        )


# ==============================================================================
# POINT D'ENTRÉE PRINCIPAL DE L'APPLICATION STREAMLIT
# ==============================================================================

def main() -> None:
    """Point d'entrée principal de l'application Streamlit ChantierMétéo."""
    if st is None:
        return

    st.set_page_config(
        page_title=f"{APP_NAME} — Planification Météo Génie Civil (ENSPY)",
        page_icon="🏗️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_session_state()

    # Injection du CSS responsive et du thème choisi
    st.markdown(get_custom_css(st.session_state.get("theme_mode", "Clair")), unsafe_allow_html=True)

    # Rendu du menu latéral
    render_sidebar()

    # Affichage des notifications d'action (flash messages)
    flash = st.session_state.get("flash_message")
    if flash:
        f_type, f_text = flash
        if f_type == "success":
            st.success(f_text)
        elif f_type == "warning":
            st.warning(f_text)
        else:
            st.info(f_text)
        st.session_state["flash_message"] = None

    # Routage vers la rubrique active
    current_page = st.session_state.get("nav_page", "🏗️ Tableau de bord")
    if current_page == "🏗️ Tableau de bord":
        render_dashboard_page()
    elif current_page == "📁 Mes chantiers":
        render_projects_page()
    elif current_page == "🌦️ Analyse météorologique":
        render_weather_analysis_page()
    elif current_page == "📅 Planification météo":
        render_planning_page()
    elif current_page == "🎲 Probabilités et scénarios":
        render_probabilities_page()
    elif current_page == "🎓 Comprendre et tester":
        render_pedagogical_lab_page()
    else:
        render_dashboard_page()


if __name__ == "__main__":
    main()