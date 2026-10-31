"""ChantierMétéo — application pédagogique de planification météo des chantiers.
Python 3.11+, Streamlit. Toutes les données utilisateur résident en session ;
la sauvegarde persistante est un export JSON volontaire.
"""
from __future__ import annotations

import html
import hashlib
import io
import json
import math
import uuid
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

st.set_page_config(page_title="ChantierMétéo | Analyse & planification", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

# --------------------------- CONFIGURATION -------------------------------
TZ = ZoneInfo("Africa/Douala")
FIELDS = {
    "precip_mm": "Précipitations (mm/j)",
    "tmin_c": "Température min. (°C)",
    "tmax_c": "Température max. (°C)",
    "wind_kmh": "Vent maximum (km/h)",
    "humidity_pct": "Humidité moyenne (%)",
    "rain_probability_pct": "Probabilité de pluie prévisionnelle (%)",
}
UNITS = {"precip_mm": "mm/j", "tmin_c": "°C", "tmax_c": "°C", "wind_kmh": "km/h", "humidity_pct": "%", "rain_probability_pct": "%"}
NUM_COLS = list(FIELDS)
NAV = ["🏠 Tableau de bord", "🏗️ Mes chantiers", "🌦️ Analyse météorologique", "📅 Planification météo", "🎲 Probabilités et scénarios", "🎓 Comprendre et tester"]
TYPE_OPTIONS = ["Bâtiment", "Route", "Ouvrage hydraulique", "Terrassement", "Autre"]
ACT_OPTIONS = ["Terrassement", "Bétonnage", "Travaux de voirie", "Peinture extérieure", "Opérations de levage", "Travaux d’assainissement", "Activité personnalisée"]
PALETTE = {"blue":"#328CCB", "orange":"#F29A38", "green":"#218A70", "red":"#D85B56", "navy":"#10243A"}
DEMO_ID = "chantier-exemple-fictif"


def today() -> date:
    return datetime.now(TZ).date()


def sid() -> str:
    return uuid.uuid4().hex[:14]


def initial_state():
    defaults = {"projects": {}, "selected_id": None, "nav": NAV[0], "theme": "Clair", "editing": False, "demo_loaded": False, "editing_activity": None}
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def jump(label: str):
    st.session_state["go_to"] = label
    st.rerun()


def active_project() -> dict | None:
    return st.session_state.projects.get(st.session_state.selected_id)


def df_project(project: dict | None) -> pd.DataFrame:
    if not project or not project.get("records"):
        return pd.DataFrame(columns=["date"] + NUM_COLS)
    return validate_weather(pd.DataFrame(project["records"]), strict=False)


def save_project(p: dict):
    st.session_state.projects[p["id"]] = p
    st.session_state.selected_id = p["id"]
    st.session_state.sync_project_picker = True


def add_demo():
    base = date(2026, 10, 20)
    rain = [0, 4, 12, 0, 8, 25, 1, 0, 16, 3]
    data = []
    for i, r in enumerate(rain):
        data.append({"date": (base + timedelta(days=i)).isoformat(), "precip_mm": float(r), "tmin_c": float(19 + i % 3), "tmax_c": float(27 + i % 4), "wind_kmh": float(12 + (i * 4) % 19), "humidity_pct": float(68 + i % 12), "rain_probability_pct": None})
    p = {"id": DEMO_ID, "name": "Chantier-école ENSPY — Démonstration", "type": "Bâtiment", "city": "Yaoundé", "latitude": 3.8667, "longitude": 11.5167, "start": base.isoformat(), "end": (base + timedelta(days=9)).isoformat(), "description": "Projet de bâtiment universitaire entièrement fictif.", "demo": True, "activities": [
        {"id": "demo-ter", "name": "Terrassement", "start": base.isoformat(), "duration": 5, "rain": 10., "tmin": None, "tmax": None, "wind": None},
        {"id": "demo-bet", "name": "Bétonnage", "start": (base + timedelta(days=5)).isoformat(), "duration": 5, "rain": 10., "tmin": None, "tmax": None, "wind": None}],
        "source": "fictif", "records": data, "retrieved_at": None, "data_note": "Série inventée pour l'enseignement, pas une mesure à Yaoundé."}
    save_project(p)
    st.session_state.demo_loaded = True


def remove_demo():
    st.session_state.projects.pop(DEMO_ID, None)
    st.session_state.demo_loaded = False
    if st.session_state.selected_id == DEMO_ID:
        st.session_state.selected_id = next(iter(st.session_state.projects), None)
    st.session_state.sync_project_picker = True


def validate_weather(input_df: pd.DataFrame, strict: bool = True) -> pd.DataFrame:
    """ISO dates, daily numeric measurements; refuse impossible/duplicate rows.
    strict=True requires ISO YYYY-MM-DD dates and useful numeric columns.
    """
    d = input_df.copy()
    d.columns = [str(x).strip().lower() for x in d.columns]
    aliases = {"time":"date", "jour":"date", "precipitation_sum":"precip_mm", "precipitations":"precip_mm", "pluie_mm":"precip_mm", "temperature_2m_min":"tmin_c", "temperature_2m_max":"tmax_c", "wind_speed_10m_max":"wind_kmh", "relative_humidity_2m_mean":"humidity_pct", "precipitation_probability_max":"rain_probability_pct"}
    d = d.rename(columns={k: v for k, v in aliases.items() if k in d.columns and v not in d.columns})
    if "date" not in d.columns:
        raise ValueError("Colonne 'date' absente (format AAAA-MM-JJ).")
    if strict and not any(col in d.columns for col in NUM_COLS):
        raise ValueError("Ajoutez au moins une colonne météo reconnue (ex. precip_mm).")
    raw_dates = d["date"].astype(str).str.strip()
    if strict and not raw_dates.str.fullmatch(r"\d{4}-\d{2}-\d{2}").all():
        raise ValueError("Les dates doivent être au format ISO AAAA-MM-JJ.")
    dt = pd.to_datetime(raw_dates, errors="coerce", format="%Y-%m-%d")
    if dt.isna().any():
        raise ValueError("Date invalide détectée.")
    d["date"] = dt.dt.strftime("%Y-%m-%d")
    if d["date"].duplicated().any():
        raise ValueError("Dates en double : une seule ligne par jour est attendue.")
    for col in NUM_COLS:
        if col not in d:
            d[col] = np.nan
        else:
            values = d[col].replace("", np.nan)
            parsed = pd.to_numeric(values, errors="coerce")
            if strict and (values.notna() & parsed.isna()).any():
                raise ValueError(f"Valeur non numérique dans {col}.")
            d[col] = parsed.astype(float).replace([np.inf, -np.inf], np.nan)
    if (d["precip_mm"].dropna() < 0).any() or (d["wind_kmh"].dropna() < 0).any():
        raise ValueError("La pluie et le vent ne peuvent pas être négatifs.")
    for col in ("humidity_pct", "rain_probability_pct"):
        if ((d[col].dropna() < 0) | (d[col].dropna() > 100)).any():
            raise ValueError(f"{col} doit appartenir à [0 ; 100].")
    both = d["tmin_c"].notna() & d["tmax_c"].notna()
    if (d.loc[both, "tmin_c"] > d.loc[both, "tmax_c"]).any():
        raise ValueError("Température minimale supérieure à la maximale.")
    d = d[["date"] + NUM_COLS].sort_values("date").reset_index(drop=True)
    return d


def df_records(df: pd.DataFrame) -> list[dict]:
    d = df.replace({np.nan: None})
    return json.loads(d.to_json(orient="records", date_format="iso"))


def summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col, label in FIELDS.items():
        vals = df[col].dropna().astype(float)
        if vals.empty:
            continue
        rows.append({"Variable": label, "n": len(vals), "Moyenne": round(vals.mean(), 2), "Médiane": round(vals.median(), 2), "Minimum": round(vals.min(), 2), "Maximum": round(vals.max(), 2), "Écart-type (échantillon)": round(vals.std(ddof=1), 2) if len(vals) > 1 else None, "Variance (échantillon)": round(vals.var(ddof=1), 2) if len(vals) > 1 else None, "Q1": round(vals.quantile(.25), 2), "Q3": round(vals.quantile(.75), 2)})
    return pd.DataFrame(rows)


def rain_exceedance(df: pd.DataFrame, threshold: float) -> tuple[int, int, float | None]:
    vals = df["precip_mm"].dropna()
    n = len(vals)
    k = int((vals >= threshold).sum())
    return k, n, k / n if n else None


def classify_weather(df: pd.DataFrame, activity: dict) -> pd.DataFrame:
    """Precedence: proven violation -> adverse; otherwise any unknown -> unknown.
    A weather limit on max wind or max temperature is violated strictly above;
    minimum temperature is violated strictly below; rain uses >=.
    """
    out = df.copy()
    states, reasons = [], []
    criteria = [("precip_mm", activity.get("rain"), ">=", "pluie"), ("tmin_c", activity.get("tmin"), "<", "température minimale"), ("tmax_c", activity.get("tmax"), ">", "température maximale"), ("wind_kmh", activity.get("wind"), ">", "vent")]
    for _, row in out.iterrows():
        bad, missing = [], []
        for col, limit, op, desc in criteria:
            if limit is None:
                continue
            val = row[col]
            if pd.isna(val):
                missing.append(desc)
            elif (op == ">=" and val >= limit) or (op == "<" and val < limit) or (op == ">" and val > limit):
                bad.append(desc)
        if bad:
            states.append("Défavorable")
            reasons.append("Dépassement : " + ", ".join(bad) + (" ; données manquantes : " + ", ".join(missing) if missing else ""))
        elif missing:
            states.append("Indéterminé")
            reasons.append("Données absentes : " + ", ".join(missing))
        else:
            states.append("Compatible")
            reasons.append("Tous les critères définis sont satisfaits")
    out["Statut"] = states
    out["Motif"] = reasons
    return out


def calendar_for_activity(df: pd.DataFrame, activity: dict, start: date | None = None) -> pd.DataFrame:
    """Preserve unavailable days by joining full daily calendar, no extrapolation."""
    start = start or date.fromisoformat(activity["start"])
    n = int(activity["duration"])
    days = pd.DataFrame({"date": [(start + timedelta(days=i)).isoformat() for i in range(n)]})
    joined = days.merge(df, on="date", how="left", validate="one_to_one")
    return classify_weather(joined, activity)


def available_windows(df: pd.DataFrame, activity: dict, days: int) -> tuple[int, int, float | None]:
    """Rolling *calendar* windows, reject missing dates and incomplete criteria."""
    if days < 1 or df.empty:
        return 0, 0, None
    cls = classify_weather(df, activity).copy()
    cls["dt"] = pd.to_datetime(cls["date"])
    cls = cls.set_index("dt").asfreq("D")
    values = cls["Statut"]
    valid_count = 0
    event_count = 0
    for i in range(len(values) - days + 1):
        block = values.iloc[i:i + days]
        if block.isna().any() or (block == "Indéterminé").any():
            continue
        valid_count += 1
        event_count += bool((block == "Défavorable").any())
    return event_count, valid_count, event_count / valid_count if valid_count else None


def valid_blocks(df: pd.DataFrame, activity: dict, length: int) -> np.ndarray:
    """Complete contiguous blocks of 0/1 weather indicators."""
    cl = classify_weather(df, activity)
    cl["dt"] = pd.to_datetime(cl["date"])
    s = cl.set_index("dt")["Statut"].asfreq("D")
    blocks = []
    for i in range(max(0, len(s) - length + 1)):
        b = s.iloc[i:i + length]
        if b.notna().all() and not (b == "Indéterminé").any():
            blocks.append((b == "Défavorable").astype(int).to_numpy())
    return np.asarray(blocks, dtype=int).reshape((-1, length)) if blocks else np.empty((0, length), dtype=int)


def monte_carlo(df: pd.DataFrame, activity: dict, horizon: int, trials: int, block: int, seed: int) -> np.ndarray:
    if min(horizon, trials, block) < 1:
        raise ValueError("La durée, le nombre de simulations et la longueur du bloc doivent être positifs.")
    if block > horizon:
        raise ValueError("La taille du bloc ne peut dépasser l'horizon simulé.")
    blocks = valid_blocks(df, activity, block)
    if len(blocks) == 0:
        raise ValueError("Pas assez de journées consécutives entièrement renseignées pour cette taille de bloc.")
    rng = np.random.default_rng(seed)
    nb = math.ceil(horizon / block)
    draws = rng.integers(0, len(blocks), size=(trials, nb))
    simulated = blocks[draws].reshape(trials, nb * block)[:, :horizon]
    return simulated.sum(axis=1)


def set_weather(p: dict, df: pd.DataFrame, source: str, note: str):
    p["records"] = df_records(df)
    p["source"] = source
    p["data_note"] = note
    p["retrieved_at"] = datetime.now(TZ).isoformat(timespec="seconds")
    save_project(p)


# ---------------------------- OPEN-METEO --------------------------------
@st.cache_data(ttl=1800, show_spinner=False)
def geocode(query: str) -> list[dict]:
    r = requests.get("https://geocoding-api.open-meteo.com/v1/search", params={"name": query, "count": 10, "language": "fr", "format": "json"}, timeout=15)
    r.raise_for_status()
    payload = r.json()
    if payload.get("error"):
        raise ValueError(payload.get("reason", "Recherche géographique impossible."))
    return payload.get("results", [])


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_weather(lat: float, lon: float, mode: str, start: str = "", end: str = "") -> pd.DataFrame:
    is_history = mode == "historique"
    endpoint = "https://archive-api.open-meteo.com/v1/archive" if is_history else "https://api.open-meteo.com/v1/forecast"
    daily = ["precipitation_sum", "temperature_2m_min", "temperature_2m_max", "wind_speed_10m_max"]
    if is_history:
        daily += ["relative_humidity_2m_mean"]
    else:
        daily += ["precipitation_probability_max"]
    params = {"latitude": round(float(lat), 5), "longitude": round(float(lon), 5), "daily": ",".join(daily), "timezone": "Africa/Douala", "wind_speed_unit": "kmh", "temperature_unit": "celsius", "precipitation_unit": "mm"}
    if is_history:
        params.update({"start_date": start, "end_date": end})
    else:
        params["forecast_days"] = 16
    response = requests.get(endpoint, params=params, timeout=25)
    response.raise_for_status()
    payload = response.json()
    if payload.get("error"):
        raise ValueError(payload.get("reason", "Réponse Open-Meteo invalide."))
    daily_payload = payload.get("daily")
    if not daily_payload or not daily_payload.get("time"):
        raise ValueError("Open-Meteo n'a renvoyé aucune journée.")
    return validate_weather(pd.DataFrame(daily_payload), strict=True)


# ------------------------- EXPORT / BACKUP -------------------------------
def make_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def make_excel(df: pd.DataFrame, project: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Donnees")
        stats = summary(df)
        if not stats.empty:
            stats.to_excel(writer, index=False, sheet_name="Statistiques")
        info = pd.DataFrame([{"Chantier": project["name"], "Ville": project["city"], "Source": project.get("source", "aucune"), "Note": project.get("data_note", "")}])
        info.to_excel(writer, index=False, sheet_name="Provenance")
        for sheet in writer.sheets.values():
            sheet.freeze_panes = "A2"
            for cells in sheet.columns:
                letter = cells[0].column_letter
                width = min(48, max(14, max(len(str(cell.value or "")) for cell in cells[:250]) + 2))
                sheet.column_dimensions[letter].width = width
            for c in sheet[1]:
                c.font = c.font.copy(bold=True, color="FFFFFF")
                from openpyxl.styles import PatternFill
                c.fill = PatternFill("solid", fgColor="10243A")
    return buf.getvalue()


def pdf_font() -> str:
    candidates = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "C:/Windows/Fonts/arial.ttf", "/Library/Fonts/Arial.ttf"]
    for path in candidates:
        try:
            pdfmetrics.registerFont(TTFont("ChantierUnicode", path))
            return "ChantierUnicode"
        except Exception:
            pass
    return "Helvetica"


def make_pdf(p: dict, df: pd.DataFrame, activity: dict | None = None) -> bytes:
    buf = io.BytesIO()
    font = pdf_font()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CTitle", parent=styles["Title"], fontName=font, fontSize=18, leading=24, textColor=colors.HexColor("#10243A"), spaceAfter=12))
    styles.add(ParagraphStyle(name="CBody", parent=styles["BodyText"], fontName=font, fontSize=9, leading=14, spaceAfter=8))
    styles.add(ParagraphStyle(name="CH2", parent=styles["Heading2"], fontName=font, fontSize=12, textColor=colors.HexColor("#328CCB"), spaceBefore=10, spaceAfter=7))
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=19 * mm, rightMargin=19 * mm, topMargin=20 * mm, bottomMargin=20 * mm)
    story = []
    def par(t: str, style="CBody"):
        story.append(Paragraph(html.escape(str(t)), styles[style]))
    def tab(rows: list[list], widths: list[float] | None = None):
        header_style = ParagraphStyle(name="pdf_head", parent=styles["CBody"], fontName=font, textColor=colors.white, fontSize=8)
        cells = [[Paragraph(html.escape(str(x)), header_style if j == 0 else styles["CBody"]) for x in row] for j, row in enumerate(rows)]
        table = Table(cells, colWidths=widths, repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#10243A")), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6FA")]), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 6), ("LINEBELOW", (0, 0), (-1, 0), .5, colors.HexColor("#328CCB"))]))
        story.append(table)
        story.append(Spacer(1, 4 * mm))
    par("CHANTIERMÉTÉO — RAPPORT D'ANALYSE", "CTitle")
    par("Rapport généré le " + datetime.now(TZ).strftime("%d/%m/%Y à %H:%M") + " (Africa/Douala).")
    par(f"Chantier : {p['name']} | Localisation : {p['city']} | Type : {p['type']}")
    par(f"Source : {p.get('source', 'aucune')} | {p.get('data_note', '')}")
    if p.get("demo"):
        par("ATTENTION : DONNÉES FICTIVES — UTILISATION PÉDAGOGIQUE UNIQUEMENT.")
    story.append(Paragraph("Statistiques des mesures", styles["CH2"]))
    par(f"Effectif total : {len(df)} journées ; période : {df['date'].min() if not df.empty else '—'} au {df['date'].max() if not df.empty else '—'}.")
    stats = summary(df)
    if not stats.empty:
        rows = [["Variable", "n", "Moyenne", "Minimum", "Maximum"]]
        for _, r in stats.iterrows():
            rows.append([r["Variable"], r["n"], r["Moyenne"], r["Minimum"], r["Maximum"]])
        tab(rows, [60 * mm, 14 * mm, 29 * mm, 29 * mm, 29 * mm])
    else:
        par("Aucune mesure exploitable.")
    if activity is not None:
        story.append(Paragraph("Classification de l'activité", styles["CH2"]))
        c = calendar_for_activity(df, activity)
        counts = c["Statut"].value_counts()
        par(f"Activité : {activity['name']} ; durée : {activity['duration']} jours ; compatible : {counts.get('Compatible', 0)} ; défavorable : {counts.get('Défavorable', 0)} ; indéterminé : {counts.get('Indéterminé', 0)}.")
        tab([["Jour", "Statut", "Motif"]] + [[row["date"], row["Statut"], row["Motif"]] for _, row in c.head(60).iterrows()], [29 * mm, 32 * mm, 100 * mm])
        if len(c) > 60:
            par("Tableau limité aux 60 premières journées. Exportez le CSV pour le détail intégral.")
    story.append(Paragraph("Méthode et limites", styles["CH2"]))
    par("Moyenne = somme des valeurs / n ; écart-type d'échantillon = racine(somme((x - moyenne)^2)/(n-1)). Les conditions compatibles signifient seulement que les limites configurées par l'utilisateur sont respectées.")
    par("Les données historiques Open-Meteo sont des réanalyses/modèles, non des observations locales garanties. Les prévisions évoluent et ne constituent pas une garantie. Les seuils sont des choix d'utilisateur, pas des normes de sécurité. Journée civile de 00:00 à 24:00, fuseau Africa/Douala.")
    def footer(canvas, document):
        canvas.saveState()
        canvas.setFont(font, 8)
        canvas.setFillColor(colors.HexColor("#607086"))
        canvas.drawString(19 * mm, 13 * mm, "ChantierMétéo • Analyse pédagogique")
        canvas.drawRightString(A4[0] - 19 * mm, 13 * mm, f"Page {document.page}")
        canvas.restoreState()
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def backup_json() -> bytes:
    data = {"version": 1, "saved_at": datetime.now(TZ).isoformat(), "projects": list(st.session_state.projects.values()), "selected_id": st.session_state.selected_id}
    return json.dumps(data, ensure_ascii=False, indent=2, allow_nan=False).encode("utf-8")


def parse_backup(raw: bytes) -> tuple[dict, str | None]:
    if len(raw) > 8_000_000:
        raise ValueError("Sauvegarde trop volumineuse (maximum 8 Mo).")
    data = json.loads(raw.decode("utf-8-sig"))
    if data.get("version") != 1 or not isinstance(data.get("projects"), list):
        raise ValueError("Sauvegarde JSON incompatible (version 1 attendue).")
    projects = {}
    for p in data["projects"]:
        if not isinstance(p, dict) or not all(k in p for k in ["id", "name", "city", "type", "start", "end", "activities", "records"]):
            raise ValueError("Chantier incomplet dans la sauvegarde.")
        if not isinstance(p["activities"], list) or not isinstance(p["records"], list):
            raise ValueError("Activités ou données mal formées.")
        if len(str(p["name"]).strip()) == 0 or len(p["name"]) > 150:
            raise ValueError("Nom de chantier invalide.")
        a, b = date.fromisoformat(p["start"]), date.fromisoformat(p["end"])
        if a > b:
            raise ValueError("Dates de chantier incohérentes.")
        p["latitude"] = float(p["latitude"])
        p["longitude"] = float(p["longitude"])
        if not (-90 <= p["latitude"] <= 90 and -180 <= p["longitude"] <= 180):
            raise ValueError("Coordonnées invalides.")
        if p["records"]:
            p["records"] = df_records(validate_weather(pd.DataFrame(p["records"]), strict=True))
        for act in p["activities"]:
            if not all(k in act for k in ["id", "name", "start", "duration", "rain", "tmin", "tmax", "wind"]):
                raise ValueError("Activité incomplète.")
            if not (1 <= int(act["duration"]) <= 730):
                raise ValueError("Durée d'activité invalide.")
            date.fromisoformat(act["start"])
        if p["id"] in projects:
            raise ValueError("Identifiant de chantier dupliqué.")
        projects[p["id"]] = p
    chosen = data.get("selected_id")
    return projects, chosen if chosen in projects else next(iter(projects), None)


# ---------------------------- DESIGN SYSTEM -------------------------------
# Control Center : navigation sombre, zone de travail claire et composants lisibles.
# Le CSS est regroupé ici pour respecter la livraison à un seul app.py.
def css(theme: str):
    dark = theme == "Sombre"
    variables = {
        "BG": "#0C1725" if dark else "#F0F4F8",
        "SURFACE": "#14283C" if dark else "#FFFFFF",
        "TEXT": "#EDF5FD" if dark else "#132C41",
        "MUTED": "#B0C2D1" if dark else "#617789",
        "BORDER": "#30465B" if dark else "#DFE7EF",
        "INPUT": "#173047" if dark else "#FFFFFF",
        "CHART": "#14283C" if dark else "#FFFFFF",
        "GRID": "#30485A" if dark else "#E6EDF2",
    }
    styles = r"""
    <style>
      :root {--cm-bg:@@BG@@;--cm-surface:@@SURFACE@@;--cm-text:@@TEXT@@;--cm-muted:@@MUTED@@;--cm-border:@@BORDER@@;--cm-orange:#F4A340;--cm-cyan:#318FBD;--cm-green:#22977A;--cm-red:#D96957;}
      .stApp {background:var(--cm-bg);color:var(--cm-text);font-family:Inter,ui-sans-serif,system-ui,'Segoe UI',sans-serif;}
      .block-container {max-width:1450px;padding-top:1.6rem;padding-bottom:3rem;}
      [data-testid="stHeader"] {background:transparent;}
      #MainMenu,footer {visibility:hidden;}
      h1,h2,h3,h4 {color:var(--cm-text);letter-spacing:-.035em;}
      h3 {font-size:1.25rem!important;margin-top:1.1rem!important;}
      p,li {line-height:1.55;}
      [data-testid="stSidebar"] {background:#10243A!important;border-right:1px solid #2B475F!important;}
      [data-testid="stSidebar"] [data-testid="stSidebarContent"] {background:transparent!important;padding-top:1.1rem;}
      [data-testid="stSidebar"] * {color:#D9E6F1;}
      [data-testid="stSidebar"] hr {border-color:#304A63!important;}
      [data-testid="stSidebar"] .stCaptionContainer, [data-testid="stSidebar"] small {color:#A9BED0!important;}
      [data-testid="stSidebar"] .stRadio [role="radiogroup"] {gap:.35rem;}
      [data-testid="stSidebar"] .stRadio [role="radiogroup"] label {border-radius:11px;padding:.56rem .6rem;border:1px solid transparent;transition:background .15s ease;}
      [data-testid="stSidebar"] .stRadio [role="radiogroup"] label:hover {background:#1C405C;}
      [data-testid="stSidebar"] .stRadio [role="radiogroup"] label:has(input:checked) {background:#23506A;border-color:#396F8A;box-shadow:inset 4px 0 #F4A340;}
      [data-testid="stSidebar"] .stRadio [role="radiogroup"] label [data-testid="stMarkdownContainer"] p {color:#EDF6FF!important;font-size:.87rem;font-weight:650;}
      [data-testid="stSidebar"] .stRadio [role="radiogroup"] label [data-baseweb="radio"] {display:none;}
      [data-testid="stSidebar"] [data-baseweb="select"] > div {background:#1C3A53!important;border-color:#426079!important;}
      [data-testid="stSidebar"] input {color:#FFFFFF!important;}
      .cm-sidebar-brand {display:flex;align-items:center;gap:.7rem;padding:.2rem .15rem 1rem;border-bottom:1px solid #304A63;margin-bottom:1rem;}
      .cm-mark {width:39px;height:39px;border-radius:12px;background:linear-gradient(145deg,#F7B65A,#F18F22);display:grid;place-items:center;color:#12253A;font-size:1rem;font-weight:900;box-shadow:0 5px 15px #030b1640;}
      .cm-brand-title {color:white;font-size:1.12rem;font-weight:850;letter-spacing:-.04em;line-height:1.2;}
      .cm-brand-title b {color:#F4A340;}
      .cm-brand-sub {font-size:.66rem;letter-spacing:.1em;text-transform:uppercase;color:#9DB6C9;margin-top:.21rem;font-weight:750;}
      .cm-sidebar-section {color:#8DA8BD;font-size:.7rem;letter-spacing:.17em;text-transform:uppercase;font-weight:800;margin:.8rem 0 .3rem;}
      .cm-sidebar-foot {border-top:1px solid #304A63;margin-top:1.1rem;padding-top:1rem;color:#9CB3C6;font-size:.72rem;line-height:1.6;}
      .cm-pagehead {padding:.45rem 0 1.2rem;border-bottom:1px solid var(--cm-border);margin-bottom:1.25rem;}
      .cm-eyebrow {font-size:.7rem;letter-spacing:.15em;color:var(--cm-cyan);font-weight:850;text-transform:uppercase;margin-bottom:.38rem;}
      .cm-pagehead h1 {font-size:clamp(1.55rem,2.5vw,2.25rem);font-weight:850;margin:0 0 .35rem;line-height:1.2;}
      .cm-pagehead p {color:var(--cm-muted);font-size:.94rem;margin:0;max-width:940px;}
      .cm-welcome {position:relative;overflow:hidden;border-radius:18px;background:linear-gradient(110deg,#15354F 5%,#22618A 67%,#298DAF 100%);padding:1.55rem 1.9rem;color:#FFFFFF;box-shadow:0 14px 30px #14395022;margin-bottom:1.05rem;min-height:176px;}
      .cm-welcome:before {content:"";position:absolute;right:3%;top:-105px;width:270px;height:270px;border-radius:50%;border:1px solid #FFFFFF32;box-shadow:0 0 0 37px #FFFFFF0D,0 0 0 78px #FFFFFF08;pointer-events:none;}
      .cm-welcome:after {content:"";position:absolute;right:-35px;bottom:-55px;width:310px;height:160px;transform:rotate(-15deg);background:repeating-linear-gradient(90deg,transparent 0 22px,#FFFFFF12 23px 24px);pointer-events:none;}
      .cm-welcome .eyebrow {font-size:.73rem;font-weight:850;letter-spacing:.14em;color:#D3EBF7;position:relative;z-index:1;}
      .cm-welcome h1 {color:#FFFFFF;font-size:clamp(1.6rem,3vw,2.2rem);margin:.5rem 0 .5rem;max-width:770px;font-weight:850;position:relative;z-index:1;line-height:1.15;}
      .cm-welcome p {color:#DAEDF8;max-width:680px;font-size:.95rem;margin:0;position:relative;z-index:1;}
      .cm-welcome .sky {position:absolute;top:25px;right:45px;font-size:3rem;z-index:1;filter:drop-shadow(0 5px 12px #081c2b45);}
      .cm-context {font-size:.81rem;color:var(--cm-muted);padding:0 0 1rem;display:flex;align-items:center;flex-wrap:wrap;gap:.55rem;}
      .cm-chip {display:inline-flex;align-items:center;border:1px solid var(--cm-border);background:var(--cm-surface);padding:.3rem .7rem;border-radius:999px;color:var(--cm-muted);font-size:.76rem;font-weight:720;}
      .cm-kpi-grid {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.85rem;margin:.9rem 0 1.45rem;}
      .cm-kpi {border-radius:15px;padding:1.02rem;background:var(--cm-surface);border:1px solid var(--cm-border);border-top:3px solid var(--kpi-accent);min-height:135px;box-shadow:0 4px 18px #18344C0A;min-width:0;}
      .cm-kpi-head {font-size:.8rem;font-weight:730;color:var(--cm-muted);display:flex;align-items:center;gap:.5rem;}
      .cm-kpi-icon {display:grid;place-items:center;background:color-mix(in srgb,var(--kpi-accent) 12%,transparent);color:var(--kpi-accent);width:29px;height:29px;border-radius:8px;font-size:1.03rem;flex-shrink:0;}
      .cm-kpi-value {font-size:clamp(1.55rem,2.3vw,2.15rem);font-weight:850;letter-spacing:-.05em;color:var(--cm-text);margin:.35rem 0 .15rem;overflow-wrap:anywhere;line-height:1.15;}
      .cm-kpi-note {font-size:.75rem;color:var(--cm-muted);line-height:1.4;}
      .metric {background:var(--cm-surface);border:1px solid var(--cm-border);border-top:3px solid var(--accent);border-radius:15px;padding:1rem;min-height:124px;box-shadow:0 4px 18px #122C4410;}
      .metric .label {color:var(--cm-muted);font-weight:720;font-size:.79rem;line-height:1.3;}
      .metric .num {color:var(--cm-text);font-size:clamp(1.4rem,2vw,2rem);font-weight:850;margin:.5rem 0 .18rem;line-height:1.15;overflow-wrap:anywhere;letter-spacing:-.035em;}
      .metric .sub {color:var(--cm-muted);font-size:.74rem;overflow-wrap:anywhere;}
      .cm-section {display:flex;align-items:center;gap:.7rem;margin:1.35rem 0 .65rem;}
      .cm-section-icon {width:34px;height:34px;background:#318FBD20;color:#318FBD;border-radius:10px;display:grid;place-items:center;flex-shrink:0;}
      .cm-section-title {font-weight:820;color:var(--cm-text);letter-spacing:-.025em;font-size:1.11rem;line-height:1.2;}
      .cm-section-sub {font-size:.8rem;color:var(--cm-muted);margin-top:.18rem;}
      .cm-panel-head {padding:.25rem .12rem .75rem;}
      .cm-panel-head b {display:block;font-size:1.04rem;color:var(--cm-text);}
      .cm-panel-head span {font-size:.78rem;color:var(--cm-muted);}
      .cm-act-row {display:flex;justify-content:space-between;align-items:center;gap:9px;padding:12px 0;border-bottom:1px solid var(--cm-border);}
      .cm-act-row:last-child {border-bottom:0;}
      .cm-act-name {font-size:.86rem;font-weight:800;color:var(--cm-text);}
      .cm-act-meta {font-size:.75rem;color:var(--cm-muted);margin-top:.15rem;}
      .cm-status {display:inline-block;border-radius:8px;padding:.35rem .5rem;font-size:.72rem;font-weight:800;white-space:nowrap;}
      .cm-status-good {background:#DDF4EA;color:#116E54;}
      .cm-status-bad {background:#FBE5DF;color:#A63A31;}
      .cm-status-unknown {background:#FFF0D4;color:#8A591D;}
      .cm-empty {border:1px dashed #91A9B9;background:var(--cm-surface);border-radius:14px;padding:1.4rem;margin:1rem 0;}
      .cm-empty b {display:block;color:var(--cm-text);font-size:1rem;}
      .cm-empty p {font-size:.88rem;color:var(--cm-muted);margin:.4rem 0 0;}
      .cm-calendar {display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:.5rem;margin:.65rem 0;}
      .cm-day {min-width:0;border:1px solid var(--cm-border);background:var(--cm-surface);border-top:4px solid var(--day-color);border-radius:10px;padding:.62rem .52rem;min-height:78px;}
      .cm-day .date {font-size:.75rem;font-weight:800;color:var(--cm-text);}
      .cm-day .rain {font-size:.76rem;color:var(--cm-muted);margin-top:.24rem;}
      .cm-day .state {font-size:.69rem;color:var(--day-color);font-weight:800;margin-top:.32rem;overflow-wrap:anywhere;}
      .cm-legend {display:flex;flex-wrap:wrap;gap:.8rem;font-size:.76rem;color:var(--cm-muted);margin-bottom:.7rem;}
      .cm-legend i {display:inline-block;width:9px;height:9px;border-radius:3px;margin-right:4px;}
      .cm-project-header {display:flex;align-items:center;gap:.6rem;margin-bottom:.65rem;}
      .cm-project-icon {background:#E5F2FA;color:#1B719C;width:38px;height:38px;border-radius:11px;display:grid;place-items:center;font-size:1.1rem;}
      .cm-project-name {font-size:1rem;font-weight:850;color:var(--cm-text);line-height:1.25;overflow-wrap:anywhere;}
      .cm-project-desc {font-size:.8rem;color:var(--cm-muted);line-height:1.55;}
      .cm-project-counts {display:flex;gap:.5rem;flex-wrap:wrap;margin:.65rem 0 .8rem;}
      .cm-project-counts span {background:var(--cm-bg);border:1px solid var(--cm-border);border-radius:7px;padding:.3rem .5rem;color:var(--cm-muted);font-size:.74rem;font-weight:750;}
      .cm-notice {padding:.86rem 1rem;border-left:4px solid var(--cm-cyan);background:var(--cm-surface);border-radius:10px;color:var(--cm-text);font-size:.88rem;margin:10px 0;}
      .cm-demo {background:#FFF0D4;color:#694514;border-left:4px solid #F4A340;border-radius:10px;padding:.85rem 1rem;margin:.8rem 0;font-size:.85rem;font-weight:750;}
      .stButton > button {border-radius:10px;min-height:42px;font-weight:720;border:1px solid var(--cm-border);transition:transform .15s ease,box-shadow .15s ease;}
      .stButton > button:hover {transform:translateY(-1px);box-shadow:0 4px 14px #10243A24;}
      .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {background:#F3A33E!important;color:#172B3C!important;border:1px solid #ED9528!important;}
      .stDownloadButton > button {border-radius:10px;min-height:41px;font-weight:740;}
      [data-testid="stVerticalBlockBorderWrapper"] > div {border-radius:15px!important;}
      [data-testid="stMain"] [data-baseweb="input"] > div,
      [data-testid="stMain"] [data-baseweb="select"] > div,
      [data-testid="stMain"] [data-baseweb="textarea"] > div {background:@@INPUT@@;color:var(--cm-text);border-color:var(--cm-border);}
      [data-testid="stMain"] [data-baseweb="input"] input,
      [data-testid="stMain"] [data-baseweb="textarea"] textarea {color:var(--cm-text);}
      [data-testid="stMain"] [data-testid="stWidgetLabel"] p {color:var(--cm-text);}
      [data-testid="stDataFrame"], [data-testid="stFileUploader"] {border-radius:11px;overflow:hidden;}
      [data-testid="stExpander"] {border-radius:12px;background:var(--cm-surface);border-color:var(--cm-border);}
      .cm-footer {margin-top:2rem;padding:1.1rem 0 0;border-top:1px solid var(--cm-border);font-size:.76rem;color:var(--cm-muted);}
      @media (max-width:1100px) {.cm-kpi-grid {grid-template-columns:repeat(2,minmax(0,1fr));}}
      @media (max-width:760px) {
        .block-container {padding:1rem .7rem 2rem;}
        .cm-welcome {padding:1.18rem 1.1rem;min-height:155px;}
        .cm-welcome .sky {font-size:2rem;right:12px;top:13px;opacity:.8;}
        .cm-welcome h1 {padding-right:30px;font-size:1.65rem;}
        .cm-welcome p {font-size:.84rem;padding-right:35px;}
        .cm-pagehead {padding-top:.2rem;}
        .cm-kpi-grid {gap:.55rem;}
        .cm-kpi {padding:.77rem;min-height:110px;}
        .cm-kpi-value {font-size:1.5rem;}
        .cm-kpi-head {font-size:.72rem;}
        .cm-calendar {grid-template-columns:repeat(3,minmax(0,1fr));gap:.38rem;}
        .cm-day {min-height:76px;}
        .cm-context {font-size:.73rem;}
        [data-testid="stHorizontalBlock"] {gap:.65rem;}
      }
      @media (max-width:370px) {.cm-kpi-grid,.cm-calendar {grid-template-columns:repeat(2,minmax(0,1fr));}}
    </style>
    """
    for key, value in variables.items():
        styles = styles.replace("@@" + key + "@@", value)
    st.markdown(styles, unsafe_allow_html=True)


def hero(kicker: str, title: str, subtitle: str):
    """En-tête compact sur les pages secondaires ; pas de bandeau répété."""
    st.markdown(
        f'<div class="cm-pagehead"><div class="cm-eyebrow">{html.escape(kicker)}</div>'
        f'<h1>{html.escape(title)}</h1><p>{html.escape(subtitle)}</p></div>',
        unsafe_allow_html=True,
    )


def dashboard_hero(p: dict | None):
    locality = html.escape(p.get("city", "CAMEROUN") if p else "CAMEROUN")
    project = html.escape(p.get("name", "AUCUN CHANTIER SÉLECTIONNÉ") if p else "DÉMARRER VOTRE PROJET")
    st.markdown(
        '<div class="cm-welcome">'
        f'<div class="eyebrow">CENTRE DE PILOTAGE · {locality.upper()}</div>'
        '<h1>Anticipez les aléas météo de vos chantiers.</h1>'
        '<p>Analysez les conditions, comparez les activités et prenez des décisions fondées sur des données.</p>'
        '<div class="sky" aria-hidden="true">🌦️</div></div>'
        f'<div class="cm-context"><span class="cm-chip">🏗️ {project}</span>'
        '<span class="cm-chip">🕒 Africa/Douala</span><span class="cm-chip">📐 Génie civil · ENSPY</span></div>',
        unsafe_allow_html=True,
    )


def section_header(icon: str, title: str, subtitle: str = ""):
    st.markdown(
        f'<div class="cm-section"><div class="cm-section-icon">{html.escape(icon)}</div>'
        f'<div><div class="cm-section-title">{html.escape(title)}</div>'
        f'<div class="cm-section-sub">{html.escape(subtitle)}</div></div></div>',
        unsafe_allow_html=True,
    )


def metric(label: str, number, accent="#318FBD", sub=""):
    accent = accent if accent in PALETTE.values() else "#318FBD"
    st.markdown(
        f'<div class="metric" style="--accent:{accent}"><div class="label">{html.escape(str(label))}</div>'
        f'<div class="num">{html.escape(str(number))}</div>'
        f'<div class="sub">{html.escape(str(sub))}</div></div>',
        unsafe_allow_html=True,
    )


def dashboard_kpis(items: list[tuple[str, str, str, str, str]]):
    cards = []
    for icon, title, value, sub, color in items:
        color = color if color in PALETTE.values() else "#318FBD"
        cards.append(
            f'<div class="cm-kpi" style="--kpi-accent:{color}"><div class="cm-kpi-head">'
            f'<span class="cm-kpi-icon">{html.escape(icon)}</span>{html.escape(title)}</div>'
            f'<div class="cm-kpi-value">{html.escape(str(value))}</div>'
            f'<div class="cm-kpi-note">{html.escape(str(sub))}</div></div>'
        )
    st.markdown('<div class="cm-kpi-grid">' + ''.join(cards) + '</div>', unsafe_allow_html=True)


def activity_overview(p: dict | None, df: pd.DataFrame):
    """Une synthèse calculée à la demande, sans résultat mis en cache."""
    if not p or not p.get('activities'):
        return '<div class="cm-empty"><b>Aucune activité configurée</b><p>Créez une activité dans « Planification météo ».</p></div>'
    rows = []
    for a in p['activities'][:5]:
        cal = calendar_for_activity(df, a)
        ok = int((cal['Statut'] == 'Compatible').sum())
        bad = int((cal['Statut'] == 'Défavorable').sum())
        unknown = int((cal['Statut'] == 'Indéterminé').sum())
        status, css_class = (f'{bad} défavorable(s)', 'cm-status-bad') if bad else ((f'{unknown} à préciser', 'cm-status-unknown') if unknown else ('Compatible', 'cm-status-good'))
        rows.append(
            f'<div class="cm-act-row"><div><div class="cm-act-name">{html.escape(a["name"])}</div>'
            f'<div class="cm-act-meta">{html.escape(a["start"])} · {int(a["duration"])} j · {ok} compatible(s)</div></div>'
            f'<span class="cm-status {css_class}">{html.escape(status)}</span></div>'
        )
    if len(p['activities']) > 5:
        rows.append(f'<div class="cm-act-meta">+ {len(p["activities"])-5} activité(s) supplémentaire(s)</div>')
    return ''.join(rows)


def calendar_strip(calendar: pd.DataFrame, max_days: int = 35):
    """Calendrier lisible et adaptatif ; les motifs sont aussi dans le tableau détaillé."""
    if calendar.empty:
        return
    color = {'Compatible': '#22977A', 'Défavorable': '#D96957', 'Indéterminé': '#E3A03A'}
    parts = []
    for _, row in calendar.head(max_days).iterrows():
        status = str(row['Statut'])
        pc = row.get('precip_mm')
        rain = f'{float(pc):g} mm' if pd.notna(pc) else 'Pluie : —'
        short = {'Compatible': 'Compatible', 'Défavorable': 'Défavorable', 'Indéterminé': 'À préciser'}.get(status, status)
        friendly = pd.Timestamp(row['date']).strftime('%d/%m')
        parts.append(
            f'<div class="cm-day" style="--day-color:{color.get(status, "#E3A03A")}" title="{html.escape(str(row.get("Motif", "")), quote=True)}">'
            f'<div class="date">{friendly}</div><div class="rain">{html.escape(rain)}</div>'
            f'<div class="state">{html.escape(short)}</div></div>'
        )
    legend = '<div class="cm-legend"><span><i style="background:#22977A"></i>Compatible</span><span><i style="background:#D96957"></i>Défavorable</span><span><i style="background:#E3A03A"></i>Indéterminé</span></div>'
    st.markdown(legend + '<div class="cm-calendar">' + ''.join(parts) + '</div>', unsafe_allow_html=True)
    if len(calendar) > max_days:
        st.caption(f'Calendrier visuel limité aux {max_days} premiers jours sur {len(calendar)}. Le tableau présente toutes les dates.')


def notice(text: str):
    st.markdown(f'<div class="cm-notice">{html.escape(text)}</div>', unsafe_allow_html=True)


def demo_warning(p: dict | None):
    if p and p.get("demo"):
        st.markdown('<div class="cm-demo">⚠️ Données fictives — démonstration pédagogique uniquement. Aucune mesure réelle de Yaoundé.</div>', unsafe_allow_html=True)


def need_project() -> dict | None:
    p = active_project()
    if p is None:
        st.markdown('<div class="cm-empty"><b>Commencez par un chantier</b><p>Ouvrez « Mes chantiers » ou utilisez l’exemple fictif pour une démonstration sans Internet.</p></div>', unsafe_allow_html=True)
        if st.button("Créer mon premier chantier", type="primary"):
            jump(NAV[1])
    return p


def pick_activity(p: dict, key="activity_pick") -> dict | None:
    acts = p.get("activities", [])
    if not acts:
        st.info("Créez d'abord une activité dans « Planification météo ».")
        return None
    ids = [a["id"] for a in acts]
    chosen = st.selectbox("Activité étudiée", ids, format_func=lambda x: next(a["name"] for a in acts if a["id"] == x), key=key)
    return next(a for a in acts if a["id"] == chosen)


def download_controls(p: dict, df: pd.DataFrame, activity: dict | None = None):
    if df.empty:
        return
    section_header("↧", "Exporter les analyses", "Téléchargez un tableau ou un rapport documentant les hypothèses.")
    cols = st.columns(3)
    cols[0].download_button("⬇️ CSV", data=make_csv(df), file_name="chantiermeteo_donnees.csv", mime="text/csv", use_container_width=True)
    cols[1].download_button("⬇️ Excel", data=make_excel(df, p), file_name="chantiermeteo_analyse.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    try:
        pdf = make_pdf(p, df, activity)
        cols[2].download_button("⬇️ Rapport PDF", data=pdf, file_name="chantiermeteo_rapport.pdf", mime="application/pdf", use_container_width=True)
    except Exception as exc:
        cols[2].error(f"PDF indisponible : {exc}")


def chart(fig, height=350):
    dark = st.session_state.get('theme') == 'Sombre'
    fg = '#EDF5FD' if dark else '#132C41'
    grid = '#30485A' if dark else '#E6EDF2'
    bg = '#14283C' if dark else '#FFFFFF'
    fig.update_layout(
        template='plotly_dark' if dark else 'plotly_white',
        height=height, margin=dict(l=12, r=12, t=47, b=34),
        font=dict(family='Segoe UI, Arial', color=fg, size=12),
        paper_bgcolor=bg, plot_bgcolor=bg,
        title=dict(font=dict(size=15, color=fg)),
        legend=dict(orientation='h', y=-.28, x=0, font=dict(size=10)),
        hoverlabel=dict(bgcolor='#10243A', font_color='white'),
    )
    fig.update_xaxes(showgrid=False, linecolor=grid, zeroline=False, automargin=True, tickfont=dict(size=10))
    fig.update_yaxes(showgrid=True, gridcolor=grid, zeroline=False, automargin=True, tickfont=dict(size=10))
    st.plotly_chart(fig, use_container_width=True, config={'displaylogo': False, 'responsive': True, 'modeBarButtonsToRemove': ['lasso2d', 'select2d']})

# ----------------------------- PAGES -------------------------------------
def dashboard():
    p = active_project()
    df = df_project(p)
    dashboard_hero(p)
    counts = pd.Series(dtype=int)
    selected = p['activities'][0] if p and p.get('activities') else None
    if selected is not None and not df.empty:
        counts = classify_weather(df, selected)['Statut'].value_counts()
    adverse = int(counts.get('Défavorable', 0))
    decidable = adverse + int(counts.get('Compatible', 0))
    dashboard_kpis([
        ('▤', 'Chantiers', len(st.session_state.projects), 'Enregistrés dans cette session', PALETTE['blue']),
        ('◷', 'Jours étudiés', len(df), f'{df.date.min()} au {df.date.max()}' if not df.empty else 'Chargez des données', PALETTE['orange']),
        ('◉', 'Défavorables', adverse if decidable else '—', 'Selon la première activité' if decidable else 'Aucune classification', PALETTE['red']),
        ('✓', 'Jours compatibles', f'{100 - int(round(100 * adverse / decidable))} %' if decidable else '—', f'{decidable} jours classables' if decidable else 'Créer une activité', PALETTE['green']),
    ])
    section_header('⚡', 'Accès rapide', 'Toutes les commandes ci-dessous sont actives.')
    one, two, three, four = st.columns(4)
    if one.button('＋ Nouveau chantier', type='primary', use_container_width=True):
        st.session_state.editing = False
        jump(NAV[1])
    if two.button('▤ Mes chantiers', use_container_width=True):
        jump(NAV[1])
    if three.button("🧪 Charger l'exemple fictif", use_container_width=True):
        add_demo()
        st.success('Exemple fictif chargé ; vos autres projets sont conservés.')
        st.rerun()
    if four.button('☁ Analyse météo', use_container_width=True):
        jump(NAV[2])
    if DEMO_ID in st.session_state.projects:
        if st.button("🧹 Réinitialiser l'exemple fictif", help='Supprime uniquement le chantier de démonstration.'):
            remove_demo()
            st.rerun()
    if not p:
        st.markdown('<div class="cm-empty"><b>Votre espace est prêt</b><p>Créez un chantier ou chargez l’exemple fictif pour visualiser le tableau de bord, les analyses et le calendrier.</p></div>', unsafe_allow_html=True)
        section_header('▧', 'Parcours conseillé', 'Trois étapes pour commencer.')
        c1, c2, c3 = st.columns(3)
        c1.info('**1 · Mes chantiers**\n\nLocalisez le projet et ajoutez ses activités.')
        c2.info('**2 · Données météo**\n\nTéléchargez des séries ou importez vos mesures.')
        c3.info('**3 · Planifier**\n\nAnalysez les journées et comparez les scénarios.')
        return
    demo_warning(p)
    st.markdown('<div class="cm-context">'
                f'<span class="cm-chip">📍 {html.escape(p["city"])}</span>'
                f'<span class="cm-chip">🗃️ {html.escape(p.get("source", "aucune"))}</span>'
                f'<span class="cm-chip">🕑 Données : {html.escape(str(p.get("retrieved_at") or "non actualisées"))}</span>'
                '</div>', unsafe_allow_html=True)
    section_header('▥', 'Vue opérationnelle', 'Pluies disponibles et compatibilité de vos activités.')
    lcol, rcol = st.columns([1.45, 1.0], gap='medium')
    with lcol:
        with st.container(border=True):
            st.markdown('<div class="cm-panel-head"><b>Précipitations journalières</b><span>Mesures chargées, en mm/jour</span></div>', unsafe_allow_html=True)
            if df.empty or df.precip_mm.dropna().empty:
                st.info('Aucune série de pluie disponible. Ouvrez « Analyse météorologique ».')
            else:
                limit = selected.get('rain') if selected else None
                display = df.copy()
                display['couleur'] = ['Défavorable' if limit is not None and pd.notna(r) and r >= limit else 'Pluie' for r in display.precip_mm]
                graph = px.bar(display, x='date', y='precip_mm', color='couleur',
                               color_discrete_map={'Pluie': PALETTE['blue'], 'Défavorable': PALETTE['orange']},
                               labels={'date': 'Date', 'precip_mm': 'Pluie (mm/j)', 'couleur': 'Critère'},
                               title=None)
                if limit is not None:
                    graph.add_hline(y=limit, line_dash='dash', line_color=PALETTE['orange'], annotation_text=f'Seuil pluie {limit:g} mm', annotation_position='top left')
                chart(graph, 305)
    with rcol:
        with st.container(border=True):
            st.markdown('<div class="cm-panel-head"><b>Activités en cours d’analyse</b><span>Classement des journées selon vos critères</span></div>', unsafe_allow_html=True)
            st.markdown(activity_overview(p, df), unsafe_allow_html=True)
            if st.button('📅 Voir la planification', use_container_width=True):
                jump(NAV[3])
    if selected is not None:
        section_header('▦', 'Calendrier de compatibilité', f'{selected["name"]} · seuils configurés par l’utilisateur')
        cal = calendar_for_activity(df, selected)
        calendar_strip(cal, max_days=28)
        with st.expander('Afficher les motifs et données détaillées du calendrier'):
            st.dataframe(cal[['date', 'precip_mm', 'Statut', 'Motif']], hide_index=True, use_container_width=True)
    section_header('↧', 'Rapport et actualisation', 'Les sorties conservent la provenance et les hypothèses de calcul.')
    a, b, c = st.columns([1, 1, 1])
    if a.button('🔄 Actualiser les données', use_container_width=True):
        if p.get('source') not in ('historique Open-Meteo', 'prévision Open-Meteo'):
            st.warning('Actualisation disponible uniquement pour les données Open-Meteo. Réimportez vos fichiers manuellement.')
        else:
            try:
                with st.spinner('Récupération des données Open-Meteo…'):
                    mode = p.get('api_mode', 'historique')
                    fresh = fetch_weather(p['latitude'], p['longitude'], mode, p.get('api_start', ''), p.get('api_end', ''))
                    set_weather(p, fresh, p['source'], 'Actualisation Open-Meteo ; données issues de réanalyses ou prévisions selon le type.')
                st.success('Données mises à jour.')
                st.rerun()
            except Exception as exc:
                st.error(f'Actualisation impossible : {exc}')
    if b.button('🎲 Analyser les probabilités', use_container_width=True):
        jump(NAV[4])
    if not df.empty:
        try:
            c.download_button('⬇️ Exporter le rapport PDF', data=make_pdf(p, df, selected),
                              file_name='chantiermeteo_rapport.pdf', mime='application/pdf', use_container_width=True)
        except Exception as exc:
            c.error(f'Export PDF impossible : {exc}')
    else:
        c.button('⬇️ Exporter le rapport PDF', disabled=True, use_container_width=True)
        st.caption('Ajoutez des données pour générer un rapport.')
    st.caption(p.get('data_note', ''))


def projects_page():
    hero("DOSSIERS", "Mes chantiers", "Créer, retrouver et modifier les chantiers. Les informations restent en mémoire jusqu'à export JSON.")
    plist = list(st.session_state.projects.values())
    section_header("▤", "Projets enregistrés", "Choisissez un dossier ou modifiez ses informations.")
    if not plist:
        st.info("Aucun projet. Utilisez le formulaire « Créer un chantier » ci-dessous.")
    for i in range(0, len(plist), 2):
        cols = st.columns(2)
        for col, p in zip(cols, plist[i:i + 2]):
            with col:
                with st.container(border=True):
                    selected_badge = ' · ACTIF' if p['id'] == st.session_state.selected_id else ''
                    st.markdown(
                        '<div class="cm-project-header"><div class="cm-project-icon">' + ('🧪' if p.get('demo') else '🏗️') + '</div>'
                        '<div><div class="cm-project-name">' + html.escape(p['name']) + '</div>'
                        '<div class="cm-project-desc">' + html.escape(p['type'] + ' · ' + p['city'] + selected_badge) + '</div></div></div>'
                        '<div class="cm-project-desc">📆 ' + html.escape(p['start'] + ' au ' + p['end']) + '</div>'
                        '<div class="cm-project-counts"><span>▥ ' + str(len(p.get('records', []))) + ' journées</span>'
                        '<span>◷ ' + str(len(p.get('activities', []))) + ' activités</span>'
                        + ('<span>FICTIF</span>' if p.get('demo') else '') + '</div>',
                        unsafe_allow_html=True,
                    )
                    c1, c2 = st.columns(2)
                    if c1.button("Ouvrir", key="open" + p["id"], use_container_width=True):
                        st.session_state.selected_id = p["id"]
                        st.session_state.sync_project_picker = True
                        jump(NAV[0])
                    if c2.button("Modifier", key="edit" + p["id"], use_container_width=True):
                        st.session_state.selected_id = p["id"]
                        st.session_state.sync_project_picker = True
                        st.session_state.editing = True
                        st.rerun()
    st.divider()
    editing = bool(st.session_state.editing and active_project())
    p = active_project() if editing else None
    title = "✏️ Modifier le chantier" if editing else "➕ Créer un chantier"
    section_header("✎", title, "Les champs marqués * sont obligatoires.")
    with st.form("project_form", clear_on_submit=False):
        name = st.text_input("Nom du chantier *", value=p["name"] if p else "", max_chars=150)
        a, b = st.columns(2)
        kind = a.selectbox("Type de projet", TYPE_OPTIONS, index=TYPE_OPTIONS.index(p["type"]) if p and p["type"] in TYPE_OPTIONS else 0)
        city = b.text_input("Ville / localisation *", value=p["city"] if p else "Yaoundé")
        c, d = st.columns(2)
        lat = c.number_input("Latitude (degrés décimaux)", min_value=-90.0, max_value=90.0, value=float(p["latitude"]) if p else 3.8667, format="%.5f")
        lon = d.number_input("Longitude (degrés décimaux)", min_value=-180.0, max_value=180.0, value=float(p["longitude"]) if p else 11.5167, format="%.5f")
        e, f = st.columns(2)
        ds = e.date_input("Début prévu *", value=date.fromisoformat(p["start"]) if p else today())
        de = f.date_input("Fin prévue *", value=date.fromisoformat(p["end"]) if p else today() + timedelta(days=30))
        desc = st.text_area("Description", value=p.get("description", "") if p else "", height=90)
        submit = st.form_submit_button("Enregistrer les modifications" if editing else "Créer le chantier", type="primary")
    if submit:
        if not name.strip() or not city.strip():
            st.error("Renseignez le nom et la ville du chantier.")
        elif ds > de:
            st.error("La date de fin ne peut précéder la date de début.")
        else:
            new = dict(p) if editing else {"id": sid(), "activities": [], "records": [], "source": "aucune", "retrieved_at": None, "data_note": "", "demo": False}
            new.update({"name": name.strip(), "type": kind, "city": city.strip(), "latitude": float(lat), "longitude": float(lon), "start": ds.isoformat(), "end": de.isoformat(), "description": desc})
            # When editing a demo, do not silently relabel fictional data as real.
            save_project(new)
            st.session_state.editing = False
            st.success("Chantier enregistré.")
            st.rerun()
    if editing:
        col1, col2 = st.columns(2)
        if col1.button("Annuler la modification"):
            st.session_state.editing = False
            st.rerun()
        with col2:
            confirm = st.checkbox("Confirmer la suppression de ce chantier")
            if st.button("🗑️ Supprimer le chantier", disabled=not confirm):
                dead = p["id"]
                st.session_state.projects.pop(dead, None)
                st.session_state.selected_id = next(iter(st.session_state.projects), None)
                st.session_state.sync_project_picker = True
                st.session_state.editing = False
                st.rerun()
    st.divider()
    section_header("⇄", "Sauvegarder ou restaurer", "Les données restent uniquement dans la session jusqu’à export JSON.")
    st.download_button("⬇️ Sauvegarder tous les chantiers (JSON)", data=backup_json(), file_name="chantiermeteo_sauvegarde.json", mime="application/json")
    uploaded = st.file_uploader("Fichier de sauvegarde JSON (max. 8 Mo)", type=["json"], key="restore_projects")
    if uploaded:
        replace = st.checkbox("Je comprends que la restauration remplacera tous les chantiers de cette session.")
        if st.button("Restaurer la sauvegarde", disabled=not replace):
            try:
                projects, sel = parse_backup(uploaded.getvalue())
                st.session_state.projects = projects
                st.session_state.selected_id = sel
                st.session_state.sync_project_picker = True
                st.session_state.demo_loaded = DEMO_ID in projects
                st.success(f"{len(projects)} chantier(s) restauré(s).")
                st.rerun()
            except Exception as exc:
                st.error(f"Restauration refusée : {exc}")


def weather_page():
    hero("DONNÉES • OBSERVER", "Analyse météorologique", "Interroger Open-Meteo, importer des séries journalières, explorer les données brutes et leurs statistiques.")
    p = need_project()
    if not p:
        return
    demo_warning(p)
    st.caption(f"Chantier sélectionné : {p['name']} · {p['city']} ({p['latitude']:.4f}, {p['longitude']:.4f})")
    section_header("☁", "Acquérir les données", "Historique de réanalyse ou prévision Open-Meteo.")
    with st.expander("📍 Rechercher des coordonnées par nom de ville"):
        q = st.text_input("Ville à rechercher", value=p["city"], key="geosearch")
        if st.button("Rechercher la ville"):
            try:
                st.session_state.geocoding_results = geocode(q.strip()) if q.strip() else []
            except Exception as exc:
                st.error(f"Géocodage impossible : {exc}")
        results = st.session_state.get("geocoding_results", [])
        if results:
            ids = list(range(len(results)))
            ind = st.selectbox("Résultats possibles", ids, format_func=lambda i: f"{results[i]['name']}, {results[i].get('admin1', '')} — {results[i].get('country', '')} ({results[i]['latitude']:.3f}, {results[i]['longitude']:.3f})")
            if st.button("Utiliser ces coordonnées"):
                choice = results[ind]
                p.update({"city": choice["name"], "latitude": float(choice["latitude"]), "longitude": float(choice["longitude"])})
                save_project(p)
                st.success("Localisation mise à jour. Rechargez les données météorologiques.")
                st.rerun()
        else:
            st.caption("Sélectionnez « Rechercher la ville ». Les homonymes doivent être vérifiés.")
    with st.container(border=True):
        mode = st.radio("Type de données à récupérer", ["Historique (réanalyse)", "Prévisions (jusqu'à 16 jours)"], horizontal=True)
        is_hist = mode.startswith("Historique")
        if is_hist:
            col1, col2 = st.columns(2)
            start = col1.date_input("Date de début de l'historique", value=today() - timedelta(days=365), key="histstart")
            end = col2.date_input("Date de fin de l'historique", value=today() - timedelta(days=8), key="histend")
        if st.button("🌐 Télécharger les données Open-Meteo", type="primary"):
            try:
                if is_hist and (start > end or end > today()):
                    raise ValueError("Choisissez un intervalle passé cohérent ; la disponibilité des jours récents dépend du service.")
                with st.spinner("Téléchargement depuis Open-Meteo..."):
                    m = "historique" if is_hist else "prevision"
                    df = fetch_weather(p["latitude"], p["longitude"], m, start.isoformat() if is_hist else "", end.isoformat() if is_hist else "")
                src = "historique Open-Meteo" if is_hist else "prévision Open-Meteo"
                note = "Réanalyses / agrégations modélisées, non mesures garanties sur le chantier." if is_hist else "Prévisions évolutives, non observées ; probabilité de pluie distincte des fréquences historiques."
                set_weather(p, df, src, note)
                p.update({"api_mode": m, "api_start": start.isoformat() if is_hist else "", "api_end": end.isoformat() if is_hist else ""})
                save_project(p)
                st.success(f"{len(df)} journées téléchargées.")
                st.rerun()
            except Exception as exc:
                st.error(f"Échec du téléchargement : {exc}. Vérifiez la période, les coordonnées et Internet.")
    section_header("↧", "Importer un jeu de données", "CSV / Excel avec contrôle des unités, dates et valeurs.")
    with st.expander("📥 CSV / Excel — format, validation et import"):
        notice("Format attendu : date (AAAA-MM-JJ), precip_mm, tmin_c, tmax_c, wind_kmh, humidity_pct, rain_probability_pct. Seule la date et une variable météo sont obligatoires. Valeurs manquantes permises.")
        blank = pd.DataFrame([{"date":"2026-01-01", "precip_mm": 2.5, "tmin_c": 20., "tmax_c": 28., "wind_kmh": 12.}])
        st.download_button("Télécharger un modèle CSV", make_csv(blank), "chantiermeteo_modele.csv", mime="text/csv")
        f = st.file_uploader("Choisir un fichier CSV ou XLSX (max. 8 Mo)", type=["csv", "xlsx"], key="weather_upload")
        if f and st.button("Importer et valider le fichier"):
            try:
                if f.size > 8_000_000:
                    raise ValueError("Fichier supérieur à 8 Mo.")
                raw = pd.read_excel(io.BytesIO(f.getvalue())) if f.name.lower().endswith("xlsx") else pd.read_csv(io.BytesIO(f.getvalue()), sep=None, engine="python", encoding="utf-8-sig")
                data = validate_weather(raw, strict=True)
                set_weather(p, data, "import CSV/Excel", f"Données fournies par l'utilisateur : {f.name}. Vérifier leur origine et leur qualité.")
                for key in ("api_mode", "api_start", "api_end"):
                    p.pop(key, None)
                save_project(p)
                st.success(f"Fichier validé : {len(data)} jours. Ancien jeu de données de ce chantier remplacé.")
                st.rerun()
            except Exception as exc:
                st.error(f"Import refusé : {exc}")
    df = df_project(p)
    st.divider()
    section_header("▥", "Explorer les mesures", "Statistiques calculées sur la période filtrée.")
    if df.empty:
        st.info("Aucune donnée météo disponible : téléchargez les données ou importez un fichier pour afficher les analyses.")
        return
    st.caption(f"**Source :** {p.get('source')} | {p.get('data_note')} | Dernière récupération : {p.get('retrieved_at') or '—'}")
    d1, d2 = st.columns(2)
    dr = d1.date_input("Filtrer à partir du", value=date.fromisoformat(df.date.min()), min_value=date.fromisoformat(df.date.min()), max_value=date.fromisoformat(df.date.max()))
    dfinal = d2.date_input("Jusqu'au", value=date.fromisoformat(df.date.max()), min_value=date.fromisoformat(df.date.min()), max_value=date.fromisoformat(df.date.max()))
    filtered = df[(df.date >= dr.isoformat()) & (df.date <= dfinal.isoformat())]
    if filtered.empty:
        st.warning("Aucune observation dans les dates choisies.")
        return
    rain = filtered.precip_mm.dropna()
    k, n, f = rain_exceedance(filtered, 10)
    ms = st.columns(4)
    with ms[0]: metric("Journées présentes", len(filtered), PALETTE["blue"])
    with ms[1]: metric("Pluie moyenne", f"{rain.mean():.1f} mm/j" if len(rain) else "—", PALETTE["orange"])
    with ms[2]: metric("Jours avec pluie > 0", f"{int((rain > 0).sum())}/{len(rain)}" if len(rain) else "—", PALETTE["green"])
    with ms[3]: metric("Valeurs manquantes", int(filtered[NUM_COLS].isna().sum().sum()), PALETTE["red"], "Sur l'ensemble des variables")
    with st.expander("🧾 Afficher les données brutes et leurs valeurs manquantes", expanded=True):
        st.dataframe(filtered.rename(columns=FIELDS), use_container_width=True, hide_index=True)
        st.dataframe(pd.DataFrame({"Variable": list(FIELDS.values()), "Valeurs manquantes": [int(filtered[col].isna().sum()) for col in NUM_COLS]}), hide_index=True, use_container_width=True)
    with st.expander("✏️ Corriger les données manuellement (facultatif)"):
        st.warning("Une correction transforme la provenance en « édition utilisateur ». L'ancienne série n'est pas modifiée à distance : exportez l'original avant tout changement.")
        if st.checkbox("Activer l'édition des données", value=False, key="enable_weather_editor"):
            edit_version = hashlib.sha256(json.dumps(p.get("records", []), sort_keys=True).encode("utf-8")).hexdigest()[:12]
            edited = st.data_editor(df, hide_index=True, use_container_width=True, num_rows="dynamic", key="weather_editor_" + edit_version)
            if st.button("Enregistrer les corrections des données"):
                try:
                    corrected = validate_weather(edited, strict=True)
                    if corrected.equals(df):
                        st.info("Aucune différence à enregistrer.")
                    else:
                        previous = p.get("source", "inconnue")
                        set_weather(p, corrected, "édition utilisateur", f"Série modifiée manuellement ; provenance antérieure : {previous}.")
                        for key in ("api_mode", "api_start", "api_end"):
                            p.pop(key, None)
                        save_project(p)
                        st.success("Données corrigées et résultats recalculés.")
                        st.rerun()
                except Exception as exc:
                    st.error(f"Corrections rejetées : {exc}")
    section_header("∑", "Statistiques descriptives", "Effectifs, dispersion et quantiles des variables disponibles.")
    stat = summary(filtered)
    if not stat.empty:
        st.dataframe(stat, use_container_width=True, hide_index=True)
        st.caption("Variance et écart-type : formules d'échantillon (diviseur n − 1) ; indisponibles pour n = 1.")
    figcols = st.columns(2)
    with figcols[0]:
        if filtered.precip_mm.notna().any():
            chart(px.bar(filtered, x="date", y="precip_mm", title="Pluies journalières", color_discrete_sequence=[PALETTE["blue"]], labels={"date": "Date", "precip_mm": "Pluie (mm/j)"}))
            chart(px.histogram(filtered.dropna(subset=["precip_mm"]), x="precip_mm", nbins=18, title="Distribution des précipitations", color_discrete_sequence=[PALETTE["orange"]], labels={"precip_mm": "Précipitations (mm/j)"}))
        if filtered.wind_kmh.notna().any():
            chart(px.histogram(filtered.dropna(subset=["wind_kmh"]), x="wind_kmh", nbins=18, title="Distribution du vent maximal", color_discrete_sequence=[PALETTE["blue"]], labels={"wind_kmh":"Vent max. (km/h)"}))
    with figcols[1]:
        if filtered[["tmin_c", "tmax_c"]].notna().any().any():
            melt = filtered.melt(id_vars="date", value_vars=["tmin_c", "tmax_c"], var_name="Série", value_name="°C")
            chart(px.line(melt, x="date", y="°C", color="Série", title="Températures journalières", color_discrete_map={"tmin_c":PALETTE["blue"], "tmax_c":PALETTE["red"]}))
        if filtered.precip_mm.notna().any():
            tmp = filtered.assign(mois=filtered.date.str.slice(0, 7)).groupby("mois", as_index=False)["precip_mm"].sum(min_count=1)
            chart(px.bar(tmp, x="mois", y="precip_mm", title="Cumul mensuel (mois présents)", color_discrete_sequence=[PALETTE["green"]], labels={"precip_mm":"Cumul (mm)", "mois":"Mois"}))
            thresholds = np.linspace(0, max(1, float(filtered.precip_mm.max()) * 1.1), 25)
            sensitivity = pd.DataFrame({"seuil": thresholds, "fréquence": [(rain >= t).mean() * 100 for t in thresholds]})
            chart(px.line(sensitivity, x="seuil", y="fréquence", title="Fréquence empirique de dépassement", labels={"seuil":"Seuil de pluie (mm/j)", "fréquence":"Journées ≥ seuil (%)"}, color_discrete_sequence=[PALETTE["red"]]))
    download_controls(p, filtered, p["activities"][0] if p.get("activities") else None)


def planning_page():
    hero("ANTICIPER • COMPARER", "Planification météo", "Définissez vos propres limites d'exécution et visualisez les journées compatibles, défavorables ou indéterminées.")
    p = need_project()
    if not p:
        return
    demo_warning(p)
    section_header("▤", "Paramétrer une activité", "Critères personnalisés pour le chantier sélectionné.")
    acts = p.get("activities", [])
    mode = st.radio("Gestion des activités", ["Ajouter une activité", "Modifier une activité"], horizontal=True, key="act_mode")
    editing = mode == "Modifier une activité" and len(acts) > 0
    a = None
    if editing:
        aid = st.selectbox("Activité à modifier", [x["id"] for x in acts], format_func=lambda x: next(z["name"] for z in acts if z["id"] == x), key="act_edit_id")
        a = next(x for x in acts if x["id"] == aid)
    elif mode == "Modifier une activité":
        st.warning("Aucune activité à modifier.")
    with st.form("activity_form"):
        name = st.selectbox("Activité", ACT_OPTIONS, index=ACT_OPTIONS.index(a["name"]) if a and a["name"] in ACT_OPTIONS else 0)
        other = st.text_input("Nom si activité personnalisée", value=a["name"] if a and a["name"] not in ACT_OPTIONS else "")
        x, y = st.columns(2)
        start = x.date_input("Date de début souhaitée", value=date.fromisoformat(a["start"]) if a else date.fromisoformat(p["start"]))
        duration = y.number_input("Durée (jours civils)", min_value=1, max_value=730, value=int(a["duration"]) if a else 5)
        st.caption("Les critères sont des hypothèses configurables, pas des limites légales ou de sécurité. Seules les cases cochées activent leur seuil.")
        limits = {}
        defaults = [("rain", "Précipitations ≥ (mm/j)", 10.0), ("tmin", "Température min. < (°C)", 15.0), ("tmax", "Température max. > (°C)", 36.0), ("wind", "Vent maximal > (km/h)", 35.0)]
        lc = st.columns(2)
        for i, (key, label, default) in enumerate(defaults):
            with lc[i % 2]:
                enabled = st.checkbox("Activer — " + label, value=(a.get(key) is not None) if a else (key == "rain"), key=f"enabled_{key}_{a['id'] if a else 'new'}")
                v = st.number_input(label, min_value=0.0 if key in ("rain", "wind") else -80.0, max_value=300.0 if key in ("rain", "wind") else 70.0, value=float(a[key]) if a and a.get(key) is not None else default, step=1.0, key=f"limit_{key}_{a['id'] if a else 'new'}")
                limits[key] = float(v) if enabled else None
        submit = st.form_submit_button("Enregistrer l'activité", type="primary")
    if submit:
        actual_name = other.strip() if name == "Activité personnalisée" else name
        if not actual_name:
            st.error("Renseignez un nom d'activité.")
        elif all(x is None for x in limits.values()):
            st.error("Activez au moins un critère météorologique.")
        elif limits["tmin"] is not None and limits["tmax"] is not None and limits["tmin"] > limits["tmax"]:
            st.error("La température minimale acceptée dépasse la maximale.")
        else:
            record = {"id": a["id"] if a else sid(), "name": actual_name, "start": start.isoformat(), "duration": int(duration), **limits}
            if editing:
                p["activities"] = [record if old["id"] == a["id"] else old for old in acts]
            else:
                p["activities"].append(record)
            save_project(p)
            st.success("Activité enregistrée ; les classifications sont recalculées à partir des données courantes.")
            st.rerun()
    if editing:
        confirm = st.checkbox("Confirmer la suppression de cette activité")
        if st.button("🗑️ Supprimer l'activité", disabled=not confirm):
            p["activities"] = [z for z in acts if z["id"] != a["id"]]
            save_project(p)
            st.rerun()
    st.divider()
    section_header("▦", "Calendrier de compatibilité", "Chaque date est classée selon les mesures et seuils disponibles.")
    if not acts:
        st.info("Enregistrez une première activité pour afficher son calendrier météo.")
        return
    selected = pick_activity(p, key="plan_act")
    if selected is None:
        return
    df = df_project(p)
    cal = calendar_for_activity(df, selected)
    count = cal["Statut"].value_counts()
    cols = st.columns(3)
    with cols[0]: metric("Compatible", int(count.get("Compatible", 0)), PALETTE["green"])
    with cols[1]: metric("Défavorable", int(count.get("Défavorable", 0)), PALETTE["red"])
    with cols[2]: metric("Indéterminé", int(count.get("Indéterminé", 0)), PALETTE["orange"])
    if df.empty:
        st.warning("Aucune donnée chargée ; toutes les journées sont indéterminées.")
    elif p.get("source") == "historique Open-Meteo":
        st.info("Vous utilisez des données historiques : ce calendrier représente une analyse rétrospective, et non une prévision future.")
    if p.get("source") == "prévision Open-Meteo":
        st.info("Prévisions limitées à la période disponible : les autres journées sont indéterminées.")
    calendar_strip(cal, max_days=35)
    with st.expander("ℹ️ Comment lire les couleurs et les motifs ?"):
        st.write("Vert : toutes les limites définies sont respectées. Rouge : au moins une violation connue. Ambre : valeurs absentes ou date hors des données.")
        st.caption("Survolez une journée sur ordinateur pour connaître son motif ; le tableau détaillé ci-dessous reste la référence complète.")
    plot_df = cal.copy()
    plot_df["Activité"] = selected["name"]
    fig = px.timeline(plot_df, x_start=pd.to_datetime(plot_df["date"]), x_end=pd.to_datetime(plot_df["date"]) + pd.Timedelta(days=1), y="Activité", color="Statut", hover_data=["date", "Motif"], title="Calendrier journalier coloré", color_discrete_map={"Compatible": PALETTE["green"], "Défavorable": PALETTE["red"], "Indéterminé": PALETTE["orange"]})
    fig.update_yaxes(title=None)
    chart(fig, 255)
    st.dataframe(cal[["date", "precip_mm", "tmin_c", "tmax_c", "wind_kmh", "Statut", "Motif"]].rename(columns={"date":"Jour", **FIELDS}), hide_index=True, use_container_width=True)
    st.caption("Jour civil 00:00–24:00 (Africa/Douala). Pluie défavorable si ≥ seuil ; températures/vent défavorables seulement si hors des limites respectives. Une donnée absente rend le résultat indéterminé sauf violation déjà démontrée.")
    section_header("⇄", "Comparer les dates de démarrage", "Même durée et critères ; dates candidates consécutives.")
    if not df.empty:
        one, two = st.columns(2)
        day = one.date_input("Première date candidate", value=date.fromisoformat(df.date.min()), key="candidate_start")
        count_cand = two.slider("Dates candidates consécutives", 2, 14, 7)
        compare = []
        for i in range(count_cand):
            d = day + timedelta(days=i)
            group = calendar_for_activity(df, selected, d)
            counts = group["Statut"].value_counts()
            compare.append({"Début candidat": d.isoformat(), "Compatible": int(counts.get("Compatible", 0)), "Défavorable": int(counts.get("Défavorable", 0)), "Indéterminé": int(counts.get("Indéterminé", 0))})
        comp = pd.DataFrame(compare)
        st.dataframe(comp, hide_index=True, use_container_width=True)
        tidy = comp.melt(id_vars="Début candidat", var_name="Statut", value_name="Journées")
        chart(px.bar(tidy, x="Début candidat", y="Journées", color="Statut", barmode="stack", title="Comparaison des scénarios sur la même durée", color_discrete_map={"Compatible":PALETTE["green"], "Défavorable":PALETTE["red"], "Indéterminé":PALETTE["orange"]}), 310)
        st.caption("Aucun classement ne prouve qu'un travail est sûr : autres risques techniques, humains et réglementaires à contrôler séparément.")
        download_controls(p, df, selected)


def probability_page():
    hero("STATISTIQUES • INCERTITUDE", "Probabilités et scénarios", "Quantifier les événements météorologiques historiques et les risques simulés sans les confondre avec des prévisions.")
    p = need_project()
    if not p:
        return
    demo_warning(p)
    df = df_project(p)
    if df.empty:
        st.info("Chargez d'abord des données dans « Analyse météorologique », ou choisissez l'exemple fictif.")
        return
    if p.get("source") == "prévision Open-Meteo":
        st.warning("Les données actuellement chargées sont des prévisions. Les fréquences ci-dessous décrivent ce jeu de prévisions, PAS des probabilités historiques. Pour une analyse fréquentiste, récupérez une série historique.")
    section_header("◉", "Fréquence empirique", "Événement : pluie quotidienne supérieure ou égale au seuil.")
    threshold = st.number_input("Seuil de pluie (mm/j), événement X ≥ seuil", min_value=0., max_value=500., value=10., step=1., key="prob_rain_threshold")
    k, n, prob = rain_exceedance(df, threshold)
    a, b, c = st.columns(3)
    with a: metric("Journées valides", n, PALETTE["blue"])
    with b: metric("Dépassements (≥ seuil)", k, PALETTE["orange"])
    with c: metric("Fréquence empirique", f"{100 * prob:.1f} %" if prob is not None else "—", PALETTE["green"])
    if prob is not None:
        width = max(0.0, min(100.0, 100.0 * prob))
        st.markdown('<div style="height:13px;background:var(--cm-border);border-radius:12px;overflow:hidden;margin:10px 0 13px;">'
                    f'<div style="height:100%;width:{width:.2f}%;background:linear-gradient(90deg,#F4A340,#D96957);border-radius:12px;"></div></div>',
                    unsafe_allow_html=True)
    st.latex(r"\widehat{P}(X\ge x)=\frac{\#\{i:X_i\ge x\}}{n_{\mathrm{valide}}}")
    st.caption("Le seuil est inclus (≥). Les jours avec pluie manquante sont exclus du dénominateur. Ce résultat est une fréquence observée ou modélisée de la série, pas la probabilité de crue ou d'accident.")
    st.divider()
    section_header("▦", "Fenêtres consécutives", "Au moins un jour défavorable dans une période de durée fixée.")
    a_sel = pick_activity(p, key="prob_act")
    if a_sel is None:
        return
    dur = st.number_input("Durée d'une fenêtre (jours civils)", min_value=1, max_value=365, value=3)
    events, valid, freq = available_windows(df, a_sel, int(dur))
    if freq is None:
        st.warning("Aucune fenêtre complète : ajoutez des données quotidiennes consécutives sans valeurs manquantes sur les critères activés.")
    else:
        st.success(f"{events} fenêtres avec au moins une journée défavorable parmi {valid} fenêtres complètes ({100 * freq:.1f} %).")
    st.caption("Une fenêtre glissante de k jours correspond à k dates civiles consécutives. Les fenêtres contenant des jours indéterminés sont exclues. Les fenêtres se chevauchent : elles ne sont pas indépendantes.")
    st.divider()
    section_header("🎲", "Simulations Monte-Carlo", "Rééchantillonnage de blocs historiques pour étudier les incertitudes.")
    st.caption("Rééchantillonnage de blocs complets de jours déjà classés. L'ordre des jours dans chaque bloc est conservé, les blocs sont tirés avec remise. Il s'agit d'un modèle simplifié, non d'une prévision datée.")
    w, x = st.columns(2)
    horizon = w.slider("Durée du scénario simulé (jours)", 1, 60, 7)
    trials = x.select_slider("Nombre de simulations", options=[100, 500, 1000, 2000, 5000, 10000], value=2000)
    y, z = st.columns(2)
    bl = y.number_input("Longueur des blocs (jours)", min_value=1, max_value=horizon, value=min(2, horizon))
    seed = z.number_input("Graine aléatoire", min_value=0, max_value=2_147_483_647, value=42)
    blockcount = len(valid_blocks(df, a_sel, int(bl)))
    st.caption(f"Blocs historiques utilisables selon les critères : {blockcount}.")
    dataset_signature = hashlib.sha256(json.dumps(p.get("records", []), sort_keys=True).encode("utf-8")).hexdigest()
    forecast_only = p.get("source") == "prévision Open-Meteo"
    if forecast_only:
        st.warning("Simulation de fréquence historique désactivée pour les prévisions : chargez un historique plutôt qu'un horizon prévisionnel.")
    if st.button("🎲 Lancer la simulation Monte-Carlo", type="primary", disabled=forecast_only):
        try:
            counts = monte_carlo(df, a_sel, int(horizon), int(trials), int(bl), int(seed))
            hits = (counts >= 1).astype(float)
            pp = hits.mean()
            se = np.sqrt(pp * (1 - pp) / len(hits))
            st.session_state["mc_result"] = {"input": [p["id"], json.dumps(a_sel, sort_keys=True), dataset_signature, horizon, trials, bl, seed], "counts": counts.tolist(), "p": float(pp), "se": float(se)}
        except Exception as exc:
            st.error(f"Simulation impossible : {exc}")
            st.session_state.pop("mc_result", None)
    result = st.session_state.get("mc_result")
    signature = [p["id"], json.dumps(a_sel, sort_keys=True), dataset_signature, horizon, trials, bl, seed]
    if result and result["input"] == signature:
        counts = np.array(result["counts"], dtype=int)
        p_est, se = result["p"], result["se"]
        b1, b2, b3 = st.columns(3)
        with b1: metric("≥ 1 jour défavorable", f"{p_est * 100:.1f} %", PALETTE["red"])
        with b2: metric("Nombre moyen de jours", f"{counts.mean():.2f}", PALETTE["blue"])
        with b3: metric("Erreur-type Monte-Carlo", f"{se * 100:.2f} points", PALETTE["orange"])
        vals, freqs = np.unique(counts, return_counts=True)
        chart(px.bar(x=vals, y=freqs, labels={"x":"Nombre de jours défavorables", "y":"Nombre de simulations"}, title="Distribution simulée des journées défavorables", color_discrete_sequence=[PALETTE["blue"]]))
        st.write("Quantiles empiriques (méthode **higher**, nombres entiers de jours) :", {"Q50": int(np.quantile(counts, .5, method="higher")), "Q90": int(np.quantile(counts, .9, method="higher")), "Q95": int(np.quantile(counts, .95, method="higher"))})
        st.caption(f"Erreur-type ≈ √[p(1−p)/N]. Environ 95 % de précision Monte-Carlo conditionnelle au modèle : ±{1.96 * se * 100:.2f} points (approximation, hors incertitude sur le climat). Les simulations où rien n'arrive sont conservées comme zéro.")
    elif result:
        st.info("Paramètres ou données modifiés : le résultat précédent est masqué. Cliquez à nouveau sur « Lancer la simulation Monte-Carlo ».")
    st.divider()
    section_header("↗", "Analyse de sensibilité", "Étudier l’impact du choix du seuil de précipitation.")
    rain = df.precip_mm.dropna()
    if rain.empty:
        st.info("Aucune valeur de pluie valide.")
    else:
        lo, hi = st.slider("Plage de seuils (mm/j)", min_value=0, max_value=max(10, int(np.ceil(rain.max())) + 20), value=(0, max(10, int(np.ceil(rain.max())))))
        thresholds = np.linspace(lo, hi, 31)
        sens = pd.DataFrame({"Seuil (mm/j)": thresholds, "Fréquence (%)": [100 * (rain >= t).mean() for t in thresholds]})
        chart(px.line(sens, x="Seuil (mm/j)", y="Fréquence (%)", title="Variation de la fréquence quand le seuil change", color_discrete_sequence=[PALETTE["orange"]]))
    with st.expander("📚 Hypothèses et limites du modèle", expanded=False):
        st.markdown("""**Événement** : au moins un jour dont une limite activée est violée ; une donnée manquante empêche de valider le jour sauf violation déjà connue. **Période** : jours civils 00:00–24:00, Africa/Douala. **Bloc** : suite de dates consécutives sans information manquante sur les critères activés. **Indépendance** : les tirages de blocs sont indépendants *entre blocs simulés* ; les jours à l'intérieur d'un bloc ne le sont pas. **Quantiles** : méthode `higher` sur les comptes entiers de jours, conservant les zéros. **Limites** : hypothèse que les motifs observés restent représentatifs ; pas de saisonnalité explicite, ni changement climatique, ni prise en compte des contraintes de chantier hors météo. Les erreurs-type décrivent seulement l'aléa des tirages, pas l'incertitude structurelle. Le Monte-Carlo ne prédit pas la date d'un événement.""")


def teaching_page():
    hero("SOUTENANCE • MÉTHODOLOGIE", "Comprendre et tester", "Un laboratoire autonome pour vérifier les calculs avec une calculatrice et défendre les résultats devant votre professeur.")
    section_header("✓", "Test déterministe", "Cas fictif vérifiable sans API ni simulation.")
    rain = np.array([0, 4, 12, 0, 8, 25, 1, 0, 16, 3], dtype=float)
    demo = pd.DataFrame({"date": pd.date_range("2026-10-20", periods=10).strftime("%Y-%m-%d"), "precip_mm": rain})
    demo = validate_weather(demo)
    threshold = st.number_input("Seuil pédagogique de pluie X ≥ x (mm/j)", min_value=0., value=10., step=1., key="ped_threshold")
    if st.button("🧮 Vérifier les calculs", type="primary"):
        k, n, f = rain_exceedance(demo, threshold)
        st.success(f"Effectif = {n} journées ; pluie moyenne = {rain.mean():.1f} mm/j ; dépassements ≥ {threshold:g} mm/j = {k} ; fréquence = {100 * f:.1f} %.")
        st.latex(r"\bar{x}=\frac{0+4+12+0+8+25+1+0+16+3}{10}=6{,}9\ \mathrm{mm/j}")
        st.info("Pour le seuil 10 mm/j : 12, 25 et 16 sont ≥ 10, donc 3 jours sur 10 = 30 %. Pour un autre seuil, le nombre change.")
    st.dataframe(demo[["date", "precip_mm"]], hide_index=True, use_container_width=True)
    with st.expander("1. Tableau de bord — ouvrir l'exemple fictif", expanded=True):
        st.markdown("""**Chemin :** `🏠 Tableau de bord` → bouton **« 🧪 Charger l'exemple fictif »**.\n\n**Données exactes :** aucune saisie.\n\n**Résultat attendu :** chantier-école fictif à Yaoundé ; 10 journées ; deux activités. Les autres projets ne sont pas modifiés. **Pourquoi :** l'application fonctionne même sans Internet.""")
    with st.expander("2. Mes chantiers — créer ou modifier"):
        st.markdown("""**Chemin :** `🏗️ Mes chantiers` → formulaire « Créer un chantier ».\n\n**Valeurs :** Nom = `Bâtiment école`, type = `Bâtiment`, ville = `Yaoundé`, latitude = `3.8667`, longitude = `11.5167`, début = `2026-11-01`, fin = `2026-11-10`.\n\n**Bouton :** « Créer le chantier ». **Attendu :** nouvelle carte cliquable. Utilisez « Modifier » puis « Enregistrer les modifications » pour le mettre à jour.""")
    with st.expander("3. Analyse météorologique — données et moyennes"):
        st.markdown("""**Chemin :** `🏠 Tableau de bord` → « 🧪 Charger l'exemple fictif » → `🌦️ Analyse météorologique`.\n\n**Valeurs exactes :** les dix pluies affichées ci-dessus, déjà chargées.\n\n**Bouton :** aucun calcul manuel : les tableaux se mettent à jour.\n\n**Attendu :** 10 journées et 6,9 mm/j de pluie moyenne. Le seuil de 10 mm/j est traité dans la page Probabilités. Source **fictif** clairement visible.""")
    with st.expander("4. Planification météo — classements par jour"):
        st.markdown("""**Chemin :** charger l'exemple → `📅 Planification météo` → « Activité étudiée » = `Terrassement`.\n\n**Valeurs exactes :** début `2026-10-20`, durée `5`, limite de précipitation activée `10 mm/j`, autres limites désactivées.\n\n**Bouton :** si nécessaire, mode « Modifier une activité » → « Enregistrer l'activité ».\n\n**Attendu :** parmi les 5 premiers jours : 4 **Compatibles**, 1 **Défavorable** (2026-10-22, 12 mm ≥ 10). Les dates absentes seraient **Indéterminées**.""")
    with st.expander("5. Probabilités et scénarios — cas connu"):
        st.markdown("""**Chemin :** charger l'exemple → `🎲 Probabilités et scénarios`.\n\n**Champs :** Seuil de pluie = `10 mm/j` ; activité = `Terrassement` ; durée de fenêtre = `3 jours`.\n\n**Attendu :** **3/10 = 30 %** de jours dépassant le seuil et **8/8 = 100 %** de fenêtres complètes de 3 jours contenant au moins un jour défavorable. Les 8 fenêtres glissantes sont les jours 1–3 jusqu'à 8–10.\n\n**Monte-Carlo :** horizon `7 jours`, simulations `2000`, blocs `2 jours`, graine `42` → bouton « 🎲 Lancer la simulation Monte-Carlo ». L'indicateur doit être reproductible lors de deux exécutions avec les mêmes paramètres. Ses résultats dépendent de l'algorithme de rééchantillonnage décrit dans l'application.""")
    with st.expander("6. Exports et sauvegardes"):
        st.markdown("""**Chemin :** charger l'exemple → `🌦️ Analyse météorologique` → rubrique « Exports du chantier » → boutons « ⬇️ CSV », « ⬇️ Excel », « ⬇️ Rapport PDF ».\n\n**Attendu :** fichiers téléchargeables, étiquetés avec la source fictive.\n\n**Sauvegarde :** `🏗️ Mes chantiers` → « ⬇️ Sauvegarder tous les chantiers (JSON) ». Pour restaurer : sélectionner le fichier, cocher la confirmation puis « Restaurer la sauvegarde ». Attention, la restauration remplace volontairement la session.""")
    st.divider()
    section_header("∑", "Formules fondamentales", "Notations de statistiques descriptives et probabilités.")
    st.latex(r"\bar{x}=\frac{1}{n}\sum_{i=1}^{n}x_i\qquad s^2=\frac{1}{n-1}\sum_{i=1}^{n}(x_i-\bar{x})^2")
    st.latex(r"\widehat{P}(X\geq x_0)=\frac{k}{n}\qquad SE_{MC}=\sqrt{\frac{\widehat{p}(1-\widehat{p})}{N}}")
    st.markdown("La **moyenne** décrit un niveau quotidien typique ; la **variance** mesure la dispersion dans l'unité au carré ; l'**écart-type** conserve l'unité d'origine ; un **quantile** est un seuil qui partage les observations ordonnées. La fréquence empirique ne démontre ni causalité ni sécurité sur un chantier.")
    st.info("**À la soutenance :** expliquez la source, les unités, le nombre de journées valides, les critères personnalisés, l'événement aléatoire, les hypothèses des simulations et leurs limites. Les réanalyses Open-Meteo ne sont pas des pluviomètres installés sur le chantier.")


def sidebar():
    css(st.session_state.theme)
    st.sidebar.markdown(
        '<div class="cm-sidebar-brand"><div class="cm-mark">CM</div>'
        '<div><div class="cm-brand-title">Chantier<b>Météo</b></div>'
        '<div class="cm-brand-sub">CONTROL CENTER</div></div></div>',
        unsafe_allow_html=True,
    )
    st.sidebar.markdown('<div class="cm-sidebar-section">Espace de travail</div>', unsafe_allow_html=True)
    if st.session_state.get('go_to') in NAV:
        st.session_state.nav = st.session_state.pop('go_to')
    page = st.sidebar.radio('Navigation', NAV, key='nav', label_visibility='collapsed')
    st.sidebar.markdown('<div class="cm-sidebar-section">Personnalisation</div>', unsafe_allow_html=True)
    st.sidebar.selectbox('Ambiance', ['Clair', 'Sombre'], key='theme')
    st.sidebar.divider()
    ids = list(st.session_state.projects)
    if ids:
        if st.session_state.pop('sync_project_picker', False) or st.session_state.get('sidebar_project_picker') not in ids:
            st.session_state.sidebar_project_picker = st.session_state.selected_id if st.session_state.selected_id in ids else ids[0]
        def selected_project_changed():
            st.session_state.selected_id = st.session_state.sidebar_project_picker
        st.sidebar.selectbox('Chantier actif', ids, key='sidebar_project_picker', on_change=selected_project_changed,
                             format_func=lambda i: st.session_state.projects[i]['name'])
    else:
        st.sidebar.caption('Aucun chantier sélectionné')
    if st.sidebar.button("🧪 Charger l'exemple fictif", use_container_width=True):
        add_demo()
        st.rerun()
    st.sidebar.download_button('⬇️ Sauvegarde JSON', data=backup_json(),
                               file_name='chantiermeteo_sauvegarde.json', mime='application/json', use_container_width=True)
    st.sidebar.markdown('<div class="cm-sidebar-foot">🏗️ ENSPY · Génie civil<br>Sources : Open-Meteo<br>Fuseau : Africa/Douala<br>Données conservées en session.</div>', unsafe_allow_html=True)
    return page


def main():
    initial_state()
    page = sidebar()
    try:
        if page == NAV[0]:
            dashboard()
        elif page == NAV[1]:
            projects_page()
        elif page == NAV[2]:
            weather_page()
        elif page == NAV[3]:
            planning_page()
        elif page == NAV[4]:
            probability_page()
        else:
            teaching_page()
    except (ValueError, KeyError) as exc:
        st.error(f"Données incohérentes : {exc}. Vérifiez le chantier et réessayez.")
    st.markdown('<div class="cm-footer">ChantierMétéo · Outil pédagogique ENSPY &nbsp;•&nbsp; Les critères personnalisés ne constituent pas des normes de sécurité ou de construction.</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
