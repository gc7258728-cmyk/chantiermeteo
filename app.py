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


# ---------------------------- PRESENTATION -------------------------------
def css(theme: str):
    dark = theme == "Sombre"
    bg, card, txt, secondary, sidebar = ("#0C1928", "#15263A", "#E7F2FF", "#ABBCD1", "#102136") if dark else ("#F3F6FA", "#FFFFFF", "#10243A", "#536579", "#FFFFFF")
    st.markdown(f"""<style>
    :root {{ --bg:{bg}; --card:{card}; --text:{txt}; --muted:{secondary}; }}
    .stApp {{background:{bg}; color:{txt};}}
    section[data-testid="stSidebar"] {{background:{sidebar}; border-right:1px solid #9aa9b433;}}
    div[data-testid="stVerticalBlockBorderWrapper"] > div {{border-radius:16px;}}
    h1,h2,h3 {{color:{txt}; letter-spacing:-.025em;}}
    .hero {{background:linear-gradient(112deg,#10243A 0%,#204F72 55%,#328CCB 100%);padding:30px 32px;border-radius:20px;color:white;margin-bottom:16px;position:relative;overflow:hidden;box-shadow:0 10px 30px #10243a18;}}
    .hero:after {{content:'◈';position:absolute;right:28px;top:-52px;font-size:200px;color:#ffffff0c;pointer-events:none;}}
    .hero h1 {{color:#fff;font-size:clamp(1.6rem,3vw,2.5rem);margin:0 0 8px;}}
    .hero p {{color:#e3f1ff;font-size:1.05rem;max-width:700px;margin:0;}}
    .pill {{display:inline-block;background:#F29A38;color:#172334;padding:5px 11px;border-radius:30px;font-weight:800;font-size:.73rem;letter-spacing:.06em;margin-bottom:11px;}}
    .metric {{background:{card};border:1px solid #7F9AB428;border-top:4px solid var(--accent);border-radius:16px;padding:19px 17px;min-height:130px;box-shadow:0 4px 20px #0c24400a;}}
    .metric .label {{color:{secondary};font-weight:650;font-size:.87rem;line-height:1.25;}}
    .metric .num {{color:{txt};font-size:clamp(1.25rem,2vw,1.9rem);font-weight:800;margin-top:13px;overflow-wrap:anywhere;}}
    .metric .sub {{color:{secondary};font-size:.76rem;margin-top:4px;}}
    .sectionhint {{color:{secondary};margin-top:-9px;margin-bottom:17px;}}
    .notice {{padding:13px 16px;background:{card};border-left:4px solid #328CCB;border-radius:9px;margin:12px 0;color:{txt};}}
    .demo {{padding:12px 16px;background:#fff0d5;color:#603900;border-left:4px solid #F29A38;border-radius:10px;margin:12px 0;font-weight:700;}}
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{background:#EE8E24;color:#172334;border:0;font-weight:750;}}
    .stButton > button {{border-radius:10px;font-weight:650;min-height:42px;}}
    [data-testid="stDataFrame"], [data-testid="stFileUploader"] {{border-radius:10px;}}
    @media (max-width:760px) {{ .hero {{padding:21px 16px;}} .hero:after {{display:none;}} .metric {{min-height:105px;padding:13px;}} .block-container {{padding-left:1rem;padding-right:1rem;}} .hero p {{font-size:.94rem;}} }}
    </style>""", unsafe_allow_html=True)


def hero(kicker: str, title: str, subtitle: str):
    st.markdown(f'<div class="hero"><span class="pill">{html.escape(kicker)}</span><h1>{html.escape(title)}</h1><p>{html.escape(subtitle)}</p></div>', unsafe_allow_html=True)


def metric(label: str, number, accent="#328CCB", sub=""):
    st.markdown(f'<div class="metric" style="--accent:{accent}"><div class="label">{html.escape(str(label))}</div><div class="num">{html.escape(str(number))}</div><div class="sub">{html.escape(str(sub))}</div></div>', unsafe_allow_html=True)


def notice(text: str):
    st.markdown(f'<div class="notice">{html.escape(text)}</div>', unsafe_allow_html=True)


def demo_warning(p: dict | None):
    if p and p.get("demo"):
        st.markdown('<div class="demo">⚠️ Données fictives — utilisation pédagogique uniquement. Aucune mesure réelle à Yaoundé.</div>', unsafe_allow_html=True)


def need_project() -> dict | None:
    p = active_project()
    if p is None:
        st.info("Aucun chantier sélectionné. Ouvrez « Mes chantiers » ou chargez l'exemple fictif.")
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
    st.markdown("#### Exports du chantier")
    cols = st.columns(3)
    cols[0].download_button("⬇️ CSV", data=make_csv(df), file_name="chantiermeteo_donnees.csv", mime="text/csv", use_container_width=True)
    cols[1].download_button("⬇️ Excel", data=make_excel(df, p), file_name="chantiermeteo_analyse.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    try:
        pdf = make_pdf(p, df, activity)
        cols[2].download_button("⬇️ Rapport PDF", data=pdf, file_name="chantiermeteo_rapport.pdf", mime="application/pdf", use_container_width=True)
    except Exception as exc:
        cols[2].error(f"PDF indisponible : {exc}")


def chart(fig, height=370):
    fig.update_layout(template="plotly_white", height=height, margin=dict(l=14, r=14, t=45, b=20), font=dict(family="Arial", color="#10243A", size=12), paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", legend=dict(orientation="h", y=-.30, x=0))
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False, "responsive": True})

# ----------------------------- PAGES -------------------------------------
def dashboard():
    hero("GESTION DES CHANTIERS • CAMEROUN", "ChantierMétéo", "Des données météorologiques aux décisions de chantier : analyser, anticiper et comparer les conditions de travail.")
    p = active_project()
    df = df_project(p)
    counts = pd.Series(dtype=int)
    if p and p.get("activities") and not df.empty:
        cls = classify_weather(df, p["activities"][0])
        counts = cls["Statut"].value_counts()
    k = int(counts.get("Défavorable", 0))
    determined = k + int(counts.get("Compatible", 0))
    share = (100 * k / determined) if determined else None
    n = len(df)
    c = st.columns(4)
    with c[0]: metric("Chantiers en session", len(st.session_state.projects), PALETTE["blue"], "Sans compte, ni connexion")
    with c[1]: metric("Chantier actif", (p["name"][:23] + "…") if p and len(p["name"]) > 24 else (p["name"] if p else "Aucun"), PALETTE["orange"], p["city"] if p else "Ajouter un chantier")
    with c[2]: metric("Journées renseignées", n, PALETTE["green"], (f"{df.date.min()} → {df.date.max()}" if n else "Aucune donnée chargée"))
    with c[3]: metric("Jours défavorables", f"{k} ({share:.0f} %)" if share is not None else "—", PALETTE["red"], f"Sur {determined} jours classables ; 1re activité" if determined else "Configurez une activité")
    st.markdown("### Démarrer une analyse")
    a, b, c = st.columns(3)
    if a.button("➕ Nouveau chantier", type="primary", use_container_width=True):
        st.session_state.editing = False
        jump(NAV[1])
    if b.button("📂 Ouvrir un chantier", use_container_width=True):
        jump(NAV[1])
    if c.button("🧪 Charger l'exemple fictif", use_container_width=True):
        add_demo()
        st.success("Exemple fictif chargé. Les autres chantiers sont conservés.")
        st.rerun()
    if DEMO_ID in st.session_state.projects:
        if st.button("🧹 Réinitialiser l'exemple fictif", help="Retire uniquement le chantier fictif."):
            remove_demo()
            st.rerun()
    st.markdown("### Projet sélectionné")
    if p:
        demo_warning(p)
        st.markdown(f"**{p['name']}** · {p['type']} · **{p['city']}**")
        st.caption(f"Période du chantier : {p['start']} au {p['end']} | Données : {p.get('source', 'aucune')} | Dernière acquisition : {p.get('retrieved_at') or '—'}")
        st.write(p.get("data_note", ""))
        if not df.empty:
            if df["precip_mm"].notna().any():
                fig = px.bar(df, x="date", y="precip_mm", labels={"date":"Date", "precip_mm":"Pluie (mm/j)"}, title="Précipitations des données chargées", color_discrete_sequence=[PALETTE["blue"]])
                chart(fig, 300)
            download_controls(p, df, p["activities"][0] if p.get("activities") else None)
        else:
            st.info("Ce chantier ne contient pas encore de données météo. Ouvrez « Analyse météorologique » pour les récupérer ou les importer.")
        col1, col2 = st.columns(2)
        if col1.button("🔄 Actualiser les données", use_container_width=True):
            if p.get("source") not in ("historique Open-Meteo", "prévision Open-Meteo"):
                st.warning("Actualisation réseau disponible uniquement pour des données téléchargées avec Open-Meteo. Pour un fichier importé, réimportez-le.")
            else:
                try:
                    with st.spinner("Récupération des nouvelles données..."):
                        mode = p.get("api_mode", "historique")
                        fresh = fetch_weather(p["latitude"], p["longitude"], mode, p.get("api_start", ""), p.get("api_end", ""))
                        set_weather(p, fresh, p["source"], "Mise à jour Open-Meteo ; données météo issues de modèles/réanalyses ou de prévisions selon le type.")
                    st.success("Données actualisées.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Actualisation impossible : {exc}")
        if col2.button("📅 Préparer les travaux", use_container_width=True):
            jump(NAV[3])
    else:
        notice("Créez un chantier pour étudier ses conditions météorologiques, ou lancez l'exemple fictif pour explorer l'application hors connexion.")
    st.markdown("### Parcours conseillé")
    one, two, three = st.columns(3)
    one.info("**1. Mes chantiers**\n\nEnregistrez votre projet et sa localisation.")
    two.info("**2. Analyse météorologique**\n\nTéléchargez ou importez les données.")
    three.info("**3. Planification / Probabilités**\n\nComparez les conditions et les incertitudes.")


def projects_page():
    hero("DOSSIERS", "Mes chantiers", "Créer, retrouver et modifier les chantiers. Les informations restent en mémoire jusqu'à export JSON.")
    plist = list(st.session_state.projects.values())
    st.markdown("### Projets enregistrés")
    if not plist:
        st.info("Aucun projet. Utilisez le formulaire « Créer un chantier » ci-dessous.")
    for i in range(0, len(plist), 2):
        cols = st.columns(2)
        for col, p in zip(cols, plist[i:i + 2]):
            with col:
                with st.container(border=True):
                    st.markdown(f"#### {'🧪 ' if p.get('demo') else '🏗️ '}{p['name']}")
                    st.caption(f"{p['type']} • {p['city']} • {p['start']} → {p['end']}")
                    st.write(f"{len(p.get('records', []))} journées · {len(p.get('activities', []))} activités")
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
    st.markdown("### " + title)
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
    st.markdown("### Sauvegarder ou restaurer mes chantiers")
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
    st.markdown("### 1. Acquisition des données")
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
    st.markdown("### 2. Importer mon propre jeu de données")
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
    st.markdown("### 3. Données et indicateurs")
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
    st.markdown("#### Statistiques descriptives")
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
    st.markdown("### 1. Activités du chantier")
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
    st.markdown("### 2. Calendrier de compatibilité")
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
    plot_df = cal.copy()
    plot_df["Activité"] = selected["name"]
    fig = px.timeline(plot_df, x_start=pd.to_datetime(plot_df["date"]), x_end=pd.to_datetime(plot_df["date"]) + pd.Timedelta(days=1), y="Activité", color="Statut", hover_data=["date", "Motif"], title="Calendrier journalier coloré", color_discrete_map={"Compatible": PALETTE["green"], "Défavorable": PALETTE["red"], "Indéterminé": PALETTE["orange"]})
    fig.update_yaxes(title=None)
    chart(fig, 255)
    st.dataframe(cal[["date", "precip_mm", "tmin_c", "tmax_c", "wind_kmh", "Statut", "Motif"]].rename(columns={"date":"Jour", **FIELDS}), hide_index=True, use_container_width=True)
    st.caption("Jour civil 00:00–24:00 (Africa/Douala). Pluie défavorable si ≥ seuil ; températures/vent défavorables seulement si hors des limites respectives. Une donnée absente rend le résultat indéterminé sauf violation déjà démontrée.")
    st.markdown("### 3. Comparer des dates de démarrage")
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
    st.markdown("### 1. Fréquence empirique de dépassement")
    threshold = st.number_input("Seuil de pluie (mm/j), événement X ≥ seuil", min_value=0., max_value=500., value=10., step=1., key="prob_rain_threshold")
    k, n, prob = rain_exceedance(df, threshold)
    a, b, c = st.columns(3)
    with a: metric("Journées valides", n, PALETTE["blue"])
    with b: metric("Dépassements (≥ seuil)", k, PALETTE["orange"])
    with c: metric("Fréquence empirique", f"{100 * prob:.1f} %" if prob is not None else "—", PALETTE["green"])
    st.latex(r"\widehat{P}(X\ge x)=\frac{\#\{i:X_i\ge x\}}{n_{\mathrm{valide}}}")
    st.caption("Le seuil est inclus (≥). Les jours avec pluie manquante sont exclus du dénominateur. Ce résultat est une fréquence observée ou modélisée de la série, pas la probabilité de crue ou d'accident.")
    st.divider()
    st.markdown("### 2. Périodes consécutives comprenant un jour défavorable")
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
    st.markdown("### 3. Monte-Carlo par blocs de journées historiques")
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
    st.markdown("### 4. Sensibilité au seuil de pluie")
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
    st.markdown("### Test déterministe entièrement reproductible")
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
    st.markdown("### Formules fondamentales")
    st.latex(r"\bar{x}=\frac{1}{n}\sum_{i=1}^{n}x_i\qquad s^2=\frac{1}{n-1}\sum_{i=1}^{n}(x_i-\bar{x})^2")
    st.latex(r"\widehat{P}(X\geq x_0)=\frac{k}{n}\qquad SE_{MC}=\sqrt{\frac{\widehat{p}(1-\widehat{p})}{N}}")
    st.markdown("La **moyenne** décrit un niveau quotidien typique ; la **variance** mesure la dispersion dans l'unité au carré ; l'**écart-type** conserve l'unité d'origine ; un **quantile** est un seuil qui partage les observations ordonnées. La fréquence empirique ne démontre ni causalité ni sécurité sur un chantier.")
    st.info("**À la soutenance :** expliquez la source, les unités, le nombre de journées valides, les critères personnalisés, l'événement aléatoire, les hypothèses des simulations et leurs limites. Les réanalyses Open-Meteo ne sont pas des pluviomètres installés sur le chantier.")


def sidebar():
    st.sidebar.markdown("## 🏗️ ChantierMétéo")
    st.sidebar.caption("Analyse et planification • ENSPY")
    st.sidebar.selectbox("Ambiance", ["Clair", "Sombre"], key="theme")
    css(st.session_state.theme)
    if st.session_state.get("go_to") in NAV:
        st.session_state.nav = st.session_state.pop("go_to")
    page = st.sidebar.radio("Navigation", NAV, key="nav")
    st.sidebar.divider()
    ids = list(st.session_state.projects)
    if ids:
        if st.session_state.pop("sync_project_picker", False) or st.session_state.get("sidebar_project_picker") not in ids:
            st.session_state.sidebar_project_picker = st.session_state.selected_id if st.session_state.selected_id in ids else ids[0]
        def selected_project_changed():
            st.session_state.selected_id = st.session_state.sidebar_project_picker
        st.sidebar.selectbox("Chantier actif", ids, key="sidebar_project_picker", on_change=selected_project_changed, format_func=lambda i: st.session_state.projects[i]["name"])
    else:
        st.sidebar.caption("Aucun chantier enregistré")
    if st.sidebar.button("🧪 Exemple fictif", use_container_width=True):
        add_demo()
        st.rerun()
    st.sidebar.download_button("⬇️ Sauvegarde JSON", data=backup_json(), file_name="chantiermeteo_sauvegarde.json", mime="application/json", use_container_width=True)
    st.sidebar.divider()
    st.sidebar.caption("Sources météo : Open-Meteo · Horaires : Africa/Douala · Aucun compte requis. Données en session uniquement.")
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
    st.markdown("---")
    st.caption("ChantierMétéo • Outil pédagogique d'aide à la décision — aucune garantie de sécurité ni de conformité aux normes de construction.")


if __name__ == "__main__":
    main()
