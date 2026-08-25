
from pathlib import Path
from datetime import date
import pandas as pd
import streamlit as st
from streamlit_image_coordinates import streamlit_image_coordinates
from PIL import Image, ImageDraw, ImageFont
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Arc, Circle

from modules.database import init_db, create_match, get_matches, insert_event, get_events, delete_event, update_event, delete_match
from modules.zones import derive_zone

st.set_page_config(page_title="AKA Goals Dashboard", page_icon="⚽", layout="wide")

st.markdown("""
<style>
:root {
    --ried:#12f2ad;
    --bg:#1d1d1d;
    --panel:#030303;
    --text:#f5f5f5;
    --muted:#b8b8b8;
}
html, body, [class*="css"] { font-family: Arial, Helvetica, sans-serif; }
.stApp {
    background:#1d1d1d;
    color:var(--text);
}
[data-testid="stHeader"] {
    background:transparent;
}
[data-testid="stSidebar"] {
    background:#151515;
    border-right:1px solid #2a2a2a;
}
[data-testid="stSidebar"] * {
    color:#f4f4f4;
}
[data-testid="stSidebar"] [role="radiogroup"] label {
    padding:.28rem 0;
}
.block-container {
    padding-top:1.0rem;
    padding-bottom:2.0rem;
    max-width:100%;
}
div[data-testid="stMetric"] {
    background:#111;
    border:1px solid #2a2a2a;
    border-radius:10px;
    padding:10px 14px;
}
div[data-testid="stMetricLabel"] p { color:#cfcfcf; }
div[data-testid="stMetricValue"] { color:white; }
div.stButton > button {
    border-radius:8px;
}
div.stButton > button[kind="primary"] {
    background:#14dca2;
    color:#08100d;
    border:none;
    font-weight:800;
}
.aka-hero {
    position:relative;
    min-height:86px;
    margin-bottom:4px;
}
.aka-leftmark {
    position:absolute;
    left:6px;
    top:0px;
    color:#149a75;
    font-size:1.3rem;
    line-height:.9;
    font-weight:900;
    font-style:italic;
    transform:rotate(-5deg);
    letter-spacing:-1px;
}
.aka-center {
    display:flex;
    align-items:center;
    justify-content:center;
    gap:18px;
    padding-top:6px;
}
.aka-balls {
    display:flex;
    gap:14px;
    align-items:center;
}
.aka-ball1,.aka-ball2 {
    width:44px;
    height:44px;
    border-radius:50%;
    box-shadow:inset 0 0 12px rgba(255,255,255,.18),0 0 8px rgba(0,0,0,.45);
}
.aka-ball1 {
    background:radial-gradient(circle at 30% 28%, #8fff9f 0%, #46dc82 36%, #2aa6a3 68%, #305d72 100%);
}
.aka-ball2 {
    background:radial-gradient(circle at 30% 28%, #8b809c 0%, #4c425f 40%, #2d253d 70%, #181821 100%);
}
.aka-title {
    color:#fff;
    font-size:2.7rem;
    font-weight:900;
    letter-spacing:-1.4px;
    text-shadow:0 2px 2px rgba(0,0,0,.55);
}
.aka-section-title {
    color:#fff;
    font-size:1.55rem;
    font-weight:900;
    margin:.15rem 0 .5rem 0;
}
.chart-shell {
    background:#020202;
    border-radius:8px;
    border:1px solid #0f0f0f;
    padding:4px 5px 0 5px;
}
.subtle-title {
    font-size:1.1rem;
    font-weight:800;
    margin:0 0 .35rem 0;
}
hr {
    border-color:#303030 !important;
}

/* Streamlit select boxes: dark, readable in sidebar and main content */
div[data-baseweb="select"] > div {
    background-color:#202020 !important;
    color:#f5f5f5 !important;
    border-color:#3a3a3a !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] input {
    color:#f5f5f5 !important;
    -webkit-text-fill-color:#f5f5f5 !important;
}
div[data-baseweb="popover"] {
    color:#f5f5f5 !important;
}
div[data-baseweb="popover"] ul {
    background:#202020 !important;
}
div[data-baseweb="popover"] li {
    color:#f5f5f5 !important;
    background:#202020 !important;
}
div[data-baseweb="popover"] li:hover {
    background:#303030 !important;
}


/* ---------- Readability fixes for forms, labels and menus ---------- */
h1, h2, h3, h4, h5, h6,
p, label, span, div {
    color: var(--text);
}
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] strong {
    color: #f5f5f5 !important;
}
label, [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p {
    color: #f5f5f5 !important;
    font-weight: 700 !important;
    opacity: 1 !important;
}
[data-testid="stCaptionContainer"] {
    color: #d7d7d7 !important;
}
.small-note, .subtle-title {
    color: #f0f0f0 !important;
}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] details summary p {
    color: #f5f5f5 !important;
    font-weight: 700 !important;
}
[data-testid="stExpander"] {
    background: #111 !important;
    border: 1px solid #2d2d2d !important;
    border-radius: 10px !important;
}
hr { border-color:#303030 !important; opacity:1 !important; }

/* Inputs */
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stDateInput"] input,
div[data-testid="stTimeInput"] input {
    background: #202020 !important;
    color: #f5f5f5 !important;
    -webkit-text-fill-color: #f5f5f5 !important;
    border: 1px solid #3a3a3a !important;
    border-radius: 10px !important;
}
div[data-testid="stTextInput"] input::placeholder,
div[data-testid="stTextArea"] textarea::placeholder {
    color: #a9a9a9 !important;
    -webkit-text-fill-color: #a9a9a9 !important;
}
div[data-testid="stTextInput"] > div,
div[data-testid="stNumberInput"] > div,
div[data-testid="stTextArea"] > div,
div[data-testid="stDateInput"] > div {
    color: #f5f5f5 !important;
}

/* Selectbox / multiselect */
div[data-baseweb="select"] > div,
div[data-baseweb="base-input"] {
    background-color:#202020 !important;
    color:#f5f5f5 !important;
    border:1px solid #3a3a3a !important;
    border-radius:10px !important;
    min-height: 42px !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] div,
div[data-baseweb="select"] input,
div[data-baseweb="base-input"] input {
    color:#f5f5f5 !important;
    -webkit-text-fill-color:#f5f5f5 !important;
    opacity:1 !important;
}
div[data-baseweb="popover"] ul,
ul[role="listbox"] {
    background:#202020 !important;
    color:#f5f5f5 !important;
    border:1px solid #3a3a3a !important;
}
div[data-baseweb="popover"] li,
ul[role="listbox"] li {
    color:#f5f5f5 !important;
    background:#202020 !important;
}
div[data-baseweb="popover"] li:hover,
ul[role="listbox"] li:hover {
    background:#303030 !important;
}

/* Segmented control / radio pills */
[data-baseweb="button-group"] button {
    background:#1b1b1b !important;
    color:#f5f5f5 !important;
    border:1px solid #444 !important;
    font-weight:700 !important;
}
[data-baseweb="button-group"] button[aria-pressed="true"] {
    background:#14dca2 !important;
    color:#08100d !important;
    border-color:#14dca2 !important;
}

/* Buttons */
div.stButton > button {
    background:#202020 !important;
    color:#f5f5f5 !important;
    border:1px solid #3a3a3a !important;
    font-weight:700 !important;
}
div.stButton > button:hover {
    border-color:#14dca2 !important;
    color:#ffffff !important;
}
div.stButton > button[kind="primary"] {
    background:#14dca2 !important;
    color:#08100d !important;
    border:none !important;
    font-weight:800 !important;
}
div.stButton > button:disabled {
    background:#2b2b2b !important;
    color:#8c8c8c !important;
    border-color:#3a3a3a !important;
    opacity:1 !important;
}

/* Sidebar select/readability */
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] p {
    color:#f5f5f5 !important;
}


/* --- v14 readability fixes for selected values --- */

/* Sidebar selectbox selected value */
[data-testid="stSidebar"] div[data-baseweb="select"] > div,
[data-testid="stSidebar"] div[data-baseweb="base-input"] {
    background:#202020 !important;
    color:#ffffff !important;
    border:1px solid #ff5b5b !important;
    border-radius:10px !important;
}
[data-testid="stSidebar"] div[data-baseweb="select"] * ,
[data-testid="stSidebar"] div[data-baseweb="base-input"] * {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
}
[data-testid="stSidebar"] div[data-baseweb="select"] input::placeholder {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
}

/* Global selected values in selectboxes */
div[data-baseweb="select"] [data-testid="stMarkdownContainer"],
div[data-baseweb="select"] [data-testid="stMarkdownContainer"] p,
div[data-baseweb="select"] [data-baseweb="tag"],
div[data-baseweb="select"] div,
div[data-baseweb="select"] span {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
}

/* Segmented control buttons: ensure inactive/disabled text remains readable */
[data-baseweb="button-group"] {
    gap: 0 !important;
}
[data-baseweb="button-group"] button {
    min-width: 120px !important;
    background:#f2f2f2 !important;
    color:#222222 !important;
    border:1px solid #d0d0d0 !important;
    opacity:1 !important;
}
[data-baseweb="button-group"] button[aria-pressed="true"] {
    background:#8b1116 !important;
    color:#ffffff !important;
    border:1px solid #ff5b5b !important;
}
[data-baseweb="button-group"] button[disabled],
[data-baseweb="button-group"] button[aria-disabled="true"] {
    background:#e6e6e6 !important;
    color:#8f8f8f !important;
    opacity:1 !important;
    border:1px solid #d0d0d0 !important;
}


/* --- v15 definitive sidebar team select fix --- */
[data-testid="stSidebar"] div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background-color:#202020 !important;
    border:1px solid #4a4a4a !important;
    color:#ffffff !important;
    min-height:42px !important;
}
[data-testid="stSidebar"] div[data-testid="stSelectbox"] div[data-baseweb="select"] div {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}
[data-testid="stSidebar"] div[data-testid="stSelectbox"] div[data-baseweb="select"] span {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}
[data-testid="stSidebar"] div[data-testid="stSelectbox"] div[data-baseweb="select"] input {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    background-color:#202020 !important;
}
[data-testid="stSidebar"] div[data-testid="stSelectbox"] svg {
    fill:#ffffff !important;
    color:#ffffff !important;
}
[data-testid="stSidebar"] div[data-testid="stSelectbox"] > label,
[data-testid="stSidebar"] div[data-testid="stSelectbox"] > label p {
    color:#ffffff !important;
    font-weight:800 !important;
}


/* --- v16 sidebar team selector as radio buttons --- */
[data-testid="stSidebar"] [role="radiogroup"] {
    gap: .25rem !important;
}
[data-testid="stSidebar"] [role="radiogroup"] label {
    background:#202020 !important;
    border:1px solid #3a3a3a !important;
    border-radius:9px !important;
    padding:.45rem .65rem !important;
    margin:0 !important;
}
[data-testid="stSidebar"] [role="radiogroup"] label p,
[data-testid="stSidebar"] [role="radiogroup"] label span,
[data-testid="stSidebar"] [role="radiogroup"] label div {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
    font-weight:700 !important;
}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
    background:#0f5f49 !important;
    border-color:#14dca2 !important;
}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) p,
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) span {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}


/* --- v18 readable click-count selector --- */
.stApp [role="radiogroup"] {
    gap:.45rem !important;
}
.stApp [role="radiogroup"] label {
    background:#202020 !important;
    border:1px solid #444 !important;
    border-radius:9px !important;
    padding:.42rem .7rem !important;
}
.stApp [role="radiogroup"] label p,
.stApp [role="radiogroup"] label span,
.stApp [role="radiogroup"] label div {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
    font-weight:700 !important;
}
.stApp [role="radiogroup"] label:has(input:checked) {
    background:#8b1116 !important;
    border-color:#ff5b5b !important;
}
.stApp [role="radiogroup"] label:has(input:checked) p,
.stApp [role="radiogroup"] label:has(input:checked) span,
.stApp [role="radiogroup"] label:has(input:checked) div {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}


/* --- v24: Dropdown-Menüs vollständig lesbar --- */

/* Popup / Menüfläche */
div[data-baseweb="popover"],
div[data-baseweb="menu"],
div[data-baseweb="select"] + div,
ul[role="listbox"],
[role="listbox"] {
    background:#202020 !important;
    color:#ffffff !important;
}

/* Jede einzelne Dropdown-Option */
li[role="option"],
div[role="option"],
[role="listbox"] [role="option"] {
    background:#202020 !important;
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
}

/* Sämtliche Textelemente innerhalb einer Option */
li[role="option"] *,
div[role="option"] *,
[role="listbox"] [role="option"] * {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
    opacity:1 !important;
}

/* Hover */
li[role="option"]:hover,
div[role="option"]:hover,
[role="listbox"] [role="option"]:hover {
    background:#353535 !important;
    color:#ffffff !important;
}

/* Aktuell fokussierte / markierte Option */
li[role="option"][aria-selected="true"],
div[role="option"][aria-selected="true"],
[role="listbox"] [role="option"][aria-selected="true"] {
    background:#0f5f49 !important;
    color:#ffffff !important;
}

li[role="option"][aria-selected="true"] *,
div[role="option"][aria-selected="true"] *,
[role="listbox"] [role="option"][aria-selected="true"] * {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}

/* BaseWeb verwendet je nach Streamlit-Version auch eigene Menücontainer */
[data-baseweb="menu"] > div,
[data-baseweb="menu"] ul,
[data-baseweb="menu"] li {
    background:#202020 !important;
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}

/* Überschreibt helle Theme-Farben in Portal-Elementen */
body > div[data-baseweb="popover"] *,
body > div[data-baseweb="menu"] * {
    color:#ffffff !important;
    -webkit-text-fill-color:#ffffff !important;
}


/* --- v29 sidebar logo smaller + more top-left --- */
[data-testid="stSidebar"] img {
    margin-top: -12px !important;
    margin-left: -10px !important;
}

</style>
""", unsafe_allow_html=True)

init_db()
TEAMS = ["U15", "U16", "U18", "JWR"]

with st.sidebar:
    st.image("assets/sv_ried_logo.png", width=85)
    page = st.radio("Navigation", ["Dashboard", "Gesamt Dashboard", "Tor / Gegentor erfassen", "Spiel anlegen"], label_visibility="collapsed")
    st.divider()
    st.markdown("**Team**")
    team = st.radio(
        "Team auswählen",
        TEAMS,
        index=0,
        label_visibility="collapsed",
        key="sidebar_team"
    )



PITCH_IMAGE_W = 650
PITCH_IMAGE_H = 1000
PITCH_LEFT = 35
PITCH_RIGHT = 615
PITCH_TOP = 35
PITCH_BOTTOM = 965
PITCH_CAPTURE_WIDTH = 620

PENALTY_SPOT_X = round((105.0 - 11.0) / 105.0 * 100.0, 2)
PENALTY_SPOT = (PENALTY_SPOT_X, 50.0)
CORNER_LEFT = (100.0, 0.0)
CORNER_RIGHT = (100.0, 100.0)
SET_PIECE_OPTIONS = ["Keine", "Eckball links", "Eckball rechts", "Elfmeter"]


def image_click_to_pitch(click, rendered_width=PITCH_CAPTURE_WIDTH):
    """
    Convert click coordinates on the vertical capture pitch to canonical
    pitch coordinates.

    Visual attacking direction: bottom -> top.
    Canonical x: 0 own goal -> 100 attacking goal.
    Canonical y: 0 left side -> 100 right side.
    """
    rendered_height = rendered_width * PITCH_IMAGE_H / PITCH_IMAGE_W
    scale_x = rendered_width / PITCH_IMAGE_W
    scale_y = rendered_height / PITCH_IMAGE_H

    left = PITCH_LEFT * scale_x
    right = PITCH_RIGHT * scale_x
    top = PITCH_TOP * scale_y
    bottom = PITCH_BOTTOM * scale_y

    # horizontal screen position maps to canonical y
    y = (float(click["x"]) - left) / (right - left) * 100.0

    # vertical screen position is inverted:
    # bottom = canonical x 0, top = canonical x 100
    x = (bottom - float(click["y"])) / (bottom - top) * 100.0

    return (
        round(max(0.0, min(100.0, x)), 2),
        round(max(0.0, min(100.0, y)), 2),
    )


def pitch_to_image_xy(point, image_width=PITCH_IMAGE_W, image_height=PITCH_IMAGE_H):
    """
    Convert canonical pitch coordinates to the vertical input image.
    x 0 -> bottom, x 100 -> top.
    y 0 -> left, y 100 -> right.
    """
    if not point:
        return None

    x, y = point
    px = PITCH_LEFT + (float(y) / 100.0) * (PITCH_RIGHT - PITCH_LEFT)
    py = PITCH_BOTTOM - (float(x) / 100.0) * (PITCH_BOTTOM - PITCH_TOP)
    return int(round(px)), int(round(py))


def apply_set_piece_points(points, set_piece_type, click_count):
    points = dict(points)

    if set_piece_type == "Elfmeter":
        points["finish"] = PENALTY_SPOT
        points["start"] = None
        points["assist"] = None
        return points, 1

    if set_piece_type == "Eckball links":
        points["assist"] = CORNER_LEFT
        if click_count < 2:
            click_count = 2
        if click_count == 2:
            points["start"] = None
        return points, click_count

    if set_piece_type == "Eckball rechts":
        points["assist"] = CORNER_RIGHT
        if click_count < 2:
            click_count = 2
        if click_count == 2:
            points["start"] = None
        return points, click_count

    return points, click_count


def build_annotated_pitch(points, key_suffix):
    """Create a temporary pitch image with visible event points and connecting lines."""
    base_path = Path(__file__).parent / "assets" / "pitch_vertical.png"
    image = Image.open(base_path).convert("RGB")
    draw = ImageDraw.Draw(image)

    ordered = [
        ("start", "1", (255, 215, 0)),
        ("assist", "2", (255, 255, 255)),
        ("finish", "3", (220, 40, 40)),
    ]

    coords = []
    for point_key, number, color in ordered:
        point = points.get(point_key)
        if point:
            x, y = pitch_to_image_xy(point, image.width, image.height)
            coords.append((x, y, color, number))

    for i in range(len(coords) - 1):
        x1, y1, _, _ = coords[i]
        x2, y2, _, _ = coords[i + 1]
        draw.line((x1, y1, x2, y2), fill=(250, 250, 250), width=8)
        draw.line((x1, y1, x2, y2), fill=(20, 20, 20), width=3)

    for x, y, color, number in coords:
        radius = 22
        draw.ellipse((x-radius, y-radius, x+radius, y+radius), fill=color, outline=(15,15,15), width=4)
        font = ImageFont.load_default()
        bbox = draw.textbbox((0,0), number, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        draw.text((x - tw/2, y - th/2 - 1), number, fill=(0,0,0), font=font)

    out_dir = Path(__file__).parent / ".tmp_pitch"
    out_dir.mkdir(exist_ok=True)
    safe_key = "".join(c if c.isalnum() or c in "-_" else "_" for c in str(key_suffix))
    out_path = out_dir / f"pitch_{safe_key}.png"
    image.save(out_path)
    return out_path


def match_label(m):
    competition = m.get("competition") or "-"
    return f"{m['match_date']} · {m['opponent']} · {competition} · {m.get('home_away') or '-'}"


def render_dashboard_header():
    st.markdown("""
    <div class="aka-hero">
        <div class="aka-leftmark">FORZA<br>RIED</div>
        <div class="aka-center">
            <div class="aka-balls">
                <div class="aka-ball1"></div>
                <div class="aka-ball2"></div>
            </div>
            <div class="aka-title">AKA Goals Dashboard</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def zone_bucket(x, y):
    if pd.isna(x) or pd.isna(y):
        return "Rest"

    if x >= 83.5 and 23 <= y <= 77:
        return "Box zentral"
    if x >= 74 and y < 23:
        return "Flügel links"
    if x >= 74 and y > 77:
        return "Flügel rechts"
    if x >= 66.67 and 23 <= y < 38:
        return "Halbraum links"
    if x >= 66.67 and 62 < y <= 77:
        return "Halbraum rechts"
    if x >= 66.67 and 38 <= y <= 62:
        return "Zone 14"
    return "Distanz"


def zone_percentages(df, coord_prefix="finish"):
    labels = ["Box zentral", "Halbraum links", "Halbraum rechts", "Zone 14", "Flügel links", "Flügel rechts", "Distanz"]
    result = {k: 0.0 for k in labels}
    if df is None or df.empty:
        return result

    x_col = f"{coord_prefix}_x"
    y_col = f"{coord_prefix}_y"
    if x_col not in df.columns or y_col not in df.columns:
        return result

    valid = df[df[x_col].notna() & df[y_col].notna()].copy()
    if valid.empty:
        return result

    bucketed = valid.apply(lambda r: zone_bucket(r.get(x_col), r.get(y_col)), axis=1)
    counts = bucketed.value_counts()
    total = len(valid)
    for label in labels:
        result[label] = round(counts.get(label, 0) / total * 100, 1) if total else 0.0
    return result




PITCH_LENGTH_M = 105.0
PITCH_WIDTH_M = 68.0
HALF_PITCH_LENGTH_M = PITCH_LENGTH_M / 2.0


def canonical_to_dashboard_xy(length_pct, width_pct):
    """
    Convert canonical 0..100 pitch coordinates to dashboard half-pitch metres.

    Canonical coordinates:
      length_pct: 0 = own goal, 100 = opponent goal
      width_pct:  0 = left touchline, 100 = right touchline

    Dashboard:
      x = width in metres, left -> right
      y = metres from opponent goal line, goal -> halfway line
    """
    if length_pct is None or width_pct is None:
        return None

    length_pct = max(0.0, min(100.0, float(length_pct)))
    width_pct = max(0.0, min(100.0, float(width_pct)))

    plot_x = (width_pct / 100.0) * PITCH_WIDTH_M
    distance_from_opponent_goal = ((100.0 - length_pct) / 100.0) * PITCH_LENGTH_M

    # The dashboard only shows the attacking half. Values from the own half
    # are clipped to the halfway line rather than being spatially compressed.
    plot_y = max(0.0, min(HALF_PITCH_LENGTH_M, distance_from_opponent_goal))

    return plot_x, plot_y


def dashboard_xy_to_canonical(plot_x, plot_y):
    """Inverse helper, useful for validation/testing."""
    width_pct = (float(plot_x) / PITCH_WIDTH_M) * 100.0
    length_pct = 100.0 - (float(plot_y) / PITCH_LENGTH_M) * 100.0
    return length_pct, width_pct


def prepare_chart_df(events_df, mirror=False):
    """
    Prepare finish / assist coordinates for chart rendering.
    mirror=True is used for Gegentore so the attacking direction is displayed
    consistently toward the shown goal.
    """
    if events_df is None or events_df.empty:
        return events_df

    df = events_df.copy()
    for col in ["finish_x", "finish_y", "assist_x", "assist_y"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if mirror:
        if "finish_x" in df.columns:
            df["finish_x"] = 100 - df["finish_x"]
            df["finish_y"] = 100 - df["finish_y"]
        if "assist_x" in df.columns:
            df["assist_x"] = 100 - df["assist_x"]
            df["assist_y"] = 100 - df["assist_y"]

    return df

def draw_half_pitch(
    ax,
    events_df,
    title,
    point_color="#58ff63",
    mirror=False,
    point_mode="finish",
    connect_assists=False,
    show_assist_points=False,
    finish_marker_color=None
):
    ax.set_facecolor("#171717")
    green = "#12f2ad"
    white = "#d8d8d8"

    ax.set_xlim(-0.5, 68.5)
    ax.set_ylim(53.5, -1.0)

    ax.add_patch(Rectangle((25.0, 0), 18.0, 16.5,
                           facecolor="#7d7010", alpha=.52, edgecolor="none", zorder=0))
    ax.add_patch(Rectangle((25.0, 16.5), 18.0, 8.2,
                           facecolor="#7f0f12", alpha=.58, edgecolor="none", zorder=0))

    ax.plot([0, 68], [0, 0], color=white, lw=1.45, zorder=2)
    ax.plot([0, 0], [0, 52.5], color=white, lw=1.45, zorder=2)
    ax.plot([68, 68], [0, 52.5], color=white, lw=1.45, zorder=2)
    ax.plot([0, 68], [52.5, 52.5], color=green, lw=2.2, zorder=2)

    ax.add_patch(Rectangle((13.84, 0), 40.32, 16.5, fill=False, ec=green, lw=2.2, zorder=3))
    ax.add_patch(Rectangle((24.84, 0), 18.32, 5.5, fill=False, ec=green, lw=2.2, zorder=3))

    ax.plot([13.84, 13.84], [0, 24.5], color=green, lw=2.0, zorder=2)
    ax.plot([54.16, 54.16], [0, 24.5], color=green, lw=2.0, zorder=2)
    ax.plot([0, 13.84], [10, 10], color=green, lw=2.0, ls="--", zorder=2)
    ax.plot([54.16, 68], [10, 10], color=green, lw=2.0, ls="--", zorder=2)
    ax.plot([24.84, 24.84], [5.5, 24.5], color=green, lw=2.0, ls="--", zorder=2)
    ax.plot([43.16, 43.16], [5.5, 24.5], color=green, lw=2.0, ls="--", zorder=2)
    ax.plot([13.84, 54.16], [16.5, 16.5], color=green, lw=2.0, zorder=2)

    ax.add_patch(Circle((34, 11), 0.35, color=green, zorder=4))
    ax.add_patch(Arc((34, 11), 18.3, 18.3, theta1=37, theta2=143,
                     color=green, lw=2.1, zorder=3))
    ax.add_patch(Circle((34, 52.5), 9.15, fill=False, ec=green, lw=2.2, zorder=3))

    chart_df = prepare_chart_df(events_df, mirror=mirror)

    def pitch_to_plot(x, y):
        return canonical_to_dashboard_xy(x, y)

    if connect_assists and chart_df is not None and not chart_df.empty:
        valid_links = chart_df[
            chart_df["assist_x"].notna() & chart_df["assist_y"].notna() &
            chart_df["finish_x"].notna() & chart_df["finish_y"].notna()
        ]
        for _, row in valid_links.iterrows():
            x1, y1 = pitch_to_plot(float(row["assist_x"]), float(row["assist_y"]))
            x2, y2 = pitch_to_plot(float(row["finish_x"]), float(row["finish_y"]))
            ax.plot([x1, x2], [y1, y2], color=(1, 1, 1, 0.34), lw=1.8, zorder=5)

        if show_assist_points and not valid_links.empty:
            assist_xy = [
                canonical_to_dashboard_xy(row["assist_x"], row["assist_y"])
                for _, row in valid_links.iterrows()
            ]
            ax.scatter(
                [p[0] for p in assist_xy],
                [p[1] for p in assist_xy],
                s=34, c="#ffffff", alpha=0.85, edgecolors="none", zorder=6
            )

    coord_prefix = "assist" if point_mode == "assist" else "finish"
    if chart_df is not None and not chart_df.empty:
        x_col = f"{coord_prefix}_x"
        y_col = f"{coord_prefix}_y"
        valid_points = chart_df[chart_df[x_col].notna() & chart_df[y_col].notna()]
        if not valid_points.empty:
            mapped_points = [
                canonical_to_dashboard_xy(row[x_col], row[y_col])
                for _, row in valid_points.iterrows()
            ]
            x_plot = [p[0] for p in mapped_points]
            y_plot = [p[1] for p in mapped_points]
            marker_color = (
                finish_marker_color
                if point_mode == "finish" and finish_marker_color
                else point_color
            )
            ax.scatter(
                x_plot, y_plot,
                s=58, c=marker_color, edgecolors="#f0f0f0",
                linewidths=1.2, zorder=7
            )

    ax.set_title(title, color="white", fontsize=19, pad=18, weight="bold")
    ax.set_xticks([])
    ax.set_yticks([])
    for spine in ax.spines.values():
        spine.set_visible(False)

    ax.text(-1.9, 49.0, "Spielfeldtiefe (m)", color="#f1f1f1",
            fontsize=9, rotation=90, va="center", ha="center")

    perc = zone_percentages(chart_df, coord_prefix=coord_prefix)

    def badge(x, y, text, fc="#0b0b0b", ec="#303030", tc="#8f8f8f"):
        ax.text(
            x, y, text,
            ha="center", va="center",
            color=tc, fontsize=11.5, weight="bold",
            bbox=dict(boxstyle="round,pad=0.28", fc=fc, ec=ec, lw=1.25, alpha=.96),
            zorder=10
        )

    badge(34, 0.1, f"{perc['Box zentral']}%",
          fc="#6e6410", ec="#d6bd00", tc="#ffe000")
    badge(20.5, 7.9, f"{perc['Halbraum links']}%")
    badge(47.7, 7.9, f"{perc['Halbraum rechts']}%")
    badge(34, 24.5, f"{perc['Zone 14']}%",
          fc="#430909", ec="#d53939", tc="#ff5656")
    badge(19.0, 20.4, f"{perc['Flügel links']}%")
    badge(49.1, 20.4, f"{perc['Flügel rechts']}%")
    badge(34, 47.0, f"{perc['Distanz']}%")

    return ax

def create_dashboard_figure(
    events_df,
    title,
    point_color="#58ff63",
    mirror=False,
    point_mode="finish",
    connect_assists=False,
    show_assist_points=False,
    finish_marker_color=None
):
    fig, ax = plt.subplots(figsize=(8.4, 6.7), facecolor="#020202")
    draw_half_pitch(
        ax,
        events_df,
        title,
        point_color=point_color,
        mirror=mirror,
        point_mode=point_mode,
        connect_assists=connect_assists,
        show_assist_points=show_assist_points,
        finish_marker_color=finish_marker_color
    )
    plt.tight_layout(pad=.65)
    return fig



def event_points_from_row(event):
    return {
        "start": (
            float(event.get("start_x") or 0),
            float(event.get("start_y") or 0),
        ) if event.get("start_x") is not None and event.get("start_y") is not None else None,
        "assist": (
            float(event.get("assist_x") or 0),
            float(event.get("assist_y") or 0),
        ) if event.get("assist_x") is not None and event.get("assist_y") is not None else None,
        "finish": (
            float(event.get("finish_x") or 0),
            float(event.get("finish_y") or 0),
        ) if event.get("finish_x") is not None and event.get("finish_y") is not None else None,
    }


def render_event_editor(event):
    event_id = event["id"]
    points_key = f"edit_points_{event_id}"
    nonce_key = f"edit_nonce_{event_id}"

    if points_key not in st.session_state:
        st.session_state[points_key] = event_points_from_row(event)
    if nonce_key not in st.session_state:
        st.session_state[nonce_key] = 0

    points = st.session_state[points_key]

    st.markdown("---")
    st.markdown(f"## Szene bearbeiten · {event['event_type']} {event.get('minute') or '-'}'")

    current_click_count = int(event.get("click_count") or 3)
    edit_click_count = st.radio(
        "Punkte der Szene",
        [1, 2, 3],
        index=[1,2,3].index(current_click_count) if current_click_count in [1,2,3] else 2,
        horizontal=True,
        format_func=lambda n: f"{n} Klick" if n == 1 else f"{n} Klicks",
        key=f"edit_click_count_{event_id}",
    )

    current_set_piece = event.get("set_piece_type") if event.get("set_piece_type") in SET_PIECE_OPTIONS else "Keine"
    edit_set_piece = st.radio(
        "Standardsituation",
        SET_PIECE_OPTIONS,
        index=SET_PIECE_OPTIONS.index(current_set_piece),
        horizontal=True,
        key=f"edit_set_piece_{event_id}",
    )

    points, edit_click_count = apply_set_piece_points(points, edit_set_piece, edit_click_count)
    st.session_state[points_key] = points

    if edit_set_piece == "Elfmeter":
        st.info("Elfmeter: Abschluss ist automatisch am Elfmeterpunkt gesetzt.")
    elif edit_set_piece.startswith("Eckball"):
        st.info("Eckball: Assistpunkt ist automatisch am gewählten Eckpunkt gesetzt.")

    edit_labels = {
        "start": "1 · Angriffsbeginn",
        "assist": "2 · Assist / letzter Pass",
        "finish": "3 · Abschluss",
    }

    if edit_click_count == 1:
        active_edit_points = ["finish"]
        points["start"] = None
        points["assist"] = None
    elif edit_click_count == 2:
        active_edit_points = ["assist", "finish"]
        points["start"] = None
    else:
        active_edit_points = ["start", "assist", "finish"]

    point_to_edit = st.radio(
        "Welchen Punkt möchtest du versetzen?",
        options=active_edit_points,
        format_func=lambda x: edit_labels[x],
        horizontal=True,
        index=len(active_edit_points)-1,
        key=f"edit_point_choice_{event_id}",
    )

    annotated = build_annotated_pitch(points, f"edit_{event_id}_{st.session_state[nonce_key]}")
    st.caption("Angriffsrichtung bei der Bearbeitung: **von unten nach oben**")
    st.caption(
        f"Klicke auf die neue Position für **{edit_labels[point_to_edit]}**. "
        "Die anderen beiden Punkte bleiben unverändert."
    )

    click = streamlit_image_coordinates(
        str(annotated),
        width=PITCH_CAPTURE_WIDTH,
        key=f"edit_pitch_{event_id}_{point_to_edit}_{st.session_state[nonce_key]}"
    )

    if click:
        new_xy = image_click_to_pitch(click, rendered_width=PITCH_CAPTURE_WIDTH)

        fixed_point = (
            (edit_set_piece == "Elfmeter" and point_to_edit == "finish") or
            (edit_set_piece.startswith("Eckball") and point_to_edit == "assist")
        )
        if fixed_point:
            st.warning("Dieser Punkt ist für die gewählte Standardsituation fest definiert.")
        elif points.get(point_to_edit) != new_xy:
            points[point_to_edit] = new_xy
            st.session_state[points_key] = points
            st.session_state[nonce_key] += 1
            st.rerun()

    point_cols = st.columns(len(active_edit_points))
    for col, key in zip(point_cols, active_edit_points):
        with col:
            p = points.get(key)
            st.markdown(f"**{edit_labels[key]}**")
            if p:
                st.write(derive_zone(*p))
                dash_xy = canonical_to_dashboard_xy(p[0], p[1])
                st.caption(
                    f"Feld: L {p[0]:.1f}% · B {p[1]:.1f}%"
                    f"  |  Dashboard: {dash_xy[0]:.1f}m / {dash_xy[1]:.1f}m"
                )
            else:
                st.caption("nicht gesetzt")

    st.markdown("### Ereignisdaten")
    ec1, ec2, ec3 = st.columns(3)

    with ec1:
        edit_minute = st.number_input(
            "Minute",
            min_value=0,
            max_value=130,
            value=int(event.get("minute") or 0),
            step=1,
            key=f"edit_minute_{event_id}",
        )
        edit_scorer = st.text_input(
            "Torschütze" if event["event_type"] == "Tor" else "Gegnerischer Torschütze",
            value=event.get("scorer") or "",
            key=f"edit_scorer_{event_id}",
        )
        edit_assister = st.text_input(
            "Assist / letzter Pass",
            value=event.get("assister") or "",
            key=f"edit_assister_{event_id}",
        )

        current_touch = event.get("finish_touch") if event.get("finish_touch") in TOUCH_OPTIONS else "One Touch"
        edit_finish_touch = st.radio(
            "Torabschluss",
            TOUCH_OPTIONS,
            horizontal=True,
            index=TOUCH_OPTIONS.index(current_touch),
            key=f"edit_finish_touch_{event_id}",
        )

    phase_options = [
        "Organisierter Ballbesitz",
        "Umschalten nach Ballgewinn",
        "Standard",
        "Zweiter Ball",
        "Sonstiges",
    ]
    creation_options = [
        "Steckpass",
        "Flanke",
        "Cutback",
        "Dribbling",
        "Fernschuss",
        "Standard",
        "Nachschuss",
        "Sonstiges",
    ]

    with ec2:
        current_phase = event.get("phase") or phase_options[0]
        if current_phase not in phase_options:
            phase_options = [current_phase] + phase_options
        edit_phase = st.selectbox(
            "Spielphase",
            phase_options,
            index=phase_options.index(current_phase),
            key=f"edit_phase_{event_id}",
        )

        current_creation = event.get("creation_type") or creation_options[0]
        if current_creation not in creation_options:
            creation_options = [current_creation] + creation_options
        edit_creation = st.selectbox(
            "Entstehung",
            creation_options,
            index=creation_options.index(current_creation),
            key=f"edit_creation_{event_id}",
        )

    with ec3:
        edit_video = st.text_input(
            "Video-Link",
            value=event.get("video_url") or "",
            key=f"edit_video_{event_id}",
        )
        edit_comment = st.text_area(
            "Kommentar",
            value=event.get("comment") or "",
            height=95,
            key=f"edit_comment_{event_id}",
        )

    a1, a2, a3 = st.columns([1.4, 1.2, 4])

    with a1:
        if st.button(
            "Änderungen speichern",
            type="primary",
            key=f"save_edit_{event_id}",
            use_container_width=True,
        ):
            s = points.get("start")
            a = points.get("assist")
            f = points.get("finish")

            required_keys = active_edit_points
            if not all(points.get(k) is not None for k in required_keys):
                st.error("Bitte alle gewählten Punkte setzen.")
            else:
                update_event(event_id, {
                    "minute": int(edit_minute),
                    "scorer": edit_scorer.strip(),
                    "assister": edit_assister.strip(),
                    "phase": "Standard" if edit_set_piece != "Keine" else edit_phase,
                    "creation_type": edit_set_piece if edit_set_piece != "Keine" else edit_creation,
                    "start_x": s[0] if s else None,
                    "start_y": s[1] if s else None,
                    "start_zone": derive_zone(*s) if s else None,
                    "assist_x": a[0] if a else None,
                    "assist_y": a[1] if a else None,
                    "assist_zone": derive_zone(*a) if a else None,
                    "finish_x": f[0] if f else None,
                    "finish_y": f[1] if f else None,
                    "finish_zone": derive_zone(*f) if f else None,
                    "video_url": edit_video.strip(),
                    "comment": edit_comment.strip(),
                    "click_count": int(edit_click_count),
                    "finish_touch": edit_finish_touch,
                    "set_piece_type": None if edit_set_piece == "Keine" else edit_set_piece,
                })

                st.session_state.pop(points_key, None)
                st.session_state.pop(nonce_key, None)
                st.session_state.pop("editing_event_id", None)
                st.success("Szene aktualisiert.")
                st.rerun()

    with a2:
        if st.button(
            "Abbrechen",
            key=f"cancel_edit_{event_id}",
            use_container_width=True,
        ):
            st.session_state.pop(points_key, None)
            st.session_state.pop(nonce_key, None)
            st.session_state.pop("editing_event_id", None)
            st.rerun()

    with a3:
        if st.button(
            "Originalpunkte wiederherstellen",
            key=f"reset_edit_{event_id}",
        ):
            st.session_state[points_key] = event_points_from_row(event)
            st.session_state[nonce_key] += 1
            st.rerun()



TOUCH_OPTIONS = ["One Touch", "Two Touch", ">2 Touches"]

def touch_summary(df):
    """Return counts and percentages for finish touch categories."""
    result = pd.DataFrame({"Abschluss": TOUCH_OPTIONS, "Anzahl": [0, 0, 0]})
    if df is None or df.empty or "finish_touch" not in df.columns:
        result["Prozent"] = 0.0
        return result

    valid = df[df["finish_touch"].isin(TOUCH_OPTIONS)]
    counts = valid["finish_touch"].value_counts()
    result["Anzahl"] = result["Abschluss"].map(counts).fillna(0).astype(int)
    total = int(result["Anzahl"].sum())
    result["Prozent"] = (result["Anzahl"] / total * 100).round(1) if total else 0.0
    return result


def render_touch_analysis(df, title):
    st.markdown(f"### {title}")
    summary = touch_summary(df)

    total = int(summary["Anzahl"].sum())
    if total == 0:
        st.caption("Noch keine Touch-Angaben erfasst.")
        return

    c1, c2, c3 = st.columns(3)
    for col, label in zip([c1, c2, c3], TOUCH_OPTIONS):
        row = summary[summary["Abschluss"] == label].iloc[0]
        with col:
            st.metric(label, int(row["Anzahl"]), f"{row['Prozent']:.1f}%")

    tc1, tc2, tc3 = st.columns([1, 2.4, 1])
    with tc2:
        fig, ax = plt.subplots(figsize=(5.4, 2.45), facecolor="#171717")
        ax.set_facecolor("#171717")
        bars = ax.bar(summary["Abschluss"], summary["Prozent"])
        ax.set_ylim(0, max(100, float(summary["Prozent"].max()) + 15))
        ax.set_ylabel("Anteil (%)", color="white", fontsize=9)
        ax.tick_params(axis="x", colors="white", labelsize=9)
        ax.tick_params(axis="y", colors="white", labelsize=8)
        for spine in ax.spines.values():
            spine.set_color("#555555")
        for bar, pct, count in zip(bars, summary["Prozent"], summary["Anzahl"]):
            ax.text(
                bar.get_x() + bar.get_width()/2,
                bar.get_height() + 2,
                f"{pct:.1f}% ({int(count)})",
                ha="center", va="bottom", color="white", fontweight="bold", fontsize=8.5
            )
        fig.tight_layout(pad=.8)
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)


def render_assist_analysis(df, title, mirror=False, full_size=False, always_show=False):
    st.markdown(f"### {title}")

    has_assist_columns = df is not None and not df.empty and "assist_x" in df.columns
    assists_df = (
        df[df["assist_x"].notna() & df["assist_y"].notna()].copy()
        if has_assist_columns
        else None
    )

    if assists_df is None or assists_df.empty:
        if not always_show:
            st.caption("Noch keine Assists erfasst.")
            return

        st.metric("Anzahl Assists", 0)
        st.caption("Noch keine Assists erfasst. Die Grafik bleibt sichtbar.")
        empty_df = df.iloc[0:0].copy() if df is not None and hasattr(df, "iloc") else None

        if full_size:
            fig = create_dashboard_figure(
                empty_df,
                title,
                point_color="#f4f4f4",
                mirror=mirror,
                point_mode="assist"
            )
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)
        else:
            ac1, ac2, ac3 = st.columns([1, 2.2, 1])
            with ac2:
                fig, ax = plt.subplots(figsize=(5.6, 4.35), facecolor="#020202")
                draw_half_pitch(
                    ax,
                    empty_df,
                    title,
                    point_color="#f4f4f4",
                    mirror=mirror,
                    point_mode="assist"
                )
                plt.tight_layout(pad=.55)
                st.pyplot(fig, use_container_width=True)
                plt.close(fig)
        return

    st.metric("Anzahl Assists", len(assists_df))
    st.caption("Die Punkte zeigen die Orte des letzten Passes / der Assistaktion.")

    if full_size:
        fig = create_dashboard_figure(
            assists_df,
            title,
            point_color="#f4f4f4",
            mirror=mirror,
            point_mode="assist"
        )
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)
    else:
        ac1, ac2, ac3 = st.columns([1, 2.2, 1])
        with ac2:
            fig, ax = plt.subplots(figsize=(5.6, 4.35), facecolor="#020202")
            draw_half_pitch(
                ax,
                assists_df,
                title,
                point_color="#f4f4f4",
                mirror=mirror,
                point_mode="assist"
            )
            plt.tight_layout(pad=.55)
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)


GOAL_TIME_BINS = [
    ("1–25", 1, 25),
    ("26–45", 26, 45),
    ("46–75", 46, 75),
    ("76–90", 76, 90),
]


def goal_time_summary(df):
    rows = []
    if df is None or df.empty or "minute" not in df.columns:
        for label, _, _ in GOAL_TIME_BINS:
            rows.append({"Zeitfenster": label, "Tore": 0, "Gegentore": 0})
        return pd.DataFrame(rows)

    work = df.copy()
    work["minute"] = pd.to_numeric(work["minute"], errors="coerce")

    for label, start_min, end_min in GOAL_TIME_BINS:
        bucket = work[(work["minute"] >= start_min) & (work["minute"] <= end_min)]
        rows.append({
            "Zeitfenster": label,
            "Tore": int((bucket["event_type"] == "Tor").sum()),
            "Gegentore": int((bucket["event_type"] == "Gegentor").sum()),
        })

    return pd.DataFrame(rows)


def render_goal_time_analysis(df, title):
    st.markdown(f"### {title}")
    summary = goal_time_summary(df)

    # compact metrics
    cols = st.columns(4)
    for col, row in zip(cols, summary.to_dict("records")):
        with col:
            st.metric(
                row["Zeitfenster"],
                f"{row['Tore']} : {row['Gegentore']}",
                help="Tore : Gegentore",
            )

    # compact grouped bar chart
    tc1, tc2, tc3 = st.columns([0.7, 3.2, 0.7])
    with tc2:
        fig, ax = plt.subplots(figsize=(6.2, 2.9), facecolor="#171717")
        ax.set_facecolor("#171717")

        positions = list(range(len(summary)))
        width = 0.36
        bars_goals = ax.bar(
            [p - width/2 for p in positions],
            summary["Tore"],
            width=width,
            label="Tore",
        )
        bars_against = ax.bar(
            [p + width/2 for p in positions],
            summary["Gegentore"],
            width=width,
            label="Gegentore",
        )

        ax.set_xticks(positions)
        ax.set_xticklabels(summary["Zeitfenster"], color="white", fontsize=9)
        ax.tick_params(axis="y", colors="white", labelsize=8)
        ax.set_ylabel("Anzahl", color="white", fontsize=9)
        ax.legend(fontsize=8, frameon=False, labelcolor="white")

        max_value = max(
            1,
            int(summary["Tore"].max()),
            int(summary["Gegentore"].max()),
        )
        ax.set_ylim(0, max_value + 1.4)

        for bars in (bars_goals, bars_against):
            for bar in bars:
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width()/2,
                    height + 0.08,
                    f"{int(height)}",
                    ha="center",
                    va="bottom",
                    color="white",
                    fontsize=8.5,
                    fontweight="bold",
                )

        for spine in ax.spines.values():
            spine.set_color("#555555")

        fig.tight_layout(pad=.7)
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)


if page == "Spiel anlegen":
    st.title(f"{team} · Spiel anlegen")
    with st.form("new_match", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            match_date = st.date_input("Datum", value=date.today())
            opponent = st.text_input("Gegner")
            home_away = st.selectbox("Ort", ["Heim", "Auswärts", "Neutral"])
        with c2:
            competition = st.selectbox("Bewerb", ["Testspiel", "Punktspiel"], index=1)
            st.info("Nach dem Speichern steht das Spiel sofort in der Tor-/Gegentor-Erfassung zur Verfügung.")
        ok = st.form_submit_button("Spiel speichern", type="primary")
    if ok:
        if not opponent.strip():
            st.error("Bitte einen Gegner eingeben.")
        else:
            create_match({
                "team": team,
                "match_date": match_date.isoformat(),
                "opponent": opponent.strip(),
                "competition": competition.strip(),
                "home_away": home_away,
                "created_by": None,
            })
            st.success("Spiel gespeichert.")

    st.divider()
    st.markdown("### Angelegte Spiele")
    existing_matches = get_matches(team)

    if not existing_matches:
        st.caption("Noch keine Spiele angelegt.")
    else:
        for m in existing_matches:
            with st.container(border=True):
                mc1, mc2, mc3 = st.columns([4, 2, 1.6])

                with mc1:
                    st.markdown(f"**{m['match_date']} · {m['opponent']}**")
                    st.caption(f"{m.get('competition') or '-'} · {m.get('home_away') or '-'}")

                with mc2:
                    match_events = get_events(team=team, match_id=m["id"])
                    goals_count = sum(1 for e in match_events if e["event_type"] == "Tor")
                    against_count = sum(1 for e in match_events if e["event_type"] == "Gegentor")
                    st.write(f"⚽ {goals_count} : {against_count} 🥅")
                    if match_events:
                        st.caption(f"{len(match_events)} Ereignis(se) gespeichert")

                with mc3:
                    confirm_match_key = f"confirm_delete_match_{m['id']}"

                    if not st.session_state.get(confirm_match_key, False):
                        if st.button(
                            "🗑️ Spiel löschen",
                            key=f"delete_match_{m['id']}",
                            use_container_width=True
                        ):
                            st.session_state[confirm_match_key] = True
                            st.rerun()
                    else:
                        st.warning("Spiel + Ereignisse löschen?" if match_events else "Spiel wirklich löschen?")

                        if st.button(
                            "Ja, löschen",
                            key=f"yes_delete_match_{m['id']}",
                            type="primary",
                            use_container_width=True
                        ):
                            delete_match(m["id"])
                            st.session_state.pop(confirm_match_key, None)

                            editing_id = st.session_state.get("editing_event_id")
                            if editing_id is not None and any(e["id"] == editing_id for e in match_events):
                                st.session_state.pop("editing_event_id", None)

                            st.success("Spiel gelöscht.")
                            st.rerun()

                        if st.button(
                            "Abbrechen",
                            key=f"cancel_delete_match_{m['id']}",
                            use_container_width=True
                        ):
                            st.session_state.pop(confirm_match_key, None)
                            st.rerun()

elif page == "Tor / Gegentor erfassen":
    st.title(f"{team} · Tor / Gegentor erfassen")
    matches = get_matches(team)
    if not matches:
        st.warning("Für dieses Team gibt es noch kein Spiel. Bitte zuerst ein Spiel anlegen.")
        st.stop()

    match_map = {match_label(m): m for m in matches}
    selected_label = st.selectbox("Spiel", list(match_map))
    match = match_map[selected_label]

    c1, c2, c3, c4 = st.columns([1.15,1,1.25,2])
    with c1:
        event_type = st.radio(
            "Ereignis",
            ["Tor", "Gegentor"],
            horizontal=True,
            index=0,
            key=f"event_type_{match['id']}"
        )
    with c2:
        minute = st.number_input("Minute", min_value=0, max_value=130, value=1, step=1)
    with c3:
        click_count = st.radio(
            "Punkte der Szene",
            [1, 2, 3],
            index=2,
            horizontal=True,
            format_func=lambda n: f"{n} Klick" if n == 1 else f"{n} Klicks",
            key=f"click_count_{match['id']}_{event_type}"
        )
    with c4:
        st.info("1 = nur Abschluss · 2 = Assist + Abschluss · 3 = Angriffsbeginn + Assist + Abschluss")

    set_piece_type = st.radio(
        "Standardsituation",
        SET_PIECE_OPTIONS,
        horizontal=True,
        index=0,
        key=f"set_piece_{match['id']}_{event_type}"
    )

    state_key = f"points_{match['id']}_{event_type}"
    if state_key not in st.session_state:
        st.session_state[state_key] = {"start": None, "assist": None, "finish": None}

    points = st.session_state[state_key]
    points, effective_click_count = apply_set_piece_points(points, set_piece_type, click_count)
    st.session_state[state_key] = points

    if set_piece_type == "Elfmeter":
        st.info("Elfmeter: Abschlussposition wird automatisch am Elfmeterpunkt gesetzt.")
    elif set_piece_type == "Eckball links":
        st.info("Eckball links: Assistpunkt wird automatisch am linken Eckpunkt gesetzt.")
    elif set_piece_type == "Eckball rechts":
        st.info("Eckball rechts: Assistpunkt wird automatisch am rechten Eckpunkt gesetzt.")

    click_count = effective_click_count

    if click_count == 1:
        steps = [("finish", "1 · Abschluss")]
        points["start"] = None
        points["assist"] = None
    elif click_count == 2:
        steps = [("assist", "1 · Assist / letzter Pass"), ("finish", "2 · Abschluss")]
        points["start"] = None
    else:
        steps = [("start", "1 · Angriffsbeginn"), ("assist", "2 · Assist / letzter Pass"), ("finish", "3 · Abschluss")]

    next_step = next((k for k, _ in steps if points[k] is None), None)

    st.subheader("Spielfeld")
    st.caption("Angriffsrichtung bei der Erfassung: **von unten nach oben**")
    pitch_key = f"{match['id']}_{event_type}_{sum(1 for p in points.values() if p)}"
    annotated_pitch = build_annotated_pitch(points, pitch_key)

    if next_step:
        step_name = dict(steps)[next_step]
        st.caption(f"Jetzt klicken: **{step_name}**")

        value = streamlit_image_coordinates(
            str(annotated_pitch),
            width=PITCH_CAPTURE_WIDTH,
            key=f"pitch_{match['id']}_{event_type}_{next_step}"
        )

        if value:
            xy = image_click_to_pitch(value, rendered_width=PITCH_CAPTURE_WIDTH)

            fixed_point = (
                (set_piece_type == "Elfmeter" and next_step == "finish") or
                (set_piece_type.startswith("Eckball") and next_step == "assist")
            )
            if fixed_point:
                st.warning("Dieser Punkt ist für die gewählte Standardsituation fest definiert.")
            elif points[next_step] != xy:
                points[next_step] = xy
                st.session_state[state_key] = points
                st.rerun()
    else:
        st.image(str(annotated_pitch), width=PITCH_CAPTURE_WIDTH)
        st.success("Alle drei Punkte sind erfasst. Der komplette Angriffsweg ist oben sichtbar.")

    cols = st.columns(3)
    for col, (key, label) in zip(cols, steps):
        with col:
            if points[key]:
                x,y = points[key]
                st.markdown(f"**{label}**")
                st.write(derive_zone(x,y))
                st.caption(f"x {x:.1f} · y {y:.1f}")
            else:
                st.markdown(f"**{label}**")
                st.caption("noch nicht gesetzt")

    if click_count == 1:
        st.caption("🔴 **1 Abschluss**")
    elif click_count == 2:
        st.caption("⚪ **1 Assist / letzter Pass**  ·  🔴 **2 Abschluss**")
    else:
        st.caption("🟡 **1 Angriffsbeginn**  ·  ⚪ **2 Assist / letzter Pass**  ·  🔴 **3 Abschluss**")

    b1,b2 = st.columns([1,3])
    with b1:
        if st.button("Punkte zurücksetzen"):
            st.session_state[state_key] = {"start":None,"assist":None,"finish":None}
            st.rerun()

    st.divider()
    c1,c2,c3 = st.columns(3)
    with c1:
        scorer = st.text_input("Torschütze" if event_type=="Tor" else "Gegnerischer Torschütze")
        assister = st.text_input("Assist / letzter Pass")
        finish_touch = st.radio(
            "Torabschluss",
            TOUCH_OPTIONS,
            horizontal=True,
            index=0,
            key=f"finish_touch_{match['id']}_{event_type}"
        )
    with c2:
        phase = st.selectbox("Spielphase", [
            "Organisierter Ballbesitz", "Umschalten nach Ballgewinn",
            "Standard", "Zweiter Ball", "Sonstiges"
        ])
        creation_type = st.selectbox("Entstehung", [
            "Steckpass", "Flanke", "Cutback", "Dribbling",
            "Fernschuss", "Standard", "Nachschuss", "Sonstiges"
        ])
    with c3:
        video_url = st.text_input("Video-Link")
        comment = st.text_area("Kommentar", height=90)

    ready = all(points[k] is not None for k, _ in steps)
    if st.button("Ereignis speichern", type="primary", disabled=not ready):
        s, a, f = points.get("start"), points.get("assist"), points.get("finish")
        insert_event({
            "match_id": match["id"], "team": team, "event_type": event_type,
            "minute": int(minute), "scorer": scorer.strip(), "assister": assister.strip(),
            "phase": "Standard" if set_piece_type != "Keine" else phase,
            "creation_type": set_piece_type if set_piece_type != "Keine" else creation_type,
            "start_x": s[0] if s else None,
            "start_y": s[1] if s else None,
            "start_zone": derive_zone(*s) if s else None,
            "assist_x": a[0] if a else None,
            "assist_y": a[1] if a else None,
            "assist_zone": derive_zone(*a) if a else None,
            "finish_x": f[0] if f else None,
            "finish_y": f[1] if f else None,
            "finish_zone": derive_zone(*f) if f else None,
            "video_url": video_url.strip(), "comment": comment.strip(),
            "click_count": int(click_count),
            "finish_touch": finish_touch,
            "set_piece_type": None if set_piece_type == "Keine" else set_piece_type,
            "created_by": None,
        })
        st.session_state[state_key] = {"start":None,"assist":None,"finish":None}
        st.success("Gespeichert.")
        st.rerun()

elif page == "Gesamt Dashboard":
    render_dashboard_header()
    st.markdown("# Gesamt Dashboard")

    all_events = get_events()
    if not all_events:
        st.info("Noch keine Tore/Gegentore erfasst.")
        st.stop()

    df_all = pd.DataFrame(all_events)
    goals_all = df_all[df_all["event_type"] == "Tor"].copy()
    against_all = df_all[df_all["event_type"] == "Gegentor"].copy()

    g1, g2, g3 = st.columns(3)
    g1.metric("Tore alle Teams", len(goals_all))
    g2.metric("Gegentore alle Teams", len(against_all))
    g3.metric("Differenz", len(goals_all) - len(against_all))

    st.caption("🔴 Abschluss · ⚪ Assist / letzter Pass · Linie = zugehörige Szene")
    p1, p2 = st.columns(2, gap="small")
    with p1:
        st.markdown('<div class="aka-section-title">Alle Teams - Eigene Tore</div>', unsafe_allow_html=True)
        fig_all_goals = create_dashboard_figure(
            goals_all,
            "Alle Teams - Eigene Tore",
            point_color="#5cff67",
            connect_assists=True,
            show_assist_points=True,
            finish_marker_color="#dc2828"
        )
        st.pyplot(fig_all_goals, use_container_width=True)
        plt.close(fig_all_goals)

    with p2:
        st.markdown('<div class="aka-section-title">Alle Teams - Gegentore</div>', unsafe_allow_html=True)
        fig_all_against = create_dashboard_figure(
            against_all,
            "Alle Teams - Gegentore",
            point_color="#ff6262",
            mirror=True,
            connect_assists=True,
            show_assist_points=True
        )
        st.pyplot(fig_all_against, use_container_width=True)
        plt.close(fig_all_against)

    st.divider()
    render_touch_analysis(goals_all, "Alle Teams – Torabschluss nach Kontakten")

    st.divider()
    render_goal_time_analysis(df_all, "Alle Teams – Tore nach Spielminute")

    st.divider()
    st.markdown("### Assists / letzter Pass")
    ga1, ga2 = st.columns(2, gap="small")
    with ga1:
        render_assist_analysis(
            goals_all,
            "Alle Teams – Assists eigene Tore",
            full_size=True,
            always_show=True
        )
    with ga2:
        render_assist_analysis(
            against_all,
            "Alle Teams – Assists Gegentore",
            mirror=True,
            full_size=True,
            always_show=True
        )

    st.markdown("### Vergleich nach Mannschaft")
    team_rows = []
    for summary_team in TEAMS:
        ev = get_events(summary_team)
        tdf = pd.DataFrame(ev) if ev else pd.DataFrame()
        if tdf.empty:
            team_goals = team_against = assists_for = assists_against = one = two = more = 0
        else:
            team_goals = int((tdf["event_type"] == "Tor").sum())
            team_against = int((tdf["event_type"] == "Gegentor").sum())
            tgoals = tdf[tdf["event_type"] == "Tor"]
            tagainst = tdf[tdf["event_type"] == "Gegentor"]
            assists_for = int(tgoals["assist_x"].notna().sum()) if "assist_x" in tgoals.columns else 0
            assists_against = int(tagainst["assist_x"].notna().sum()) if "assist_x" in tagainst.columns else 0
            one = int((tgoals.get("finish_touch", pd.Series(dtype=str)) == "One Touch").sum())
            two = int((tgoals.get("finish_touch", pd.Series(dtype=str)) == "Two Touch").sum())
            more = int((tgoals.get("finish_touch", pd.Series(dtype=str)) == ">2 Touches").sum())

        team_rows.append({
            "Team": summary_team,
            "Tore": team_goals,
            "Gegentore": team_against,
            "Differenz": team_goals - team_against,
            "Assists eigene Tore": assists_for,
            "Assists Gegentore": assists_against,
            "One Touch": one,
            "Two Touch": two,
            ">2 Touches": more,
        })

    st.dataframe(pd.DataFrame(team_rows), use_container_width=True, hide_index=True)

else:
    render_dashboard_header()

    all_events = get_events()
    team_events = get_events(team)

    if not all_events:
        st.info("Noch keine Tore/Gegentore erfasst.")
        st.stop()

    df_all = pd.DataFrame(all_events)
    df_team = pd.DataFrame(team_events) if team_events else pd.DataFrame(columns=df_all.columns)

    goals_team = df_team[df_team["event_type"] == "Tor"].copy() if not df_team.empty else pd.DataFrame(columns=df_all.columns)
    against_team = df_team[df_team["event_type"] == "Gegentor"].copy() if not df_team.empty else pd.DataFrame(columns=df_all.columns)
    against_all = df_all[df_all["event_type"] == "Gegentor"].copy()
    goals_all = df_all[df_all["event_type"] == "Tor"].copy()

    st.caption("🔴 Abschluss · ⚪ Assist / letzter Pass · Linie = zugehörige Szene")
    c1, c2 = st.columns(2, gap="small")

    with c1:
        st.markdown(f'<div class="aka-section-title">{team} - Eigene Tore</div>', unsafe_allow_html=True)
        fig1 = create_dashboard_figure(
            goals_team,
            f"{team} - Eigene Tore",
            point_color="#4cff42",
            connect_assists=True,
            show_assist_points=True,
            finish_marker_color="#dc2828"
        )
        st.pyplot(fig1, use_container_width=True)
        plt.close(fig1)

    with c2:
        st.markdown(f'<div class="aka-section-title">{team} - Gegentore</div>', unsafe_allow_html=True)
        fig2 = create_dashboard_figure(
            against_team,
            f"{team} - Gegentore",
            point_color="#ff5656",
            mirror=True,
            connect_assists=True,
            show_assist_points=True
        )
        st.pyplot(fig2, use_container_width=True)
        plt.close(fig2)

    # Compact stats beneath the two main charts
    s1, s2, s3, s4 = st.columns(4)
    s1.metric(f"{team} Tore", len(goals_team))
    s2.metric(f"{team} Gegentore", len(against_team))
    s3.metric(f"{team} Tore gesamt", len(goals_team))
    s4.metric(f"{team} Gegentore gesamt", len(against_team))

    st.divider()
    render_touch_analysis(goals_team, f"{team} – Torabschluss nach Kontakten")

    st.divider()
    render_goal_time_analysis(df_team, f"{team} – Tore nach Spielminute")

    st.divider()
    st.markdown("### Assists / letzter Pass")
    ta1, ta2 = st.columns(2, gap="small")
    with ta1:
        render_assist_analysis(
            goals_team,
            f"{team} – Assists eigene Tore",
            full_size=True,
            always_show=True
        )
    with ta2:
        render_assist_analysis(
            against_team,
            f"{team} – Assists Gegentore",
            mirror=True,
            full_size=True,
            always_show=True
        )

    with st.expander("Weitere Auswertungen", expanded=False):
        d1, d2 = st.columns(2)
        with d1:
            st.markdown(f"#### {team} - Gegentore")
            fig3 = create_dashboard_figure(
                against_team,
                f"{team} - Gegentore",
                point_color="#ff7474",
                mirror=True,
                connect_assists=True,
                show_assist_points=True
            )
            st.pyplot(fig3, use_container_width=True)
            plt.close(fig3)

        with d2:
            st.markdown(f"#### {team} - Eigene Tore")
            fig4 = create_dashboard_figure(
                goals_team,
                f"{team} - Eigene Tore",
                point_color="#5cff67",
                connect_assists=True,
                show_assist_points=True
            )
            st.pyplot(fig4, use_container_width=True)
            plt.close(fig4)

    st.markdown("### Ereignisse verwalten")
    st.caption("Tore und Gegentore können hier weiterhin gelöscht werden.")

    match_lookup = {m["id"]: m for m in get_matches()}

    for event in reversed(team_events):
        match_info = match_lookup.get(event["match_id"], {})
        opponent = match_info.get("opponent", "Unbekannt")
        match_date = match_info.get("match_date", "")
        event_icon = "⚽" if event["event_type"] == "Tor" else "🥅"

        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([1.0, 2.4, 3.3, 1.2])

            with c1:
                st.markdown(f"**{event_icon} {event['event_type']}**")
                st.caption(f"{event.get('minute') or '-'}'")

            with c2:
                st.markdown(f"**{match_date} · {opponent}**")
                st.write(event.get("scorer") or "–")
                if event.get("assister"):
                    st.caption(f"Assist: {event['assister']}")

            with c3:
                st.write(event.get("finish_zone") or "Keine Abschlusszone")
                details = " · ".join(
                    x for x in [
                        event.get("phase"),
                        event.get("creation_type"),
                        event.get("finish_touch")
                    ] if x
                )
                if details:
                    st.caption(details)
                if event.get("comment"):
                    st.caption(event["comment"])
                if event.get("video_url"):
                    st.link_button("Video öffnen", event["video_url"], key=f"video_{event['id']}")

            with c4:
                if st.button(
                    "👁️ Ansehen / Bearbeiten",
                    key=f"edit_{event['id']}",
                    use_container_width=True,
                ):
                    st.session_state["editing_event_id"] = event["id"]
                    st.session_state.pop(f"edit_points_{event['id']}", None)
                    st.session_state.pop(f"edit_nonce_{event['id']}", None)
                    st.rerun()

                confirm_key = f"confirm_delete_{event['id']}"
                if not st.session_state.get(confirm_key, False):
                    if st.button("🗑️ Löschen", key=f"delete_{event['id']}", use_container_width=True):
                        st.session_state[confirm_key] = True
                        st.rerun()
                else:
                    st.warning("Wirklich?")
                    if st.button("Ja", key=f"yes_delete_{event['id']}", type="primary", use_container_width=True):
                        delete_event(event["id"])
                        st.session_state.pop(confirm_key, None)
                        if st.session_state.get("editing_event_id") == event["id"]:
                            st.session_state.pop("editing_event_id", None)
                        st.rerun()
                    if st.button("Nein", key=f"cancel_delete_{event['id']}", use_container_width=True):
                        st.session_state.pop(confirm_key, None)
                        st.rerun()

    editing_id = st.session_state.get("editing_event_id")
    if editing_id is not None:
        current_event = next(
            (e for e in team_events if e["id"] == editing_id),
            None
        )
        if current_event:
            render_event_editor(current_event)
        else:
            st.session_state.pop("editing_event_id", None)
