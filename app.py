"""
Study Performance Analytics System
A Streamlit app to track study sessions, analyze productivity,
detect weak subjects, and visualize weekly performance.
"""

import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
from datetime import datetime, date, timedelta
import json
import os
import pathlib
import calendar
import streamlit.components.v1 as components

# Google Sheets — imported lazily so the app works even without credentials
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

# ─────────────────────────────────────────────
# File-based persistence — saved on the server,
# no download needed on your computer.
# ─────────────────────────────────────────────
DATA_FILE = pathlib.Path(__file__).parent / "study_data.json"

def load_data() -> dict:
    """Load saved study data from disk. Returns defaults if file doesn't exist."""
    if DATA_FILE.exists():
        try:
            with open(DATA_FILE, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {"study_log": [], "streak": 0, "study_goal": 8.0}

def save_data():
    """
    Persist session data to disk when possible.
    On Streamlit Community Cloud the filesystem is ephemeral, so this is
    best-effort only — Google Sheets is the durable store there.
    """
    payload = {
        "study_log": st.session_state.study_log,
        "streak": st.session_state.streak,
        "study_goal": st.session_state.study_goal,
        "sheet_url": st.session_state.get("sheet_url", ""),
    }
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(payload, f, indent=2)
    except OSError:
        pass  # read-only filesystem (e.g. Streamlit Community Cloud)

# ─────────────────────────────────────────────
# Google Sheets integration
# Credentials come from the GOOGLE_SERVICE_ACCOUNT_KEY environment
# variable (a JSON string). The user also pastes their Sheet URL once.
# ─────────────────────────────────────────────
SHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]
SHEET_TAB_NAME = "StudySessions"
SHEET_HEADERS  = ["Date", "Total Hours", "Productivity (%)", "Goal Achieved", "Subjects & Hours"]


def _get_sa_key() -> dict | None:
    """
    Read service account credentials from env or st.secrets.

    Handles three formats:
      1. Env var  GOOGLE_SERVICE_ACCOUNT_KEY = '{"type":"service_account",...}'
      2. SCC TOML string   [GOOGLE_SERVICE_ACCOUNT_KEY] as a quoted JSON string
      3. SCC TOML section  [GOOGLE_SERVICE_ACCOUNT_KEY] as nested key/value pairs
    """
    # 1 — plain environment variable (Replit Secrets, Docker, etc.)
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_KEY", "")
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return None

    # 2 & 3 — Streamlit Community Cloud secrets
    try:
        secret = st.secrets["GOOGLE_SERVICE_ACCOUNT_KEY"]
        # Case 2: it's already a string (user pasted the JSON as one value)
        if isinstance(secret, str):
            return json.loads(secret)
        # Case 3: it's a TOML section (AttrDict / dict-like)
        return dict(secret)
    except Exception:
        return None


@st.cache_resource(show_spinner=False)
def get_gsheets_client():
    """Return an authenticated gspread client, or None if credentials missing."""
    if not GSHEETS_AVAILABLE:
        return None
    key = _get_sa_key()
    if not key:
        return None
    try:
        creds = Credentials.from_service_account_info(key, scopes=SHEETS_SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None


def _extract_sheet_id(url_or_id: str) -> str:
    """Pull the spreadsheet ID out of a full Google Sheets URL or return as-is."""
    if "/spreadsheets/d/" in url_or_id:
        part = url_or_id.split("/spreadsheets/d/")[1]
        return part.split("/")[0]
    return url_or_id.strip()


def _get_worksheet(client, sheet_id: str):
    """Open the spreadsheet and return (or create) the StudySessions worksheet."""
    try:
        ss = client.open_by_key(sheet_id)
    except gspread.exceptions.SpreadsheetNotFound:
        return None, "Spreadsheet not found. Make sure you shared it with the service account email."
    except Exception as e:
        return None, str(e)

    # Get or create the tab
    try:
        ws = ss.worksheet(SHEET_TAB_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=SHEET_TAB_NAME, rows=1000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)
    return ws, None


def sync_session_to_sheets(entry: dict) -> tuple[bool, str]:
    """
    Write (or overwrite) a session row in the connected Google Sheet.
    Each date occupies exactly one row. If a row for that date exists it is
    replaced; otherwise a new row is appended.
    Returns (success, message).
    """
    client = get_gsheets_client()
    if client is None:
        return False, "No Google Sheets credentials configured."

    sheet_url = st.session_state.get("sheet_url", "")
    if not sheet_url:
        return False, "No Google Sheets URL provided."

    sheet_id = _extract_sheet_id(sheet_url)
    ws, err = _get_worksheet(client, sheet_id)
    if err:
        return False, err

    # Build the row
    subjects_str = " | ".join(f"{s}: {h}h" for s, h in entry["subject_hours"].items())
    new_row = [
        entry["date"],
        entry["total_hours"],
        entry["productivity"],
        "Yes" if entry["goal_achieved"] else "No",
        subjects_str,
    ]

    # Find existing row for this date and replace, or append
    try:
        cell = ws.find(entry["date"], in_column=1)
        ws.update(f"A{cell.row}:E{cell.row}", [new_row])
    except gspread.exceptions.CellNotFound:
        ws.append_row(new_row)
    except Exception as e:
        return False, str(e)

    return True, "Synced to Google Sheets ✅"


def load_from_sheets() -> list:
    """
    Load all sessions from Google Sheets and merge with local data.
    Sheets data wins for dates that exist in both (cloud is source of truth).
    Returns the merged list sorted by date, or empty list on failure.
    """
    client = get_gsheets_client()
    if client is None:
        return []

    sheet_url = st.session_state.get("sheet_url", "")
    if not sheet_url:
        return []

    sheet_id = _extract_sheet_id(sheet_url)
    ws, err = _get_worksheet(client, sheet_id)
    if err or ws is None:
        return []

    try:
        rows = ws.get_all_records()
    except Exception:
        return []

    sessions = []
    for row in rows:
        try:
            # Reconstruct subject_hours from the "Subjects & Hours" column
            subject_hours = {}
            for part in str(row.get("Subjects & Hours", "")).split(" | "):
                if ":" in part:
                    subj, hrs_str = part.rsplit(":", 1)
                    subject_hours[subj.strip()] = float(hrs_str.replace("h", "").strip())
            sessions.append({
                "date": str(row["Date"]),
                "total_hours": float(row["Total Hours"]),
                "productivity": float(row["Productivity (%)"]),
                "goal_achieved": str(row["Goal Achieved"]).strip().lower() == "yes",
                "subject_hours": subject_hours,
            })
        except (KeyError, ValueError):
            continue

    return sorted(sessions, key=lambda e: e["date"])

# ─────────────────────────────────────────────
# Page configuration
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Study Performance Analytics",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# Session-state initialisation
# All persistent data lives here so it survives reruns.
# ─────────────────────────────────────────────
def init_state():
    # Load from disk on the very first run of this browser session.
    # Subsequent reruns skip this because the keys are already in session_state.
    if "study_log" not in st.session_state:
        saved = load_data()
        st.session_state.study_log  = saved.get("study_log", [])
        st.session_state.streak     = saved.get("streak", 0)
        st.session_state.study_goal = saved.get("study_goal", 8.0)
        # Restore Google Sheet URL if saved locally
        st.session_state.sheet_url  = saved.get("sheet_url", "")

    # If a sheet URL is configured, merge cloud data on first load
    if "sheets_loaded" not in st.session_state:
        st.session_state.sheets_loaded = False
    if not st.session_state.sheets_loaded and st.session_state.get("sheet_url"):
        cloud_sessions = load_from_sheets()
        if cloud_sessions:
            # Merge: cloud wins for any date that appears in both
            local_dates = {e["date"] for e in st.session_state.study_log}
            cloud_dates = {e["date"] for e in cloud_sessions}
            merged = [e for e in st.session_state.study_log if e["date"] not in cloud_dates]
            merged.extend(cloud_sessions)
            merged.sort(key=lambda e: e["date"])
            st.session_state.study_log = merged
            # Recalculate streak from merged log
            st.session_state.streak = sum(
                1 for e in reversed(merged) if e["goal_achieved"]
            )
        st.session_state.sheets_loaded = True

init_state()

# ─────────────────────────────────────────────
# Core calculation functions
# ─────────────────────────────────────────────

def calculate_total_hours(subject_hours: dict) -> float:
    """Sum all hours studied across subjects for a session."""
    return sum(subject_hours.values())


def calculate_productivity(total_hours: float, goal: float) -> float:
    """
    Productivity score = (hours studied / daily goal) * 100
    Capped at 100 % for display purposes but can exceed in raw value.
    """
    if goal <= 0:
        return 0.0
    return round((total_hours / goal) * 100, 1)


def goal_achieved(total_hours: float, goal: float) -> bool:
    """Return True if the student met or exceeded the daily study goal."""
    return total_hours >= goal


def detect_weak_subjects(subject_hours: dict, threshold_pct: float = 15.0) -> list:
    """
    Identify subjects whose share of today's study time is below threshold_pct %.
    threshold_pct defaults to 15 % of total study time.
    """
    total = sum(subject_hours.values())
    if total == 0:
        return list(subject_hours.keys())
    weak = [
        subj
        for subj, hrs in subject_hours.items()
        if (hrs / total * 100) < threshold_pct
    ]
    return weak


def generate_recommendations(
    weak_subjects: list,
    productivity: float,
    goal: float,
    total_hours: float,
) -> list[str]:
    """
    Build a short list of personalised study recommendations based on
    today's performance. Returns plain-text strings.
    """
    recs = []

    if productivity < 50:
        recs.append(
            f"Your productivity is {productivity}%. Try to study at least "
            f"{goal / 2:.1f} hours tomorrow to build momentum."
        )
    elif productivity < 100:
        deficit = round(goal - total_hours, 1)
        recs.append(
            f"You are {deficit} hour(s) short of your daily goal. "
            "A focused evening session could close the gap."
        )
    else:
        recs.append("Great job! You met your daily goal. Keep the streak alive tomorrow.")

    if weak_subjects:
        recs.append(
            f"Give more attention to: **{', '.join(weak_subjects)}**. "
            "Try dedicating at least 20 % of your study time to each subject."
        )

    if total_hours > goal * 1.3:
        recs.append(
            "You studied significantly more than your goal today. "
            "Remember to rest — recovery improves long-term retention."
        )

    if not recs:
        recs.append("Everything looks balanced. Stay consistent!")

    return recs


def update_streak(achieved: bool):
    """Increment or reset the streak counter based on goal achievement."""
    if achieved:
        st.session_state.streak += 1
    else:
        st.session_state.streak = 0


def get_weekly_data() -> pd.DataFrame:
    """
    Aggregate the study log into a DataFrame suitable for weekly bar charts.
    Returns columns: date, subject, hours.
    """
    rows = []
    # Only look at the last 7 logged days
    recent = st.session_state.study_log[-7:]
    for entry in recent:
        for subj, hrs in entry["subject_hours"].items():
            rows.append({"date": entry["date"], "subject": subj, "hours": hrs})
    if not rows:
        return pd.DataFrame(columns=["date", "subject", "hours"])
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────
# Chart builders
# ─────────────────────────────────────────────

def productivity_pie_chart(subject_hours: dict):
    """Donut chart showing distribution of study hours per subject."""
    labels = list(subject_hours.keys())
    values = list(subject_hours.values())
    fig = go.Figure(
        go.Pie(
            labels=labels,
            values=values,
            hole=0.45,
            textinfo="label+percent",
            hovertemplate="%{label}: %{value} hr(s)<extra></extra>",
        )
    )
    fig.update_layout(
        title_text="Study Time Distribution",
        showlegend=True,
        margin=dict(t=60, b=20, l=20, r=20),
        height=380,
    )
    return fig


def weekly_bar_chart(df: pd.DataFrame):
    """Grouped bar chart of subject hours per day for the last 7 sessions."""
    fig = px.bar(
        df,
        x="date",
        y="hours",
        color="subject",
        barmode="group",
        labels={"hours": "Hours Studied", "date": "Date", "subject": "Subject"},
        title="Weekly Subject-Wise Study Hours",
        height=420,
    )
    fig.update_layout(
        xaxis_tickangle=-30,
        legend_title_text="Subject",
        margin=dict(t=60, b=60, l=40, r=20),
    )
    return fig


def build_streak_calendar_html(study_log: list) -> str:
    """
    Generate a Duolingo-style streak calendar as an HTML string.
    Green glow + 🔥 for goal-achieved days, 3D red flag for missed days.
    Rendered via st.components.v1.html so no download is required.
    """
    # Build a quick lookup: "YYYY-MM-DD" -> "achieved" or "missed"
    log_map = {
        e["date"]: ("achieved" if e["goal_achieved"] else "missed")
        for e in study_log
    }
    log_map_json = json.dumps(log_map)

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: transparent;
    font-family: 'Segoe UI', system-ui, sans-serif;
    color: #f9fafb;
  }}
  .cal-wrapper {{
    background: linear-gradient(145deg, #0f172a, #1e293b);
    border-radius: 24px;
    padding: 28px 24px 24px;
    max-width: 640px;
    margin: 0 auto;
    box-shadow: 0 20px 60px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.06);
    border: 1px solid rgba(255,255,255,0.07);
  }}
  .cal-header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 22px;
  }}
  .cal-nav {{
    background: rgba(255,255,255,0.07);
    border: 1px solid rgba(255,255,255,0.1);
    color: #f9fafb;
    border-radius: 10px;
    padding: 7px 16px;
    cursor: pointer;
    font-size: 18px;
    line-height: 1;
    transition: background 0.15s, transform 0.1s;
  }}
  .cal-nav:hover {{ background: rgba(255,255,255,0.14); transform: scale(1.05); }}
  .cal-nav:active {{ transform: scale(0.97); }}
  .cal-month {{
    font-size: 20px;
    font-weight: 800;
    letter-spacing: 0.3px;
    color: #f1f5f9;
  }}
  .cal-grid {{
    display: grid;
    grid-template-columns: repeat(7, 1fr);
    gap: 7px;
  }}
  .cal-weekday {{
    text-align: center;
    font-size: 10px;
    font-weight: 700;
    color: #4b5563;
    text-transform: uppercase;
    padding: 2px 0 12px;
    letter-spacing: 1px;
  }}
  .cal-day {{
    aspect-ratio: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    border-radius: 14px;
    background: rgba(255,255,255,0.04);
    font-size: 13px;
    font-weight: 600;
    color: #9ca3af;
    position: relative;
    cursor: default;
    transition: transform 0.15s;
    border: 1px solid rgba(255,255,255,0.04);
  }}
  .cal-day.empty {{
    background: transparent;
    border: none;
  }}
  .cal-day.future {{
    opacity: 0.28;
  }}
  .cal-day.today {{
    border: 2px solid #facc15;
    color: #fde047;
    font-weight: 800;
    background: rgba(250,204,21,0.08);
  }}
  /* Goal achieved — green glow */
  .cal-day.achieved {{
    background: linear-gradient(145deg, #14532d, #166534);
    color: #4ade80;
    font-weight: 800;
    box-shadow:
      0 0 0 2px #22c55e,
      0 0 14px rgba(34,197,94,0.55),
      0 0 28px rgba(34,197,94,0.25);
    border: none;
  }}
  .cal-day.achieved.today {{
    box-shadow:
      0 0 0 2px #facc15,
      0 0 14px rgba(250,204,21,0.5),
      0 0 28px rgba(34,197,94,0.3);
  }}
  .day-num {{ line-height: 1.1; }}
  .flame {{ font-size: 11px; line-height: 1; margin-top: 1px; }}

  /* 3-D red flag */
  .flag-wrap {{
    position: relative;
    width: 14px;
    height: 15px;
    margin-top: 3px;
    flex-shrink: 0;
  }}
  .flag-pole {{
    position: absolute;
    left: 2px;
    top: 0;
    width: 2px;
    height: 100%;
    background: linear-gradient(to right, #7f1d1d, #ef4444, #7f1d1d);
    border-radius: 1px;
    box-shadow: 1px 1px 2px rgba(0,0,0,0.6);
  }}
  .flag-fabric {{
    position: absolute;
    left: 4px;
    top: 1px;
    width: 10px;
    height: 7px;
    background: linear-gradient(135deg, #ef4444 0%, #dc2626 60%, #991b1b 100%);
    border-radius: 0 3px 2px 0;
    box-shadow:
      1px 2px 4px rgba(0,0,0,0.55),
      inset 0 1px 1px rgba(255,255,255,0.25),
      inset 0 -1px 1px rgba(0,0,0,0.3);
    transform: perspective(24px) rotateY(-8deg) rotateX(4deg);
    clip-path: polygon(0 0, 100% 15%, 100% 85%, 0 100%);
  }}

  .legend {{
    display: flex;
    gap: 18px;
    justify-content: center;
    margin-top: 20px;
    flex-wrap: wrap;
  }}
  .legend-item {{
    display: flex;
    align-items: center;
    gap: 7px;
    font-size: 11px;
    color: #6b7280;
    font-weight: 500;
  }}
  .legend-dot {{
    width: 14px;
    height: 14px;
    border-radius: 5px;
    flex-shrink: 0;
  }}
  .legend-dot.g {{
    background: #166534;
    box-shadow: 0 0 6px rgba(34,197,94,0.5);
    border: 1px solid #22c55e;
  }}
  .legend-dot.r {{ background: #1f2937; border: 1px solid #374151; position:relative; overflow:hidden; }}
  .legend-dot.t {{
    background: rgba(250,204,21,0.08);
    border: 2px solid #facc15;
  }}
  .legend-dot.e {{ background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08); }}
</style>
</head>
<body>
<div class="cal-wrapper">
  <div class="cal-header">
    <button class="cal-nav" onclick="changeMonth(-1)">&#8249;</button>
    <div class="cal-month" id="cal-title"></div>
    <button class="cal-nav" onclick="changeMonth(1)">&#8250;</button>
  </div>
  <div class="cal-grid" id="cal-grid"></div>
  <div class="legend">
    <div class="legend-item"><div class="legend-dot g"></div> Goal achieved</div>
    <div class="legend-item">
      <div style="position:relative;width:14px;height:14px;border-radius:5px;background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);">
        <div style="position:absolute;left:3px;top:1px;width:2px;height:11px;background:#ef4444;border-radius:1px;"></div>
        <div style="position:absolute;left:5px;top:1px;width:7px;height:5px;background:#dc2626;border-radius:0 2px 2px 0;clip-path:polygon(0 0,100% 15%,100% 85%,0 100%);"></div>
      </div>
      Goal missed
    </div>
    <div class="legend-item"><div class="legend-dot t"></div> Today</div>
    <div class="legend-item"><div class="legend-dot e"></div> No data</div>
  </div>
</div>

<script>
const LOG = {log_map_json};
const WEEKDAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
const MONTHS   = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];

const now = new Date();
let viewYear  = now.getFullYear();
let viewMonth = now.getMonth();

function pad(n) {{ return String(n).padStart(2,'0'); }}

function toDateStr(y, m, d) {{
  return y + '-' + pad(m+1) + '-' + pad(d);
}}

function changeMonth(delta) {{
  viewMonth += delta;
  if (viewMonth < 0)  {{ viewMonth = 11; viewYear--; }}
  if (viewMonth > 11) {{ viewMonth = 0;  viewYear++; }}
  render();
}}

function render() {{
  document.getElementById('cal-title').textContent =
    MONTHS[viewMonth] + ' ' + viewYear;

  const grid = document.getElementById('cal-grid');
  grid.innerHTML = '';

  // Weekday headers
  WEEKDAYS.forEach(d => {{
    const el = document.createElement('div');
    el.className = 'cal-weekday';
    el.textContent = d;
    grid.appendChild(el);
  }});

  const todayStr  = toDateStr(now.getFullYear(), now.getMonth(), now.getDate());
  const firstDay  = new Date(viewYear, viewMonth, 1).getDay();
  const offset    = firstDay === 0 ? 6 : firstDay - 1; // shift to Mon=0
  const daysInMo  = new Date(viewYear, viewMonth + 1, 0).getDate();

  // Empty leading cells
  for (let i = 0; i < offset; i++) {{
    const el = document.createElement('div');
    el.className = 'cal-day empty';
    grid.appendChild(el);
  }}

  // Day cells
  for (let d = 1; d <= daysInMo; d++) {{
    const dateStr = toDateStr(viewYear, viewMonth, d);
    const status  = LOG[dateStr]; // 'achieved' | 'missed' | undefined
    const isToday = dateStr === todayStr;
    const isFuture = new Date(viewYear, viewMonth, d) > now;

    const el = document.createElement('div');
    const classes = ['cal-day'];
    if (isFuture)           classes.push('future');
    if (isToday)            classes.push('today');
    if (status === 'achieved') classes.push('achieved');
    if (status === 'missed')   classes.push('missed');
    el.className = classes.join(' ');

    // Date number
    const num = document.createElement('div');
    num.className = 'day-num';
    num.textContent = d;
    el.appendChild(num);

    // Status icon
    if (status === 'achieved') {{
      const flame = document.createElement('div');
      flame.className = 'flame';
      flame.textContent = '🔥';
      el.appendChild(flame);
    }} else if (status === 'missed') {{
      const fw = document.createElement('div');
      fw.className = 'flag-wrap';
      fw.innerHTML =
        '<div class="flag-pole"></div>' +
        '<div class="flag-fabric"></div>';
      el.appendChild(fw);
    }}

    grid.appendChild(el);
  }}
}}

render();
</script>
</body>
</html>"""

# ─────────────────────────────────────────────
# Sidebar — global settings
# ─────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Settings")
    st.session_state.study_goal = st.number_input(
        "Daily Study Goal (hours)",
        min_value=0.5,
        max_value=24.0,
        value=float(st.session_state.study_goal),
        step=0.5,
        help="Set the number of hours you aim to study each day.",
    )
    st.divider()

    # ── Google Sheets Sync ──────────────────────
    st.subheader("📋 Google Sheets Sync")
    client = get_gsheets_client()
    if client is None:
        st.caption(
            "To enable Google Sheets sync, add your service account key as "
            "the **GOOGLE_SERVICE_ACCOUNT_KEY** secret. "
            "See the setup guide below the main app."
        )
    else:
        st.success("Google credentials connected ✅")
        sheet_url_input = st.text_input(
            "Paste your Google Sheet URL",
            value=st.session_state.get("sheet_url", ""),
            placeholder="https://docs.google.com/spreadsheets/d/...",
            help="Create a sheet, share it with the service account email, then paste the URL here.",
        )
        if sheet_url_input != st.session_state.get("sheet_url", ""):
            st.session_state.sheet_url = sheet_url_input
            st.session_state.sheets_loaded = False  # trigger re-load
            save_data()
            st.rerun()

        if st.session_state.get("sheet_url"):
            if st.button("☁️ Sync all sessions to Sheet", use_container_width=True):
                with st.spinner("Syncing…"):
                    errors = []
                    for entry in st.session_state.study_log:
                        ok, msg = sync_session_to_sheets(entry)
                        if not ok:
                            errors.append(msg)
                if errors:
                    st.error(errors[0])
                else:
                    st.success(f"Synced {len(st.session_state.study_log)} session(s) ✅")

    st.divider()
    st.subheader("📊 Session History")
    if st.session_state.study_log:
        for entry in reversed(st.session_state.study_log[-5:]):
            icon = "✅" if entry["goal_achieved"] else "❌"
            st.caption(f"{icon} {entry['date']} — {entry['total_hours']} hr(s)")
    else:
        st.caption("No sessions logged yet.")

    st.divider()
    if st.button("🗑️ Clear All Data", use_container_width=True):
        st.session_state.study_log = []
        st.session_state.streak = 0
        save_data()
        st.rerun()

# ─────────────────────────────────────────────
# Main layout — tab navigation
# ─────────────────────────────────────────────
st.title("📚 Study Performance Analytics System")
st.caption("Track your study sessions, measure productivity, and stay consistent.")

tab_input, tab_dashboard, tab_weak, tab_weekly, tab_streak = st.tabs([
    "📝 Daily Input",
    "📈 Productivity Dashboard",
    "⚠️ Weak Subject Analysis",
    "📅 Weekly Reports",
    "🔥 Streak Tracker",
])

# ══════════════════════════════════════════════
# TAB 1 — Daily Input
# ══════════════════════════════════════════════
with tab_input:
    st.header("Log Today's Study Session")

    col_date, col_goal = st.columns([2, 1])
    with col_date:
        selected_date = st.date_input("Session Date", value=date.today())
    with col_goal:
        st.metric("Daily Goal", f"{st.session_state.study_goal} hrs")

    st.divider()
    st.subheader("Subjects & Hours Studied")

    # Dynamic subject rows
    if "num_subjects" not in st.session_state:
        st.session_state.num_subjects = 3

    col_add, col_remove, _ = st.columns([1, 1, 4])
    with col_add:
        if st.button("➕ Add subject"):
            st.session_state.num_subjects += 1
            st.rerun()
    with col_remove:
        if st.button("➖ Remove subject") and st.session_state.num_subjects > 1:
            st.session_state.num_subjects -= 1
            st.rerun()

    subject_hours: dict[str, float] = {}
    cols_per_row = 2

    for i in range(st.session_state.num_subjects):
        if i % cols_per_row == 0:
            row_cols = st.columns(cols_per_row)
        with row_cols[i % cols_per_row]:
            subj_name = st.text_input(
                f"Subject {i + 1}",
                value=f"Subject {i + 1}",
                key=f"subj_name_{i}",
                placeholder="e.g. Mathematics",
            )
            hours_val = st.number_input(
                f"Hours for {subj_name}",
                min_value=0.0,
                max_value=24.0,
                value=0.0,
                step=0.25,
                key=f"subj_hrs_{i}",
            )
            if subj_name.strip() and hours_val > 0:
                subject_hours[subj_name.strip()] = hours_val

    st.divider()

    if st.button("✅ Log Session", type="primary", use_container_width=True):
        if not subject_hours:
            st.error("Please enter at least one subject with hours > 0.")
        else:
            total = calculate_total_hours(subject_hours)
            productivity = calculate_productivity(total, st.session_state.study_goal)
            achieved = goal_achieved(total, st.session_state.study_goal)
            update_streak(achieved)

            entry = {
                "date": str(selected_date),
                "subject_hours": subject_hours,
                "total_hours": total,
                "productivity": productivity,
                "goal_achieved": achieved,
            }
            # Prevent duplicate dates — replace existing entry for same date
            st.session_state.study_log = [
                e for e in st.session_state.study_log if e["date"] != str(selected_date)
            ]
            st.session_state.study_log.append(entry)
            # Keep log sorted by date
            st.session_state.study_log.sort(key=lambda e: e["date"])
            # Persist to disk so data survives restarts
            save_data()
            # Auto-sync to Google Sheets if connected
            sheets_msg = ""
            if st.session_state.get("sheet_url") and get_gsheets_client():
                ok, sheets_msg = sync_session_to_sheets(entry)

            if achieved:
                st.success(
                    f"Session logged! You studied **{total:.1f} hrs** "
                    f"(Productivity: **{productivity}%**). Goal achieved! 🎉"
                )
            else:
                deficit = round(st.session_state.study_goal - total, 1)
                st.warning(
                    f"Session logged! You studied **{total:.1f} hrs** "
                    f"(Productivity: **{productivity}%**). "
                    f"You were **{deficit} hr(s)** short of your goal."
                )
            if sheets_msg:
                st.caption(sheets_msg)

# ══════════════════════════════════════════════
# TAB 2 — Productivity Dashboard
# ══════════════════════════════════════════════
with tab_dashboard:
    st.header("Productivity Dashboard")

    if not st.session_state.study_log:
        st.info("No sessions logged yet. Head to **Daily Input** to record your first session.")
    else:
        # Show metrics for the most recent session
        latest = st.session_state.study_log[-1]
        productivity = latest["productivity"]
        total_hours = latest["total_hours"]
        achieved = latest["goal_achieved"]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Latest Session Date", latest["date"])
        m2.metric("Hours Studied", f"{total_hours:.1f} hr(s)")
        m3.metric("Productivity Score", f"{productivity}%")
        m4.metric("Daily Goal Met", "Yes ✅" if achieved else "No ❌")

        st.divider()
        st.subheader("Study Time Distribution")

        subject_hours = latest["subject_hours"]
        if subject_hours:
            col_pie, col_recs = st.columns([1.5, 1])
            with col_pie:
                fig = productivity_pie_chart(subject_hours)
                st.plotly_chart(fig, use_container_width=True)
            with col_recs:
                st.subheader("💡 Recommendations")
                weak = detect_weak_subjects(subject_hours)
                recs = generate_recommendations(
                    weak,
                    productivity,
                    st.session_state.study_goal,
                    total_hours,
                )
                for rec in recs:
                    st.info(rec)

        # Aggregate stats across all logged sessions
        st.divider()
        st.subheader("All-Time Summary")
        all_totals = [e["total_hours"] for e in st.session_state.study_log]
        all_prod = [e["productivity"] for e in st.session_state.study_log]
        s1, s2, s3 = st.columns(3)
        s1.metric("Total Sessions", len(st.session_state.study_log))
        s2.metric("Avg. Hours / Session", f"{sum(all_totals)/len(all_totals):.1f}")
        s3.metric("Avg. Productivity", f"{sum(all_prod)/len(all_prod):.1f}%")

# ══════════════════════════════════════════════
# TAB 3 — Weak Subject Analysis
# ══════════════════════════════════════════════
with tab_weak:
    st.header("Weak Subject Analysis")

    if not st.session_state.study_log:
        st.info("Log at least one session to see weak subject analysis.")
    else:
        latest = st.session_state.study_log[-1]
        subject_hours = latest["subject_hours"]
        total = latest["total_hours"]

        st.caption(f"Analysis based on your latest session: **{latest['date']}**")

        # Build subject-level table
        rows = []
        for subj, hrs in subject_hours.items():
            pct = round(hrs / total * 100, 1) if total > 0 else 0
            rows.append({"Subject": subj, "Hours": hrs, "Share (%)": pct})
        df_subj = pd.DataFrame(rows).sort_values("Share (%)", ascending=True)

        weak_subjects = detect_weak_subjects(subject_hours)

        col_table, col_bar = st.columns([1, 1.5])
        with col_table:
            st.subheader("Subject Breakdown")
            # Colour weak subjects red in the table
            def highlight_weak(row):
                if row["Subject"] in weak_subjects:
                    return ["background-color: #fee2e2"] * len(row)
                return [""] * len(row)

            st.dataframe(
                df_subj.style.apply(highlight_weak, axis=1),
                use_container_width=True,
                hide_index=True,
            )

        with col_bar:
            st.subheader("Hours per Subject")
            colours = [
                "#ef4444" if s in weak_subjects else "#3b82f6"
                for s in df_subj["Subject"]
            ]
            fig_bar = go.Figure(
                go.Bar(
                    x=df_subj["Hours"],
                    y=df_subj["Subject"],
                    orientation="h",
                    marker_color=colours,
                    hovertemplate="%{y}: %{x} hr(s)<extra></extra>",
                )
            )
            fig_bar.update_layout(
                xaxis_title="Hours",
                height=300,
                margin=dict(t=20, b=20, l=10, r=20),
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        if weak_subjects:
            st.error(
                f"**Weak subjects detected** (< 15% of today's study time): "
                f"{', '.join(weak_subjects)}"
            )
            st.subheader("💡 Improvement Tips")
            for subj in weak_subjects:
                target = round(total * 0.2, 1)
                st.markdown(
                    f"- **{subj}**: Aim for at least **{target} hr(s)** "
                    f"(20% of your daily total) in tomorrow's session."
                )
        else:
            st.success("No weak subjects detected. Your study time is well-distributed!")

# ══════════════════════════════════════════════
# TAB 4 — Weekly Reports
# ══════════════════════════════════════════════
with tab_weekly:
    st.header("Weekly Reports")

    if len(st.session_state.study_log) < 2:
        st.info("Log at least 2 sessions to see weekly reports.")
    else:
        df_weekly = get_weekly_data()

        if df_weekly.empty:
            st.info("No data available for the weekly report.")
        else:
            # Subject-wise grouped bar chart
            fig_weekly = weekly_bar_chart(df_weekly)
            st.plotly_chart(fig_weekly, use_container_width=True)

            st.divider()

            # Aggregate table: total hours per subject across the week
            st.subheader("Subject Totals (Last 7 Sessions)")
            df_totals = (
                df_weekly.groupby("subject")["hours"]
                .sum()
                .reset_index()
                .rename(columns={"subject": "Subject", "hours": "Total Hours"})
                .sort_values("Total Hours", ascending=False)
            )
            st.dataframe(df_totals, use_container_width=True, hide_index=True)

            # Daily total comparison against goal
            st.subheader("Daily Totals vs. Goal")
            recent_entries = st.session_state.study_log[-7:]
            dates_list = [e["date"] for e in recent_entries]
            totals_list = [e["total_hours"] for e in recent_entries]
            goal_line = [st.session_state.study_goal] * len(recent_entries)

            fig_goal = go.Figure()
            fig_goal.add_trace(
                go.Bar(
                    x=dates_list,
                    y=totals_list,
                    name="Hours Studied",
                    marker_color="#3b82f6",
                )
            )
            fig_goal.add_trace(
                go.Scatter(
                    x=dates_list,
                    y=goal_line,
                    name="Daily Goal",
                    mode="lines",
                    line=dict(color="#ef4444", width=2, dash="dash"),
                )
            )
            fig_goal.update_layout(
                title="Daily Hours vs. Goal",
                xaxis_title="Date",
                yaxis_title="Hours",
                height=360,
                margin=dict(t=60, b=60, l=40, r=20),
                xaxis_tickangle=-30,
            )
            st.plotly_chart(fig_goal, use_container_width=True)

# ══════════════════════════════════════════════
# TAB 5 — Streak Tracker
# ══════════════════════════════════════════════
with tab_streak:
    st.header("Streak Tracker")

    # ── Top row: streak counter + motivational message ──
    col_streak, col_msg = st.columns([1, 2])
    with col_streak:
        st.metric("🔥 Current Streak", f"{st.session_state.streak} day(s)")

    with col_msg:
        s = st.session_state.streak
        if s == 0:
            st.warning("No active streak. Log a session today to start one!")
        elif s < 3:
            st.info(f"You're on a {s}-day streak — keep going!")
        elif s < 7:
            st.success(f"Awesome! {s}-day streak. Halfway to a week!")
        else:
            st.success(f"🏆 {s}-day streak! Outstanding consistency!")

    st.divider()

    # ── Duolingo-style streak calendar ──
    st.subheader("Streak Calendar")
    cal_html = build_streak_calendar_html(st.session_state.study_log)
    components.html(cal_html, height=420, scrolling=False)

    # ── Recent sessions list below the calendar ──
    if st.session_state.study_log:
        st.divider()
        st.subheader("Recent Sessions")
        recent = st.session_state.study_log[-7:]
        for entry in reversed(recent):
            if entry["goal_achieved"]:
                st.success(
                    f"✅ {entry['date']} — {entry['total_hours']} hr(s) "
                    f"| Productivity: {entry['productivity']}% | Goal met"
                )
            else:
                st.error(
                    f"❌ {entry['date']} — {entry['total_hours']} hr(s) "
                    f"| Productivity: {entry['productivity']}% | Goal missed"
                )

# ─────────────────────────────────────────────
# Google Sheets Setup Guide (shown when no credentials)
# ─────────────────────────────────────────────
if get_gsheets_client() is None:
    st.divider()
    with st.expander("📋 How to connect Google Sheets — deploy to Streamlit Community Cloud", expanded=False):
        st.markdown("""
### Deploy this app + connect Google Sheets in 6 steps

Streamlit Community Cloud is **free** and gives this app a permanent public URL.
Google Sheets becomes your durable database — sessions survive restarts and redeploys.
Setup takes about 15 minutes and only needs to be done once.

---

#### Part A — Push your code to GitHub
1. Create a free account at [github.com](https://github.com) if you don't have one
2. Create a **new repository** (public or private)
3. Push this project to that repo (Replit has a built-in Git panel in the left sidebar)

---

#### Part B — Create Google credentials
**Step 1 — Google Cloud project**
1. Go to [console.cloud.google.com](https://console.cloud.google.com) → **New Project** → any name → **Create**

**Step 2 — Enable APIs**
1. **APIs & Services → Library** → search **Google Sheets API** → **Enable**
2. Same place → search **Google Drive API** → **Enable**

**Step 3 — Create a Service Account**
1. **APIs & Services → Credentials → Create Credentials → Service Account**
2. Give it any name → **Create and Continue** → skip optional steps → **Done**
3. Click the new service account → **Keys → Add Key → Create new key → JSON**
4. A `.json` file downloads — open it and copy the `client_email` value

**Step 4 — Share your Google Sheet with the service account**
1. Create a new sheet at [sheets.google.com](https://sheets.google.com)
2. Click **Share** → paste the `client_email` → set to **Editor** → **Send**

---

#### Part C — Deploy on Streamlit Community Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io) → sign in with GitHub
2. Click **New app** → pick your repo → set **Main file path** to `study-app/app.py`
3. Before clicking Deploy, open **Advanced settings → Secrets** and paste this
   (replace everything inside the braces with the contents of your downloaded `.json` file):

```toml
[GOOGLE_SERVICE_ACCOUNT_KEY]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\\n...\\n-----END RSA PRIVATE KEY-----\\n"
client_email = "your-service-account@project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

4. Click **Deploy** — your app gets a permanent `https://your-name.streamlit.app` URL!
5. Paste your Google Sheet URL into the sidebar → start logging sessions

---
**Tip:** Every time you push a new commit to GitHub, Streamlit Community Cloud auto-redeploys. Your study data stays safe in Google Sheets throughout.
        """)
else:
    # Show the service account email so the user knows what to share their sheet with
    key = _get_sa_key()
    if key and "client_email" in key:
        st.sidebar.caption(f"Service account: `{key['client_email']}`")
