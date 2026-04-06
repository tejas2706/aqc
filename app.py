from pathlib import Path
from html import escape

import streamlit as st

from aqc_data import DATA_PATH, add_attendee, load_or_create_dataset, search_attendees


st.set_page_config(page_title="AQC Attendee Tag Finder", page_icon="🏷️", layout="centered")

st.markdown(
    """
    <style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(34, 197, 94, 0.08), transparent 28%),
            linear-gradient(180deg, #f8faf7 0%, #eef2f6 100%);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 980px;
    }
    .app-shell {
        background: rgba(255, 255, 255, 0.82);
        border: 1px solid rgba(209, 213, 219, 0.85);
        border-radius: 24px;
        padding: 1.2rem 1.2rem 1rem;
        box-shadow: 0 20px 50px rgba(15, 23, 42, 0.07);
        backdrop-filter: blur(8px);
        margin-bottom: 1.2rem;
    }
    .app-heading {
        margin: 0;
        font-size: 1.3rem;
        line-height: 1.1;
        color: #1f2937;
        letter-spacing: -0.03em;
        font-weight: 700;
    }
    .app-subtitle {
        margin: 0.45rem 0 0;
        color: #5b6473;
        font-size: 0.98rem;
    }
    .section-title {
        margin: 0 0 0.75rem;
        font-size: 1.15rem;
        font-weight: 700;
        color: #253047;
    }
    .tab-copy {
        margin: 0 0 1rem;
        color: #5b6473;
        font-size: 0.96rem;
    }
    .result-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
        gap: 0.9rem;
        margin-top: 1rem;
    }
    .result-card {
        border: 1px solid #cfd6df;
        border-radius: 18px;
        padding: 1rem 1rem 0.9rem;
        background: linear-gradient(180deg, #f2f4f7 0%, #e7ebf0 100%);
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.08);
        min-height: 220px;
    }
    .result-name {
        margin: 0 0 0.75rem;
        font-size: 1.25rem;
        line-height: 1.3;
        color: #2b2d42;
    }
    .tag-pill {
        display: inline-block;
        padding: 0.2rem 0.65rem;
        border: 2px solid #1faa59;
        border-radius: 999px;
        background: #dff7e8;
        color: #0f6b36;
        font-weight: 700;
        margin-bottom: 0.85rem;
    }
    .result-meta {
        margin: 0.35rem 0;
        color: #404457;
        font-size: 0.95rem;
        line-height: 1.45;
        word-break: break-word;
    }
    .result-source {
        margin-top: 0.8rem;
        color: #7a8194;
        font-size: 0.82rem;
    }
    .metric-panel {
        background: rgba(255, 255, 255, 0.74);
        border: 1px solid #d7dde5;
        border-radius: 18px;
        padding: 0.3rem 0.85rem;
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05);
    }
    div[data-testid="stMetric"] {
        background: transparent;
        border: none;
        padding: 0.2rem 0;
    }
    div[data-testid="stForm"] {
        background: rgba(255, 255, 255, 0.72);
        border: 1px solid #d7dde5;
        border-radius: 18px;
        padding: 0.7rem;
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="stAlert"] {
        border-radius: 16px;
    }
    div[data-testid="stTabs"] button[role="tab"] {
        border-radius: 999px;
        padding: 0.55rem 0.9rem;
        font-size: 0.95rem;
        white-space: nowrap;
    }
    div[data-testid="stTextInput"] input {
        min-height: 3rem;
        font-size: 1rem;
    }
    div[data-testid="stButton"] button,
    div[data-testid="stFormSubmitButton"] button {
        min-height: 3rem;
        border-radius: 14px;
        font-weight: 600;
    }
    .spacer-sm {
        height: 0.35rem;
    }
    @media (max-width: 1024px) {
        .block-container {
            max-width: 100%;
            padding-left: 1.25rem;
            padding-right: 1.25rem;
        }
        .app-shell {
            padding: 1rem;
            border-radius: 20px;
        }
        .app-heading {
            font-size: 1.18rem;
        }
        .result-grid {
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        }
    }
    @media (max-width: 768px) {
        .block-container {
            padding-top: 1rem;
            padding-bottom: 1.25rem;
            padding-left: 0.9rem;
            padding-right: 0.9rem;
        }
        .app-shell {
            padding: 0.95rem;
            border-radius: 18px;
        }
        .app-heading {
            font-size: 1.08rem;
        }
        .app-subtitle {
            font-size: 0.92rem;
        }
        .section-title {
            font-size: 1.05rem;
        }
        .result-grid {
            grid-template-columns: 1fr;
            gap: 0.75rem;
        }
        .result-card {
            min-height: auto;
            padding: 0.9rem;
            border-radius: 16px;
        }
        .result-name {
            font-size: 1.12rem;
            margin-bottom: 0.6rem;
        }
        .result-meta {
            font-size: 0.92rem;
            line-height: 1.35;
        }
        .tag-pill {
            padding: 0.18rem 0.55rem;
            margin-bottom: 0.7rem;
        }
        div[data-testid="stTabs"] button[role="tab"] {
            padding: 0.48rem 0.7rem;
            font-size: 0.86rem;
        }
    }
    @media (max-width: 480px) {
        .block-container {
            padding-left: 0.7rem;
            padding-right: 0.7rem;
        }
        .app-heading {
            font-size: 1rem;
        }
        .app-subtitle,
        .tab-copy {
            font-size: 0.88rem;
        }
        div[data-testid="stTabs"] {
            overflow-x: auto;
        }
        div[data-testid="stTabs"] button[role="tab"] {
            min-width: max-content;
        }
        .metric-panel {
            padding: 0.25rem 0.7rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def refresh_dataset(force_refresh: bool = False) -> dict:
    return load_or_create_dataset(force_refresh=force_refresh)


def render_result_cards(results: list[dict]) -> None:
    cards = []
    for result in results:
        cards.append(
            (
                '<div class="result-card">'
                f'<h3 class="result-name">{escape(result.get("name", "Unknown"))}</h3>'
                f'<div class="tag-pill">{escape(result.get("tag", "-"))}</div>'
                f'<p class="result-meta"><strong>Email ID:</strong> {escape(result.get("emailId", "-") or "-")}</p>'
                f'<p class="result-meta"><strong>Institute:</strong> {escape(result.get("institute", "-") or "-")}</p>'
                f'<p class="result-meta"><strong>Hub:</strong> {escape(result.get("hub", "-") or "-")}</p>'
                f'<div class="result-source">Source: {escape(result.get("source", "unknown"))}</div>'
                "</div>"
            )
        )

    st.markdown(
        '<div class="result-grid">' + "".join(cards) + "</div>",
        unsafe_allow_html=True,
    )


st.markdown(
    """
    <div class="app-shell">
        <div class="app-heading">AQC Attendee Tag Finder</div>
        <p class="app-subtitle">Search by attendee name or email ID, and manage additions.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if "dataset" not in st.session_state:
    st.session_state.dataset = refresh_dataset()
if "show_add_form" not in st.session_state:
    st.session_state.show_add_form = False


search_tab, sync_tab, add_tab = st.tabs(
    ["Search", "Sync with excel", "Add attendee"]
)

with search_tab:
    st.markdown('<div class="section-title">Search attendee:</div>', unsafe_allow_html=True)
    query = st.text_input(
        "Enter a name or email ID",
        placeholder="e.g. Akshay Naik or anaik@iisc.ac.in",
        key="search_query",
    )
    search_clicked = st.button("Search", use_container_width=True, key="search_button")

    if search_clicked:
        if not query.strip():
            st.warning("Enter a name or email ID to search.")
        else:
            results = search_attendees(query, st.session_state.dataset.get("records", []))
            if results:
                st.success(f"Found {len(results)} matching attendee(s).")
                render_result_cards(results)
            else:
                st.warning("No attendee matched that search.")

with sync_tab:
    st.markdown('<div class="section-title">Sync with Excel</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="tab-copy">Refresh the local JSON cache from the Excel workbook whenever the source sheet changes.</div>',
        unsafe_allow_html=True,
    )
    col1, col2 = st.columns([1.2, 1])
    with col1:
        if st.button("Refresh from Excel", use_container_width=True, key="refresh_button"):
            st.session_state.dataset = refresh_dataset(force_refresh=True)
            st.success(f"Reloaded attendee data into {DATA_PATH.name}")
    with col2:
        st.markdown('<div class="metric-panel">', unsafe_allow_html=True)
        st.metric("People loaded", len(st.session_state.dataset.get("records", [])))
        st.markdown("</div>", unsafe_allow_html=True)

with add_tab:
    st.markdown('<div class="section-title">Add attendee</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="tab-copy">Add a new attendee to the local JSON list without modifying the Excel workbook.</div>',
        unsafe_allow_html=True,
    )
    toggle_label = "Hide add attendee form" if st.session_state.show_add_form else "Add attendee"
    if st.button(toggle_label, use_container_width=True, key="toggle_add_form"):
        st.session_state.show_add_form = not st.session_state.show_add_form

    if st.session_state.show_add_form:
        with st.form("add_attendee_form", clear_on_submit=True):
            name = st.text_input("Name *")
            tag = st.text_input("Tag *")
            institute = st.text_input("Institute")
            hub = st.text_input("Hub")
            email = st.text_input("emailId")
            submitted = st.form_submit_button("Add to local list", use_container_width=True)

        if submitted:
            if not name.strip() or not tag.strip():
                st.error("Name and Tag are required.")
            else:
                attendee = add_attendee(name=name, tag=tag, institute=institute, hub=hub, email=email)
                st.session_state.dataset = refresh_dataset()
                st.session_state.show_add_form = False
                st.success(f"Added {attendee['name']} with tag {attendee['tag']} to {Path(DATA_PATH).name}.")


# with st.expander("Dataset details"):
#     st.json(
#         {
#             "source_file": st.session_state.dataset.get("source_file"),
#             "sheet_name": st.session_state.dataset.get("sheet_name"),
#             "generated_at": st.session_state.dataset.get("generated_at"),
#             "json_path": str(DATA_PATH),
#             "manual_entries": len(st.session_state.dataset.get("added_records", [])),
#         }
#     )
