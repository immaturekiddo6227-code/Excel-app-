import io, json, re, os, ast, operator
from typing import Any
import pandas as pd
import plotly.express as px
import streamlit as st
import requests

# Chart palette: emerald/mint first, then distinct supporting hues for readability
CHART_COLORWAY = ["#34d399", "#2dd4bf", "#a3e635", "#38bdf8", "#fbbf24", "#a78bfa", "#fb7185", "#94a3b8"]
px.defaults.template = "plotly_dark"
px.defaults.color_discrete_sequence = CHART_COLORWAY

st.set_page_config(page_title="Excel AI Studio", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
    /* Prevent the main content from sitting underneath the toolbar */
    [data-testid="stAppViewContainer"] {
        overflow: visible;
    }

    [data-testid="stMain"] {
        padding-top: 1.5rem !important;
        overflow: visible;
    }

    [data-testid="stMainBlockContainer"] {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 1500px;
    }

    /* Keep the Streamlit toolbar above the page without covering content */
    [data-testid="stToolbar"] {
        z-index: 999;
    }

    /* Ensure the hero/banner is fully displayed */
    .hero, .hero-banner, .hero-section {
        margin-top: 1rem !important;
        position: relative;
        z-index: 1;
        overflow: visible !important;
        box-sizing: border-box;
    }

    /* Avoid accidental clipping of rounded banner containers */
    div[data-testid="stVerticalBlock"] {
        overflow: visible;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- Excel AI Studio: Aurora Glass UI ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    /* Emerald + mint design tokens */
    --bg: #06100d;
    --surface: #0b1a15;
    --surface-2: #0e1e1a;
    --border: rgba(167, 243, 208, .13);
    --border-strong: rgba(110, 231, 183, .34);
    --text: #ecf6f1;
    --text-2: #c9ddd4;
    --muted: #8fa89d;
    --emerald-600: #059669;
    --emerald-500: #10b981;
    --emerald-400: #34d399;
    --mint-300: #6ee7b7;
    --mint-200: #a7f3d0;
    --ring: rgba(52, 211, 153, .32);
    --font: "Inter", "Segoe UI", system-ui, -apple-system, "Helvetica Neue", Arial, sans-serif;
    --ease: cubic-bezier(.2, .7, .2, 1);
    color-scheme: dark;
}

.stApp {
    color: var(--text);
    background:
        radial-gradient(ellipse 60% 40% at 8% -5%, rgba(16, 185, 129, .15), transparent 62%),
        radial-gradient(ellipse 50% 35% at 96% 0%, rgba(52, 211, 153, .07), transparent 62%),
        linear-gradient(rgba(167, 243, 208, .018) 1px, transparent 1px),
        linear-gradient(90deg, rgba(167, 243, 208, .018) 1px, transparent 1px),
        var(--bg);
    background-size: auto, auto, 34px 34px, 34px 34px, auto;
}

/* Keep the dark emerald canvas visible whatever the device or Streamlit theme is */
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background: transparent;
}

[data-testid="stSidebar"] {
    color: var(--text);
}

::selection {
    background: rgba(52, 211, 153, .3);
}

.block-container {
    max-width: 1580px;
    padding-top: 2rem;
    padding-bottom: 3rem;
    padding-left: clamp(1rem, 3vw, 2.5rem);
    padding-right: clamp(1rem, 3vw, 2.5rem);
    box-sizing: border-box;
}
/* Streamlit top header and toolbar */
header[data-testid="stHeader"] {
    background: rgba(9, 14, 27, 0.92) !important;
    box-shadow: none !important;
}

/* Keep toolbar controls above page content */
div[data-testid="stToolbar"] {
    z-index: 1001 !important;
}

/* Prevent the hero from competing with the toolbar */
/* Stable Streamlit header */
header[data-testid="stHeader"] {
    background: rgba(9, 14, 27, 0.96) !important;
    box-shadow: none !important;
}

/* Keep toolbar controls above page content */
div[data-testid="stToolbar"] {
    z-index: 1001;
}

/* Consistent main content spacing */
[data-testid="stMainBlockContainer"] {
    box-sizing: border-box;
    max-width: 1580px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}

/* Keep the hero in normal document flow */
.hero {
    position: relative;
    isolation: isolate;
    z-index: auto;
    margin-top: 0;
}
/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(170deg, #0c1c17 0%, #07130f 80%);
    border-right: 1px solid var(--border);
}

[data-testid="stSidebar"] [data-testid="stHeading"] h3,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
    color: var(--mint-200);
    font-size: 1.02rem;
    font-weight: 650;
    letter-spacing: -.005em;
}

/* ---------- Typography ---------- */
.stApp :is(button, input, textarea, select, label,
           [data-baseweb="tab"], [data-baseweb="select"],
           [data-testid="stCaptionContainer"], [data-testid="stHeading"],
           [data-testid="stMetric"], [data-testid="stAlert"],
           [data-testid="stToast"], [data-testid="stExpander"] summary),
[data-testid="stMarkdownContainer"]:not(:has(.hero)),
[data-baseweb="popover"] :is(li, [role="option"]) {
    font-family: var(--font);
}

[data-testid="stMarkdownContainer"]:not(:has(.hero)) h3,
[data-testid="stHeading"] h3 {
    color: var(--text);
    font-size: 1.3rem;
    font-weight: 700;
    letter-spacing: -.015em;
    line-height: 1.3;
    padding: .35rem 0 .6rem;
}

[data-testid="stMarkdownContainer"]:not(:has(.hero)) strong {
    color: #f6fcf9;
    font-weight: 650;
}

[data-testid="stMarkdownContainer"]:not(:has(.hero)) a,
[data-testid="stCaptionContainer"] a {
    color: var(--mint-300);
    text-decoration-color: rgba(110, 231, 183, .4);
    text-underline-offset: 3px;
    transition: color .18s ease, text-decoration-color .18s ease;
}

[data-testid="stMarkdownContainer"]:not(:has(.hero)) a:hover,
[data-testid="stCaptionContainer"] a:hover {
    color: var(--mint-200);
    text-decoration-color: currentColor;
}

[data-testid="stMarkdownContainer"]:not(:has(.hero)) li::marker {
    color: var(--emerald-400);
}

[data-testid="stMarkdownContainer"]:not(:has(.hero)) code {
    padding: .1em .4em;
    border-radius: 6px;
    color: var(--mint-200);
    background: rgba(16, 185, 129, .12);
}

[data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] p {
    color: var(--muted);
    font-size: .86rem;
    line-height: 1.5;
}

[data-testid="stWidgetLabel"] p {
    color: var(--text-2);
    font-size: .88rem;
    font-weight: 600;
    letter-spacing: .01em;
}

.stApp hr {
    border-color: var(--border);
}

/* ---------- Theme-proof text colours ----------
   The app always uses its dark emerald palette, so every text-bearing
   widget gets an explicit light colour. Text stays readable even when
   the device, browser or Streamlit itself is in light mode. */
[data-testid="stMarkdownContainer"]:not(:has(.hero)),
[data-testid="stMarkdownContainer"]:not(:has(.hero)) :is(p, li, em, td, th) {
    color: inherit;
}

[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] label {
    color: var(--text-2);
}

[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
    color: var(--text);
    caret-color: var(--emerald-400);
}

[data-testid="stSelectbox"] [data-baseweb="select"] :is(div, span, input),
[data-testid="stMultiSelect"] [data-baseweb="select"] :is(div, span, input) {
    color: var(--text);
}

[data-baseweb="select"] svg {
    color: var(--muted);
    fill: var(--muted);
}

[data-baseweb="select"] input::placeholder {
    color: var(--muted);
}

[data-baseweb="popover"] > div,
[data-baseweb="popover"] [data-baseweb="menu"],
[data-baseweb="popover"] ul[role="listbox"] {
    background-color: var(--surface-2);
}

[data-baseweb="popover"] :is(li, [role="option"], [role="option"] *) {
    color: var(--text);
}

[data-testid="stCheckbox"] label,
[data-testid="stRadio"] label,
[data-testid="stExpander"] summary {
    color: var(--text);
}

[data-testid="stCheckbox"] label *,
[data-testid="stRadio"] label *,
[data-testid="stExpander"] summary * {
    color: inherit;
}

[data-testid="stFileUploader"] {
    color: var(--text-2);
}

[data-testid="stFileUploader"] :is(span, p, div, section, li, label) {
    color: inherit;
}

[data-testid="stFileUploader"] small {
    color: var(--muted);
}

[data-testid="stAlert"],
[data-testid="stToast"] {
    color: var(--text);
}

[data-testid="stAlert"] :is(p, span, div, li, strong, code),
[data-testid="stToast"] :is(p, span, div) {
    color: inherit;
}

div[data-testid="stMetricLabel"] *,
div[data-testid="stMetricValue"] * {
    color: inherit;
}

/* Header controls (sidebar toggle, toolbar, menu) stay readable on the dark canvas */
[data-testid="stSidebarCollapsedControl"] :is(button, span, svg),
[data-testid="stSidebarCollapseButton"] :is(button, span, svg),
[data-testid="stToolbar"] :is(button, a, span, svg, p) {
    color: var(--text-2);
}

[data-testid="stSidebarCollapsedControl"] button:hover,
[data-testid="stSidebarCollapseButton"] button:hover,
[data-testid="stToolbar"] button:hover {
    color: var(--mint-200);
    background-color: rgba(16, 185, 129, .1);
}

/* ---------- Metric cards ---------- */
div[data-testid="stMetric"] {
    padding: 1.05rem 1.2rem;
    border-radius: 16px;
    border: 1px solid var(--border);
    background: linear-gradient(150deg, rgba(19, 40, 34, .85), rgba(10, 24, 19, .92));
    box-shadow: 0 8px 24px rgba(0, 0, 0, .18),
                inset 0 1px 0 rgba(255, 255, 255, .03);
    transition: transform .22s var(--ease),
                border-color .22s ease,
                box-shadow .22s ease;
}

div[data-testid="stMetric"]:hover {
    transform: translateY(-2px);
    border-color: var(--border-strong);
    box-shadow: 0 12px 30px rgba(0, 0, 0, .24);
}

div[data-testid="stMetricLabel"],
div[data-testid="stMetricLabel"] p {
    color: var(--muted);
    font-size: .84rem;
    font-weight: 600;
    letter-spacing: .03em;
}

div[data-testid="stMetricValue"] {
    color: #fff;
    font-weight: 750;
    letter-spacing: -.02em;
    font-variant-numeric: tabular-nums;
}

/* ---------- Buttons ---------- */
.stButton > button,
.stDownloadButton > button,
[data-testid="stFileUploaderDropzone"] button {
    min-height: 44px;
    padding: .5rem 1.1rem;
    border-radius: 12px;
    border: 1px solid var(--border-strong);
    background: rgba(255, 255, 255, .035);
    color: var(--text);
    font-size: .93rem;
    font-weight: 600;
    line-height: 1.25;
    letter-spacing: .005em;
    box-shadow: 0 2px 8px rgba(0, 0, 0, .18),
                inset 0 1px 0 rgba(255, 255, 255, .04);
    transition: transform .18s var(--ease),
                box-shadow .18s ease,
                border-color .18s ease,
                background-color .18s ease,
                color .18s ease,
                filter .18s ease;
}

.stButton > button p,
.stDownloadButton > button p {
    margin: 0;
    color: inherit;
    font-weight: inherit;
}

.stButton > button:hover:not(:disabled),
.stDownloadButton > button:hover:not(:disabled),
[data-testid="stFileUploaderDropzone"] button:hover:not(:disabled) {
    transform: translateY(-1px);
    border-color: rgba(52, 211, 153, .65);
    background: rgba(16, 185, 129, .12);
    color: #f4fcf8;
    box-shadow: 0 6px 18px rgba(0, 0, 0, .26);
}

.stButton > button:active:not(:disabled),
.stDownloadButton > button:active:not(:disabled) {
    transform: translateY(0);
    box-shadow: 0 2px 6px rgba(0, 0, 0, .22);
}

.stButton > button:focus-visible,
.stDownloadButton > button:focus-visible,
[data-testid="stFileUploaderDropzone"] button:focus-visible {
    outline: 2px solid var(--emerald-400);
    outline-offset: 2px;
}

.stButton > button:disabled,
.stDownloadButton > button:disabled {
    color: var(--muted);
    background: rgba(255, 255, 255, .02);
    border-color: var(--border);
    cursor: not-allowed;
    box-shadow: none;
}

/* Primary actions */
.stButton > button[kind="primary"],
.stButton > button[data-testid="stBaseButton-primary"] {
    color: #04140d;
    background: linear-gradient(135deg, #10b981 0%, #34d399 100%);
    border-color: rgba(110, 231, 183, .55);
    box-shadow: 0 6px 18px rgba(16, 185, 129, .22),
                inset 0 1px 0 rgba(255, 255, 255, .25);
}

.stButton > button[kind="primary"]:hover:not(:disabled),
.stButton > button[data-testid="stBaseButton-primary"]:hover:not(:disabled) {
    color: #04140d;
    background: linear-gradient(135deg, #10b981 0%, #34d399 100%);
    border-color: rgba(167, 243, 208, .8);
    filter: brightness(1.07);
    box-shadow: 0 10px 26px rgba(16, 185, 129, .3);
}

/* ---------- Inputs ---------- */
[data-testid="stTextInput"] [data-baseweb="input"],
[data-testid="stTextArea"] [data-baseweb="textarea"],
[data-testid="stSelectbox"] [data-baseweb="select"] > div,
[data-testid="stMultiSelect"] [data-baseweb="select"] > div {
    background-color: rgba(255, 255, 255, .035);
    border: 1px solid var(--border);
    border-radius: 12px;
    transition: border-color .18s ease,
                box-shadow .18s ease,
                background-color .18s ease;
}

[data-testid="stTextInput"] [data-baseweb="base-input"],
[data-testid="stTextArea"] [data-baseweb="base-input"],
[data-testid="stTextInput"] input,
[data-testid="stTextArea"] textarea {
    background-color: transparent;
    color: var(--text);
}

[data-testid="stTextInput"] input::placeholder,
[data-testid="stTextArea"] textarea::placeholder {
    color: var(--muted);
    opacity: .8;
}

[data-testid="stTextInput"] [data-baseweb="input"]:hover,
[data-testid="stTextArea"] [data-baseweb="textarea"]:hover,
[data-testid="stSelectbox"] [data-baseweb="select"] > div:hover,
[data-testid="stMultiSelect"] [data-baseweb="select"] > div:hover {
    border-color: var(--border-strong);
}

[data-testid="stTextInput"] [data-baseweb="input"]:focus-within,
[data-testid="stTextArea"] [data-baseweb="textarea"]:focus-within,
[data-testid="stSelectbox"] [data-baseweb="select"] > div:focus-within,
[data-testid="stMultiSelect"] [data-baseweb="select"] > div:focus-within {
    border-color: var(--emerald-400);
    background-color: rgba(16, 185, 129, .05);
    box-shadow: 0 0 0 3px var(--ring);
}

/* Dropdown menus */
[data-baseweb="popover"] [data-baseweb="menu"],
[data-baseweb="popover"] ul[role="listbox"] {
    background: var(--surface-2);
    border: 1px solid var(--border-strong);
    border-radius: 12px;
    box-shadow: 0 16px 40px rgba(0, 0, 0, .45);
}

[data-baseweb="popover"] li[role="option"] {
    transition: background-color .15s ease;
}

[data-baseweb="popover"] li[role="option"]:hover,
[data-baseweb="popover"] li[role="option"][aria-selected="true"] {
    background-color: rgba(16, 185, 129, .16);
}

/* Multi-select tags */
span[data-baseweb="tag"] {
    background-color: rgba(16, 185, 129, .18);
    border-radius: 8px;
    color: var(--mint-200);
}

span[data-baseweb="tag"] span {
    color: inherit;
}

/* Checkboxes and radios */
[data-testid="stCheckbox"] label,
[data-testid="stRadio"] label {
    transition: color .18s ease;
}

[data-testid="stCheckbox"] label:hover,
[data-testid="stRadio"] label:hover {
    color: var(--mint-200);
}

/* ---------- Tabs ---------- */
[data-baseweb="tab-list"] {
    gap: .25rem;
}

div[data-baseweb="tab-border"] {
    background-color: var(--border);
}

button[data-baseweb="tab"] {
    padding: .65rem 1rem;
    border-radius: 10px 10px 0 0;
    color: var(--muted);
    font-weight: 600;
    transition: color .2s ease, background-color .2s ease;
}

button[data-baseweb="tab"]:hover {
    color: var(--text);
    background-color: rgba(16, 185, 129, .06);
}

button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--mint-300);
}

button[data-baseweb="tab"]:focus-visible {
    outline: 2px solid var(--emerald-400);
    outline-offset: -2px;
}

div[data-baseweb="tab-highlight"] {
    background: linear-gradient(90deg, var(--emerald-500), var(--mint-300));
    height: 3px;
    border-radius: 99px;
}

/* ---------- Data tables ---------- */
[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
    border: 1px solid var(--border);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 8px 28px rgba(0, 0, 0, .2);
    transition: border-color .2s ease;
}

[data-testid="stDataFrame"]:hover,
[data-testid="stDataEditor"]:hover {
    border-color: var(--border-strong);
}

/* ---------- Expanders ---------- */
[data-testid="stExpander"] {
    border: 1px solid var(--border);
    border-radius: 14px;
    background: rgba(13, 28, 23, .55);
    overflow: hidden;
}

[data-testid="stExpander"] summary {
    font-weight: 600;
    transition: background-color .18s ease, color .18s ease;
}

[data-testid="stExpander"] summary:hover {
    color: var(--mint-300);
    background-color: rgba(16, 185, 129, .06);
}

/* ---------- File upload ---------- */
[data-testid="stFileUploader"] {
    padding: 10px;
    border: 1px dashed rgba(52, 211, 153, .4);
    border-radius: 16px;
    background: rgba(16, 185, 129, .04);
    transition: border-color .2s ease, background-color .2s ease;
}

[data-testid="stFileUploader"]:hover {
    border-color: var(--emerald-400);
    background: rgba(16, 185, 129, .08);
}

[data-testid="stFileUploaderDropzone"] {
    background: transparent;
    border-radius: 12px;
}

/* ---------- Status messages ---------- */
[data-testid="stAlert"] > div,
[data-testid="stAlert"] [data-baseweb="notification"] {
    border-radius: 12px;
    border: 1px solid var(--border);
    background-color: rgba(255, 255, 255, .04);
}

[data-testid="stAlert"] p {
    color: var(--text);
}

[data-testid="stAlert"] [data-baseweb="notification"]:has([data-testid="stAlertContentSuccess"]) {
    background-color: rgba(16, 185, 129, .1);
    border-color: rgba(52, 211, 153, .32);
}

[data-testid="stAlert"] [data-baseweb="notification"]:has([data-testid="stAlertContentInfo"]) {
    background-color: rgba(56, 189, 248, .08);
    border-color: rgba(56, 189, 248, .28);
}

[data-testid="stAlert"] [data-baseweb="notification"]:has([data-testid="stAlertContentWarning"]) {
    background-color: rgba(251, 191, 36, .09);
    border-color: rgba(251, 191, 36, .3);
}

[data-testid="stAlert"] [data-baseweb="notification"]:has([data-testid="stAlertContentError"]) {
    background-color: rgba(251, 113, 133, .09);
    border-color: rgba(251, 113, 133, .32);
}

[data-testid="stToast"] {
    background-color: var(--surface-2);
    border: 1px solid var(--border-strong);
    border-radius: 12px;
    box-shadow: 0 12px 32px rgba(0, 0, 0, .4);
}

/* ---------- Code / text output ---------- */
[data-testid="stCode"] {
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
}

[data-testid="stCode"] pre,
[data-testid="stCode"] code {
    background: #07130f !important;
    color: var(--text-2) !important;
    font-size: .86rem;
    line-height: 1.55;
}

[data-testid="stCode"] code span {
    color: inherit !important;
}

/* Hero / branding */
.hero {
    position: relative;
    isolation: isolate;
    overflow: hidden;
    z-index: 0;
    padding: 2.6rem 2.4rem;
    margin-bottom: 1.5rem;
    border-radius: 25px;
    border: 1px solid rgba(101, 245, 192, .25);
    background:
        radial-gradient(ellipse at 90% 0%, rgba(167,139,250,.23), transparent 38%),
        linear-gradient(125deg, rgba(17,39,48,.97), rgba(19,23,48,.97));
    box-shadow: 0 20px 60px rgba(0,0,0,.24),
                inset 0 1px rgba(255,255,255,.04);
    animation: none;
}

.hero::before {
    content: "";
    position: absolute;
    z-index: -1;
    width: 250px;
    height: 250px;
    right: -100px;
    top: -140px;
    border-radius: 50%;
    background: rgba(101,245,192,.24);
    filter: blur(48px);
    animation: aurora 7s ease-in-out infinite alternate;
}

.hero::after {
    content: "";
    position: absolute;
    top: 0;
    bottom: 0;
    left: -65%;
    width: 38%;
    transform: skewX(-18deg);
    background: linear-gradient(90deg, transparent,
                rgba(255,255,255,.045), transparent);
    animation: light-sweep 9s ease-in-out infinite;
    pointer-events: none;
}

.hero-brand {
    display: flex;
    align-items: center;
    gap: 19px;
    position: relative;
}

.hero-logo {
    width: 82px;
    height: 82px;
    flex-shrink: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 2.7rem;
    border-radius: 23px;
    background: linear-gradient(140deg, #173f39, #29234c);
    border: 1px solid rgba(101,245,192,.38);
    box-shadow: 0 0 35px rgba(101,245,192,.13),
                inset 0 1px rgba(255,255,255,.12);
    animation: logo-float 4.5s ease-in-out infinite;
}

.hero h1 {
    color: #fff;
    font-size: clamp(2.1rem, 4vw, 3.3rem);
    font-weight: 850;
    line-height: 1.25;
    padding-top: 0.08em;
    padding-bottom: 0.08em;
    letter-spacing: -1.8px;
    margin: 0;
}

.hero h1 span {
    background: linear-gradient(100deg, #65f5c0, #a3f4d8 45%, #b8a5ff);
    -webkit-background-clip: text;
    background-clip: text;
    color: transparent;
}

.hero p {
    color: #b8c6dc;
    font-size: 1rem;
    line-height: 1.75;
    margin: .95rem 0 0;
    max-width: 700px;
}

.hero-badge {
    display: inline-block;
    margin-top: 1.1rem;
    padding: 8px 13px;
    border-radius: 999px;
    color: #a8f8da;
    background: rgba(101,245,192,.075);
    border: 1px solid rgba(101,245,192,.23);
    font-size: .75rem;
    font-weight: 700;
    letter-spacing: .8px;
}

/* Stabilize the main page while scrolling */
[data-testid="stAppViewContainer"] {
    overflow-x: hidden;
}

[data-testid="stMain"] {
    overflow-x: hidden;
}

[data-testid="stMainBlockContainer"] {
    box-sizing: border-box;
}

/* Keep the top toolbar above the page */
header[data-testid="stHeader"] {
    position: sticky;
    top: 0;
    z-index: 1000;
    background: #090e1b !important;
    box-shadow: none !important;
}
/* ---------- Charts ---------- */
[data-testid="stPlotlyChart"] {
    padding: .4rem .5rem;
    border: 1px solid var(--border);
    border-radius: 16px;
    background-color: var(--surface);
    box-shadow: 0 8px 28px rgba(0, 0, 0, .2);
    overflow: hidden;
    transition: border-color .22s ease, box-shadow .22s ease;
}

[data-testid="stPlotlyChart"]:hover {
    border-color: var(--border-strong);
    box-shadow: 0 12px 34px rgba(0, 0, 0, .26);
}

.js-plotly-plot .plotly .modebar {
    opacity: .65;
    transition: opacity .2s ease;
}

.js-plotly-plot:hover .plotly .modebar {
    opacity: 1;
}

.small-note {
    color: var(--muted);
    font-size: .86rem;
    text-align: center;
    padding-top: 1.2rem;
    letter-spacing: .15px;
}

/* ---------- Gentle appearance transition ---------- */
@keyframes fade-up {
    from { opacity: 0; transform: translateY(6px); }
    to { opacity: 1; transform: none; }
}

[data-testid="stTabs"] [role="tabpanel"] {
    animation: fade-up .28s var(--ease) backwards;
}

/* ---------- Responsive refinements ---------- */
@media (max-width: 900px) {
    div[data-testid="stMetric"] {
        padding: .9rem 1rem;
    }

    [data-testid="stMarkdownContainer"]:not(:has(.hero)) h3,
    [data-testid="stHeading"] h3 {
        font-size: 1.15rem;
    }
}

@media (max-width: 640px) {
    .stButton > button,
    .stDownloadButton > button {
        min-height: 46px;
    }

    button[data-baseweb="tab"] {
        padding: .55rem .75rem;
    }
}

@keyframes hero-enter {
    from { opacity: 0; transform: translateY(14px); }
    to { opacity: 1; transform: translateY(0); }
}

@keyframes logo-float {
    0%, 100% { transform: translateY(0) rotate(0deg); }
    50% { transform: translateY(-5px) rotate(-2deg); }
}

@keyframes aurora {
    from { transform: scale(.88) translate(0,0); opacity: .55; }
    to { transform: scale(1.12) translate(-15px,12px); opacity: 1; }
}

@keyframes light-sweep {
    0%, 18% { left: -65%; }
    55%, 100% { left: 135%; }
}

@media (max-width: 640px) {
    .block-container {
        padding-top: 1rem !important;
        padding-right: 1rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        box-sizing: border-box;
    }

    .hero {
        padding: 1.3rem;
        border-radius: 19px;
    }

    .hero-brand {
        gap: 12px;
        align-items: flex-start;
    }

    .hero-logo {
        width: 55px;
        height: 55px;
        font-size: 1.8rem;
        border-radius: 15px;
    }

    .hero h1 {
        font-size: 1.8rem;
        letter-spacing: -1px;
    }

    .hero p {
        font-size: .9rem;
    }
}

@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
        animation: none !important;
        transition: none !important;
        scroll-behavior: auto !important;
    }
}
/* Keep Streamlit header visible so the sidebar toggle works */
header[data-testid="stHeader"] {
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
    background: transparent !important;
    box-shadow: none !important;
}

/* Keep the toolbar and sidebar controls available */
div[data-testid="stToolbar"] {
    display: flex !important;
    visibility: visible !important;
    opacity: 1 !important;
    z-index: 1001 !important;
}

/* Keep sidebar open/close controls visible */
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapseButton"] {
    display: flex !important;
    visibility: visible !important;
    opacity: 1 !important;
}

/* Hide only unnecessary decoration and footer */
[data-testid="stDecoration"],
#MainMenu,
footer {
    display: none !important;
}

/* ---------- Consistent emerald theme in every device mode ---------- */

/* Keep the existing dark emerald palette even when the device is in light mode */
:root[data-theme="light"] {
    color-scheme: dark;
    --bg: #06100d;
    --surface: #0b1a15;
    --surface-2: #0e1e1a;
    --border: rgba(167, 243, 208, .13);
    --border-strong: rgba(110, 231, 183, .34);
    --text: #ecf6f1;
    --text-2: #c9ddd4;
    --muted: #b4c9bf;
    --ring: rgba(52, 211, 153, .32);
}

/* Keep the app background greenish and all standard text readable */
:root[data-theme="light"] .stApp {
    color: var(--text) !important;
    color-scheme: dark;
}

/* Main content, paragraphs, headings, lists and bold text */
:root[data-theme="light"] [data-testid="stMain"],
:root[data-theme="light"] [data-testid="stMarkdownContainer"],
:root[data-theme="light"] [data-testid="stMarkdownContainer"] p,
:root[data-theme="light"] [data-testid="stMarkdownContainer"] li,
:root[data-theme="light"] [data-testid="stMarkdownContainer"] h1,
:root[data-theme="light"] [data-testid="stMarkdownContainer"] h2,
:root[data-theme="light"] [data-testid="stMarkdownContainer"] h3,
:root[data-theme="light"] [data-testid="stMarkdownContainer"] strong,
:root[data-theme="light"] [data-testid="stHeading"] {
    color: var(--text) !important;
}

/* Sidebar text and headings */
:root[data-theme="light"] [data-testid="stSidebar"] {
    background: linear-gradient(170deg, #0c1c17 0%, #07130f 80%) !important;
    color: var(--text) !important;
}

:root[data-theme="light"] [data-testid="stSidebar"] p,
:root[data-theme="light"] [data-testid="stSidebar"] label,
:root[data-theme="light"] [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
:root[data-theme="light"] [data-testid="stSidebar"] [data-testid="stHeading"] h3 {
    color: var(--text) !important;
}

/* Captions and widget labels */
:root[data-theme="light"] [data-testid="stCaptionContainer"],
:root[data-theme="light"] [data-testid="stCaptionContainer"] p {
    color: var(--muted) !important;
}

:root[data-theme="light"] [data-testid="stWidgetLabel"],
:root[data-theme="light"] [data-testid="stWidgetLabel"] p {
    color: var(--text-2) !important;
}

/* Rows, Columns, Missing cells and Duplicate rows */
:root[data-theme="light"] [data-testid="stMetricLabel"],
:root[data-theme="light"] [data-testid="stMetricLabel"] p {
    color: var(--muted) !important;
}

:root[data-theme="light"] [data-testid="stMetricValue"],
:root[data-theme="light"] [data-testid="stMetricValue"] div {
    color: #ffffff !important;
}

/* Text inputs, text areas and selected values */
:root[data-theme="light"] [data-testid="stTextInput"] input,
:root[data-theme="light"] [data-testid="stTextArea"] textarea,
:root[data-theme="light"] [data-baseweb="select"] span {
    color: var(--text) !important;
    -webkit-text-fill-color: var(--text);
}

:root[data-theme="light"] [data-testid="stTextInput"] input::placeholder,
:root[data-theme="light"] [data-testid="stTextArea"] textarea::placeholder {
    color: var(--muted) !important;
    -webkit-text-fill-color: var(--muted);
}

/* Tabs */
:root[data-theme="light"] button[data-baseweb="tab"] {
    color: var(--muted) !important;
}

:root[data-theme="light"] button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--mint-300) !important;
}

/* Links remain mint-colored against the dark green background */
:root[data-theme="light"] [data-testid="stMarkdownContainer"]:not(:has(.hero)) a,
:root[data-theme="light"] [data-testid="stCaptionContainer"] a {
    color: var(--mint-300) !important;
}

/* Expanders and status messages */
:root[data-theme="light"] [data-testid="stExpander"],
:root[data-theme="light"] [data-testid="stExpander"] summary,
:root[data-theme="light"] [data-testid="stAlert"] p {
    color: var(--text) !important;
}

/* Keep the primary button text readable */
:root[data-theme="light"] .stButton > button[kind="primary"],
:root[data-theme="light"] .stButton > button[data-testid="stBaseButton-primary"] {
    color: #04140d !important;
}

/* Keep keyboard focus clearly visible */
:root[data-theme="light"] :is(button, input, textarea, select):focus-visible {
    outline-color: var(--emerald-400);
}

/* Labels inside buttons and tabs follow the button colour (not the forced paragraph colour) */
:root[data-theme="light"] .stButton > button[kind="primary"] p,
:root[data-theme="light"] .stButton > button[data-testid="stBaseButton-primary"] p {
    color: #04140d !important;
}

:root[data-theme="light"] button[data-baseweb="tab"] p {
    color: inherit !important;
}
/* Keep the hero box compact on mobile */
@media (max-width: 640px) {
    .hero {
        padding: 1rem !important;
        margin-top: 0 !important;
        margin-bottom: 0.5rem !important;
    }

    .hero-brand {
        gap: 10px !important;
        align-items: center !important;
    }

    .hero-logo {
        width: 46px !important;
        height: 46px !important;
        font-size: 1.5rem !important;
    }

    .hero h1 {
        font-size: 1.55rem !important;
        line-height: 1.15 !important;
    }

    .hero p {
        font-size: 0.82rem !important;
        line-height: 1.35 !important;
    }

    .hero-badge {
        font-size: 0.65rem !important;
        padding: 6px 9px !important;
    }
}
</style>

<div class="hero">
    <div class="hero-brand">
        <div class="hero-logo">📊</div>
        <div>
            <h1>Excel <span>AI Studio</span></h1>
            <p>
                Turn spreadsheets into insights.
                Clean data, explore trends, create charts
                and analyse your workbook with AI.
            </p>
            <div class="hero-badge">
                ✦ SMART DATA WORKSPACE &nbsp; • &nbsp; AI-POWERED
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

SAMPLE = pd.DataFrame({
    "Month": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    "Region": ["North", "South", "East", "West"] * 3,
    "Sales": [9200, 10150, 11400, 11900, 13200, 14500, 15100, 16400, 17100, 18400, 19300, 21100],
    "Cost": [6100, 6500, 7200, 7600, 8300, 9100, 9400, 10100, 10600, 11200, 11900, 12700],
    "Units": [182, 195, 214, 226, 245, 269, 281, 297, 310, 329, 344, 371],
    "Returns": [145, 182, 121, 210, 188, 176, 230, 201, 192, 175, 164, 153],
})

if "sheets" not in st.session_state:
    st.session_state.sheets = {"Sales": SAMPLE.copy()}
    st.session_state.active_sheet = "Sales"
if "history" not in st.session_state: st.session_state.history = []
if "ai_log" not in st.session_state: st.session_state.ai_log = []
if "last_analysis_report" not in st.session_state:
    st.session_state.last_analysis_report = ""

if "chart_spec" not in st.session_state: st.session_state.chart_spec = None

sheets = st.session_state.sheets
active_sheet = st.session_state.active_sheet
if active_sheet not in sheets:
    active_sheet = next(iter(sheets))
    st.session_state.active_sheet = active_sheet

def df_now() -> pd.DataFrame:
    return st.session_state.sheets[st.session_state.active_sheet]

def set_df(df: pd.DataFrame, label: str = "Updated data"):
    st.session_state.history.append((st.session_state.active_sheet, df_now().copy(), label))
    st.session_state.sheets[st.session_state.active_sheet] = df.reset_index(drop=True)

def log(role: str, msg: str):
    st.session_state.ai_log.append({"role": role, "message": str(msg)})

def to_excel_bytes(sheet_map: dict[str, pd.DataFrame]) -> bytes:
    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        for name, data in sheet_map.items():
            safe_name = str(name)[:31] or "Sheet1"
            data.to_excel(writer, sheet_name=safe_name, index=False)
    return buff.getvalue()

def analysis_excel_bytes(report_text, df, sheet_name, include_data, include_stats, include_quality):
    from datetime import datetime

    buff = io.BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        overview = pd.DataFrame({
            "Field": ["Report Date", "Active Sheet", "Rows", "Columns"],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                sheet_name, len(df), len(df.columns)
            ]
        })
        overview.to_excel(writer, sheet_name="Overview", index=False)

        pd.DataFrame({
            "Analysis Report": report_text.splitlines()
        }).to_excel(writer, sheet_name="Analysis Report", index=False)

        if include_stats:
            numeric = df.select_dtypes(include="number")
            if not numeric.empty:
                numeric.describe().round(3).to_excel(
                    writer, sheet_name="Statistics"
                )

        if include_quality:
            quality = pd.DataFrame({
                "Column": df.columns.astype(str),
                "Data Type": df.dtypes.astype(str).values,
                "Missing Values": df.isna().sum().values,
                "Missing %": (df.isna().mean().values * 100).round(2),
                "Unique Values": df.nunique().values
            })
            quality.to_excel(writer, sheet_name="Data Quality", index=False)

        if include_data:
            df.to_excel(writer, sheet_name="Dataset"[:31], index=False)

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                max_length = max(
                    (len(str(cell.value or "")) for cell in column_cells),
                    default=10
                )
                ws.column_dimensions[column_cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 45
                )
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)

    return buff.getvalue()


def analysis_pdf_bytes(report_text, df, sheet_name, include_stats, include_quality):
    from datetime import datetime
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    )

    buff = io.BytesIO()
    doc = SimpleDocTemplate(
        buff, pagesize=A4,
        rightMargin=0.55 * inch, leftMargin=0.55 * inch,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Excel AI Studio — Analysis Report", styles["Title"]),
        Paragraph(f"Generated: {datetime.now():%Y-%m-%d %H:%M}", styles["Normal"]),
        Paragraph(f"Sheet: {escape(str(sheet_name))}", styles["Normal"]),
        Paragraph(f"Dataset: {len(df)} rows × {len(df.columns)} columns", styles["Normal"]),
        Spacer(1, 14),
        Paragraph("Analysis Results", styles["Heading2"])
    ]

    for line in report_text.splitlines():
        safe_line = escape(line.strip())
        if safe_line:
            story.append(Paragraph(safe_line[:3000], styles["BodyText"]))
            story.append(Spacer(1, 4))

    if include_quality:
        story.extend([
            Spacer(1, 10),
            Paragraph("Data Quality", styles["Heading2"])
        ])
        rows = [["Column", "Type", "Missing", "Unique"]]
        for col in df.columns:
            rows.append([
                str(col)[:35], str(df[col].dtype),
                str(int(df[col].isna().sum())), str(int(df[col].nunique()))
            ])
        table = Table(rows, repeatRows=1, colWidths=[2.5*inch, 1.1*inch, 0.9*inch, 0.9*inch])
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#263b59")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#edf2f7")])
        ]))
        story.append(table)

    if include_stats:
        numeric = df.select_dtypes(include="number")
        if not numeric.empty:
            story.extend([
                Spacer(1, 10),
                Paragraph("Summary Statistics", styles["Heading2"])
            ])
            stats = numeric.describe().round(2).reset_index()
            rows = [[str(c) for c in stats.columns]] + [
                [str(v)[:35] for v in row] for row in stats.values.tolist()
            ]
            table = Table(rows, repeatRows=1)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#263b59")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP")
            ]))
            story.append(table)

    def add_page_number(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(A4[0] - 0.55*inch, 0.3*inch, f"Page {document.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
    return buff.getvalue()

def numeric_columns(df: pd.DataFrame) -> list[str]:
    return list(df.select_dtypes(include="number").columns)

def style_figure(fig):
    """Apply the emerald visual theme to a Plotly figure (appearance only)."""
    grid = "rgba(167, 243, 208, 0.08)"
    axis_line = "rgba(167, 243, 208, 0.20)"
    fig.update_layout(
        template="plotly_dark",
        colorway=CHART_COLORWAY,
        paper_bgcolor="#0b1a15",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, 'Segoe UI', system-ui, -apple-system, sans-serif", size=13, color="#c9ddd4"),
        title_font=dict(size=18, color="#ecf6f1"),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=12, color="#c9ddd4"),
                    title_font=dict(color="#c9ddd4")),
        coloraxis_colorbar=dict(tickfont=dict(color="#c9ddd4"), title_font=dict(color="#c9ddd4")),
        hoverlabel=dict(bgcolor="#0e1e1a", bordercolor="#34d399", font=dict(size=12, color="#ecf6f1")),
        margin=dict(l=20, r=20, t=60, b=20),
    )
    tick = dict(size=12, color="#c9ddd4")
    axis_title = dict(size=13, color="#c9ddd4")
    fig.update_xaxes(gridcolor=grid, zerolinecolor=axis_line, linecolor=axis_line,
                     tickfont=tick, title_font=axis_title)
    fig.update_yaxes(gridcolor=grid, zerolinecolor=axis_line, linecolor=axis_line,
                     tickfont=tick, title_font=axis_title)
    return fig

def get_ai_plan(prompt: str, df: pd.DataFrame, api_key: str, model: str) -> dict[str, Any]:
    """Ask Gemini for a constrained JSON action plan; only allow-listed actions are executed."""
    context = {
        "columns": [{"name": str(c), "dtype": str(df[c].dtype)} for c in df.columns],
        "sample_rows": df.head(8).replace({float("nan"): None}).to_dict(orient="records"),
        "row_count": int(len(df)),
    }

    system = """You are an Excel data assistant. Return ONLY valid JSON with exactly these top-level keys: "say" and "actions".

CRITICAL ACTION FORMAT:
Every item in "actions" MUST be a JSON object containing a non-empty "action" field.
Correct example: {"say":"Calculating statistics","actions":[{"action":"sum","columns":["Sales"]},{"action":"mean","columns":["Sales"]},{"action":"min","columns":["Sales"]},{"action":"max","columns":["Sales"]},{"action":"median","columns":["Sales"]}]}
Incorrect example: {"say":"Calculating","actions":[{"columns":["Sales"]},{"action":"sum"}]}

Available actions and schemas:
- sort: {"action":"sort","column":"existing column name","ascending":true}
- drop_duplicates: {"action":"drop_duplicates"}
- drop_missing: {"action":"drop_missing","how":"any"}
- fill_missing: {"action":"fill_missing","value":"replacement value"}
- rename_column: {"action":"rename_column","old":"existing name","new":"new name"}
- drop_column: {"action":"drop_column","column":"existing column name"}
- filter: {"action":"filter","column":"existing column name","operator":"eq","value":"value"}
- add_calculated_column: {"action":"add_calculated_column","name":"new column","expression":"Sales - Cost"}
- chart: {"action":"chart","x":"existing column","ys":["numeric column"],"type":"bar"}
- describe: {"action":"describe"}
- sum: {"action":"sum","columns":["numeric column names"]}
- mean: {"action":"mean","columns":["numeric column names"]}
- median: {"action":"median","columns":["numeric column names"]}
- min: {"action":"min","columns":["numeric column names"]}
- max: {"action":"max","columns":["numeric column names"]}
- std: {"action":"std","columns":["numeric column names"]}
- data_quality: {"action":"data_quality"}

For totals use sum. For averages use mean. If the user asks for every numeric column, omit "columns" or include all actual numeric column names. For full data analysis, use data_quality and describe as appropriate.

Use only column names that exist in the supplied data context. Never invent values or columns. For charts, allowed types are bar, line, area, scatter, pie, histogram, and box. For calculations, use only column names, numbers, and basic arithmetic operators.

Return an empty actions array only if the request cannot be performed with the available actions. Ensure every action object includes its "action" field."""

    import time

    payload = {
        "systemInstruction": {"parts": [{"text": system}]},
        "contents": [{"role": "user", "parts": [{"text": "DATA CONTEXT:\n" + json.dumps(context, default=str) + "\n\nUSER REQUEST:\n" + prompt}]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0.1}
    }

    # Try the chosen model first; if it is rate-limited, fall back to lighter models.
    fallback_models = ["gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]
    models_to_try = [model] + [m for m in fallback_models if m != model]

    for m in models_to_try:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"
        for attempt in range(2):
            resp = requests.post(url, params={"key": api_key}, json=payload, timeout=45)
            if resp.status_code in (429, 503):
                time.sleep(3 * (attempt + 1))  # short wait, then retry once
                continue
            if resp.status_code == 404 and m != model:
                break  # fallback model not available, try the next one
            resp.raise_for_status()
            data = resp.json()
            raw = data["candidates"][0]["content"]["parts"][0]["text"]
            return json.loads(raw)

    raise RuntimeError(
        "Gemini quota or rate limit reached on all models. "
        "Please wait about a minute and try again, or check your quota in Google AI Studio."
    )

def safe_calculation(df: pd.DataFrame, expression: str) -> pd.Series:
    # Parse a restricted arithmetic expression; never use Python eval on model output.
    expr = expression.strip()
    aliases = {}
    for i, col in enumerate(sorted(map(str, df.columns), key=len, reverse=True)):
        alias = f"COL{i}"
        replaced, count = re.subn(rf"(?<![\w]){re.escape(col)}(?![\w])", alias, expr)
        if count:
            expr = replaced
            aliases[alias] = pd.to_numeric(df[col], errors="coerce")
    tree = ast.parse(expr, mode="eval")
    binary_ops = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
                  ast.Div: operator.truediv, ast.Mod: operator.mod, ast.Pow: operator.pow}
    unary_ops = {ast.UAdd: operator.pos, ast.USub: operator.neg}
    def visit(node):
        if isinstance(node, ast.Expression): return visit(node.body)
        if isinstance(node, ast.Name) and node.id in aliases: return aliases[node.id]
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)): return node.value
        if isinstance(node, ast.BinOp) and type(node.op) in binary_ops:
            return binary_ops[type(node.op)](visit(node.left), visit(node.right))
        if isinstance(node, ast.UnaryOp) and type(node.op) in unary_ops:
            return unary_ops[type(node.op)](visit(node.operand))
        raise ValueError("Use only column names, numbers, and +, -, *, /, %, or **.")
    result = visit(tree)
    if not isinstance(result, pd.Series): result = pd.Series([result] * len(df))
    return result


def execute_actions(plan: dict[str, Any], df: pd.DataFrame) -> tuple[pd.DataFrame, list[str], dict | None]:
    messages, chart = [], None

    for action in plan.get("actions", []):
        kind = str(action.get("action", "")).lower()

        try:
            if kind == "sort":
                col = action["column"]
                if col not in df.columns:
                    raise ValueError(f"Column not found: {col}")
                df = df.sort_values(
                    col,
                    ascending=bool(action.get("ascending", True)),
                    kind="stable"
                )
                messages.append(f"Sorted by {col}.")

            elif kind == "drop_duplicates":
                before = len(df)
                df = df.drop_duplicates()
                messages.append(f"Removed {before - len(df)} duplicate row(s).")

            elif kind == "drop_missing":
                before = len(df)
                df = df.dropna(how=action.get("how", "any"))
                messages.append(f"Removed {before - len(df)} row(s) with missing values.")

            elif kind == "fill_missing":
                value = action.get("value", "")
                df = df.fillna(value)
                messages.append(f"Filled missing cells with {value!r}.")

            elif kind == "rename_column":
                old, new = action["old"], action["new"]
                if old not in df.columns:
                    raise ValueError(f"Column not found: {old}")
                df = df.rename(columns={old: new})
                messages.append(f"Renamed {old} to {new}.")

            elif kind == "drop_column":
                col = action["column"]
                if col not in df.columns:
                    raise ValueError(f"Column not found: {col}")
                df = df.drop(columns=[col])
                messages.append(f"Removed column {col}.")

            elif kind == "filter":
                col = action["column"]
                op = action.get("operator", "eq")
                val = action.get("value")

                if col not in df.columns:
                    raise ValueError(f"Column not found: {col}")

                s = df[col]
                if op in ("gt", "ge", "lt", "le"):
                    s = pd.to_numeric(s, errors="coerce")
                    val = float(val)

                operations = {
                    "eq": lambda: s.astype(str) == str(val),
                    "ne": lambda: s.astype(str) != str(val),
                    "gt": lambda: s > val,
                    "ge": lambda: s >= val,
                    "lt": lambda: s < val,
                    "le": lambda: s <= val,
                    "contains": lambda: s.astype(str).str.contains(
                        str(val), case=False, na=False
                    ),
                }

                if op not in operations:
                    raise ValueError(f"Unsupported filter operator: {op}")

                df = df[operations[op]()]
                messages.append(f"Filtered {col} ({op} {val}).")

            elif kind == "add_calculated_column":
                name = action["name"]
                expr = action["expression"]
                df[name] = safe_calculation(df, expr)
                messages.append(f"Created calculated column {name}.")

            elif kind == "chart":
                chart = {
                    "x": action.get("x"),
                    "ys": action.get("ys", []),
                    "type": action.get("type", "bar"),
                }
                messages.append("Prepared a chart from your request.")

            elif kind == "describe":
                if df.empty:
                    messages.append("The spreadsheet has no data rows.")
                else:
                    messages.append(
                        "Summary statistics for numeric columns:\n"
                        + df.describe().to_string()
                    )

            elif kind in ("sum", "mean", "median", "min", "max", "std"):
                cols = action.get("columns") or numeric_columns(df)
                if isinstance(cols, str):
                    cols = [cols]

                cols = [
                    c for c in cols
                    if c in df.columns
                    and pd.api.types.is_numeric_dtype(df[c])
                ]

                if not cols:
                    raise ValueError("No numeric columns matched.")

                result = getattr(df[cols], kind)()
                messages.append(
                    f"{kind.upper() if kind == 'std' else kind.title()} "
                    f"for each numeric column:\n{result.to_string()}"
                )

            elif kind == "data_quality":
                report = pd.DataFrame({
                    "Metric": [
                        "Total rows",
                        "Total columns",
                        "Duplicate rows",
                        "Total missing cells",
                        "Columns containing missing values",
                    ],
                    "Value": [
                        len(df),
                        len(df.columns),
                        int(df.duplicated().sum()),
                        int(df.isna().sum().sum()),
                        int(df.isna().any().sum()),
                    ],
                })

                messages.append(
                    "Data quality report:\n" + report.to_string(index=False)
                )
                messages.append(
                    "Missing values by column:\n"
                    + df.isna().sum().to_string()
                )

                numeric = numeric_columns(df)
                if numeric:
                    messages.append(
                        "Numeric-column insights:\n"
                        + df[numeric].describe().to_string()
                    )

            else:
                messages.append(
                    f"Skipped unsupported action: {kind or '(empty)'}"
                )

        except Exception as exc:
            messages.append(f"Could not complete {kind or 'action'}: {exc}")

    return df.reset_index(drop=True), messages, chart


# Sidebar: data input and settings
with st.sidebar:
    st.subheader("📂 Data")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "xls", "csv", "tsv"])
    if uploaded is not None and st.button("Load uploaded file", use_container_width=True, type="primary"):
        try:
            if uploaded.name.lower().endswith((".csv", ".tsv")):
                sep = "\t" if uploaded.name.lower().endswith(".tsv") else ","
                st.session_state.sheets = {os.path.splitext(uploaded.name)[0][:31]: pd.read_csv(uploaded, sep=sep)}
            else:
                book = pd.ExcelFile(uploaded)
                st.session_state.sheets = {name: pd.read_excel(book, sheet_name=name) for name in book.sheet_names}
            st.session_state.active_sheet = next(iter(st.session_state.sheets))
            st.session_state.history = []
            st.success("File loaded.")
            st.rerun()
        except Exception as exc:
            st.error(f"Could not read file: {exc}")
    a, b = st.columns(2)
    if a.button("Load sample", use_container_width=True):
        st.session_state.sheets = {"Sales": SAMPLE.copy()}; st.session_state.active_sheet = "Sales"; st.session_state.history = []; st.rerun()
    if b.button("New sheet", use_container_width=True):
        st.session_state.sheets = {"Sheet1": pd.DataFrame(columns=["Column 1", "Column 2"])}; st.session_state.active_sheet = "Sheet1"; st.session_state.history = []; st.rerun()
    st.divider()
    st.subheader("✨ AI settings")

    api_key = os.getenv("GEMINI_API_KEY", "").strip()

    try:
        secret_key = st.secrets["GEMINI_API_KEY"]
        if secret_key:
            api_key = str(secret_key).strip()
    except Exception:
        pass

    model = st.text_input(
        "Gemini model",
        value="gemini-3.8-flash"
    )

    st.caption(
        "Without an API key, the built-in command helper "
        "still supports common data tasks."
    )
    st.divider()
    st.caption(
        "Privacy tip: avoid sending confidential spreadsheet "
        "rows to external AI services."
    )

# choose sheet
if len(st.session_state.sheets) > 1:
    names = list(st.session_state.sheets)
    selected = st.selectbox("Workbook sheet", names, index=names.index(st.session_state.active_sheet) if st.session_state.active_sheet in names else 0)
    if selected != st.session_state.active_sheet:
        st.session_state.active_sheet = selected
        st.rerun()

df = df_now()
# metrics
m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows", f"{len(df):,}")
m2.metric("Columns", f"{len(df.columns):,}")
m3.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")
m4.metric("Duplicate rows", f"{int(df.duplicated().sum()):,}")

main_tab, chart_tab, clean_tab, analyze_tab, ai_tab, export_tab = st.tabs(["🧾 Spreadsheet", "📈 Charts", "🧹 Clean & Edit", "🔎 Analyze", "🤖 AI Assistant", "⬇️ Export"])

with main_tab:
    top1, top2 = st.columns([4, 1])
    with top1: st.markdown("### Spreadsheet editor")
    with top2:
        if st.button("↩ Undo last change", use_container_width=True, disabled=not st.session_state.history):
            sheet_name, previous, label = st.session_state.history.pop()
            st.session_state.sheets[sheet_name] = previous
            st.toast(f"Undid: {label}")
            st.rerun()
    st.caption("Edit cells directly in the table. Changes are applied when you click outside the edited cell.")
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, hide_index=True, key=f"editor_{active_sheet}", height=420)
    if not edited.equals(df):
        st.session_state.history.append((active_sheet, df.copy(), "Edited spreadsheet"))
        st.session_state.sheets[active_sheet] = edited.reset_index(drop=True)
        st.toast("Spreadsheet changes applied")
        df = df_now()
    search = st.text_input("🔍 Search rows", placeholder="Type to find a value in any column…", key="grid_search")
    if search:
        mask = df.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        st.caption(f"Showing {int(mask.sum())} matching rows out of {len(df)}")
        st.dataframe(df[mask], use_container_width=True, hide_index=True)

with chart_tab:
    st.markdown("### Build a chart")
    if df.empty or len(df.columns) == 0:
        st.info("Load some data to build a chart.")
    else:
        all_cols = list(df.columns); nums = numeric_columns(df)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1: chart_type = st.selectbox("Chart type", ["Bar", "Line", "Area", "Scatter", "Pie", "Histogram", "Box"], index=0)
        with c2: xcol = st.selectbox("Category / X-axis", all_cols, index=0)
        with c3: ycols = st.multiselect("Value column(s)", nums, default=nums[:min(2, len(nums))])
        color_col = st.selectbox("Optional group / color", ["(none)"] + all_cols, index=0)
        title = st.text_input("Chart title", value=f"{chart_type} chart")
        if st.button("Generate chart", type="primary"):
            st.session_state.chart_spec = {"type": chart_type, "x": xcol, "ys": ycols, "color": None if color_col == "(none)" else color_col, "title": title}
        spec = st.session_state.chart_spec
        if spec:
            try:
                chart_df = df.copy()
                kind = spec["type"]
                if kind == "Histogram":
                    col = (spec.get("ys") or nums or all_cols)[0]
                    fig = px.histogram(chart_df, x=col, title=spec.get("title", "Histogram"))
                elif kind == "Box":
                    col = (spec.get("ys") or nums or all_cols)[0]
                    fig = px.box(chart_df, x=spec.get("x"), y=col, color=spec.get("color"), title=spec.get("title"))
                elif kind == "Pie":
                    y = (spec.get("ys") or nums)
                    if not y: raise ValueError("Choose a numeric value column for the pie chart.")
                    fig = px.pie(chart_df, names=spec["x"], values=y[0], title=spec.get("title"))
                else:

                    ys = [c for c in spec.get("ys", []) if c in chart_df.columns]
                    if not ys:
                        raise ValueError("Choose at least one numeric value column.")

                    fn = {"Bar": px.bar, "Line": px.line, "Area": px.area, "Scatter": px.scatter}[kind]

                    if kind == "Line":
                        fig = fn(
                            chart_df,
                            x=spec["x"],
                            y=ys,
                            color=spec.get("color"),
                            title=spec.get("title"),
                            markers=True
                        )
                    else:
                        fig = fn(
                            chart_df,
                            x=spec["x"],
                            y=ys,
                            color=spec.get("color"),
                            title=spec.get("title")
                        )

                style_figure(fig)

                st.plotly_chart(fig, use_container_width=True, theme=None)
                st.download_button("Download chart data (CSV)", chart_df.to_csv(index=False).encode("utf-8-sig"), file_name="chart_data.csv", mime="text/csv")

                st.download_button(
                  "⬇ Download Interactive Chart (.html)",
                  data=fig.to_html(full_html=True, include_plotlyjs=True),
                  file_name="excel_ai_chart.html",
                  mime="text/html",
                  use_container_width=True
               )

            except Exception as exc: st.warning(str(exc))
        st.caption("Tip: select multiple value columns to compare several metrics in one chart.")

with clean_tab:
    st.markdown("### Clean, transform, and calculate")
    left, right = st.columns(2)
    with left:
        st.markdown("**Missing values**")
        missing_mode = st.selectbox("Action", ["Drop rows with any missing cells", "Drop rows only when all cells are missing", "Fill missing cells with a value"])
        fill_value = st.text_input("Fill value", value="0")
        if st.button("Apply missing-value action"):
            new = df.copy()
            if missing_mode.startswith("Drop rows with any"): new = new.dropna(how="any")
            elif missing_mode.startswith("Drop rows only"): new = new.dropna(how="all")
            else:
                value = fill_value
                try: value = float(fill_value)
                except ValueError: pass
                new = new.fillna(value)
            set_df(new, "Missing-value cleanup"); st.success("Cleanup applied."); st.rerun()
        if st.button("Remove duplicate rows"):
            set_df(df.drop_duplicates(), "Remove duplicates"); st.success("Duplicate rows removed."); st.rerun()
    with right:
        st.markdown("**Column tools**")
        if len(df.columns):
            col = st.selectbox("Column", list(df.columns), key="clean_col")
            new_name = st.text_input("Rename to", value=str(col))
            if st.button("Rename column"):
                if new_name.strip() and new_name != col:
                    set_df(df.rename(columns={col: new_name.strip()}), "Rename column"); st.rerun()
            if st.button("Delete selected column"):
                set_df(df.drop(columns=[col]), "Delete column"); st.rerun()
        st.markdown("**Add a calculated column**")
        calc_name = st.text_input("New column name", value="Profit")
        if len(df.columns):
            calc_left = st.selectbox("First numeric column", numeric_columns(df) or list(df.columns), key="calc_left")
            calc_op = st.selectbox("Operation", ["minus (−)", "plus (+)", "multiply (×)", "divide (÷)"], key="calc_op")
            calc_right = st.selectbox("Second column", numeric_columns(df) or list(df.columns), key="calc_right")
            if st.button("Create calculated column"):
                a = pd.to_numeric(df[calc_left], errors="coerce")
                b = pd.to_numeric(df[calc_right], errors="coerce")
                if calc_op.startswith("minus"): result = a - b
                elif calc_op.startswith("plus"): result = a + b
                elif calc_op.startswith("multiply"): result = a * b
                else: result = a / b.replace(0, float("nan"))
                new = df.copy(); new[calc_name.strip() or "Calculated"] = result
                set_df(new, "Add calculated column"); st.success("Calculated column created."); st.rerun()

with analyze_tab:
    st.markdown("### Data overview")
    if df.empty: st.info("No rows to analyze yet.")
    else:
        st.markdown("**Column profile**")
        profile = pd.DataFrame({"Column": df.columns.astype(str), "Type": [str(x) for x in df.dtypes], "Non-empty": [int(df[c].notna().sum()) for c in df.columns], "Missing": [int(df[c].isna().sum()) for c in df.columns], "Unique values": [int(df[c].nunique(dropna=True)) for c in df.columns]})
        st.dataframe(profile, use_container_width=True, hide_index=True)
        nums = numeric_columns(df)
        if nums:
            st.markdown("**Summary statistics**")
            st.dataframe(df[nums].describe().T, use_container_width=True)
            st.markdown("**Quick calculations**")
            chosen = st.multiselect("Numeric columns", nums, default=nums)
            calc_kind = st.selectbox("Calculation", ["Sum", "Mean", "Median", "Minimum", "Maximum", "Standard deviation"])
            funcs = {"Sum":"sum", "Mean":"mean", "Median":"median", "Minimum":"min", "Maximum":"max", "Standard deviation":"std"}
            if chosen: st.dataframe(df[chosen].agg(funcs[calc_kind]).rename("Result").to_frame(), use_container_width=True)
        st.markdown("**Correlation matrix**")
        if len(nums) >= 2:
            corr = df[nums].corr(numeric_only=True)
            fig = px.imshow(corr, text_auto=True, color_continuous_scale="RdBu_r", zmin=-1, zmax=1, title="Numeric column correlations")
            style_figure(fig)
            st.plotly_chart(fig, use_container_width=True, theme=None)
        else: st.caption("Add at least two numeric columns to view correlations.")

with ai_tab:
    st.markdown("### Tell the app what to do")
    st.caption("The app executes supported actions on the active sheet. For more flexible prompts, add a Gemini API key in the sidebar.")
    examples = ["Sort Sales from highest to lowest", "Remove duplicate rows", "Create a Profit column = Sales - Cost", "Make a bar chart of Sales by Region", "Show summary statistics"]
    cols = st.columns(len(examples))
    for i, ex in enumerate(examples):
        with cols[i]:
            if st.button(ex, key=f"example_{i}", use_container_width=True): st.session_state.ai_prompt = ex
    prompt = st.text_area("Your instruction", key="ai_prompt", placeholder="e.g. Sort Sales from highest to lowest and create a Profit column = Sales - Cost", height=100)
    if st.button("Run instruction", type="primary"):
        if not prompt.strip(): st.warning("Enter an instruction first.")
        elif df.empty: st.warning("Load or create data first.")
        else:
            try:
                if api_key:
                    plan = get_ai_plan(prompt, df, api_key, model.strip() or "gemini-3.8-flash")
                else:
                    p = prompt.lower(); actions = []
                    if "duplicate" in p: actions.append({"action":"drop_duplicates"})
                    elif "missing" in p and ("drop" in p or "remove" in p): actions.append({"action":"drop_missing", "how":"any"})
                    elif "sort" in p:
                        match = next((str(c) for c in df.columns if str(c).lower() in p), None)
                        if match: actions.append({"action":"sort", "column":match, "ascending": not any(w in p for w in ["descending", "highest to lowest", "high to low", "largest first"])})
                    elif "profit" in p and "sales" in [str(c).lower() for c in df.columns] and "cost" in [str(c).lower() for c in df.columns]:
                        sales = next(c for c in df.columns if str(c).lower()=="sales"); cost = next(c for c in df.columns if str(c).lower()=="cost")
                        actions.append({"action":"add_calculated_column", "name":"Profit", "expression":f"{sales} - {cost}"})
                    elif any(w in p for w in ["summary", "statistics", "describe", "analyze"]): actions.append({"action":"describe"})
                    elif "chart" in p or "plot" in p:
                        nums = numeric_columns(df); x = next((c for c in df.columns if str(c).lower() in p and c not in nums), str(df.columns[0]) if len(df.columns) else None)
                        ys = [c for c in nums if str(c).lower() in p] or nums[:1]
                        actions.append({"action":"chart", "x":x, "ys":ys, "type":"line" if "line" in p else "pie" if "pie" in p else "bar"})
                    plan = {"say": "Using the built-in command helper. Add a Gemini API key for more flexible natural-language requests.", "actions": actions}
                log("user", prompt)
                updated, messages, chart_action = execute_actions(plan, df.copy())
                changed = not updated.equals(df)
                if changed: set_df(updated, "AI instruction")
                log("assistant", plan.get("say", "Completed request."))
                for msg in messages: log("assistant", msg)
                st.session_state.last_analysis_report = (
                    f"EXCEL AI STUDIO - ANALYSIS REPORT\n"
                    f"{'=' * 42}\n\n"
                    f"Instruction: {prompt}\n"
                    f"Active sheet: {active_sheet}\n"
                    f"Rows: {len(updated)}\n"
                    f"Columns: {len(updated.columns)}\n\n"
                    + "\n\n".join(str(msg) for msg in messages)
                )

                if chart_action:
                    st.session_state.chart_spec = {"type": str(chart_action.get("type", "bar")).title(), "x": chart_action.get("x") or (list(df.columns)[0] if len(df.columns) else None), "ys": chart_action.get("ys") or numeric_columns(df)[:1], "color": None, "title": f"{chart_action.get('type','bar').title()} chart from AI"}
                    st.info("Chart prepared. Open the Charts tab to view it.")
                if not messages and not plan.get("actions"): st.warning("I couldn't map that request to an action. Try one of the examples or configure Gemini.")
                else:
                    st.success("Instruction processed. Review the results below and check the Spreadsheet tab for changes.")

                    for msg in messages:
                        if ":" in msg and "\n" in msg:
                            heading, details = msg.split(":", 1)
                            st.markdown(f"**{heading.strip()}**")
                            try:
                                table = pd.read_csv(
                                    io.StringIO(details.strip()),
                                    sep=r"\s{2,}",
                                    engine="python"
                                )
                                if table.shape[1] > 1:
                                    st.dataframe(
                                        table,
                                        use_container_width=True,
                                        hide_index=True
                                    )
                                else:
                                    st.code(details.strip(), language=None)
                            except Exception:
                                st.code(details.strip(), language=None)
                        else:
                            st.markdown(f"- {msg}")

            except Exception as exc:
                log("error", str(exc)); st.error(f"AI request failed: {exc}")


if st.session_state.last_analysis_report:
    st.markdown("### 📄 Export Analysis Report")
    st.caption("Choose the sections you want included in your exports.")

    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        include_data = st.checkbox("Include dataset", value=False)
    with rc2:
        include_stats = st.checkbox("Include statistics", value=True)
    with rc3:
        include_quality = st.checkbox("Include data quality", value=True)

    report_text = st.session_state.last_analysis_report
    current_df = df_now()
    current_name = st.session_state.active_sheet

    try:
        excel_report = analysis_excel_bytes(
            report_text, current_df, current_name,
            include_data, include_stats, include_quality
        )
        pdf_report = analysis_pdf_bytes(
            report_text, current_df, current_name,
            include_stats, include_quality
        )

        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button(
                "⬇ Excel Report (.xlsx)",
                data=excel_report,
                file_name="excel_ai_analysis_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with d2:
            st.download_button(
                "⬇ PDF Report",
                data=pdf_report,
                file_name="excel_ai_analysis_report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with d3:
            st.download_button(
                "⬇ Text Report (.txt)",
                data=report_text,
                file_name="excel_ai_analysis_report.txt",
                mime="text/plain",
                use_container_width=True
            )
    except Exception as exc:
        st.error(f"Could not generate report: {exc}")

    if st.session_state.ai_log:
        with st.expander("AI activity log", expanded=True):
            for item in st.session_state.ai_log[-20:]:
                st.markdown(f"**{item['role'].title()}:** {item['message']}")
            if st.button("Clear activity log"):
                st.session_state.ai_log = []; st.rerun()

with export_tab:
    st.markdown("### Export your workbook")
    st.write("Download the edited active sheet as CSV, or all currently loaded sheets as an Excel workbook.")
    ec1, ec2 = st.columns(2)
    with ec1:
        st.download_button("⬇ Download active sheet as CSV", df.to_csv(index=False).encode("utf-8-sig"), file_name=f"{active_sheet}.csv", mime="text/csv", use_container_width=True)
    with ec2:
        st.download_button("⬇ Download workbook as Excel", to_excel_bytes(st.session_state.sheets), file_name="excel_ai_studio_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.divider()
    st.markdown("**Before exporting**")
    st.write(f"- Active sheet: **{active_sheet}**\n- Rows: **{len(df):,}**\n- Columns: **{len(df.columns):,}**\n- Loaded sheets: **{len(st.session_state.sheets)}**")

st.markdown("<p class='small-note'>Excel AI Studio · Streamlit edition · Your spreadsheet stays in this app session unless you explicitly send context to Gemini.</p>", unsafe_allow_html=True)
