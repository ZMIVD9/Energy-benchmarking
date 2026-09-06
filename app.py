"""
Energy Model Benchmarking & QA/QC Platform
===========================================
Run:  streamlit run app.py
Deps: pip install streamlit pandas plotly openpyxl reportlab

IMPORTANT: Keep benchmarks.xlsx in the same folder as app.py

Branding
--------
The interface follows the Stantec visual identity and design guidelines:

  Colour       Wit / Mist / Kiezelsteen / Fossiel carry the surfaces, Zwart
               carries the type, and Stantec-oranje (#ED6631) is the house
               colour used for primary actions, accents and the Lens-tag.
               Secondary colours (Fern, Marine, Kurkuma, Himalayazout) are
               reserved for status and data visualisation.
  Type         Source Serif 4 for headings and figures, Roboto / Roboto
               Condensed for body copy, labels and UI chrome.
  Contrast     Text is black or white only. Type on orange is black, per the
               WCAG table in the guidelines.

Logo: place the approved artwork at ``assets/stantec_logo.svg`` (or .png)
next to this file. The header loads it automatically; until then a
Lens-tag style orange placeholder is shown.
"""

import io
import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import date

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Energy Benchmarking Platform",
    page_icon="🟧",   # Lens-tag stand-in; swap for assets/stantec_logo.png if preferred
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,opsz,wght@0,8..60,400;0,8..60,600;0,8..60,700;1,8..60,400&family=Roboto:ital,wght@0,300;0,400;0,500;0,700;1,400&family=Roboto+Condensed:wght@400;500;700&display=swap');

/* ═══════════════════════════════════════════════════════════════════════════
   STANTEC VISUAL IDENTITY — colour tokens
   Sampled from "Visuele identiteit & ontwerprichtlijnen" (kleurenpalet, p.15)

   Hoofdkleuren    Wit #FFFFFF · Mist #F2EFEC · Kiezelsteen #E2DDDB
                   Fossiel #C9C4BD · Zwart #000000 · Stantec-oranje #ED6631
   Secundair       Fern #1A845C · Marine #2A73D9 · Kurkuma #F3BE0A
                   Himalayazout #F4B5B5

   Rules applied:  white space carries the layout; orange is the house colour
                   and stays prominent (lens-tag, buttons, accents); text is
                   black or white only; primary colours are used at full
                   strength — tints appear only in data visualisation.
   ═══════════════════════════════════════════════════════════════════════════ */
:root{
  /* Hoofdkleuren */
  --white:#FFFFFF; --mist:#F2EFEC; --pebble:#E2DDDB; --fossil:#C9C4BD;
  --black:#000000; --orange:#ED6631;
  /* Secundaire kleuren */
  --fern:#1A845C; --marine:#2A73D9; --kurkuma:#F3BE0A; --salt:#F4B5B5;

  /* Surfaces */
  --bg-0:var(--white); --bg-1:var(--mist); --card:var(--white); --side:var(--mist);

  /* Semantic status — Fern / Kurkuma / Marine come straight from the palette.
     --err is a functional alert colour: the brand palette carries no red. */
  --ok:var(--fern); --warn:var(--kurkuma); --info:var(--marine); --err:#B3261E;

  /* Text — black on light, per the WCAG table in the guidelines */
  --tx:#000000; --tx2:#26221F; --tx3:#5F5850;

  --bd:#DED8D3; --bd2:#C9C4BD; --hov:rgba(0,0,0,0.04);
  --r:4px;                       /* the identity is square-cornered, not pill-soft */
  --serif:'Source Serif 4',Georgia,'Times New Roman',serif;
  --sans:'Roboto',Arial,-apple-system,'Segoe UI',system-ui,sans-serif;
  --cond:'Roboto Condensed','Arial Narrow',Arial,sans-serif;
}

/* ── Canvas ─────────────────────────────────────────────────────────── */
/* White space is the foundation of the identity — the canvas stays white and
   the warm greys are used only on cards, the sidebar and section fields. */
.stApp{background-color:var(--white)}
.main .block-container{padding-top:1.1rem; padding-bottom:3rem; max-width:1500px;}
html,body,[class*="css"],.stApp,p,li,span,label,div,input,select,textarea,button{
  font-family:var(--sans);
}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-track{background:var(--mist)}
::-webkit-scrollbar-thumb{background:var(--fossil);border-radius:0;border:2px solid var(--mist)}
::-webkit-scrollbar-thumb:hover{background:#ADA79F}

/* ── Typography ─────────────────────────────────────────────────────── */
/* Source Serif 4 for titles and headings, Roboto for everything else. */
h1,h2,h3,h4,h5{color:var(--tx)!important;letter-spacing:-0.005em;font-family:var(--serif)!important}
h1{font-size:34px!important;font-weight:700!important;margin-bottom:.15rem!important;line-height:1.15!important}
h2{font-size:25px!important;font-weight:700!important}
h3{font-size:19px!important;font-weight:600!important}
h4{font-size:15px!important;font-weight:600!important;color:var(--tx)!important;font-family:var(--sans)!important}
p,li,label,.stMarkdown{color:var(--tx2);font-size:15px;line-height:1.62;font-weight:400}
.stCaption,[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p{color:var(--tx3)!important;font-size:13px!important}
small{color:var(--tx3)}
strong,b{color:var(--tx);font-weight:700}
code{background:var(--mist)!important;color:var(--tx)!important;border:1px solid var(--bd);
     border-radius:2px;padding:1px 6px;font-size:13px}
a{color:var(--tx)!important;text-decoration:underline;text-decoration-color:var(--orange);
  text-underline-offset:3px;text-decoration-thickness:2px}
a:hover{color:var(--orange)!important}
hr{border:none!important;border-top:1px solid var(--bd)!important;margin:1.5rem 0!important}
[data-testid="stMarkdownContainer"] table{border:1px solid var(--bd);border-radius:0;overflow:hidden;
     border-collapse:separate;border-spacing:0;font-size:13px}
[data-testid="stMarkdownContainer"] th{background:var(--mist)!important;color:var(--tx)!important;
     font-weight:700!important;border-bottom:1px solid var(--bd)!important;padding:8px 12px!important;
     text-transform:uppercase;letter-spacing:.06em;font-size:11px!important}
[data-testid="stMarkdownContainer"] td{color:var(--tx2)!important;border-bottom:1px solid var(--bd)!important;padding:8px 12px!important}

/* ── Motion ─────────────────────────────────────────────────────────── */
@keyframes riseIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes pulseDot{0%,100%{box-shadow:0 0 0 0 rgba(26,132,92,.45)}70%{box-shadow:0 0 0 7px rgba(26,132,92,0)}}
@keyframes flowLine{to{background-position:200% 0}}
@keyframes shimmer{0%{background-position:-500px 0}100%{background-position:500px 0}}
.rise{animation:riseIn .45s cubic-bezier(.22,1,.36,1) both}
.rise-1{animation-delay:.04s}.rise-2{animation-delay:.08s}.rise-3{animation-delay:.12s}
.rise-4{animation-delay:.16s}.rise-5{animation-delay:.20s}.rise-6{animation-delay:.24s}

/* ── Top header ─────────────────────────────────────────────────────── */
.ei-header{
  display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;
  background:var(--white);
  border:1px solid var(--bd);border-top:3px solid var(--orange);border-radius:0;
  padding:15px 22px;margin-bottom:18px;
  box-shadow:0 1px 3px rgba(0,0,0,.06);
  animation:riseIn .5s cubic-bezier(.22,1,.36,1) both;
}
.ei-brand{display:flex;align-items:center;gap:13px}
.ei-mark{width:40px;height:40px;flex-shrink:0}
.ei-name{font-family:var(--serif);font-size:21px;font-weight:700;color:var(--tx);letter-spacing:-.01em;line-height:1.15}
.ei-tag{font-size:10.5px;color:var(--tx3);letter-spacing:.14em;text-transform:uppercase;margin-top:3px;font-weight:700;font-family:var(--cond)}
.ei-hactions{display:flex;align-items:center;gap:9px;flex-wrap:wrap}
.ei-pill{
  display:inline-flex;align-items:center;gap:7px;padding:6px 12px;border-radius:2px;
  border:1px solid var(--bd2);background:var(--white);
  font-size:11px;font-weight:700;color:var(--tx2);letter-spacing:.06em;text-transform:uppercase;
  transition:all .22s ease;white-space:nowrap;font-family:var(--cond);
}
.ei-pill:hover{background:var(--mist);border-color:var(--tx3)}
.ei-pill.live{color:var(--tx);border-color:var(--fern);background:var(--white)}
.ei-pill.ver{color:var(--black);border-color:transparent;background:var(--orange)}
.ei-dot{width:7px;height:7px;border-radius:50%;background:var(--fern);animation:pulseDot 2.4s infinite}
.ei-ico{width:33px;height:33px;border-radius:2px;display:inline-flex;align-items:center;justify-content:center;
  border:1px solid var(--bd);background:var(--white);color:var(--tx2);transition:all .22s ease;cursor:default}
.ei-ico:hover{background:var(--mist);color:var(--black);border-color:var(--bd2)}
.ei-avatar{width:33px;height:33px;border-radius:2px;display:inline-flex;align-items:center;justify-content:center;
  background:var(--orange);color:var(--black);font-weight:700;font-size:12.5px;font-family:var(--cond);letter-spacing:.04em}

/* ── Page title block ───────────────────────────────────────────────── */
.ei-page{margin:2px 0 18px 0;animation:riseIn .5s cubic-bezier(.22,1,.36,1) both}
.ei-crumb{font-size:11px;color:var(--tx3);letter-spacing:.12em;text-transform:uppercase;font-weight:700;
  margin-bottom:9px;font-family:var(--cond)}
.ei-crumb span{color:var(--orange)}
.ei-h1{font-family:var(--serif);font-size:38px;font-weight:700;color:var(--tx);letter-spacing:-.01em;line-height:1.14}
.ei-sub{font-size:15px;color:var(--tx2);margin-top:9px;max-width:820px;line-height:1.6}

/* ── Section header ─────────────────────────────────────────────────── */
.ei-sec{display:flex;align-items:center;gap:11px;margin:26px 0 13px 0;
  padding-bottom:9px;border-bottom:1px solid var(--bd)}
.ei-sec-bar{width:4px;height:21px;border-radius:0;background:var(--orange)}
.ei-sec-t{font-family:var(--serif);font-size:20px;font-weight:700;color:var(--tx);letter-spacing:-.005em}
.ei-sec-d{font-size:13px;color:var(--tx3);margin-left:2px}

/* ── Cards ──────────────────────────────────────────────────────────── */
.ei-card{
  background:var(--white);
  border:1px solid var(--bd);border-radius:var(--r);padding:18px 20px;
  box-shadow:0 1px 3px rgba(0,0,0,.05);transition:all .26s cubic-bezier(.22,1,.36,1);height:100%;
}
.ei-card:hover{border-color:var(--fossil);box-shadow:0 6px 18px rgba(0,0,0,.09);transform:translateY(-2px)}

/* KPI card — warm-grey field, orange rule on top */
.kpi{
  position:relative;overflow:hidden;
  background:var(--mist);
  border:1px solid var(--bd);border-radius:var(--r);padding:16px 18px 14px 18px;
  box-shadow:none;transition:all .26s cubic-bezier(.22,1,.36,1);height:100%;
}
.kpi:hover{transform:translateY(-3px);border-color:var(--fossil);box-shadow:0 8px 22px rgba(0,0,0,.10)}
.kpi::after{content:'';position:absolute;top:0;left:0;right:0;height:3px;
  background:var(--accent-c,var(--orange));opacity:1}
.kpi-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:9px}
.kpi-l{font-size:11px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.1em;
  line-height:1.35;font-family:var(--cond)}
.kpi-i{width:30px;height:30px;border-radius:2px;display:flex;align-items:center;justify-content:center;flex-shrink:0;
  background:var(--white);color:var(--accent-c,var(--orange));border:1px solid var(--bd)}
.kpi-v{font-family:var(--serif);font-size:42px;font-weight:700;color:var(--tx);line-height:1;
  letter-spacing:-.02em;font-variant-numeric:tabular-nums}
.kpi-v.sm{font-size:32px}
.kpi-u{font-family:var(--sans);font-size:12px;color:var(--tx3);font-weight:500;margin-left:6px;letter-spacing:.01em}
.kpi-foot{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:10px;
  padding-top:9px;border-top:1px solid var(--bd)}
.kpi-cmp{font-size:11.5px;color:var(--tx3);line-height:1.4}
.kpi-trend{font-size:11.5px;font-weight:700;display:inline-flex;align-items:center;gap:4px;white-space:nowrap}

/* ── Status pills ───────────────────────────────────────────────────── */
.pill{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:2px;
  font-size:10.5px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;white-space:nowrap;
  font-family:var(--cond)}
.pill i{width:6px;height:6px;border-radius:50%;background:currentColor;display:inline-block}
.pill.ok{color:var(--black);background:var(--white);border:1px solid var(--fern);box-shadow:inset 3px 0 0 var(--fern)}
.pill.info{color:var(--black);background:var(--white);border:1px solid var(--marine);box-shadow:inset 3px 0 0 var(--marine)}
.pill.warn{color:var(--black);background:var(--white);border:1px solid var(--kurkuma);box-shadow:inset 3px 0 0 var(--kurkuma)}
.pill.err{color:var(--black);background:var(--white);border:1px solid var(--err);box-shadow:inset 3px 0 0 var(--err)}
.pill.idle{color:var(--tx3);background:var(--mist);border:1px solid var(--bd)}

/* ── Workflow stepper ───────────────────────────────────────────────── */
.wf{display:flex;align-items:stretch;gap:0;margin:6px 0 20px 0;flex-wrap:nowrap;overflow-x:auto;padding-bottom:4px}
.wf-s{flex:1 1 0;min-width:132px;background:var(--white);
  border:1px solid var(--bd);border-radius:var(--r);padding:13px 14px;transition:all .28s cubic-bezier(.22,1,.36,1)}
.wf-s:hover{transform:translateY(-2px);border-color:var(--fossil)}
.wf-s.done{border-color:var(--fern);background:var(--white)}
.wf-s.act{border-color:var(--orange);background:var(--orange);
  box-shadow:0 6px 18px rgba(237,102,49,.28)}
.wf-n{width:25px;height:25px;border-radius:2px;display:flex;align-items:center;justify-content:center;
  font-size:11.5px;font-weight:700;margin-bottom:9px;border:1px solid var(--bd);
  background:var(--mist);color:var(--tx3);font-family:var(--cond)}
.wf-s.done .wf-n{background:var(--fern);color:var(--white);border-color:var(--fern)}
.wf-s.act .wf-n{background:var(--black);color:var(--white);border-color:var(--black)}
.wf-t{font-size:12.5px;font-weight:700;color:var(--tx);letter-spacing:-.01em;line-height:1.3}
.wf-s.act .wf-t{color:var(--black)}
.wf-d{font-size:10.5px;color:var(--tx3);margin-top:3px;line-height:1.35}
.wf-s.act .wf-d{color:var(--black);opacity:.72}
.wf-c{flex:0 0 26px;display:flex;align-items:center;justify-content:center;min-width:26px}
.wf-c i{display:block;width:100%;height:2px;border-radius:0;background:var(--bd2)}
.wf-c.on i{background:linear-gradient(90deg,var(--orange),#F59A72,var(--orange));
  background-size:200% 100%;animation:flowLine 2.2s linear infinite}

/* ── Validation feed ────────────────────────────────────────────────── */
.vf{border:1px solid var(--bd);border-radius:var(--r);padding:12px 15px;margin-bottom:8px;
  display:flex;gap:12px;align-items:flex-start;transition:all .22s ease;background:var(--white)}
.vf:hover{border-color:var(--fossil);transform:translateX(2px);box-shadow:0 3px 10px rgba(0,0,0,.06)}
.vf-b{width:4px;border-radius:0;align-self:stretch;flex-shrink:0}
.vf-i{width:22px;height:22px;border-radius:2px;display:flex;align-items:center;justify-content:center;
  font-size:12px;font-weight:700;flex-shrink:0;margin-top:1px}
.vf-body{flex:1;min-width:0}
.vf-h{display:flex;align-items:center;gap:9px;flex-wrap:wrap;margin-bottom:3px}
.vf-cat{font-size:10.5px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.09em;font-family:var(--cond)}
.vf-m{font-size:13.5px;color:var(--tx2);line-height:1.55}
.vf.pass{border-left:none}
.vf.pass .vf-b{background:var(--fern)} .vf.pass .vf-i{background:var(--fern);color:var(--white)}
.vf.warn .vf-b{background:var(--kurkuma)} .vf.warn .vf-i{background:var(--kurkuma);color:var(--black)}
.vf.fail .vf-b{background:var(--err)}    .vf.fail .vf-i{background:var(--err);color:var(--white)}
.vf.info .vf-b{background:var(--marine)} .vf.info .vf-i{background:var(--marine);color:var(--white)}

/* summary counters */
.vsum{display:flex;gap:11px;flex-wrap:wrap;margin-bottom:14px}
.vsum-c{flex:1;min-width:118px;border:1px solid var(--bd);border-radius:var(--r);padding:13px 16px;
  background:var(--mist);transition:all .24s ease}
.vsum-c:hover{transform:translateY(-2px);border-color:var(--fossil)}
.vsum-n{font-family:var(--serif);font-size:30px;font-weight:700;line-height:1;letter-spacing:-.02em;
  font-variant-numeric:tabular-nums}
.vsum-l{font-size:10.5px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.1em;
  margin-top:6px;font-family:var(--cond)}

/* ── Upload dropzone ────────────────────────────────────────────────── */
.dz{border:1.5px dashed var(--fossil);border-radius:var(--r);padding:30px 24px;text-align:center;
  background:var(--mist);transition:all .3s cubic-bezier(.22,1,.36,1);margin-bottom:10px}
.dz:hover{border-color:var(--orange);background:var(--pebble);
  box-shadow:0 10px 30px rgba(0,0,0,.08);transform:translateY(-2px)}
.dz-t{font-family:var(--serif);font-size:18px;font-weight:700;color:var(--tx);margin-top:11px}
.dz-d{font-size:13px;color:var(--tx3);margin-top:5px}
.dz-f{display:flex;gap:7px;justify-content:center;flex-wrap:wrap;margin-top:15px}
.dz-f span{font-size:10.5px;font-weight:700;color:var(--tx2);padding:5px 11px;border-radius:2px;
  border:1px solid var(--bd);background:var(--white);letter-spacing:.06em;transition:all .2s ease;font-family:var(--cond)}
.dz-f span:hover{border-color:var(--orange);color:var(--black)}

/* ── Empty state ────────────────────────────────────────────────────── */
.es{text-align:center;padding:42px 24px;border:1px dashed var(--bd2);border-radius:var(--r);background:var(--mist)}
.es-t{font-family:var(--serif);font-size:17px;font-weight:700;color:var(--tx);margin-top:13px}
.es-d{font-size:13.5px;color:var(--tx3);margin-top:6px;max-width:430px;margin-left:auto;margin-right:auto;line-height:1.6}

/* ── Footer ─────────────────────────────────────────────────────────── */
.ei-foot{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;
  margin-top:32px;padding:15px 20px;border-top:3px solid var(--orange);border-radius:0;
  background:var(--mist);font-size:11.5px;color:var(--tx3)}
.ei-foot b{color:var(--tx);font-weight:700}
.ei-foot-r{display:flex;gap:16px;flex-wrap:wrap;align-items:center}

/* ── Sidebar ────────────────────────────────────────────────────────── */
section[data-testid="stSidebar"]{
  background:var(--mist);
  border-right:1px solid var(--bd);
}
section[data-testid="stSidebar"] > div{padding-top:1.1rem}
section[data-testid="stSidebar"] *{color:var(--tx2)}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3{
  color:var(--tx)!important;font-family:var(--serif)!important}
section[data-testid="stSidebar"] hr{border-top:1px solid var(--bd)!important;margin:.9rem 0!important}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p{color:var(--tx3)!important}

/* nav item */
section[data-testid="stSidebar"] div[role="radiogroup"]{gap:5px!important;display:flex;flex-direction:column}
section[data-testid="stSidebar"] div[role="radiogroup"] > label{
  display:flex!important;align-items:center;gap:11px;padding:10px 12px;border-radius:var(--r);
  border:1px solid transparent;cursor:pointer;transition:all .22s cubic-bezier(.22,1,.36,1);
  margin:0!important;position:relative;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{background:var(--white);border-color:var(--bd)}
/* Hide the native radio glyph — the orange Lens-tag icon stands in for it.
   Two selectors so this holds across Streamlit's older and newer radio DOMs. */
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child,
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div > div > div:first-child{
  display:none!important}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:last-child,
section[data-testid="stSidebar"] div[role="radiogroup"] > label [data-testid="stMarkdownContainer"] p{
  font-size:13.5px!important;font-weight:500!important;color:var(--tx2)!important;letter-spacing:-.01em;
  margin:0!important}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){
  background:var(--white);border-color:var(--bd2);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::before{
  content:'';position:absolute;left:0;top:15%;height:70%;width:4px;border-radius:0;
  background:var(--orange);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) > div:last-child,
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) [data-testid="stMarkdownContainer"] p{
  color:var(--tx)!important;font-weight:700!important}
section[data-testid="stSidebar"] div[role="radiogroup"] input{accent-color:var(--orange)}

section[data-testid="stSidebar"] button{
  background:var(--white)!important;border:1px solid var(--bd2)!important;
  color:var(--tx)!important;font-weight:500!important;border-radius:var(--r)!important;font-size:13px!important;
  transition:all .22s ease!important}
section[data-testid="stSidebar"] button:hover{
  background:var(--orange)!important;border-color:var(--orange)!important;
  color:var(--black)!important;transform:translateY(-1px)}

/* ── Inputs ─────────────────────────────────────────────────────────── */
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div{
  background:var(--white)!important;border:1px solid var(--bd2)!important;border-radius:var(--r)!important;
  color:var(--tx)!important;transition:all .2s ease!important;min-height:40px}
.stSelectbox div[data-baseweb="select"] > div:hover{border-color:var(--orange)!important;background:var(--white)!important}
div[data-baseweb="select"] svg{color:var(--tx3)!important}
div[data-baseweb="popover"] li{background:var(--white)!important;color:var(--tx2)!important;font-size:13.5px!important}
div[data-baseweb="popover"] li:hover{background:var(--mist)!important;color:var(--black)!important}
div[data-baseweb="popover"] ul{background:var(--white)!important;border:1px solid var(--bd2)!important;border-radius:var(--r)!important}

.stTextInput input,.stNumberInput input,.stTextArea textarea{
  background:var(--white)!important;border:1px solid var(--bd2)!important;border-radius:var(--r)!important;
  color:var(--tx)!important;font-size:14px!important;transition:all .2s ease!important}
.stTextInput input:focus,.stNumberInput input:focus,.stTextArea textarea:focus{
  border-color:var(--orange)!important;box-shadow:0 0 0 3px rgba(237,102,49,.22)!important}
.stTextInput input::placeholder,.stTextArea textarea::placeholder{color:#8C857D!important}
.stNumberInput button{background:var(--mist)!important;border-color:var(--bd2)!important;color:var(--tx2)!important}
[data-testid="stWidgetLabel"] p,[data-testid="stWidgetLabel"] label{
  font-size:11.5px!important;font-weight:700!important;color:var(--tx2)!important;
  letter-spacing:.06em!important;text-transform:uppercase!important;font-family:var(--cond)!important}

/* ── Buttons ────────────────────────────────────────────────────────── */
.stButton button,.stDownloadButton button,.stFormSubmitButton button{
  border-radius:var(--r)!important;font-weight:500!important;font-size:13.5px!important;
  transition:all .24s cubic-bezier(.22,1,.36,1)!important;border:1px solid var(--black)!important;
  background:var(--white)!important;color:var(--tx)!important;padding:.5rem 1.1rem!important}
.stButton button:hover,.stDownloadButton button:hover,.stFormSubmitButton button:hover{
  background:var(--mist)!important;border-color:var(--black)!important;transform:translateY(-1px)}
/* Orange is the house colour — primary actions carry it, with black type (WCAG) */
.stButton button[kind="primary"],.stFormSubmitButton button[kind="primary"],
.stButton button[data-testid="baseButton-primary"]{
  background:var(--orange)!important;border-color:var(--orange)!important;
  color:var(--black)!important;font-weight:700!important;box-shadow:0 2px 8px rgba(237,102,49,.30)!important}
.stButton button[kind="primary"]:hover,.stFormSubmitButton button[kind="primary"]:hover,
.stButton button[data-testid="baseButton-primary"]:hover{
  box-shadow:0 6px 18px rgba(237,102,49,.42)!important;transform:translateY(-2px)!important;
  background:var(--black)!important;border-color:var(--black)!important;color:var(--white)!important}
.stDownloadButton button{background:var(--white)!important;border-color:var(--orange)!important;color:var(--black)!important}
.stDownloadButton button:hover{background:var(--orange)!important;color:var(--black)!important}

/* ── Tabs ───────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"]{gap:5px;background:transparent;border-bottom:1px solid var(--bd);padding-bottom:0}
.stTabs [data-baseweb="tab"]{
  background:transparent!important;border-radius:0!important;padding:9px 17px!important;
  font-weight:500!important;font-size:13.5px!important;color:var(--tx3)!important;transition:all .22s ease!important}
.stTabs [data-baseweb="tab"]:hover{background:var(--mist)!important;color:var(--tx)!important}
.stTabs [aria-selected="true"]{color:var(--tx)!important;background:transparent!important;font-weight:700!important}
.stTabs [data-baseweb="tab-highlight"]{background:var(--orange)!important;height:3px!important}
.stTabs [data-baseweb="tab-border"]{background:transparent!important}

/* ── Data grid ──────────────────────────────────────────────────────── */
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{
  border:1px solid var(--bd)!important;border-radius:var(--r)!important;overflow:hidden!important;
  box-shadow:0 1px 3px rgba(0,0,0,.05)!important;background:var(--white)!important}
[data-testid="stDataFrame"] div[role="columnheader"],[data-testid="stDataEditor"] div[role="columnheader"]{
  background:var(--mist)!important;color:var(--tx)!important;font-weight:700!important;
  text-transform:uppercase;letter-spacing:.06em;font-size:11px!important}

/* ── Alerts ─────────────────────────────────────────────────────────── */
div[data-testid="stAlert"]{border-radius:var(--r)!important;border:1px solid var(--bd)!important;
  border-left:4px solid var(--orange)!important;background:var(--mist)!important}
div[data-testid="stAlert"] p{color:var(--tx2)!important;font-size:13.5px!important}

/* ── Expander / file uploader / metric / progress ───────────────────── */
[data-testid="stExpander"]{border:1px solid var(--bd)!important;border-radius:var(--r)!important;
  background:var(--white)!important;overflow:hidden}
[data-testid="stExpander"] summary{font-size:13.5px!important;font-weight:600!important;color:var(--tx)!important}
[data-testid="stExpander"] summary:hover{background:var(--mist)!important}
[data-testid="stFileUploader"] section{
  background:var(--mist)!important;border:1.5px dashed var(--fossil)!important;
  border-radius:var(--r)!important;transition:all .28s cubic-bezier(.22,1,.36,1)!important;padding:18px!important}
[data-testid="stFileUploader"] section:hover{
  border-color:var(--orange)!important;background:var(--pebble)!important;
  box-shadow:0 10px 30px rgba(0,0,0,.08)!important;transform:translateY(-2px)}
[data-testid="stFileUploader"] section small,[data-testid="stFileUploader"] section span{color:var(--tx3)!important}
[data-testid="stFileUploader"] button{background:var(--orange)!important;border-color:var(--orange)!important;color:var(--black)!important}
[data-testid="stMetric"]{background:var(--mist);
  border:1px solid var(--bd);border-radius:var(--r);padding:15px 17px;transition:all .26s ease}
[data-testid="stMetric"]:hover{transform:translateY(-2px);border-color:var(--fossil)}
[data-testid="stMetricLabel"] p{font-size:11px!important;font-weight:700!important;color:var(--tx3)!important;
  text-transform:uppercase!important;letter-spacing:.1em!important}
[data-testid="stMetricValue"]{font-family:var(--serif)!important;font-size:28px!important;font-weight:700!important;
  color:var(--tx)!important;letter-spacing:-.02em}
.stProgress > div > div > div{background:var(--orange)!important}
.stSpinner > div{border-top-color:var(--orange)!important}
[data-testid="stForm"]{border:1px solid var(--bd)!important;border-radius:var(--r)!important;
  background:var(--white)!important;padding:20px!important}
.js-plotly-plot .plotly .modebar{background:transparent!important}
.js-plotly-plot{border-radius:var(--r);overflow:hidden}

/* ── Streamlit DOM compatibility ────────────────────────────────────────
   Newer Streamlit builds render selects as a react-aria ComboBox and wrap
   alerts in their own tinted container. These rules re-apply the brand
   surface for those builds without disturbing the older selectors above. */
[data-testid="stSelectbox"] div[role="group"],
[data-testid="stMultiSelect"] div[role="group"],
[data-testid="stDateInput"] div[role="group"]{
  background:var(--white)!important;border:1px solid var(--bd2)!important;
  border-radius:var(--r)!important;color:var(--tx)!important}
[data-testid="stSelectbox"] div[role="group"]:focus-within,
[data-testid="stMultiSelect"] div[role="group"]:focus-within{
  border-color:var(--orange)!important;box-shadow:0 0 0 3px rgba(237,102,49,.22)!important}
[data-testid="stSelectbox"] input,[data-testid="stMultiSelect"] input{
  background:transparent!important;color:var(--tx)!important}
.react-aria-Popover,[role="listbox"]{
  background:var(--white)!important;border:1px solid var(--bd2)!important;border-radius:var(--r)!important}
[role="option"]{background:var(--white)!important;color:var(--tx2)!important}
[role="option"]:hover,[role="option"][data-focused]{background:var(--mist)!important;color:var(--black)!important}

/* Alerts: warm-grey field with a coloured rule, instead of Streamlit's tints */
[data-testid="stAlertContainer"]{background:var(--mist)!important;color:var(--tx2)!important;
  border-radius:var(--r)!important;box-shadow:inset 4px 0 0 var(--orange)}
[data-testid="stAlertContentInfo"]    ~ *,[data-testid="stAlertContentInfo"]{color:var(--tx2)!important}
div[data-testid="stAlert"]:has([data-testid="stAlertContentInfo"])    [data-testid="stAlertContainer"]{box-shadow:inset 4px 0 0 var(--marine)}
div[data-testid="stAlert"]:has([data-testid="stAlertContentSuccess"]) [data-testid="stAlertContainer"]{box-shadow:inset 4px 0 0 var(--fern)}
div[data-testid="stAlert"]:has([data-testid="stAlertContentWarning"]) [data-testid="stAlertContainer"]{box-shadow:inset 4px 0 0 var(--kurkuma)}
div[data-testid="stAlert"]:has([data-testid="stAlertContentError"])   [data-testid="stAlertContainer"]{box-shadow:inset 4px 0 0 var(--err)}

/* Streamlit's own chrome — the "Running…" status pill, toasts, tooltips and
   the connection banner — is painted by the base theme, not by this stylesheet.
   .streamlit/config.toml sets that base theme to light; these rules make the
   pieces that stay visible match the identity. */
[data-testid="stStatusWidget"]{
  visibility:visible!important;height:auto!important;
  background:var(--mist)!important;border:1px solid var(--bd)!important;
  border-left:4px solid var(--orange)!important;border-radius:var(--r)!important;
  box-shadow:0 2px 8px rgba(0,0,0,.10)!important}
[data-testid="stStatusWidget"] *,[data-testid="stStatusWidget"] label{
  color:var(--tx)!important;font-family:var(--sans)!important}
[data-testid="stStatusWidget"] code{background:var(--white)!important;color:var(--tx)!important;
  border:1px solid var(--bd)!important}
[data-testid="stStatusWidget"] svg{fill:var(--orange)!important;color:var(--orange)!important}
[data-testid="stToast"]{background:var(--mist)!important;color:var(--tx)!important;
  border:1px solid var(--bd)!important;border-radius:var(--r)!important}
[data-testid="stTooltipContent"],[data-baseweb="tooltip"]{
  background:var(--black)!important;color:var(--white)!important;border-radius:var(--r)!important;
  font-family:var(--sans)!important}
[data-testid="stSpinner"] p,[data-testid="stSpinner"] div{color:var(--tx2)!important}

/* Hide the menu, Deploy button and footer, but keep the header itself in flow:
   the "Running…" status pill lives inside it and is the only feedback the user
   gets while the benchmark sheet loads. */
#MainMenu,[data-testid="stMainMenu"],[data-testid="stAppDeployButton"],footer{display:none!important}
header[data-testid="stHeader"]{background:transparent!important;box-shadow:none!important}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM — tokens, Lucide icons and reusable UI components
#  (Presentation layer only — no analytical behaviour lives here.)
# ══════════════════════════════════════════════════════════════════════════════
# Stantec brand palette — see "Visuele identiteit & ontwerprichtlijnen", kleurenpalet.
# Hoofdkleuren: Wit, Mist, Kiezelsteen, Fossiel, Zwart, Stantec-oranje.
# Secundaire kleuren: Fern, Marine, Kurkuma, Himalayazout.
BRAND = {
    "orange": "#ED6631",   # Stantec-oranje — house colour
    "white":  "#FFFFFF",   # Wit
    "mist":   "#F2EFEC",   # Mist (extra-lichtgrijs)
    "pebble": "#E2DDDB",   # Kiezelsteen (lichtgrijs)
    "fossil": "#C9C4BD",   # Fossiel (grijs)
    "black":  "#000000",   # Zwart
    "fern":   "#1A845C",   # Fern (groen)
    "marine": "#2A73D9",   # Marine (blauw)
    "kurkuma": "#F3BE0A",  # Kurkuma (geel)
    "salt":   "#F4B5B5",   # Himalayazout (roze)
}

UI = {
    # Surfaces
    "bg": BRAND["white"], "bg2": BRAND["mist"],
    "card": BRAND["white"], "side": BRAND["mist"],
    # Accents — "blue"/"cyan" are kept as key names for backwards compatibility,
    # but both now resolve to the brand's house colour and its supporting blue.
    "blue": BRAND["orange"], "cyan": BRAND["orange"], "accent2": BRAND["marine"],
    # Status. Fern / Kurkuma / Marine come from the palette; the brand carries no
    # red, so a functional alert red is used for hard failures only.
    "ok": BRAND["fern"], "warn": BRAND["kurkuma"], "err": "#B3261E",
    "info": BRAND["marine"],
    # Type — black on light, per the WCAG contrast table in the guidelines.
    "tx": BRAND["black"], "tx2": "#26221F", "tx3": "#5F5850",
    "grid": "#E2DDDB",
    "bd": "#DED8D3",
}
UI.update(BRAND)

APP_VERSION = "v2.0"

# Typeface stacks — Source Serif 4 for headings, Roboto for body copy.
FONT_SERIF = "Source Serif 4, Georgia, serif"
FONT_SANS  = "Roboto, Arial, sans-serif"

# Path to the official Stantec logo. Drop the asset from the brand portal here;
# the header falls back to a Lens-tag style orange block if it is absent.
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "stantec_logo.svg")

# Navigation destinations (each maps to a real, implemented page).
PAGE_QA       = "QA/QC Validation"
PAGE_EXPLORER = "Benchmark Explorer"
PAGE_DB       = "Benchmark Database"

# Chart palette — brand colours first, then tints. The guidelines allow tints of
# the primary colours in data visualisation only.
CHART_SEQ = [BRAND["orange"], BRAND["marine"], BRAND["fern"], BRAND["kurkuma"],
             BRAND["salt"], BRAND["black"], BRAND["fossil"],
             "#F59A72", "#7FA9E8"]

# Lucide-style outline icons (24x24, currentColor, stroke 2).
_L = {
    "activity":  "M22 12h-4l-3 9L9 3l-3 9H2",
    "upload":    "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12",
    "columns":   "M12 3v18M3 3h18v18H3z",
    "building":  "M3 21h18M5 21V7l8-4v18M19 21V11l-6-4M9 9v.01M9 12v.01M9 15v.01M9 18v.01",
    "shield":    "M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z",
    "bar":       "M12 20V10M18 20V4M6 20v-4",
    "database":  "M12 5c4.42 0 8-1.12 8-2.5S16.42 0 12 0 4 1.12 4 2.5 7.58 5 12 5z"
                 "M20 12c0 1.38-3.58 2.5-8 2.5S4 13.38 4 12M20 5v14c0 1.38-3.58 2.5-8 2.5s-8-1.12-8-2.5V5",
    "search":    "M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16zM21 21l-4.35-4.35",
    "bell":      "M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9M13.73 21a2 2 0 0 1-3.46 0",
    "book":      "M4 19.5A2.5 2.5 0 0 1 6.5 17H20M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z",
    "settings":  "M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z"
                 "M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33"
                 " 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06"
                 "a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.6 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09"
                 "A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.6"
                 "a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06"
                 "a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09"
                 "a1.65 1.65 0 0 0-1.51 1z",
    "zap":       "M13 2 3 14h9l-1 8 10-12h-9l1-8z",
    "flame":     "M12 23a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5s-3 3.5-3 5.5a7 7 0 0 0 7 7z",
    "cloud":     "M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z",
    "leaf":      "M11 20A7 7 0 0 1 9.8 6.1C15.5 5 17 4.48 19 2c1 2 2 4.18 2 8 0 5.5-4.78 10-10 10zM2 21c0-3 1.85-5.36 5.08-6",
    "gauge":     "M12 14l4-4M20.5 15a9 9 0 1 0-17 0",
    "clock":     "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 6v6l4 2",
    "check":     "M20 6 9 17l-5-5",
    "alert":     "M12 9v4M12 17h.01M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z",
    "x":         "M18 6 6 18M6 6l12 12",
    "info":      "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 16v-4M12 8h.01",
    "trend-up":  "M23 6l-9.5 9.5-5-5L1 18M17 6h6v6",
    "trend-down":"M23 18l-9.5-9.5-5 5L1 6M17 18h6v-6",
    "layers":    "M12 2 2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5",
    "file":      "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM14 2v6h6",
    "download":  "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3",
    "filter":    "M22 3H2l8 9.46V19l4 2v-8.54L22 3z",
    "target":    "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM12 18a6 6 0 1 0 0-12 6 6 0 0 0 0 12zM12 14a2 2 0 1 0 0-4 2 2 0 0 0 0 4z",
    "help":      "M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20zM9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3M12 17h.01",
    "cpu":       "M4 4h16v16H4zM9 9h6v6H9zM9 1v3M15 1v3M9 20v3M15 20v3M20 9h3M20 14h3M1 9h3M1 14h3",
    "inbox":     "M22 12h-6l-2 3h-4l-2-3H2M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z",
}

def icon(name, size=16, color="currentColor", sw=2):
    """Inline Lucide-style outline SVG."""
    d = _L.get(name, _L["activity"])
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" '
            f'fill="none" stroke="{color}" stroke-width="{sw}" stroke-linecap="round" '
            f'stroke-linejoin="round"><path d="{d}"/></svg>')

def _svg_uri(name, stroke="%235F5850"):
    """URL-encoded single-path SVG for use inside a CSS data: URI (sidebar nav icons)."""
    d = _L.get(name, _L["activity"])
    svg = ("<svg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' "
           "fill='none' stroke='STROKE' stroke-width='2' stroke-linecap='round' "
           "stroke-linejoin='round'><path d='" + d + "'/></svg>")
    svg = svg.replace("stroke='STROKE'", "stroke='" + stroke + "'")
    # Percent-encode the characters that are unsafe inside a CSS url("...") data URI.
    for a, b in (("%", "%25"), ("#", "%23"), ("<", "%3C"), (">", "%3E"),
                 ('"', "%22"), ("{", "%7B"), ("}", "%7D")):
        if a == "%":
            continue
        svg = svg.replace(a, b)
    return svg

@st.cache_data(show_spinner=False)
def _load_logo_asset(path):
    """Read the official Stantec logo from disk once. Returns (kind, payload)."""
    if not os.path.exists(path):
        return None, None
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".svg":
            with open(path, "r", encoding="utf-8") as f:
                return "svg", f.read()
        with open(path, "rb") as f:
            import base64
            mime = "image/png" if ext == ".png" else "image/jpeg"
            return "img", f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
    except Exception:
        return None, None


def logo_mark(size=40):
    """Stantec logo lockup.

    The official mark is never redrawn here — it is loaded from
    ``assets/stantec_logo.svg`` (or .png) so the approved artwork is always the
    one that ships. Until that asset is dropped in, a Lens-tag style orange
    block stands in: an orange square with the tool's initials, matching the
    'Lens-tag' guidance for constrained spaces.
    """
    kind, payload = _load_logo_asset(LOGO_PATH)
    if kind == "svg":
        return (f'<span class="ei-mark" style="display:inline-flex;align-items:center;'
                f'height:{size}px;width:auto">{payload}</span>')
    if kind == "img":
        return (f'<img class="ei-mark" src="{payload}" alt="Stantec" '
                f'style="height:{size}px;width:auto;object-fit:contain"/>')
    # Fallback placeholder — a solid Stantec-orange tag, black type (WCAG-safe).
    return (f'<span class="ei-mark" title="Place the official logo at assets/stantec_logo.svg" '
            f'style="width:{size}px;height:{size}px;border-radius:2px;background:{BRAND["orange"]};'
            f'display:inline-flex;align-items:center;justify-content:center;'
            f'font-family:{FONT_SERIF};font-weight:700;font-size:{max(11, int(size*0.42))}px;'
            f'color:{BRAND["black"]};letter-spacing:-.02em;flex-shrink:0">EI</span>')

def render_header(bm_count, page_label):
    """Top application header: brand, environment badges, utility icons, avatar."""
    st.markdown(f'''<div class="ei-header">
      <div class="ei-brand">
        {logo_mark(40)}
        <div>
          <div class="ei-name">Energy Intelligence</div>
          <div class="ei-tag">Benchmarking &middot; QA/QC &middot; Analytics</div>
        </div>
      </div>
      <div class="ei-hactions">
        <span class="ei-pill ver">{APP_VERSION}</span>
        <span class="ei-pill live"><span class="ei-dot"></span>Live</span>
        <span class="ei-pill">{icon("database",13)}&nbsp;{bm_count} Benchmarks</span>
        <span class="ei-pill">{icon("cpu",13)}&nbsp;AI Ready</span>
        <span class="ei-ico" title="Search">{icon("search",16)}</span>
        <span class="ei-ico" title="Notifications">{icon("bell",16)}</span>
        <span class="ei-ico" title="Documentation">{icon("book",16)}</span>
        <span class="ei-ico" title="Help">{icon("help",16)}</span>
        <span class="ei-avatar" title="Signed in">EI</span>
      </div>
    </div>''', unsafe_allow_html=True)

def page_head(title, subtitle, crumbs):
    """Breadcrumbed page title block."""
    trail = ' <span>/</span> '.join(crumbs)
    st.markdown(f'''<div class="ei-page">
      <div class="ei-crumb">{trail}</div>
      <div class="ei-h1">{title}</div>
      <div class="ei-sub">{subtitle}</div>
    </div>''', unsafe_allow_html=True)

def section(title, desc="", ic="bar"):
    """Accent-barred section header."""
    d = f'<span class="ei-sec-d">{desc}</span>' if desc else ""
    st.markdown(f'''<div class="ei-sec"><div class="ei-sec-bar"></div>
      <span style="color:{UI['cyan']};display:flex;align-items:center">{icon(ic,17)}</span>
      <span class="ei-sec-t">{title}</span>{d}</div>''', unsafe_allow_html=True)

def sparkline(values, color="#ED6631", w=104, h=26):
    """Minimal inline SVG sparkline with a soft gradient area fill."""
    vals = [float(v) for v in values if v is not None]
    if len(vals) < 2:
        return ""
    lo, hi = min(vals), max(vals)
    rng = (hi - lo) or 1.0
    n = len(vals)
    pts = [(i * w / (n - 1), h - 3 - ((v - lo) / rng) * (h - 6)) for i, v in enumerate(vals)]
    line = " ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
    area = f"0,{h} " + line + f" {w},{h}"
    uid = abs(hash((tuple(vals), color))) % 100000
    return (f'<svg width="{w}" height="{h}" viewBox="0 0 {w} {h}" fill="none">'
            f'<defs><linearGradient id="sp{uid}" x1="0" y1="0" x2="0" y2="1">'
            f'<stop stop-color="{color}" stop-opacity=".38"/>'
            f'<stop offset="1" stop-color="{color}" stop-opacity="0"/></linearGradient></defs>'
            f'<polygon points="{area}" fill="url(#sp{uid})"/>'
            f'<polyline points="{line}" stroke="{color}" stroke-width="1.7" fill="none" '
            f'stroke-linecap="round" stroke-linejoin="round"/>'
            f'<circle cx="{pts[-1][0]:.1f}" cy="{pts[-1][1]:.1f}" r="2.4" fill="{color}"/></svg>')

# KPI accent tones: (rule / icon colour, icon field, border).
# Icon fields stay white so no colour is used as a tint of itself.
_TONE = {
    "ok":   (BRAND["fern"],   BRAND["white"], "#DED8D3"),
    "warn": (BRAND["kurkuma"], BRAND["white"], "#DED8D3"),
    "err":  ("#B3261E",       BRAND["white"], "#DED8D3"),
    "blue": (BRAND["orange"], BRAND["white"], "#DED8D3"),   # legacy key → house colour
    "cyan": (BRAND["orange"], BRAND["white"], "#DED8D3"),   # legacy key → house colour
    "marine": (BRAND["marine"], BRAND["white"], "#DED8D3"),
    "idle": (BRAND["fossil"], BRAND["white"], "#DED8D3"),
}

def kpi_html(label, value, unit="", ic="bar", tone="cyan",
             compare="", trend=None, spark=None, small=False):
    """Power BI-style KPI card: metric, unit, icon, comparison line, trend, sparkline."""
    c, bg, bd = _TONE.get(tone, _TONE["cyan"])
    trend_html = ""
    if trend is not None:
        tv = float(trend)
        tc = UI["ok"] if tv < 0 else UI["err"] if tv > 0 else UI["tx3"]
        ti = "trend-down" if tv < 0 else "trend-up" if tv > 0 else "activity"
        trend_html = (f'<span class="kpi-trend" style="color:{tc}">{icon(ti,12)}{tv:+.0f}%</span>')
    spark_html = sparkline(spark, c) if spark else ""
    foot = ""
    if compare or trend_html or spark_html:
        right = spark_html or trend_html
        left = f'<span class="kpi-cmp">{compare}</span>' if compare else "<span></span>"
        if spark_html and trend_html:
            left = f'<span class="kpi-cmp">{compare} {trend_html}</span>'
        foot = f'<div class="kpi-foot">{left}{right}</div>'
    u = f'<span class="kpi-u">{unit}</span>' if unit else ""
    return (f'<div class="kpi" style="--accent-c:{c};--accent-bg:{bg};--accent-bd:{bd}">'
            f'<div class="kpi-top"><div class="kpi-l">{label}</div>'
            f'<div class="kpi-i">{icon(ic,15)}</div></div>'
            f'<div class="kpi-v{" sm" if small else ""}">{value}{u}</div>{foot}</div>')

def kpi(col, *a, **kw):
    """Render a KPI card into a Streamlit column."""
    col.markdown(kpi_html(*a, **kw), unsafe_allow_html=True)

def pill(text, tone="idle"):
    return f'<span class="pill {tone}"><i></i>{text}</span>'

def workflow(current):
    """Connected workflow cards for the QA/QC pipeline (replaces the numbered step list)."""
    steps = [("Upload", "Simulation results"), ("Column Mapping", "Match export fields"),
             ("Building Details", "Metadata & benchmark"), ("Validation", "Automated QA/QC"),
             ("Results", "Dashboard & report")]
    html = '<div class="wf">'
    for i, (t, d) in enumerate(steps, 1):
        cls = "done" if i < current else "act" if i == current else ""
        num = icon("check", 13) if i < current else str(i)
        html += (f'<div class="wf-s {cls}"><div class="wf-n">{num}</div>'
                 f'<div class="wf-t">{t}</div><div class="wf-d">{d}</div></div>')
        if i < len(steps):
            html += f'<div class="wf-c {"on" if i < current else ""}"><i></i></div>'
    return html + "</div>"

def empty_state(title, desc, ic="inbox"):
    st.markdown(f'''<div class="es">
      <span style="color:{UI['tx3']};display:inline-flex">{icon(ic,40,sw=1.4)}</span>
      <div class="es-t">{title}</div><div class="es-d">{desc}</div></div>''', unsafe_allow_html=True)

def footer(bm_count, synced=True):
    dot = UI["ok"] if synced else UI["warn"]
    state = "Connected" if synced else "Check credentials"
    st.markdown(f'''<div class="ei-foot">
      <div>Energy Intelligence Platform <b>{APP_VERSION}</b> &nbsp;&middot;&nbsp; Buildings Practice</div>
      <div class="ei-foot-r">
        <span><span style="color:{dot}">&#9679;</span> Google Sheets <b>{state}</b></span>
        <span>Benchmarks <b>{bm_count}</b></span>
        <span>Last sync <b>&lt; 60 s</b></span>
        <span>API <b style="color:{UI['ok']}">Operational</b></span>
        <span>Internal use only</span>
      </div></div>''', unsafe_allow_html=True)

# ── Plotly theming ────────────────────────────────────────────────────────────
def style_fig(fig, height=330, title=None, legend=True, ytitle=None, xtitle=None):
    """Apply the Stantec light chart theme to any Plotly figure."""
    fig.update_layout(
        template="plotly_white",
        colorway=CHART_SEQ,
        height=height,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT_SANS, color=UI["tx2"], size=12),
        **({"title": dict(text=title, font=dict(size=15, color=UI["tx"], family=FONT_SERIF),
                          x=0, xanchor="left", y=.97)} if title else {}),
        margin=dict(t=48 if title else 18, b=12, l=10, r=12),
        showlegend=legend,
        legend=dict(orientation="h", y=-0.20, x=0, font=dict(size=11, color=UI["tx3"]),
                    bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor=BRAND["white"], bordercolor=BRAND["fossil"],
                        font=dict(color=BRAND["black"], size=12, family=FONT_SANS)),
        yaxis_title=ytitle, xaxis_title=xtitle,
    )
    fig.update_xaxes(gridcolor=UI["grid"], zerolinecolor=UI["grid"], linecolor=UI["grid"],
                     tickfont=dict(color=UI["tx3"], size=11),
                     title_font=dict(color=UI["tx3"], size=11.5))
    fig.update_yaxes(gridcolor=UI["grid"], zerolinecolor=UI["grid"], linecolor=UI["grid"],
                     tickfont=dict(color=UI["tx3"], size=11),
                     title_font=dict(color=UI["tx3"], size=11.5))
    return fig

def make_gauge(value, title, vmax=100, good_high=True):
    """Radial gauge for score-style KPIs (Performance / QA / Compliance)."""
    v = max(0.0, min(float(value), float(vmax)))
    ratio = v / vmax if vmax else 0
    lvl = ratio if good_high else 1 - ratio
    col = (UI["ok"] if lvl >= 0.75 else BRAND["orange"] if lvl >= 0.5
           else UI["warn"] if lvl >= 0.25 else UI["err"])
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=round(v),
        number=dict(font=dict(size=38, color=BRAND["black"], family=FONT_SERIF), suffix=""),
        gauge=dict(
            axis=dict(range=[0, vmax], tickwidth=1, tickcolor=BRAND["fossil"],
                      tickfont=dict(color=UI["tx3"], size=10)),
            bar=dict(color=col, thickness=0.30),
            bgcolor=BRAND["white"], borderwidth=1, bordercolor=UI["bd"],
            # Warm-grey bands: the scale reads through the bar colour, not tinted fields.
            steps=[dict(range=[0, vmax * .25], color=BRAND["pebble"]),
                   dict(range=[vmax * .25, vmax * .5], color="#EAE6E3"),
                   dict(range=[vmax * .5, vmax * .75], color=BRAND["mist"]),
                   dict(range=[vmax * .75, vmax], color=BRAND["white"])],
            threshold=dict(line=dict(color=BRAND["black"], width=2),
                           thickness=0.8, value=v)),
    ))
    fig.update_layout(
        height=190, margin=dict(t=42, b=6, l=22, r=22),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT_SANS, color=UI["tx2"]),
        title=dict(text=title, font=dict(size=12.5, color=UI["tx3"], family=FONT_SANS),
                   x=.5, xanchor="center", y=.97),
    )
    return fig

# ── Constants ─────────────────────────────────────────────────────────────────
# Google Sheet ID — replace this with your own Sheet ID
# Get it from your sheet URL: docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit
SHEET_ID = st.secrets.get("SHEET_ID", "YOUR_SHEET_ID_HERE")
SHEET_NAME = "Benchmarks"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"
# NECB end-use savings benchmarks live on their own tab. Set this to match the tab name
# in your sheet (the screenshot shows "NECB Savings"; the spec calls it "NECB").
NECB_SHEET_NAME = "NECB Savings"
NECB_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={NECB_SHEET_NAME.replace(' ', '%20')}"
NECB_TOL = 10  # percentage-point tolerance: savings within 10 pts of (or above) the target → pass
MODEL_TYPES      = ["Proposed","Baseline","Existing"]
PROJECT_PHASES   = ["Concept","Schematic Design","Design Development","100% Design","As-Built"]
SOFTWARE_OPTIONS = ["IES VE","EnergyPlus","OpenStudio","eQUEST","Manual / Excel template","Other"]
DHW_BUILDINGS    = ["School","Office","Hospital","Residential","Community Centre","Library"]
END_USE_COLORS   = ["#ef4444","#3b82f6","#8b5cf6","#f59e0b","#06b6d4","#10b981","#f97316"]

# ── GHG emission factors (Canada 2026 provincial values) ──────────────────────
# Natural gas energy content used to convert g CO₂/m³ → kg CO₂/kWh (HHV ≈ 38 MJ/m³).
NG_KWH_PER_M3 = 10.55
# Marketable natural-gas CO₂ factor by province/territory, g CO₂/m³ (Table 1.3, 2026).
NG_CO2_G_PER_M3 = {
    "British Columbia":1966, "Alberta":1962, "Saskatchewan":1920, "Manitoba":1915,
    "Ontario":1921, "Quebec":1926, "New Brunswick":1919, "Nova Scotia":1919,
    "Prince Edward Island":1919, "Newfoundland and Labrador":1919,
    "Yukon":1966, "Northwest Territories":1966, "Nunavut":1966,
}
# Electricity consumption intensity by province/territory, g CO₂e/kWh (Table 5.3, 2026).
ELEC_CO2E_G_PER_KWH = {
    "British Columbia":18, "Alberta":438, "Saskatchewan":631, "Manitoba":2.5,
    "Ontario":59, "Quebec":1.9, "New Brunswick":234, "Nova Scotia":581,
    "Prince Edward Island":234, "Newfoundland and Labrador":17,
    "Yukon":74, "Northwest Territories":420, "Nunavut":800,
}
# Map the cities used in the tool to their province/territory.
CITY_PROVINCE = {
    "Edmonton":"Alberta", "Calgary":"Alberta", "Red Deer":"Alberta", "Lethbridge":"Alberta",
    "Vancouver":"British Columbia", "Victoria":"British Columbia", "Kelowna":"British Columbia",
    "Toronto":"Ontario", "Ottawa":"Ontario", "Hamilton":"Ontario", "London":"Ontario",
    "Winnipeg":"Manitoba", "Regina":"Saskatchewan", "Saskatoon":"Saskatchewan",
    "Montreal":"Quebec", "Quebec City":"Quebec",
    "Halifax":"Nova Scotia", "Fredericton":"New Brunswick",
    "St. John's":"Newfoundland and Labrador", "Charlottetown":"Prince Edward Island",
    "Whitehorse":"Yukon", "Yellowknife":"Northwest Territories", "Iqaluit":"Nunavut",
}

def province_of_city(city):
    """Province for a city: prefer the sheet's Province column, fall back to the built-in map."""
    try:
        for (bt, c, z, sub), v in BENCHMARKS.items():
            if c == city and v.get("province"):
                return v["province"]
    except NameError:
        pass
    return CITY_PROVINCE.get(city)

def gas_factor_for(city):
    """Default natural-gas factor (kg CO₂/kWh) for a city, converted from g CO₂/m³."""
    g = NG_CO2_G_PER_M3.get(province_of_city(city) or "")
    return round(g / NG_KWH_PER_M3 / 1000, 4) if g else 0.0

def elec_factor_for(city):
    """Default electricity factor (kg CO₂e/kWh) for a city."""
    g = ELEC_CO2E_G_PER_KWH.get(province_of_city(city) or "")
    return round(g / 1000, 4) if g is not None else 0.0

def average_benchmarks(bm_list):
    """Average a list of benchmark dicts into a single aggregated benchmark
    (used for province-wide views, e.g. all Schools in Alberta)."""
    n = len(bm_list)
    def mean(key):
        vals = [b[key] for b in bm_list if b.get(key) is not None]
        return round(sum(vals) / len(vals), 1) if vals else 0
    def mean_pos(key):  # average only non-zero values (0 = not provided, e.g. TEDI)
        vals = [b[key] for b in bm_list if b.get(key)]
        return round(sum(vals) / len(vals), 1) if vals else 0
    def mean_array(key):
        arrs = [b[key] for b in bm_list if b.get(key)]
        if not arrs:
            return []
        L = min(len(a) for a in arrs)
        return [round(sum(a[i] for a in arrs) / len(arrs), 1) for i in range(L)]
    med_eui  = mean("median_eui")
    med_tedi = mean_pos("median_tedi")
    return {
        "median_eui":  med_eui,
        "good_eui":    round(med_eui * 0.85, 1),
        "high_eui":    round(med_eui * 1.15, 1),
        "median_tedi": med_tedi,
        "good_tedi":   round(med_tedi * 0.85, 1) if med_tedi else 0,
        "high_tedi":   round(med_tedi * 1.15, 1) if med_tedi else 0,
        "median_ghgi": mean("median_ghgi"),
        "heat_pct": mean("heat_pct"), "cool_pct": mean("cool_pct"), "fan_pct": mean("fan_pct"),
        "ltg_pct": mean("ltg_pct"), "dhw_pct": mean("dhw_pct"), "recept_pct": mean("recept_pct"),
        "pumps_pct": mean("pumps_pct"), "elec_pct": mean("elec_pct"),
        "gas_pct": mean("gas_pct"), "heat_rej_pct": mean("heat_rej_pct"),
        "misc_pct": mean("misc_pct"),
        "pct_data":      mean_array("pct_data"),
        "tedi_pct_data": mean_array("tedi_pct_data"),
        "subtype": "All", "_n": n,
    }

AUTO_MAP = {
    # ── Main energy sources ──
    "electricity_kwh":  ["electricity","total electricity","elec","electricity_kwh","electricity (kwh)",
                         "total elec","elec_kwh","annual electricity","total_kwh","total electricity use"],
    "gas_kwh":          ["gas","natural gas","gas consumption","gas_kwh","naturalgas","natural gas (kwh)",
                         "gas (kwh)","annual gas","ng_kwh","natural_gas","natural_gas_kwh"],
    "area_m2":          ["area","floor area","gfa","area_m2","floor area (m2)","area (m2)",
                         "gross floor area","conditioned area","heated floor area","area_m2"],
    # ── Heating ──
    "heating_kwh":      ["heating","heat","heating_kwh","heating (kwh)","htg","htg energy",
                         "space heating","space_heating","space heating (kwh)","space_heating_kwh",
                         "annual heating","zone heating","heat energy"],
    # ── Cooling ──
    "cooling_kwh":      ["cooling","cool","cooling_kwh","cooling (kwh)","clg","clg energy",
                         "space cooling","space_cooling","space cooling (kwh)","space_cooling_kwh",
                         "annual cooling","zone cooling","heat rejection","heat_rejection",
                         "heat rejection (kwh)","heat_rejection_kwh"],
    # ── Fans — three separate components ──
    "central_fan_kwh":  ["central fan","central_fan","central_fan_kwh","central fan (kwh)",
                         "interior central fan","interior_central_fan","interior_central_fan_kwh",
                         "interior central fans","interior_central_fans","interior_central_fans_kwh",
                         "supply fan","supply_fan","supply_fan_kwh","ahu fan","ahu_fan",
                         "air handling unit","air handling","central air handling",
                         "fans","fan","fans_kwh","fan energy","supply fans","total fan energy"],
    "local_fan_kwh":    ["local fan","local_fan","local_fan_kwh","local fan (kwh)",
                         "interior local fan","interior_local_fan","interior_local_fan_kwh",
                         "interior local fans","interior_local_fans","interior_local_fans_kwh",
                         "fan coil","fan coil unit","fcu","fan_coil","fan_coil_kwh",
                         "zone fan","local ventilation","unit ventilator","local unit"],
    "exhaust_fan_kwh":  ["exhaust fan","exhaust_fan","exhaust_fan_kwh","exhaust fan (kwh)",
                         "exhaust fans","exhaust_fans","exhaust energy","exhaust fan energy",
                         "exhaust_fan_energy","building exhaust","toilet exhaust",
                         "kitchen exhaust","parking exhaust","general exhaust"],
    # ── Lighting ──
    "lighting_kwh":     ["lighting","lights","lighting_kwh","lighting (kwh)","ltg",
                         "interior lighting","interior_lighting","interior_lighting_kwh",
                         "interior lighting (kwh)","annual lighting","zone lighting",
                         "lighting energy","interior lights"],
    # ── DHW ──
    "dhw_kwh":          ["dhw","domestic hot water","dhw_kwh","dhw (kwh)","hot water",
                         "service hot water","shw","dhw energy","dhw_energy",
                         "domestic hot water (kwh)","hot water energy"],
    # ── Pumps ──
    "pumps_kwh":        ["pumps","pump","pumps_kwh","pump energy","pumps (kwh)","heating pumps",
                         "pump_kwh","pumps energy","chw pump","hw pump","condenser pump",
                         "heating pump","cooling pump","pumping energy"],
    # ── Receptacle / plug loads ──
    "receptacle_kwh":   ["receptacle","receptacles","receptacle_kwh","receptacle (kwh)",
                         "plug loads","plug_loads","plug load","plug_load_kwh",
                         "equipment","equipment (kwh)","equipment_kwh","process load",
                         "interior equipment","interior_equipment","interior_equipment_kwh",
                         "miscellaneous","misc load","misc_load","other loads"],
    # ── Interior ceiling / heat rejection misc ──
    "heat_rejection_kwh":["heat rejection","heat_rejection","heat_rejection_kwh",
                          "heat rejection (kwh)","cooling tower","condenser heat",
                          "interior ceiling","interior_ceiling","interior_ce"],
    # ── Exterior lighting ──
    "ext_lighting_kwh": ["exterior lighting","exterior_lighting","exterior_lighting_kwh",
                         "ext lighting","ext_lighting","outdoor lighting","site lighting"],
    # ── Process / other ──
    "process_kwh":      ["process","process load","process_kwh","process (kwh)",
                         "process energy","other","other energy","other_kwh"],
    # ── Other fuels / biomass / district energy (counted in total energy + GHGI) ──
    "other_fuel_kwh":   ["biomass","biomass_kwh","other fuel","other_fuel","other_fuel_kwh",
                         "other resource","other resources","district heating","district energy",
                         "oil","propane","wood","wood pellets","other fuel (kwh)"],
    # ── TEDI — Thermal Energy Demand Intensity (already kWh/m²·yr, NOT divided by area) ──
    "tedi":             ["tedi","tedi (kwh/m2)","tedi_kwh_m2","tedi kwh/m2","tedi (kwh/m2·yr)",
                         "thermal energy demand intensity","thermal demand intensity","tedi_kwh/m2"],
    # ── Unmet (out-of-range comfort) hours — counts, NOT divided by area ──
    "unmet_hours_heating": ["unmet_hours_heating","unmet hours heating","heating unmet hours","unmet heating hours"],
    "unmet_hours_cooling": ["unmet_hours_cooling","unmet hours cooling","cooling unmet hours","unmet cooling hours"],
    "unmet_hours_total":   ["unmet_hours_total","unmet hours total","total unmet hours","unmet hours","unmet_hours"],
}

FIELD_LABELS = {
    "electricity_kwh":   "Electricity (kWh/yr)",
    "gas_kwh":           "Natural Gas (kWh/yr)",
    "area_m2":           "Floor Area (m²)",
    "heating_kwh":       "Space Heating (kWh/yr)",
    "cooling_kwh":       "Space Cooling (kWh/yr)",
    "central_fan_kwh":   "Interior Central Fan / AHU (kWh/yr)",
    "local_fan_kwh":     "Interior Local Fan (kWh/yr)",
    "exhaust_fan_kwh":   "Exhaust Fan (kWh/yr)",
    "lighting_kwh":      "Interior Lighting (kWh/yr)",
    "dhw_kwh":           "DHW (kWh/yr)",
    "pumps_kwh":         "Pumps (kWh/yr)",
    "receptacle_kwh":    "Receptacle / Plug Loads (kWh/yr)",
    "heat_rejection_kwh":"Heat Rejection (kWh/yr)",
    "ext_lighting_kwh":  "Exterior Lighting (kWh/yr)",
    "process_kwh":       "Process / Other (kWh/yr)",
    "other_fuel_kwh":    "Other Fuel / Biomass (kWh/yr)",
    "tedi":              "TEDI (kWh/m²·yr)",
    "unmet_hours_heating": "Unmet Hours — Heating",
    "unmet_hours_cooling": "Unmet Hours — Cooling",
    "unmet_hours_total":   "Unmet Hours — Total",
}

# ── Load benchmarks from Google Sheets ────────────────────────────────────────
@st.cache_data(ttl=60)  # cache for 60 seconds — edits in Google Sheets appear within 1 minute
def load_benchmarks():
    """Read from Google Sheets — keyed by (building_type, city, zone, subtype).
    Subtype allows multiple benchmarks for the same building type + city
    e.g. School · Edmonton · Zone 7 · Boiler+VAV  vs  School · Edmonton · Zone 7 · Heat Pump+DOAS
    """
    try:
        df = pd.read_csv(SHEET_URL, dtype=str)
    except Exception as e:
        st.error(f"❌ Could not load benchmark data from Google Sheets. Check your SHEET_ID. Error: {e}")
        st.stop()
    df.columns = [c.strip() for c in df.columns]
    benchmarks = {}
    for _, row in df.iterrows():
        try:
            btype   = str(row["Building Type"]).strip()
            city    = str(row["City"]).strip()
            zone    = str(row["Climate Zone"]).strip()
            # Province from the sheet (accepts the common "Provience" misspelling);
            # falls back to the built-in city→province map only if the column is blank.
            province = str(row.get("Province", row.get("Provience", "")) or "").strip() or CITY_PROVINCE.get(city, "")
            # Subtype is optional — falls back to "General" if column missing or empty
            subtype = str(row.get("Subtype", "") or "").strip() or "General"
            pct_raw = str(row["Percentile Data (comma separated)"]).strip()
            pct_data = [float(x.strip()) for x in pct_raw.split(",") if x.strip()]
            median = float(row["Median EUI"])
            try:
                med_tedi = float(row.get("Median TEDI", "") or 0)
            except (TypeError, ValueError):
                med_tedi = 0.0
            # Optional TEDI percentile distribution. If a dedicated column isn't provided
            # but a Median TEDI is, approximate the shape from the EUI distribution scaled
            # to the TEDI median so the chart still renders.
            tedi_pct_raw = str(row.get("TEDI Percentile Data (comma separated)", "") or "").strip()
            tedi_pct_data = [float(x.strip()) for x in tedi_pct_raw.split(",") if x.strip()]
            if not tedi_pct_data and med_tedi and pct_data and median:
                tedi_pct_data = [round(v * med_tedi / median, 1) for v in pct_data]
            benchmarks[(btype, city, zone, subtype)] = {
                "median_eui":  median,
                "good_eui":    round(median * 0.85, 1),   # −15% of median
                "high_eui":    round(median * 1.15, 1),   # +15% of median
                "median_tedi": med_tedi,
                "good_tedi":   round(med_tedi * 0.85, 1) if med_tedi else 0,
                "high_tedi":   round(med_tedi * 1.15, 1) if med_tedi else 0,
                "median_ghgi": float(row["Median GHGI"]),
                "heat_pct":    float(row["Heating %"]),
                "cool_pct":    float(row["Cooling %"]),
                "fan_pct":     float(row["Fan %"]),
                "ltg_pct":     float(row["Lighting %"]),
                "dhw_pct":     float(row["DHW %"]),
                "recept_pct":  float(row["Receptacle %"]),
                "pumps_pct":   float(row["Pumps %"]),
                "elec_pct":    float(row.get("Electricity %")   or row.get("Electricity")  or 0),
                "gas_pct":     float(row.get("NaturalGas %")    or row.get("Natural Gas %") or row.get("NaturalGas") or 0),
                "heat_rej_pct":float(row.get("Heat Rejection %") or row.get("HeatRejection %") or 0),
                "misc_pct":    float(row.get("Miscellaneous %")  or row.get("Miscellaneous") or 0),
                "pct_data":    pct_data,
                "tedi_pct_data": tedi_pct_data,
                "subtype":     subtype,
                "province":    province,
                "project_year":   str(row.get("Project Year/Date", row.get("Project Year", "")) or "").strip(),
                "audit_new":      str(row.get("Audit/New", "") or "").strip(),
                "heating_system": str(row.get("Heating System", "") or "").strip(),
            }
        except Exception:
            continue
    return benchmarks

def reload_benchmarks():
    """Clear cache and reload from Google Sheets."""
    load_benchmarks.clear()
    st.rerun()

# ── NECB end-use savings benchmarks (separate tab) ───────────────────────────
# Maps each internal energy field to the NECB end-use name(s) used on the sheet.
NECB_ENDUSES = [
    ("electricity_kwh",    "Electricity",       ["electricity","elec","total electricity"]),
    ("gas_kwh",            "Natural Gas",       ["natural gas","gas","naturalgas","natural_gas"]),
    ("lighting_kwh",       "Interior Lighting", ["interior lighting","lighting","lights","ltg","interior_lighting"]),
    ("heating_kwh",        "Space Heating",     ["space heating","heating","heat","space_heating"]),
    ("cooling_kwh",        "Space Cooling",     ["space cooling","cooling","cool","space_cooling"]),
    ("pumps_kwh",          "Pumps",             ["pumps","pump"]),
    ("heat_rejection_kwh", "Heat Rejection",    ["heat rejection","heat_rejection","heat rej"]),
    ("central_fan_kwh",    "Central Fan",       ["central fan","interior central fans","central fans","ahu","interior_central_fans"]),
    ("local_fan_kwh",      "Local Fan",         ["local fan","interior local fans","local fans","interior_local_fans"]),
    ("exhaust_fan_kwh",    "Exhaust Fan",       ["exhaust fan","exhaust fans","exhaust_fans"]),
    ("dhw_kwh",            "DHW",               ["dhw","domestic hot water","hot water","service hot water"]),
    ("receptacle_kwh",     "Receptacle",        ["receptacle","receptacle loads","plug loads","receptacles"]),
    ("total_kwh",          "Total Energy",      ["total energy","total","total energy use","total_kwh","total kwh"]),
]

# Fallback savings targets (used if the NECB tab can't be read or is missing a value).
DEFAULT_NECB_SAVINGS = {
    "electricity_kwh":30, "gas_kwh":40, "lighting_kwh":76, "heating_kwh":76, "cooling_kwh":30,
    "pumps_kwh":40, "heat_rejection_kwh":76, "central_fan_kwh":76, "local_fan_kwh":30,
    "exhaust_fan_kwh":40, "dhw_kwh":76, "receptacle_kwh":76, "total_kwh":40,
}

def _match_necb_key(name):
    """Resolve a sheet end-use label to an internal field key."""
    n = str(name).strip().lower()
    for key, _label, aliases in NECB_ENDUSES:   # exact match first
        if n == key or n in aliases:
            return key
    for key, _label, aliases in NECB_ENDUSES:   # then substring
        if any(a in n for a in aliases):
            return key
    return None

# NECB Savings tab end-use columns → internal field keys, matched on a normalized
# (lowercase, alphanumeric-only, trailing "kwh"/"%" stripped) form so it tolerates the
# per-end-use names (electricity, natural_gas, space_heating, …), the legacy "Heating %"
# names, and small spelling/casing differences.
NECB_NORM_ALIASES = {
    "electricity": ["electricity_kwh"], "elec": ["electricity_kwh"], "totalelectricity": ["electricity_kwh"],
    "naturalgas": ["gas_kwh"], "gas": ["gas_kwh"], "natgas": ["gas_kwh"],
    "interiorlighting": ["lighting_kwh"], "lighting": ["lighting_kwh"], "lights": ["lighting_kwh"],
    "spaceheating": ["heating_kwh"], "heating": ["heating_kwh"], "heat": ["heating_kwh"],
    "spacecooling": ["cooling_kwh"], "cooling": ["cooling_kwh"], "cool": ["cooling_kwh"],
    "pumps": ["pumps_kwh"], "pump": ["pumps_kwh"],
    "heatrejection": ["heat_rejection_kwh"], "heatrej": ["heat_rejection_kwh"],
    "interiorcentralfans": ["central_fan_kwh"], "interiorcentralfan": ["central_fan_kwh"],
    "centralfans": ["central_fan_kwh"], "centralfan": ["central_fan_kwh"],
    "interiorlocalfans": ["local_fan_kwh"], "interiorlocalfan": ["local_fan_kwh"],
    "localfans": ["local_fan_kwh"], "localfan": ["local_fan_kwh"],
    "exhaustfans": ["exhaust_fan_kwh"], "exhaustfan": ["exhaust_fan_kwh"],
    "fan": ["central_fan_kwh", "local_fan_kwh", "exhaust_fan_kwh"],   # legacy combined "Fan %"
    "fans": ["central_fan_kwh", "local_fan_kwh", "exhaust_fan_kwh"],
    "dhw": ["dhw_kwh"], "domestichotwater": ["dhw_kwh"], "hotwater": ["dhw_kwh"],
    "receptacle": ["receptacle_kwh"], "receptacles": ["receptacle_kwh"],
    "receptacleloads": ["receptacle_kwh"], "plugloads": ["receptacle_kwh"],
    "total": ["total_kwh"], "totalenergy": ["total_kwh"], "totalkwh": ["total_kwh"],
}

def _necb_norm(s):
    n = "".join(ch for ch in str(s).lower() if ch.isalnum())
    if n.endswith("kwh"):
        n = n[:-3]
    return n

def _necb_col_keys(colname):
    """Internal field key(s) a NECB-sheet end-use column maps to, or None for metadata cols."""
    return NECB_NORM_ALIASES.get(_necb_norm(colname))

def _necb_code_from_version(v):
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    try:    s = str(int(float(s)))   # "2020" / "2020.0" → "2020"
    except Exception: pass
    return s if s.upper().startswith("NECB") else f"NECB {s}"

@st.cache_data(ttl=60)
def load_necb_savings():
    """Read the NECB Savings tab (wide format: one row per building type / city / NECB
    version, with per-end-use savings columns). Returns {"rows":[...], "codes":[...]}."""
    empty = {"rows": [], "codes": []}
    try:
        df = pd.read_csv(NECB_SHEET_URL, dtype=str)
    except Exception:
        return empty
    df.columns = [str(c).strip() for c in df.columns]
    low = {c.lower(): c for c in df.columns}
    bt_col   = next((low[c] for c in ["building type","buildingtype","type"] if c in low), None)
    city_col = next((low[c] for c in ["city"] if c in low), None)
    prov_col = next((low[c] for c in ["province","provience"] if c in low), None)
    ver_col  = next((low[c] for c in ["necb version","version","compliance code","code","necb"] if c in low), None)
    # Any column that resolves to an end-use key is treated as a savings target.
    sav_cols = {c: _necb_col_keys(c) for c in df.columns if _necb_col_keys(c)}
    if not sav_cols:
        return empty
    rows, codes = [], []
    for _, r in df.iterrows():
        code = (_necb_code_from_version(r[ver_col]) if ver_col else None) or "NECB"
        if code not in codes:
            codes.append(code)
        savings = {}
        for actual, keys in sav_cols.items():
            raw = str(r[actual]).replace("%", "").replace(",", "").strip()
            if raw == "" or raw.lower() == "nan":
                continue
            try:
                pctv = float(raw)
            except Exception:
                continue
            if pctv != pctv:   # NaN guard
                continue
            for key in keys:
                savings[key] = pctv
        rows.append({
            "building_type": (str(r[bt_col]).strip()   if bt_col   else ""),
            "city":          (str(r[city_col]).strip() if city_col else ""),
            "province":      (str(r[prov_col]).strip() if prov_col else ""),
            "code":          code,
            "savings":       savings,
        })
    return {"rows": rows, "codes": codes or ["NECB 2020"]}

def _avg_savings(maps):
    keys = set().union(*[m.keys() for m in maps]) if maps else set()
    out = {}
    for k in keys:
        vals = [m[k] for m in maps if k in m]
        if vals:
            out[k] = round(sum(vals) / len(vals), 1)
    return out

def necb_savings_for(necb, building_type, city, code):
    """Resolve the savings target map for a building type / city / code, with fallbacks:
    exact (type+city+code) → type+code (avg across cities) → type → any."""
    rows = necb.get("rows", [])
    bt = (building_type or "").strip().lower()
    ct = (city or "").strip().lower()
    exact = [r["savings"] for r in rows if r["building_type"].lower()==bt and r["city"].lower()==ct and r["code"]==code]
    if exact:
        return exact[0]
    bc = [r["savings"] for r in rows if r["building_type"].lower()==bt and r["code"]==code]
    if bc:
        return _avg_savings(bc)
    b = [r["savings"] for r in rows if r["building_type"].lower()==bt]
    if b:
        return _avg_savings(b)
    return _avg_savings([r["savings"] for r in rows]) if rows else {}

def build_necb_rows(vals, ref_vals, savings_map):
    """Per-end-use proposed-vs-reference savings rows for the NECB QA table."""
    def g(d, k):
        try: return float((d or {}).get(k) or 0)
        except Exception: return 0.0
    def total(d): return g(d,"electricity_kwh") + g(d,"gas_kwh") + g(d,"other_fuel_kwh")
    rows = []
    for key, label, _aliases in NECB_ENDUSES:
        prop = total(vals)     if key == "total_kwh" else g(vals, key)
        ref  = total(ref_vals) if key == "total_kwh" else g(ref_vals, key)
        bench = savings_map.get(key)
        if ref:
            savings = (ref - prop) / ref * 100
            sav_txt = f"{savings:+.0f}%"
            auto = "⚪" if bench is None else ("🟢" if savings >= bench - NECB_TOL else "🔴")
        else:
            sav_txt = "n/a"; auto = "⚪"
        rows.append({
            "End Use": label,
            "Proposed": round(prop, 1),
            "Reference Model": round(ref, 1),
            "Savings": sav_txt,
            "Benchmark": (f"{bench:.0f}%" if bench is not None else "—"),
            "Auto Flag": auto,
        })
    return rows


# ── Writing back to Google Sheets ─────────────────────────────────────────────
# Two supported routes, tried in this order:
#   1. SHEETS_WEBHOOK_URL  — an Apps Script web app bound to the spreadsheet.
#      No Google Cloud project needed, so it works where IT restricts GCP.
#   2. GOOGLE_SERVICE_ACCOUNT (or gcp_service_account) — the Sheets REST API
#      with a service-account JWT. The standard route.
# Both are configured in .streamlit/secrets.toml; see GOOGLE_SHEETS_SETUP.md.

# Submissions land here as Pending and are promoted to PROJECT_SHEET_NAME on
# approval. The app only ever writes to the submissions tab.
SUBMISSIONS_SHEET_NAME = "Submissions"
PROJECT_SHEET_NAME = "Project Results"

# Columns the app sends, in order. The Apps Script prepends Submission ID and
# Submitted At and appends Status / Reviewed By / Reviewed On / Review Comment,
# so a submitter cannot set their own approval state.
SUBMISSION_HEADERS = [
    "Modeller", "Modeller Email",
    "Project Name", "Building Type", "City", "Climate Zone", "Subtype",
    "Software", "Model Type", "Phase", "Date", "Area (m²)", "Total EUI", "TEDI",
    "Elec EUI", "Gas EUI", "Heating EUI", "Cooling EUI", "Fan EUI",
    "Lighting EUI", "DHW EUI", "Receptacle EUI", "Pumps EUI", "GHGI",
    "Percentile", "QA Flags", "Notes",
]
# Columns the Apps Script adds around the client's. The app only needs these
# when writing directly with a service account, where nothing stamps them.
SUBMISSION_LEADING = ["Submission ID", "Submitted At"]
SUBMISSION_TRAILING = ["Status", "Reviewed By", "Reviewed On", "Review Comment"]
SUBMISSION_SHEET_HEADERS = SUBMISSION_LEADING + SUBMISSION_HEADERS + SUBMISSION_TRAILING

# Kept for the manual-copy fallback and any older sheets.
PROJECT_RESULT_HEADERS = SUBMISSION_HEADERS


def _valid_email(value):
    """Loose sanity check — catches typos, not a deliverability guarantee."""
    import re
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$", (value or "").strip()))

_HTTP_TIMEOUT = 20            # seconds — never let a write hang the app
_TOKEN_CACHE = {"token": None, "expires": 0.0}


def _service_account_creds():
    """Service-account mapping from secrets, under either accepted key name."""
    import json
    for key in ("GOOGLE_SERVICE_ACCOUNT", "gcp_service_account"):
        try:
            raw = st.secrets.get(key, None)
        except Exception:
            raw = None
        if raw:
            try:
                return json.loads(raw) if isinstance(raw, str) else dict(raw)
            except Exception as e:
                raise ValueError(f"{key} in secrets is not valid JSON: {e}")
    return None


def sheets_account_email():
    """Email to share the spreadsheet with, or '' when no key is configured."""
    try:
        creds = _service_account_creds()
        return (creds or {}).get("client_email", "")
    except Exception:
        return ""


def _sheets_access_token(creds):
    """Signed JWT exchanged for an OAuth access token, cached until it expires."""
    import time, json, base64, requests
    if _TOKEN_CACHE["token"] and _TOKEN_CACHE["expires"] > time.time() + 60:
        return _TOKEN_CACHE["token"]

    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import padding

    def _b64(obj):
        return base64.urlsafe_b64encode(json.dumps(obj).encode()).rstrip(b"=").decode()

    now = int(time.time())
    header = _b64({"alg": "RS256", "typ": "JWT"})
    claim = _b64({
        "iss": creds["client_email"],
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now, "exp": now + 3600,
    })
    # Secrets managers often turn the key's "\n" escapes into literal backslash-n.
    pem = creds["private_key"].replace("\\n", "\n").encode()
    key = serialization.load_pem_private_key(pem, password=None)
    sig = base64.urlsafe_b64encode(
        key.sign(f"{header}.{claim}".encode(), padding.PKCS1v15(), hashes.SHA256())
    ).rstrip(b"=").decode()

    resp = requests.post(
        "https://oauth2.googleapis.com/token",
        data={"grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
              "assertion": f"{header}.{claim}.{sig}"},
        timeout=_HTTP_TIMEOUT,
    )
    token = resp.json().get("access_token") if resp.ok else None
    if not token:
        detail = resp.text[:300]
        if "invalid_grant" in detail:
            detail += ("  (an 'invalid_grant' usually means the private key was "
                       "pasted incorrectly or the server clock is off)")
        raise RuntimeError(f"Could not get an access token: {detail}")

    _TOKEN_CACHE.update(token=token, expires=time.time() + 3300)
    return token


def _ensure_tab(sheet_name, token, headers):
    """Create the tab and write its header row if it does not exist yet."""
    import requests
    meta = requests.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
        "?fields=sheets.properties.title",
        headers={"Authorization": f"Bearer {token}"}, timeout=_HTTP_TIMEOUT)
    if not meta.ok:
        return  # let the append call report the real error
    titles = [s["properties"]["title"] for s in meta.json().get("sheets", [])]
    if sheet_name in titles:
        return
    requests.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}:batchUpdate",
        headers={"Authorization": f"Bearer {token}"},
        json={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
        timeout=_HTTP_TIMEOUT)
    if headers:
        from urllib.parse import quote
        requests.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/"
            f"{quote(f'{sheet_name}!A1')}:append"
            "?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
            headers={"Authorization": f"Bearer {token}"},
            json={"values": [headers]}, timeout=_HTTP_TIMEOUT)


def _append_via_webhook(webhook, sheet_name, row, headers=None):
    """POST the row to an Apps Script web app bound to the spreadsheet."""
    import requests
    payload = {"sheet": sheet_name, "row": row, "headers": headers or []}
    try:
        secret = st.secrets.get("SHEETS_WEBHOOK_TOKEN", None)
    except Exception:
        secret = None
    if secret:
        payload["token"] = secret
    try:
        resp = requests.post(webhook, json=payload, timeout=_HTTP_TIMEOUT)
    except Exception as e:
        return False, f"Could not reach the webhook: {e}"
    body = (resp.text or "").strip()
    # Apps Script answers 200 with a JSON body even for its own errors.
    if resp.ok and ('"ok":true' in body.replace(" ", "") or body.lower() == "ok"):
        try:
            import json
            return True, json.loads(body).get("id") or "Success"
        except Exception:
            return True, "Success"
    if resp.status_code in (301, 302) or "accounts.google.com" in body:
        return False, ("The web app redirected to a Google sign-in — redeploy it "
                       "with Execute as: Me and Who has access: Anyone.")
    return False, f"Webhook responded {resp.status_code}: {body[:300]}"


def append_to_google_sheet(sheet_name: str, row: list, headers: list = None,
                          pending_workflow: bool = False):
    """Append one row to the spreadsheet. Returns (ok, message).

    With ``pending_workflow`` the row is a review submission. The Apps Script
    stamps the id, timestamp and Pending status itself; when writing directly
    with a service account there is no script, so the app fills those columns
    in — otherwise the row would arrive with no status at all.
    """
    try:
        try:
            webhook = st.secrets.get("SHEETS_WEBHOOK_URL", None)
        except Exception:
            webhook = None
        if webhook:
            return _append_via_webhook(webhook, sheet_name, row, headers)

        creds = _service_account_creds()
        if not creds:
            return False, ("No Google credentials in secrets — add either "
                           "SHEETS_WEBHOOK_URL or GOOGLE_SERVICE_ACCOUNT to "
                           ".streamlit/secrets.toml (see GOOGLE_SHEETS_SETUP.md).")
        if not SHEET_ID or SHEET_ID == "YOUR_SHEET_ID_HERE":
            return False, "SHEET_ID is not set in secrets."

        import requests
        from urllib.parse import quote
        if pending_workflow:
            from datetime import datetime
            row = ([f"MAN-{datetime.now():%Y%m%d-%H%M%S}",
                    datetime.now().isoformat(timespec="seconds")]
                   + list(row) + ["Pending", "", "", ""])
            headers = SUBMISSION_SHEET_HEADERS

        token = _sheets_access_token(creds)
        _ensure_tab(sheet_name, token, headers)

        # The range must be quoted — a tab name containing a space (such as
        # "Project Results") is otherwise rejected as an unparseable range.
        rng = quote(f"'{sheet_name}'!A1")
        resp = requests.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{rng}"
            ":append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS",
            headers={"Authorization": f"Bearer {token}"},
            json={"values": [row]}, timeout=_HTTP_TIMEOUT)

        if resp.ok:
            return True, "Success"

        detail = resp.text[:400]
        email = creds.get("client_email", "the service account")
        if resp.status_code == 403 and "PERMISSION_DENIED" in detail:
            if "has not been used" in detail or "disabled" in detail:
                return False, ("The Google Sheets API is not enabled for this "
                               "service account's project. Enable it in the Google "
                               "Cloud console, wait a minute, and try again.")
            return False, (f"{email} does not have access to the spreadsheet. "
                           f"Open the sheet, press Share, and give that address "
                           f"Editor access.")
        if resp.status_code == 404:
            return False, (f"Spreadsheet {SHEET_ID} not found — check SHEET_ID "
                           f"in secrets.")
        return False, f"Sheets API error {resp.status_code}: {detail}"
    except Exception as e:
        return False, str(e)

# ── Helper functions ───────────────────────────────────────────────────────────
def guess_mapping(headers):
    mapping = {}
    for field, aliases in AUTO_MAP.items():
        match = next((h for h in headers if h.strip().lower() in aliases), "")
        mapping[field] = match
    return mapping

def calc_percentile(eui, pct_data):
    above = sum(1 for v in sorted(pct_data) if v > eui)
    return round((above / len(pct_data)) * 100)

def calculate_kpis(vals, area_override=None, ef_elec=0.0, ef_gas=0.0, ef_other=0.0):
    area    = float(area_override) if area_override else float(vals.get("area_m2") or 1)
    elec    = float(vals.get("electricity_kwh")    or 0)
    gas     = float(vals.get("gas_kwh")            or 0)
    other_fuel = float(vals.get("other_fuel_kwh")  or 0)   # biomass / district energy / oil etc.
    heat    = float(vals.get("heating_kwh")        or 0)
    cool    = float(vals.get("cooling_kwh")        or 0)
    central_fan  = float(vals.get("central_fan_kwh")  or 0)
    local_fan    = float(vals.get("local_fan_kwh")    or 0)
    exhaust_fan  = float(vals.get("exhaust_fan_kwh")  or 0)
    fans         = central_fan + local_fan + exhaust_fan  # total fans = central + local + exhaust
    ltg     = float(vals.get("lighting_kwh")       or 0)
    dhw     = float(vals.get("dhw_kwh")            or 0)
    pumps   = float(vals.get("pumps_kwh")          or 0)
    recept  = float(vals.get("receptacle_kwh")     or 0)
    heat_rej= float(vals.get("heat_rejection_kwh") or 0)
    ext_ltg = float(vals.get("ext_lighting_kwh")   or 0)
    process = float(vals.get("process_kwh")        or 0)
    tedi    = float(vals.get("tedi")               or 0)   # already kWh/m²·yr — not divided by area
    # Unmet (out-of-range comfort) hours — raw counts, NOT divided by area
    unmet_h = float(vals.get("unmet_hours_heating") or 0)
    unmet_c = float(vals.get("unmet_hours_cooling") or 0)
    unmet_t = float(vals.get("unmet_hours_total")   or 0)
    total   = elec + gas + other_fuel
    def eui(v): return round(v / area, 1) if area else 0

    # GHGI uses user-supplied emission factors (kgCO₂e/kWh). If none are provided,
    # GHGI is left undefined (None) rather than shown as a misleading value.
    ef_elec, ef_gas, ef_other = float(ef_elec or 0), float(ef_gas or 0), float(ef_other or 0)
    factors_on = (ef_elec > 0) or (ef_gas > 0) or (ef_other > 0)
    ghgi = round((elec*ef_elec + gas*ef_gas + other_fuel*ef_other) / area, 1) if (factors_on and area) else None

    return {
        "area": area, "total_energy": total,
        "total_eui":      eui(total),
        "tedi":           round(tedi, 1),
        "unmet_heating":  int(round(unmet_h)),
        "unmet_cooling":  int(round(unmet_c)),
        "unmet_total":    int(round(unmet_t)),
        "elec_eui":       eui(elec),
        "gas_eui":        eui(gas),
        "other_fuel_eui": eui(other_fuel),
        "heat_eui":       eui(heat),
        "cool_eui":       eui(cool),
        "fan_eui":         eui(fans),
        "central_fan_eui": eui(central_fan),
        "local_fan_eui":   eui(local_fan),
        "exhaust_fan_eui": eui(exhaust_fan),
        "ltg_eui":        eui(ltg),
        "dhw_eui":        eui(dhw),
        "pumps_eui":      eui(pumps),
        "recept_eui":     eui(recept),
        "heat_rej_eui":   eui(heat_rej),
        "ext_ltg_eui":    eui(ext_ltg),
        "process_eui":    eui(process),
        "ghgi":           ghgi,
    }

def generate_flags(kpis, bm, building_type):
    flags = []
    if not bm:
        flags.append(("info","ℹ️","No benchmark found for this combination. KPIs calculated but no comparison available."))
        return flags
    # All thresholds use ±15% of median EUI
    med       = bm["median_eui"]
    good      = bm["good_eui"]    # median * 0.85
    high      = bm["high_eui"]    # median * 1.15

    def end_med(pct): return med * pct / 100
    def end_good(pct): return end_med(pct) * 0.85
    def end_high(pct): return end_med(pct) * 1.15

    fan_med  = end_med(bm["fan_pct"])
    heat_med = end_med(bm["heat_pct"])
    cool_med = end_med(bm["cool_pct"])
    ltg_med  = end_med(bm["ltg_pct"])

    # Total EUI
    if kpis["total_eui"] > high:
        flags.append(("fail","✗",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is more than 15% above the benchmark median "
            f"({med} kWh/m²·yr). High flag threshold: {high} kWh/m²·yr."))
    elif kpis["total_eui"] < good * 0.6:
        flags.append(("warn","⚠",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is unusually low — confirm all end-uses are modelled."))
    elif kpis["total_eui"] <= good:
        flags.append(("pass","✓",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is below the benchmark median ({med} kWh/m²·yr) — good performance."))
    else:
        flags.append(("pass","✓",
            f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is within ±15% of the benchmark median ({med} kWh/m²·yr)."))

    # Fan energy
    if fan_med > 0 and kpis["fan_eui"] > end_high(bm["fan_pct"]):
        flags.append(("fail","✗",
            f"Fan energy ({kpis['fan_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(fan_med,1)} kWh/m²·yr) — verify AHU schedules and fan sizing."))
    elif kpis["fan_eui"] == 0:
        flags.append(("warn","⚠","Fan energy is zero — confirm fan systems are included in the model."))
    else:
        flags.append(("pass","✓",f"Fan energy ({kpis['fan_eui']} kWh/m²·yr) is within expected range."))

    # Cooling
    if kpis["cool_eui"] == 0:
        flags.append(("warn","⚠","No cooling energy — confirm whether a mechanical cooling system exists."))
    elif kpis["cool_eui"] < cool_med * 0.25:
        flags.append(("warn","⚠",
            f"Cooling energy ({kpis['cool_eui']} kWh/m²·yr) is very low compared to benchmark — verify cooling system modelling."))
    elif kpis["cool_eui"] > end_high(bm["cool_pct"]):
        flags.append(("fail","✗",
            f"Cooling energy ({kpis['cool_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(cool_med,1)} kWh/m²·yr)."))

    # Heating
    if heat_med > 0 and kpis["heat_eui"] > end_high(bm["heat_pct"]):
        flags.append(("fail","✗",
            f"Heating energy ({kpis['heat_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(heat_med,1)} kWh/m²·yr) — review envelope and heating schedules."))
    elif heat_med > 0 and kpis["heat_eui"] <= end_high(bm["heat_pct"]):
        flags.append(("pass","✓",f"Heating energy ({kpis['heat_eui']} kWh/m²·yr) is within expected range."))

    # Simultaneous heating + cooling
    if kpis["heat_eui"] > end_high(bm["heat_pct"]) and kpis["cool_eui"] > end_high(bm["cool_pct"]):
        flags.append(("warn","⚠","Both heating and cooling are above the high threshold — possible simultaneous heating/cooling or control issue."))

    # DHW
    if building_type in DHW_BUILDINGS and kpis["dhw_eui"] == 0:
        flags.append(("fail","✗",f"DHW energy is zero for a {building_type} — domestic hot water is typically required."))
    elif kpis["dhw_eui"] > 0:
        flags.append(("pass","✓",f"DHW energy present ({kpis['dhw_eui']} kWh/m²·yr)."))

    # Lighting
    if ltg_med > 0 and kpis["ltg_eui"] > end_high(bm["ltg_pct"]):
        flags.append(("fail","✗",
            f"Lighting energy ({kpis['ltg_eui']} kWh/m²·yr) is more than 15% above benchmark median "
            f"({round(ltg_med,1)} kWh/m²·yr) — verify LPD values against NECB."))

    # GHGI (only when emission factors were provided)
    if kpis.get("ghgi") is not None:
        if kpis["ghgi"] > bm["median_ghgi"] * 1.15:
            flags.append(("fail","✗",
                f"GHGI ({kpis['ghgi']} kgCO₂e/m²·yr) is more than 15% above benchmark median "
                f"({bm['median_ghgi']} kgCO₂e/m²·yr) — review fuel mix and emission factors."))
        else:
            flags.append(("pass","✓",
                f"GHGI ({kpis['ghgi']} kgCO₂e/m²·yr) is within acceptable range (benchmark median: {bm['median_ghgi']})."))
    return flags

def status_color(val, good, median, high):
    if val <= good:   return "🟢"
    if val <= median: return "🟡"
    if val <= high:   return "🟠"
    return "🔴"

def build_comparison_rows(kpis, bm):
    """Per-metric comparison rows with the threshold-vs-median QA Status (the Auto Flag)."""
    def be(pct): return round(bm["median_eui"] * pct / 100, 1)
    def bs(val, pct):
        med = bm["median_eui"] * pct / 100
        return status_color(val, med*0.85, med, med*1.15)
    def pdiff(your, med):  # percentage difference vs benchmark median
        return f"{(your-med)/med*100:+.0f}%" if med else "n/a"
    rows = [
        {"End Use":"Total EUI (kWh/m²·yr)","Your Model":kpis["total_eui"],"Benchmark Median":bm["median_eui"],
         "Difference":pdiff(kpis["total_eui"], bm["median_eui"]),
         "QA Status":status_color(kpis["total_eui"],bm["good_eui"],bm["median_eui"],bm["high_eui"])},
    ]
    # TEDI sits right next to Total EUI when both a benchmark and a model value are present.
    if bm.get("median_tedi") and kpis.get("tedi", 0) > 0:
        rows.append(
            {"End Use":"TEDI (kWh/m²·yr)","Your Model":kpis["tedi"],"Benchmark Median":bm["median_tedi"],
             "Difference":pdiff(kpis["tedi"], bm["median_tedi"]),
             "QA Status":status_color(kpis["tedi"],bm["good_tedi"],bm["median_tedi"],bm["high_tedi"])})
    rows += [
        {"End Use":"Heating EUI (kWh/m²·yr)","Your Model":kpis["heat_eui"],"Benchmark Median":be(bm["heat_pct"]),
         "Difference":pdiff(kpis["heat_eui"], be(bm["heat_pct"])),"QA Status":bs(kpis["heat_eui"],bm["heat_pct"])},
        {"End Use":"Cooling EUI (kWh/m²·yr)","Your Model":kpis["cool_eui"],"Benchmark Median":be(bm["cool_pct"]),
         "Difference":pdiff(kpis["cool_eui"], be(bm["cool_pct"])),"QA Status":bs(kpis["cool_eui"],bm["cool_pct"])},
        {"End Use":"Fan EUI (kWh/m²·yr)","Your Model":kpis["fan_eui"],"Benchmark Median":be(bm["fan_pct"]),
         "Difference":pdiff(kpis["fan_eui"], be(bm["fan_pct"])),"QA Status":bs(kpis["fan_eui"],bm["fan_pct"])},
        {"End Use":"Lighting EUI (kWh/m²·yr)","Your Model":kpis["ltg_eui"],"Benchmark Median":be(bm["ltg_pct"]),
         "Difference":pdiff(kpis["ltg_eui"], be(bm["ltg_pct"])),"QA Status":bs(kpis["ltg_eui"],bm["ltg_pct"])},
        {"End Use":"DHW EUI (kWh/m²·yr)","Your Model":kpis["dhw_eui"],"Benchmark Median":be(bm["dhw_pct"]),
         "Difference":pdiff(kpis["dhw_eui"], be(bm["dhw_pct"])),"QA Status":bs(kpis["dhw_eui"],bm["dhw_pct"])},
        {"End Use":"Receptacle EUI (kWh/m²·yr)","Your Model":kpis.get("recept_eui",0),"Benchmark Median":be(bm["recept_pct"]),
         "Difference":pdiff(kpis.get("recept_eui",0), be(bm["recept_pct"])),"QA Status":bs(kpis.get("recept_eui",0),bm["recept_pct"])},
        {"End Use":"Pumps EUI (kWh/m²·yr)","Your Model":kpis["pumps_eui"],"Benchmark Median":be(bm["pumps_pct"]),
         "Difference":pdiff(kpis["pumps_eui"], be(bm["pumps_pct"])),"QA Status":bs(kpis["pumps_eui"],bm["pumps_pct"])},
    ]
    if kpis.get("ghgi") is not None:
        rows.append(
            {"End Use":"GHGI (kgCO₂e/m²·yr)","Your Model":kpis["ghgi"],"Benchmark Median":bm["median_ghgi"],
             "Difference":pdiff(kpis["ghgi"], bm["median_ghgi"]),
             "QA Status":status_color(kpis["ghgi"],bm["median_ghgi"]*0.85,bm["median_ghgi"],bm["median_ghgi"]*1.15)})
    return rows

def flags_from_comparison(rows, kpis, bm, building_type):
    """Build the QA/QC Flags list directly from the Auto Flag column so the two never
    disagree. Special-cases a missing required system (e.g. DHW = 0), which a pure
    threshold check would mistakenly read as "low = good"."""
    LEVEL = {"🟢":"pass", "🟡":"pass", "🟠":"warn", "🔴":"fail"}
    ICON  = {"pass":"✓", "warn":"⚠", "fail":"✗"}
    HINT  = {
        "Total EUI":  "review overall model inputs, schedules and envelope.",
        "TEDI":       "review envelope, airtightness and ventilation heat recovery.",
        "Heating":    "review envelope inputs and heating schedules.",
        "Cooling":    "verify cooling system sizing and controls.",
        "Fan":        "verify AHU schedules and fan sizing.",
        "Lighting":   "verify LPD values against NECB.",
        "DHW":        "review DHW demand and system efficiency.",
        "Receptacle": "verify plug-load assumptions.",
        "Pumps":      "verify pump sizing and run hours.",
        "GHGI":       "review fuel mix and emission factors.",
    }
    def hint_for(name):
        return next((v for k, v in HINT.items() if name.startswith(k)), "")

    flags = []
    for r in rows:
        name, status = r["End Use"], r["QA Status"]
        your, bench, diff = r["Your Model"], r["Benchmark Median"], r["Difference"]

        # Missing required systems — thresholds alone would call a 0 value "good".
        if name.startswith("DHW") and kpis["dhw_eui"] == 0 and building_type in DHW_BUILDINGS:
            flags.append(("fail","✗", f"DHW energy is zero for a {building_type} — domestic hot water is typically required."))
            continue
        if name.startswith("Cooling") and kpis["cool_eui"] == 0:
            flags.append(("warn","⚠", "No cooling energy — confirm whether a mechanical cooling system exists."))
            continue
        if name.startswith("Total EUI") and kpis["total_eui"] < bm["good_eui"] * 0.6:
            flags.append(("warn","⚠", f"Total EUI ({kpis['total_eui']} kWh/m²·yr) is unusually low — confirm all end-uses are modelled."))
            continue

        level = LEVEL.get(status, "warn")
        if level == "fail":
            msg = f"{name}: {your} vs benchmark median {bench} ({diff}) — more than 15% above median; {hint_for(name)}"
        elif level == "warn":
            msg = f"{name}: {your} vs benchmark median {bench} ({diff}) — near the +15% high threshold; review."
        else:
            msg = f"{name}: {your} vs benchmark median {bench} ({diff}) — within the expected range."
        flags.append((level, ICON[level], msg))

    # TEDI not provided but a TEDI benchmark exists.
    if bm.get("median_tedi") and kpis.get("tedi", 0) == 0:
        flags.append(("warn","⚠", "TEDI not provided — add Thermal Energy Demand Intensity to compare against the benchmark."))

    # Cross-check: heating and cooling both above the high threshold.
    heat_red = any(r["End Use"].startswith("Heating") and r["QA Status"] == "🔴" for r in rows)
    cool_red = any(r["End Use"].startswith("Cooling") and r["QA Status"] == "🔴" for r in rows)
    if heat_red and cool_red:
        flags.append(("warn","⚠", "Both heating and cooling are above the high threshold — possible simultaneous heating/cooling or control issue."))
    return flags

def make_pie(labels, values, title):
    """Brand donut chart with the aggregate total rendered in the centre."""
    total = sum(values) if values else 0
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.62, sort=False,
        marker=dict(colors=CHART_SEQ[:len(labels)],
                    line=dict(color=BRAND["white"], width=2.5)),
        textinfo="percent", textposition="outside",
        textfont=dict(size=11.5, color=UI["tx2"], family=FONT_SANS),
        hovertemplate="<b>%{label}</b><br>%{value:.1f} kWh/m\u00b2\u00b7yr &middot; %{percent}<extra></extra>",
    ))
    fig.add_annotation(text=f"<b>{total:,.0f}</b>", showarrow=False,
                       font=dict(size=27, color=BRAND["black"], family=FONT_SERIF), y=.545)
    fig.add_annotation(text="kWh/m\u00b2\u00b7yr", showarrow=False,
                       font=dict(size=10.5, color=UI["tx3"], family=FONT_SANS), y=.40)
    return style_fig(fig, height=340, title=title, legend=True)

def build_pdf_report(meta, kpis, bm, flags, pct, comparison_df):
    """Build a one-file PDF QA/QC report (replaces the previous Excel export).

    Emoji status dots and subscript characters are mapped to plain text so they
    render reliably in the PDF's base fonts. The reviewer's overridden QA Status
    and Comments (from the editable Benchmark Comparison table) are included.
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape as _xml

    STATUS_WORD = {"🟢": "Pass", "🟡": "OK", "🟠": "Watch", "🔴": "Fail"}
    FLAG_WORD   = {"pass": "PASS", "warn": "REVIEW", "fail": "FAIL", "info": "INFO"}

    def desub(x):                       # subscript 2 -> 2 (base fonts can't render ₂)
        return str(x).replace("₂", "2")

    def destatus(x):                    # emoji dot -> word
        s = desub(x)
        for k, v in STATUS_WORD.items():
            s = s.replace(k, v)
        return s.strip()

    def para(x, style):                 # XML-safe paragraph
        return Paragraph(_xml(desub(x)), style)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        topMargin=1.4*cm, bottomMargin=1.4*cm, leftMargin=1.4*cm, rightMargin=1.4*cm,
        title=f"QA/QC Report - {meta['project_name'] or 'Project'}",
    )
    styles = getSampleStyleSheet()
    # Report styling follows the Stantec palette: black type, orange table headers,
    # warm-grey banding (Mist) instead of the cool blue-grey of the previous theme.
    h1    = ParagraphStyle("h1", parent=styles["Heading1"], textColor=colors.HexColor("#000000"), fontSize=17, spaceAfter=2)
    h2    = ParagraphStyle("h2", parent=styles["Heading2"], textColor=colors.HexColor("#000000"), fontSize=12, spaceBefore=8, spaceAfter=4)
    body  = styles["BodyText"]
    small = ParagraphStyle("small", parent=body, fontSize=9, textColor=colors.HexColor("#5F5850"))
    cell  = ParagraphStyle("cell", parent=body, fontSize=8, leading=10)

    elems = []

    # Title + metadata line
    elems.append(para(meta["project_name"] or "Project Results", h1))
    sub = meta['building_type']
    if meta.get('city'):
        sub += f" · {meta['city']}"
    sub += f" · Climate Zone {meta['climate_zone']}"
    if meta.get("subtype", "General") not in ("", "General", "All"):
        sub += f" · {meta['subtype']}"
    if not meta.get('city'):
        sub += " · zone average"
    sub += f" · {meta['model_type']}"
    if meta.get("phase"):
        sub += f" · {meta['phase']}"
    sub += f" · {meta['software']} · {meta['date']}"
    elems.append(para(sub, small))
    elems.append(Spacer(1, 8))

    # Overall result
    fc = {"pass": 0, "warn": 0, "fail": 0}
    for f in flags:
        fc[f[0]] = fc.get(f[0], 0) + 1
    overall = "Issues Found" if fc["fail"] > 0 else "Review Required" if fc["warn"] > 0 else "All Clear"
    pct_line = f"  |  Benchmark percentile: {pct}th (lower = better)" if pct else ""
    elems.append(Paragraph(
        desub(f"<b>Overall QA/QC:</b> {overall} — Pass {fc['pass']} · Review {fc['warn']} · Fail {fc['fail']}{pct_line}"),
        body))
    elems.append(Spacer(1, 10))

    # Energy summary
    elems.append(para("Energy Summary", h2))
    summ = [
        ["Metric", "Value", "Benchmark median"],
        ["Total EUI (kWh/m2/yr)",  kpis["total_eui"], bm["median_eui"]  if bm else "-"],
        ["TEDI (kWh/m2/yr)",       kpis.get("tedi", 0) or "-", (bm.get("median_tedi") or "-") if bm else "-"],
        ["GHGI (kgCO2e/m2/yr)",    kpis.get("ghgi") if kpis.get("ghgi") is not None else "-",      bm["median_ghgi"] if bm else "-"],
        ["Electricity EUI",        kpis["elec_eui"],  "-"],
        ["Gas EUI",                kpis["gas_eui"],   "-"],
        ["Floor area (m2)",        round(kpis["area"]), "-"],
        ["Unmet hours (total)",    kpis.get("unmet_total", 0), "-"],
        ["Unmet hours (heating / cooling)", f"{kpis.get('unmet_heating',0)} / {kpis.get('unmet_cooling',0)}", "-"],
    ]
    t = Table(summ, hAlign="LEFT", colWidths=[7*cm, 4*cm, 5*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ED6631")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.black),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#DED8D3")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2EFEC")]),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elems.append(t)

    # Benchmark comparison with reviewer overrides + comments
    if bm and comparison_df is not None:
        elems.append(para("Benchmark Comparison & Reviewer Notes", h2))
        header = ["End Use", "Your Model", "Bench. Median", "Diff.", "Auto Flag", "QA Status", "Comment"]
        data = [header]
        for _, rr in comparison_df.iterrows():
            data.append([
                para(rr["End Use"], cell),
                rr["Your Model"],
                rr["Benchmark Median"],
                str(rr["Difference"]),
                destatus(rr["Auto Flag"]),
                destatus(rr["QA Status"]),
                para(rr.get("Comment", "") or "", cell),
            ])
        ct = Table(data, hAlign="LEFT", repeatRows=1,
                   colWidths=[4.0*cm, 2.0*cm, 2.4*cm, 1.5*cm, 1.7*cm, 1.7*cm, 4.0*cm])
        ct.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#ED6631")),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.black),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#DED8D3")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2EFEC")]),
            ("VALIGN",     (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        elems.append(ct)

    # QA/QC flags
    elems.append(para("QA/QC Flags", h2))
    for level, _ic, msg in flags:
        safe_msg = _xml(desub(msg))
        elems.append(Paragraph(f"<b>[{FLAG_WORD.get(level, level.upper())}]</b> {safe_msg}", body))
        elems.append(Spacer(1, 2))

    doc.build(elems)
    return buf.getvalue()

# ── Session state ─────────────────────────────────────────────────────────────
if "step"    not in st.session_state: st.session_state.step    = 1
if "vals"    not in st.session_state: st.session_state.vals    = {}
if "results" not in st.session_state: st.session_state.results = None
if "page"    not in st.session_state: st.session_state.page    = PAGE_QA
if "bm_reload" not in st.session_state: st.session_state.bm_reload = 0

# ── Load benchmarks ───────────────────────────────────────────────────────────
BENCHMARKS = load_benchmarks()
ALL_BUILDING_TYPES = sorted(set(k[0] for k in BENCHMARKS)) or ["School","Office","Retail","Hospital","Residential","Warehouse"]
CITIES             = sorted(set(k[1] for k in BENCHMARKS)) or ["Edmonton","Calgary","Vancouver","Toronto"]
CLIMATE_ZONES      = sorted(set(k[2] for k in BENCHMARKS)) or ["4","5","6","7","8"]
# Subtypes derived dynamically per selection in Step 3

# ── Sidebar: enterprise navigation ────────────────────────────────────────────
NAV = [
    (PAGE_QA,       "shield",   "Validate a model against benchmarks"),
    (PAGE_EXPLORER, "bar",      "Explore benchmark distributions"),
    (PAGE_DB,       "database", "Manage the benchmark records"),
]
_NAV_ICONS = "".join(
    f'section[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type({i}) '
    f'> div:first-child{{display:none}}'
    f'section[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type({i})::after{{'
    f'content:"";width:30px;height:30px;border-radius:2px;order:-1;flex-shrink:0;'
    f'border:1px solid #DED8D3;background-color:#FFFFFF;'
    f'background-image:url("data:image/svg+xml;utf8,{_svg_uri(ic)}");'
    f'background-repeat:no-repeat;background-position:center;background-size:15px 15px;'
    f'transition:all .22s ease}}'
    # Selected nav item carries the Lens-tag: solid Stantec orange, black glyph.
    f'section[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type({i}):has(input:checked)::after{{'
    f'background-image:url("data:image/svg+xml;utf8,{_svg_uri(ic, "%23000000")}");'
    f'border-color:#ED6631;background-color:#ED6631}}'
    for i, (_lbl, ic, _d) in enumerate(NAV, 1)
)
st.markdown(f"<style>{_NAV_ICONS}</style>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown(f'''<div style="display:flex;align-items:center;gap:11px;padding:2px 2px 16px 2px">
      {logo_mark(36)}
      <div><div style="font-family:{FONT_SERIF};font-size:17px;font-weight:700;color:{UI['tx']};letter-spacing:-.01em;line-height:1.15">Energy Intelligence</div>
      <div style="font-size:9.5px;color:{UI['tx3']};letter-spacing:.14em;text-transform:uppercase;margin-top:3px;font-weight:700">Stantec Buildings</div></div>
    </div>''', unsafe_allow_html=True)

    st.markdown(f'''<div style="font-size:10px;font-weight:700;color:{UI['tx3']};letter-spacing:.14em;
        text-transform:uppercase;margin:2px 0 8px 2px">Workspace</div>
      <div style="display:flex;align-items:center;gap:9px;padding:9px 11px;border-radius:2px;
        border:1px solid {UI['bd']};background:{BRAND['white']};margin-bottom:14px">
        <span style="width:24px;height:24px;border-radius:2px;background:{BRAND['orange']};
          display:flex;align-items:center;justify-content:center;color:{BRAND['black']};font-weight:700;font-size:10.5px">BP</span>
        <div style="line-height:1.25"><div style="font-size:12.5px;font-weight:700;color:{UI['tx']}">Buildings Practice</div>
        <div style="font-size:10px;color:{UI['tx3']}">Western Canada</div></div>
      </div>''', unsafe_allow_html=True)

    st.markdown(f'''<div style="font-size:10px;font-weight:700;color:{UI['tx3']};letter-spacing:.14em;
        text-transform:uppercase;margin:2px 0 6px 2px">Navigation</div>''', unsafe_allow_html=True)
    page = st.radio("Navigate", [n[0] for n in NAV], label_visibility="collapsed")
    st.session_state.page = page

    if page == PAGE_QA:
        st.markdown(f'''<div style="font-size:10px;font-weight:700;color:{UI['tx3']};letter-spacing:.14em;
            text-transform:uppercase;margin:18px 0 8px 2px">Pipeline</div>''', unsafe_allow_html=True)
        _steps = ["Upload", "Column Mapping", "Building Details", "Validation & Results"]
        _cur = min(st.session_state.step, 4)
        _h = '<div style="display:flex;flex-direction:column;gap:7px;margin-bottom:14px">'
        for _i, _s in enumerate(_steps, 1):
            if _i < _cur:
                _bg, _fg, _tc, _lb, _bd = BRAND["fern"], BRAND["white"], UI["tx2"], icon("check", 12), BRAND["fern"]
            elif _i == _cur:
                # Current step wears the house colour, with black type (WCAG-safe).
                _bg, _fg, _tc, _lb, _bd = BRAND["orange"], BRAND["black"], UI["tx"], str(_i), BRAND["orange"]
            else:
                _bg, _fg, _tc, _lb, _bd = BRAND["white"], UI["tx3"], UI["tx3"], str(_i), UI["bd"]
            _h += (f'<div style="display:flex;align-items:center;gap:10px">'
                   f'<span style="width:22px;height:22px;border-radius:2px;background:{_bg};color:{_fg};'
                   f'border:1px solid {_bd};display:flex;align-items:center;justify-content:center;'
                   f'font-size:10.5px;font-weight:800;flex-shrink:0">{_lb}</span>'
                   f'<span style="font-size:12.5px;color:{_tc};font-weight:{"700" if _i == _cur else "500"}">{_s}</span></div>')
        st.markdown(_h + "</div>", unsafe_allow_html=True)

        if st.button("Start Over", use_container_width=True):
            for k in ["step","vals","results","headers","csv_df","mapping","meta","ref_csv_df","ref_vals","compliance_code"]:
                if k in st.session_state: del st.session_state[k]
            st.session_state.step = 1
            st.rerun()

    st.divider()
    st.markdown(f'''<div style="display:flex;flex-direction:column;gap:7px">
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:11px;color:{UI['tx3']};font-weight:600">Benchmarks</span>
        <span style="font-size:12.5px;color:{UI['tx']};font-weight:700">{len(BENCHMARKS)}</span></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:11px;color:{UI['tx3']};font-weight:600">Data source</span>
        <span style="font-size:11px;color:{UI['ok']};font-weight:700">&#9679; Synced</span></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:11px;color:{UI['tx3']};font-weight:600">Environment</span>
        <span style="font-size:11px;color:{BRAND['orange']};font-weight:700">Production</span></div>
    </div>''', unsafe_allow_html=True)

# ── Application header ────────────────────────────────────────────────────────
render_header(len(BENCHMARKS), st.session_state.page)



# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: MANAGE BENCHMARKS
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.page == PAGE_DB:
    page_head("Benchmark Database",
              "The single source of truth for every benchmark record. Data is hosted in Google Sheets "
              "and synced live \u2014 edits appear in the platform within 60 seconds.",
              ["Platform", "Data", "<span>Database</span>"])

    tb1, tb2, tb3, tb4 = st.columns([2.2, 1, 1, 1])
    with tb1:
        st.markdown(f'''<div style="display:flex;align-items:center;gap:9px;flex-wrap:wrap;padding-top:4px">
          {pill(f"{len(BENCHMARKS)} records", "info")}
          {pill("Google Sheets", "ok")}
          {pill("Live sync", "ok")}
          {pill("Read / write", "idle")}
        </div>''', unsafe_allow_html=True)
    with tb2:
        st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit",
                       use_container_width=True)
    with tb3:
        _sync = st.button("Sync now", use_container_width=True, type="primary")
    with tb4:
        st.markdown(f'''<div style="text-align:right;padding-top:9px;font-size:11px;color:{UI['tx3']}">
          Last sync<br><b style="color:{UI['tx2']}">&lt; 60 seconds ago</b></div>''', unsafe_allow_html=True)
    if _sync:
        load_benchmarks.clear()
        st.rerun()

    tab_view, tab_howto = st.tabs(["Records", "How to Edit"])

    # ── VIEW ──
    with tab_view:
        section("All Benchmarks", "Sortable, searchable enterprise data grid", "database")
        if st.button("Refresh from Google Sheets", use_container_width=False):
            load_benchmarks.clear()
            st.rerun()
        rows = []
        for (btype, city, zone, subtype), bm in BENCHMARKS.items():
            rows.append({
                "Building Type": btype, "Province": bm.get("province",""), "City": city, "Zone": zone,
                "Subtype": subtype,
                "Median EUI": bm["median_eui"],
                "Median TEDI": bm.get("median_tedi", 0),
                "Good (−15%)": bm["good_eui"],
                "High Flag (+15%)": bm["high_eui"],
                "GHGI": bm["median_ghgi"],
                "Heating %": bm["heat_pct"], "Cooling %": bm["cool_pct"],
                "Fan %": bm["fan_pct"], "Lighting %": bm["ltg_pct"],
                "DHW %": bm["dhw_pct"], "Receptacle %": bm["recept_pct"],
                "Pumps %": bm["pumps_pct"], "Electricity %": bm.get("elec_pct",0),
                "NaturalGas %": bm.get("gas_pct",0), "Heat Rejection %": bm.get("heat_rej_pct",0),
                "Miscellaneous %": bm.get("misc_pct",0),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        st.caption(f"Total: {len(rows)} benchmark records — refreshes automatically every 60 seconds")

    # ── HOW TO EDIT ──
    with tab_howto:
        section("How to add or edit benchmarks", "", "book")
        st.markdown(f"**[Open the Google Sheet]"
                    f"(https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)**")
        st.markdown("""
**To add a new benchmark:**
1. Open the Google Sheet using the link above
2. Scroll to the next empty row
3. Fill in all columns — Building Type, City, Climate Zone, EUI values, percentages, and percentile data
4. Save — the app picks up the change within 60 seconds

**To edit an existing benchmark:**
1. Open the Google Sheet
2. Find the row you want to change
3. Edit the cell directly
4. Save — changes appear in the app within 60 seconds

**To delete a benchmark:**
1. Open the Google Sheet
2. Right-click the row number → Delete row
3. Save

**Column guide:**
- Percentile Data: 10 numbers separated by commas e.g. `95,110,125,141,155,165,180,195,210,230`
- All percentages should sum to approximately 100%
        """)


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: BENCHMARK EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.page == PAGE_EXPLORER:
    page_head("Benchmark Explorer",
              "Interrogate the benchmark portfolio by typology, climate and system. "
              "Distributions, end-use fingerprints and percentile rankings \u2014 no upload required.",
              ["Platform", "Analytics", "<span>Explorer</span>"])

    section("Filters", "Refine the comparison set", "filter")

    # ── Row 1: required filters ───────────────────────────────────────────────
    cf1, cf2, cf3 = st.columns(3)
    with cf1:
        bx_type = st.selectbox("Building Type", ALL_BUILDING_TYPES)
    all_zones = sorted({k[2] for k,v in BENCHMARKS.items() if k[0]==bx_type and k[2]})
    with cf2:
        bx_zone = st.selectbox("Climate Zone", all_zones if all_zones else ["—"])
    # HVAC System — required; list all distinct values for this type + zone
    all_hvac = sorted({str(v.get("heating_system","")) for k,v in BENCHMARKS.items()
                       if k[0]==bx_type and k[2]==bx_zone and v.get("heating_system")})
    all_hvac_label = "All HVAC systems"
    with cf3:
        bx_hvac = st.selectbox("HVAC System", [all_hvac_label] + all_hvac)

    # ── Row 2: optional refinement filters (all in one row) ──────────────────
    opt1, opt2, opt3, opt4 = st.columns(4)

    provs = sorted({v["province"] for k,v in BENCHMARKS.items()
                    if k[0]==bx_type and k[2]==bx_zone
                    and v.get("province") and str(v["province"]).strip().lower() not in ("","nan")})
    all_provs_label = "All provinces"
    with opt1:
        bx_prov = st.selectbox("Province (optional)", [all_provs_label] + provs)

    def _city_filter(k, v):
        if k[0]!=bx_type or k[2]!=bx_zone: return False
        if bx_prov != all_provs_label and v.get("province")!=bx_prov: return False
        if bx_hvac != all_hvac_label and str(v.get("heating_system",""))!=bx_hvac: return False
        return True
    cities = sorted({k[1] for k,v in BENCHMARKS.items() if _city_filter(k,v)})
    all_cities_label = "All cities"
    with opt2:
        bx_city = st.selectbox("City (optional)", [all_cities_label] + cities)

    def _bm_filter(k, v):
        if not _city_filter(k,v): return False
        if bx_city != all_cities_label and k[1]!=bx_city: return False
        return True
    filtered_bms = {k:v for k,v in BENCHMARKS.items() if _bm_filter(k,v)}

    proj_years  = sorted({str(v.get("project_year","")) for v in filtered_bms.values() if v.get("project_year")})
    audit_types = sorted({str(v.get("audit_new",""))    for v in filtered_bms.values() if v.get("audit_new")})
    all_label = "All"
    with opt3:
        bx_year  = st.selectbox("Project Year (optional)", [all_label] + proj_years)
    with opt4:
        bx_audit = st.selectbox("Audit / New (optional)",  [all_label] + audit_types)

    # ── Resolve matching benchmarks ─────────────────────────────────────────
    def _final_filter(k, v):
        if not _bm_filter(k,v): return False
        if bx_year  != all_label and str(v.get("project_year","")) != bx_year:  return False
        if bx_audit != all_label and str(v.get("audit_new",""))    != bx_audit: return False
        return True
    final_bms = {k:v for k,v in BENCHMARKS.items() if _final_filter(k,v)}

    if not final_bms:
        scope = f"{bx_type} · Climate Zone {bx_zone}"
        if bx_hvac  != all_hvac_label:  scope += f" · {bx_hvac}"
        if bx_prov  != all_provs_label: scope += f" · {bx_prov}"
        if bx_city  != all_cities_label: scope += f" · {bx_city}"
        st.warning(f"No benchmark data found for **{scope}** with the selected filters. Try broadening your selection.")
        st.stop()

    agg = average_benchmarks(list(final_bms.values()))
    scope_parts = [f"Climate Zone {bx_zone}"]
    if bx_hvac  != all_hvac_label:  scope_parts.append(bx_hvac)
    if bx_prov  != all_provs_label:  scope_parts.append(bx_prov)
    if bx_city  != all_cities_label: scope_parts.append(bx_city)
    if bx_year  != all_label:        scope_parts.append(f"Year {bx_year}")
    if bx_audit != all_label:        scope_parts.append(bx_audit)
    scope_str = " · ".join(scope_parts)
    matches = {(bx_type, bx_city if bx_city!=all_cities_label else "", bx_zone, "filtered"): agg}
    if agg["_n"] == 1:
        st.info(f"Showing 1 benchmark for **{bx_type} · {scope_str}**.")
    else:
        st.info(f"Showing the **average of {agg['_n']} benchmarks** for **{bx_type} · {scope_str}**. Refine with the optional filters above.")

    for (btype, bcity, bzone, bsubtype), bm in matches.items():
        section(f"{btype} \u00b7 {scope_str}", f"{agg.get('_n', 1)} benchmark(s) in scope", "building")

        m1, m2, m3, m4 = st.columns(4)
        kpi(m1, "Median EUI", bm["median_eui"], "kWh/m\u00b2\u00b7yr", "zap", "cyan",
            compare="Portfolio typical", spark=sorted(bm["pct_data"])[:6] if bm.get("pct_data") else None)
        kpi(m2, "Good Practice", bm["good_eui"], "kWh/m\u00b2\u00b7yr", "check", "ok",
            compare="\u221215% \u00b7 QA pass threshold")
        kpi(m3, "High Flag", bm["high_eui"], "kWh/m\u00b2\u00b7yr", "alert", "err",
            compare="+15% \u00b7 QA fail threshold")
        kpi(m4, "Median GHGI", bm.get("median_ghgi", 0), "kgCO\u2082e/m\u00b2", "leaf", "blue",
            compare="Carbon intensity")

        if bm.get("median_tedi"):
            st.markdown("")
            t1, t2, t3, t4 = st.columns(4)
            kpi(t1, "Median TEDI", bm["median_tedi"], "kWh/m\u00b2\u00b7yr", "flame", "warn",
                compare="Thermal demand intensity",
                spark=sorted(bm["tedi_pct_data"])[:6] if bm.get("tedi_pct_data") else None)
            kpi(t2, "Good Practice", bm["good_tedi"], "kWh/m\u00b2\u00b7yr", "check", "ok",
                compare="\u221215% \u00b7 QA pass threshold")
            kpi(t3, "High Flag", bm["high_tedi"], "kWh/m\u00b2\u00b7yr", "alert", "err",
                compare="+15% \u00b7 QA fail threshold")
            kpi(t4, "Records", agg.get("_n", 1), "in scope", "database", "idle",
                compare="Matching benchmarks")

        # ── end-use data ──────────────────────────────────────────────────────
        eu_labels = ["Heating","Cooling","Fans","Lighting","DHW","Receptacle","Pumps","Heat Rejection","Miscellaneous"]
        eu_pcts   = [bm["heat_pct"], bm["cool_pct"], bm["fan_pct"], bm["ltg_pct"],
                     bm["dhw_pct"], bm["recept_pct"], bm["pumps_pct"],
                     bm.get("heat_rej_pct",0), bm.get("misc_pct",0)]
        med_vals  = [round(bm["median_eui"]*p/100,1) for p in eu_pcts]
        good_vals = [round(bm["good_eui"]*p/100,1)   for p in eu_pcts]
        high_vals = [round(bm["high_eui"]*p/100,1)   for p in eu_pcts]

        # ── energy source data ─────────────────────────────────────────────────
        src_labels = ["Electricity","Natural Gas"]
        src_pcts   = [bm.get("elec_pct",0), bm.get("gas_pct",0)]
        src_med    = [round(bm["median_eui"]*p/100,1) for p in src_pcts]

        cp1, cp2 = st.columns(2)
        with cp1:
            subtype_label = f" ({bsubtype})" if bsubtype not in ("General", "All", "filtered") else ""
            # Left: energy source pie
            src_l = [l for l,v in zip(src_labels, src_med) if v>0]
            src_v = [v for v in src_med if v>0]
            if src_l:
                fig_src = make_pie(src_l, src_v, "Energy Source Split")
                fig_src.update_traces(marker=dict(colors=[UI["cyan"], UI["warn"]][:len(src_l)],
                                                  line=dict(color="rgba(14,17,23,.85)", width=2.5)))
                st.plotly_chart(fig_src, use_container_width=True, config={"displayModeBar": False})
            else:
                st.caption("No Electricity % / NaturalGas % data in the sheet for this benchmark.")
        with cp2:
            # Right: end-use pie (includes Miscellaneous)
            pie_l = [l for l,v in zip(eu_labels, med_vals) if v>0]
            pie_v = [v for v in med_vals if v>0]
            st.plotly_chart(make_pie(pie_l, pie_v, f"Median End-Use Split — {btype} · {bcity}{subtype_label}"), use_container_width=True)

        # ── bar chart — end-uses only; heat rejection + miscellaneous included ──
        bar_pairs = [(l,v) for l,v in zip(eu_labels, med_vals) if v>0]
        bar_labels_f = [l for l,v in bar_pairs]
        bar_med_f    = [v for l,v in bar_pairs]
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(name="Median EUI", x=bar_labels_f, y=bar_med_f,
                                 marker=dict(color=bar_med_f, colorscale=[[0, "#F7B79A"], [1, BRAND["orange"]]],
                                             line=dict(width=0)),
                                 hovertemplate="<b>%{x}</b><br>%{y:.1f} kWh/m\u00b2\u00b7yr<extra></extra>",
                                 error_y=dict(type="data", symmetric=True,
                                              array=[round(v*0.15,1) for v in bar_med_f],
                                              color=BRAND["black"], thickness=1.2, width=4)))
        fig_bar.update_traces(marker_cornerradius=2)
        st.plotly_chart(style_fig(fig_bar, 330, "End-Use Median EUI \u00b7 error bars show \u00b115%",
                                  legend=False, ytitle="kWh/m\u00b2\u00b7yr"),
                        use_container_width=True, config={"displayModeBar": False})

        section("End-Use Breakdown", "Share of median EUI by system", "layers")
        st.caption("Good and High Flag are calculated as ±15% of median.")
        st.dataframe(pd.DataFrame({
            "End Use": eu_labels,
            "Share (%)": eu_pcts,
            "Median (kWh/m²·yr)": med_vals,
        }), use_container_width=True, hide_index=True)

        section("Percentile Distribution \u00b7 EUI", "Where projects sit across the portfolio", "bar")
        sorted_pcts = sorted(bm["pct_data"])
        bar_colors  = [UI["ok"] if v<=bm["good_eui"] else BRAND["marine"] if v<=bm["median_eui"] else UI["warn"] if v<=bm["high_eui"] else UI["err"] for v in sorted_pcts]
        fig_pct = go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)], y=sorted_pcts, marker_color=bar_colors,
            marker_cornerradius=2,
            hovertemplate="<b>%{x} percentile</b><br>EUI: %{y} kWh/m\u00b2\u00b7yr<extra></extra>"))
        fig_pct.add_hline(y=bm["good_eui"],   line_dash="dot",  line_color=UI["ok"],   line_width=1.4, opacity=.8)
        fig_pct.add_hline(y=bm["median_eui"], line_dash="dash", line_color=BRAND["black"], line_width=1.8, opacity=.9)
        fig_pct.add_hline(y=bm["high_eui"],   line_dash="dot",  line_color=UI["err"],  line_width=1.4, opacity=.8)
        st.plotly_chart(style_fig(fig_pct, 310, legend=False, xtitle="Percentile", ytitle="EUI (kWh/m\u00b2\u00b7yr)"),
                        use_container_width=True, config={"displayModeBar": False})
        st.caption(f"Good (\u221215%): {bm['good_eui']}  \u00b7  Median: {bm['median_eui']}  \u00b7  High flag (+15%): {bm['high_eui']}  kWh/m\u00b2\u00b7yr")

        if bm.get("median_tedi") and bm.get("tedi_pct_data"):
            section("Percentile Distribution \u00b7 TEDI", "Thermal demand across the portfolio", "flame")
            sorted_tedi = sorted(bm["tedi_pct_data"])
            tedi_colors = [UI["ok"] if v<=bm["good_tedi"] else BRAND["marine"] if v<=bm["median_tedi"] else UI["warn"] if v<=bm["high_tedi"] else UI["err"] for v in sorted_tedi]
            fig_tedi = go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)], y=sorted_tedi, marker_color=tedi_colors,
                marker_cornerradius=2,
                hovertemplate="<b>%{x} percentile</b><br>TEDI: %{y} kWh/m\u00b2\u00b7yr<extra></extra>"))
            fig_tedi.add_hline(y=bm["good_tedi"],   line_dash="dot",  line_color=UI["ok"],   line_width=1.4, opacity=.8)
            fig_tedi.add_hline(y=bm["median_tedi"], line_dash="dash", line_color=BRAND["black"], line_width=1.8, opacity=.9)
            fig_tedi.add_hline(y=bm["high_tedi"],   line_dash="dot",  line_color=UI["err"],  line_width=1.4, opacity=.8)
            st.plotly_chart(style_fig(fig_tedi, 310, legend=False, xtitle="Percentile", ytitle="TEDI (kWh/m\u00b2\u00b7yr)"),
                            use_container_width=True, config={"displayModeBar": False})
            st.caption(f"Good (\u221215%): {bm['good_tedi']}  \u00b7  Median: {bm['median_tedi']}  \u00b7  High flag (+15%): {bm['high_tedi']}  kWh/m\u00b2\u00b7yr")

        st.divider()


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE: QA/QC TOOL
# ══════════════════════════════════════════════════════════════════════════════
else:
    page_head("Energy Intelligence Platform",
              "Benchmark, validate and optimize building energy models using automated QA/QC "
              "against a live portfolio of completed projects.",
              ["Platform", "Validation", "<span>QA/QC</span>"])

    st.markdown(workflow(min(st.session_state.step, 5)), unsafe_allow_html=True)

    if st.session_state.step == 1:
        # Portfolio snapshot — derived live from the loaded benchmark database.
        _n_types = len({k[0] for k in BENCHMARKS})
        _n_zones = len({k[2] for k in BENCHMARKS if k[2]})
        _n_cities = len({k[1] for k in BENCHMARKS if k[1]})
        h1, h2c, h3, h4, h5 = st.columns(5)
        kpi(h1, "Benchmarks", len(BENCHMARKS), "", "database", "cyan", compare="Live from Google Sheets")
        kpi(h2c, "Building Types", _n_types, "", "building", "blue", compare="Typologies covered")
        kpi(h3, "Climate Zones", _n_zones, "", "layers", "blue", compare="NECB zones")
        kpi(h4, "Cities", _n_cities, "", "target", "ok", compare="Locations in database")
        kpi(h5, "Validation Engine", "Ready", "", "cpu", "ok", compare="Automated QA/QC rules")
        st.markdown("")

        section("Upload Simulation Results", "IES VE \u00b7 EnergyPlus \u00b7 OpenStudio \u00b7 CSV", "upload")
        tab_upload, tab_manual = st.tabs(["Upload CSV", "Enter Manually"])
        with tab_upload:
            st.markdown(f'''<div class="dz">
              <span style="color:{UI['cyan']};display:inline-flex">{icon("cloud", 44, sw=1.3)}</span>
              <div class="dz-t">Drag &amp; drop your simulation export</div>
              <div class="dz-d">or use Browse files below &middot; single CSV, up to 200&nbsp;MB</div>
              <div class="dz-f"><span>IES VE</span><span>EnergyPlus</span><span>OpenStudio</span>
                <span>eQUEST</span><span>CSV</span></div>
            </div>''', unsafe_allow_html=True)
            uploaded = st.file_uploader("Proposed Model CSV (required)", type=["csv"], label_visibility="collapsed")
            if uploaded:
                df = pd.read_csv(uploaded)
                st.session_state.csv_df  = df
                st.session_state.headers = list(df.columns)
                fu1, fu2, fu3, fu4 = st.columns(4)
                kpi(fu1, "File", uploaded.name.rsplit(".", 1)[0][:16], "", "file", "cyan", small=True)
                kpi(fu2, "Rows", f"{len(df):,}", "", "layers", "blue", small=True)
                kpi(fu3, "Columns", len(df.columns), "", "columns", "blue", small=True)
                kpi(fu4, "Status", "Ready", "", "check", "ok", small=True)
                st.markdown("")
                st.dataframe(df.head(3), use_container_width=True)

                # ── Optional: Reference model + NECB savings check ──
                st.markdown("---")
                section("Optional \u2014 NECB Savings Check", "Add a reference model to test code compliance", "target")
                st.caption("Add a Reference (baseline) model to compare end-use savings against NECB targets. "
                           "Leave blank to run the tool exactly as before.")
                necb_codes = load_necb_savings().get("codes") or ["NECB 2020", "NECB 2017"]
                prev_code = st.session_state.get("compliance_code")
                st.session_state.compliance_code = st.selectbox(
                    "Compliance Code", necb_codes,
                    index=necb_codes.index(prev_code) if prev_code in necb_codes else 0)
                ref_up = st.file_uploader("Reference Model CSV (optional)", type=["csv"], key="ref_uploader")
                if ref_up:
                    rdf = pd.read_csv(ref_up)
                    st.session_state.ref_csv_df = rdf
                    st.success(f"✅ Reference: {ref_up.name} — {len(rdf)} rows, {len(rdf.columns)} columns")
                elif st.session_state.get("ref_csv_df") is not None:
                    st.caption("✅ Reference model already loaded — re-upload above to replace it.")

                if st.button("Continue to Column Mapping", type="primary"):
                    st.session_state.step = 2; st.rerun()
        with tab_manual:
            c1,c2,c3 = st.columns(3)
            manual = {}
            with c1:
                section("Energy Sources", "", "zap")
                manual["electricity_kwh"]   = st.number_input("Electricity (kWh/yr)",        min_value=0.0, step=1000.0, format="%.0f")
                manual["gas_kwh"]           = st.number_input("Natural Gas (kWh/yr)",        min_value=0.0, step=1000.0, format="%.0f")
                manual["other_fuel_kwh"]    = st.number_input("Other Fuel / Biomass (kWh/yr)", min_value=0.0, step=1000.0, format="%.0f",
                                                              help="Biomass, district energy, oil, propane, etc. Counted in total energy and GHGI.")
                manual["area_m2"]           = st.number_input("Floor Area (m²)",              min_value=0.0, step=100.0,  format="%.0f")
                manual["tedi"]              = st.number_input("TEDI (kWh/m²·yr)",             min_value=0.0, step=10.0,   format="%.1f",
                                                              help="Thermal Energy Demand Intensity — enter as an intensity (already per m²).")
            with c2:
                section("HVAC End Uses", "", "flame")
                manual["heating_kwh"]       = st.number_input("Space Heating (kWh/yr)",      min_value=0.0, step=1000.0, format="%.0f")
                manual["cooling_kwh"]       = st.number_input("Space Cooling (kWh/yr)",      min_value=0.0, step=1000.0, format="%.0f")
                manual["central_fan_kwh"]  = st.number_input("Interior Central Fan / AHU (kWh/yr)", min_value=0.0, step=1000.0, format="%.0f")
                manual["local_fan_kwh"]    = st.number_input("Interior Local Fan (kWh/yr)",        min_value=0.0, step=1000.0, format="%.0f")
                manual["exhaust_fan_kwh"]  = st.number_input("Exhaust Fan (kWh/yr)",               min_value=0.0, step=1000.0, format="%.0f")
                manual["pumps_kwh"]         = st.number_input("Pumps (kWh/yr)",              min_value=0.0, step=500.0,  format="%.0f")
                manual["heat_rejection_kwh"]= st.number_input("Heat Rejection (kWh/yr)",     min_value=0.0, step=500.0,  format="%.0f")
            with c3:
                section("Other End Uses", "", "zap")
                manual["lighting_kwh"]      = st.number_input("Interior Lighting (kWh/yr)",  min_value=0.0, step=1000.0, format="%.0f")
                manual["dhw_kwh"]           = st.number_input("DHW (kWh/yr)",                min_value=0.0, step=500.0,  format="%.0f")
                manual["receptacle_kwh"]    = st.number_input("Receptacle / Plug Loads (kWh/yr)", min_value=0.0, step=500.0, format="%.0f")
                manual["ext_lighting_kwh"]  = st.number_input("Exterior Lighting (kWh/yr)",  min_value=0.0, step=500.0,  format="%.0f")
                manual["process_kwh"]       = st.number_input("Process / Other (kWh/yr)",    min_value=0.0, step=500.0,  format="%.0f")
                section("Unmet Hours", "", "clock")
                manual["unmet_hours_heating"] = st.number_input("Unmet Hours — Heating", min_value=0.0, step=1.0, format="%.0f",
                                                                help="Count of occupied hours outside the heating setpoint range (not divided by area).")
                manual["unmet_hours_cooling"] = st.number_input("Unmet Hours — Cooling", min_value=0.0, step=1.0, format="%.0f")
                manual["unmet_hours_total"]   = st.number_input("Unmet Hours — Total",   min_value=0.0, step=1.0, format="%.0f")
            if st.button("Continue to Building Details", type="primary"):
                st.session_state.vals = manual; st.session_state.ref_vals = None; st.session_state.step = 3; st.rerun()

    elif st.session_state.step == 2:
        section("Column Mapping", "Match your export fields to platform variables", "columns")
        headers = st.session_state.headers
        auto_map = guess_mapping(headers)
        options  = ["— not mapped —"] + headers
        mapping  = {}
        auto_matched = sum(1 for v in auto_map.values() if v)
        st.info(f"Auto-matched {auto_matched} of {len(FIELD_LABELS)} fields. Review and fix any that say '— not mapped —'.")

        # Group fields into sections for clarity
        sections = {
            ("Energy Sources", "zap"):    ["electricity_kwh","gas_kwh","other_fuel_kwh","area_m2","tedi"],
            ("HVAC End Uses", "flame"):   ["heating_kwh","cooling_kwh","central_fan_kwh","local_fan_kwh","exhaust_fan_kwh","pumps_kwh","heat_rejection_kwh"],
            ("Other End Uses", "zap"):    ["lighting_kwh","dhw_kwh","receptacle_kwh","ext_lighting_kwh","process_kwh"],
            ("Unmet Hours", "clock"):     ["unmet_hours_heating","unmet_hours_cooling","unmet_hours_total"],
        }
        for (section_title, section_icon), keys in sections.items():
            section(section_title, "", section_icon)
            cols = st.columns(3)
            for i, key in enumerate(keys):
                label = FIELD_LABELS.get(key, key)
                with cols[i % 3]:
                    default_idx = options.index(auto_map[key]) if auto_map.get(key) and auto_map[key] in options else 0
                    sel = st.selectbox(label, options, index=default_idx, key=f"map_{key}")
                    mapping[key] = sel if sel != "— not mapped —" else None
            st.markdown("")
        cb,cn = st.columns([1,4])
        with cb:
            if st.button("Back"): st.session_state.step=1; st.rerun()
        with cn:
            if st.button("Continue to Building Details", type="primary"):
                df = st.session_state.csv_df
                vals = {key: pd.to_numeric(df[col], errors="coerce").sum() if col else 0 for key,col in mapping.items()}
                st.session_state.vals=vals; st.session_state.mapping=mapping
                # Reference model (optional) — same export format, so reuse the same mapping
                rdf = st.session_state.get("ref_csv_df")
                if rdf is not None:
                    st.session_state.ref_vals = {
                        key: (pd.to_numeric(rdf[col], errors="coerce").sum() if (col and col in rdf.columns) else 0)
                        for key, col in mapping.items()
                    }
                else:
                    st.session_state.ref_vals = None
                st.session_state.step=3; st.rerun()

    elif st.session_state.step == 3:
        section("Building Details", "Project metadata, benchmark selection and emission factors", "building")
        # ── Required project metadata (visible & mandatory) ──
        m1, m2, m3 = st.columns(3)
        with m1:
            project_name = st.text_input("Project Name", placeholder="e.g. New School A")
        with m2:
            software     = st.selectbox("Simulation Software *", SOFTWARE_OPTIONS)
        with m3:
            model_type   = st.selectbox("Model Type", MODEL_TYPES)
        phase = ""   # Project Phase removed; kept blank for downstream compatibility

        # ── Benchmark selection — identical filtering to the Benchmark Explorer ──
        st.markdown("**Benchmark to compare against** — same filters as the Benchmark Explorer. "
                    "Pick a city to compare against that exact benchmark model, or leave City on the zone average.")
        b1, b2, b3, b4 = st.columns(4)
        with b1:
            building_type = st.selectbox("Building Type", ALL_BUILDING_TYPES)
        provs = sorted({v["province"] for k,v in BENCHMARKS.items() if k[0]==building_type and v.get("province")})
        with b2:
            province = st.selectbox("Province", provs if provs else ["—"])
        zones_avail = sorted({k[2] for k,v in BENCHMARKS.items()
                              if k[0]==building_type and v.get("province")==province and k[2]})
        with b3:
            climate_zone = st.selectbox("Climate Zone", zones_avail if zones_avail else ["—"])
        cities_in_zone = sorted({k[1] for k,v in BENCHMARKS.items()
                                 if k[0]==building_type and v.get("province")==province and k[2]==climate_zone})
        zone_avg_label = f"Climate Zone {climate_zone} average"
        with b4:
            city_sel = st.selectbox("City (optional)", [zone_avg_label] + cities_in_zone,
                                    help="Leave on the zone average, or pick a city to compare against that exact benchmark model.")

        # Resolve the benchmark baseline used by all comparison tables, charts and QA checks
        if city_sel == zone_avg_label:
            zone_bms = [v for k,v in BENCHMARKS.items()
                        if k[0]==building_type and v.get("province")==province and k[2]==climate_zone]
            bm = average_benchmarks(zone_bms) if zone_bms else None
            city, subtype = "", "All"
            if bm:
                st.info(f"Comparing against the **Climate Zone {climate_zone} average** for {building_type} in {province} "
                        f"— average of {bm['_n']} benchmark(s). Select a city to use a specific benchmark model.")
            else:
                st.warning("No benchmark found for this combination. KPIs will be calculated but no comparison will be shown.")
        else:
            city = city_sel
            available_subtypes = sorted({k[3] for k in BENCHMARKS
                                         if k[0]==building_type and k[1]==city and k[2]==climate_zone})
            if len(available_subtypes) > 1:
                subtype = st.selectbox("Benchmark Model (subtype)", available_subtypes,
                                       help="Pick the exact benchmark model to compare against.")
            elif len(available_subtypes) == 1:
                subtype = available_subtypes[0]
            else:
                subtype = "General"
            bm = BENCHMARKS.get((building_type, city, climate_zone, subtype))
            if bm:
                lbl = f"{building_type} · {city} · Climate Zone {climate_zone}" + (f" · {subtype}" if subtype not in ("General","All") else "")
                st.info(f"Comparing against benchmark model: **{lbl}**.")
            else:
                st.warning("No benchmark found for this exact combination. KPIs will be calculated but no comparison will be shown.")

        # ── GHG emission factors (defaults follow the selected province) ──
        elec_opts  = {p: round(g/1000, 4)              for p, g in ELEC_CO2E_G_PER_KWH.items()}
        gas_opts   = {p: round(g/NG_KWH_PER_M3/1000,4) for p, g in NG_CO2_G_PER_M3.items()}
        other_opts = {"None": 0.0, "Propane": 0.2197, "Butane": 0.2229}
        CUSTOM = "Custom value…"
        elec_default = round(ELEC_CO2E_G_PER_KWH.get(province, 0)/1000, 4)
        gas_default  = round(NG_CO2_G_PER_M3.get(province, 0)/NG_KWH_PER_M3/1000, 4) if NG_CO2_G_PER_M3.get(province) else 0.0

        st.markdown("**GHG Emission Factors (kg CO₂e/kWh)** — select a province/fuel value (Canada 2026 tables) or choose *Custom value…* to enter your own. Choose 0 / None to skip GHGI.")
        ef1, ef2, ef3 = st.columns(3)
        with ef1:
            opts = [f"{p} — {v:.4f}" for p, v in elec_opts.items()] + [CUSTOM]
            di = list(elec_opts).index(province) if province in elec_opts else 0
            sel = st.selectbox("Electricity", opts, index=di, key=f"ef_elec_sel_{province}",
                               help="Provincial electricity consumption intensity (Table 5.3, 2026).")
            ef_elec = (st.number_input("Electricity — custom", min_value=0.0, step=0.001, format="%.4f",
                                       value=elec_default, key=f"ef_elec_cust_{province}")
                       if sel == CUSTOM else elec_opts[sel.rsplit(" — ", 1)[0]])
        with ef2:
            opts = [f"{p} — {v:.4f}" for p, v in gas_opts.items()] + [CUSTOM]
            di = list(gas_opts).index(province) if province in gas_opts else 0
            sel = st.selectbox("Natural Gas", opts, index=di, key=f"ef_gas_sel_{province}",
                               help="Marketable natural-gas CO₂ (Table 1.3, 2026, g CO₂/m³) ÷ 10.55 kWh/m³.")
            ef_gas = (st.number_input("Natural Gas — custom", min_value=0.0, step=0.001, format="%.4f",
                                      value=gas_default, key=f"ef_gas_cust_{province}")
                      if sel == CUSTOM else gas_opts[sel.rsplit(" — ", 1)[0]])
        with ef3:
            opts = [f"{p} — {v:.4f}" for p, v in other_opts.items()] + [CUSTOM]
            sel = st.selectbox("Other Resources / Biomass", opts, index=0, key=f"ef_other_sel_{province}",
                               help="Propane/Butane converted from Table 3.3; or pick Custom for biomass/other.")
            ef_other = (st.number_input("Other — custom", min_value=0.0, step=0.001, format="%.4f",
                                        value=0.0, key=f"ef_other_cust_{province}")
                        if sel == CUSTOM else other_opts[sel.rsplit(" — ", 1)[0]])

        cb,cn = st.columns([1,4])
        with cb:
            if st.button("Back"):
                st.session_state.step = 2 if "headers" in st.session_state else 1; st.rerun()
        with cn:
            if st.button("Run Validation", type="primary"):
                if not software:
                    st.error("⚠️ Simulation Software is required.")
                else:
                    kpis   = calculate_kpis(st.session_state.vals, ef_elec=ef_elec, ef_gas=ef_gas, ef_other=ef_other)
                    flags  = generate_flags(kpis, bm, building_type)
                    pct    = calc_percentile(kpis["total_eui"], bm["pct_data"]) if bm else None
                    st.session_state.results = {"kpis":kpis,"bm":bm,"flags":flags,"percentile":pct,
                        "meta":{"project_name":project_name,"building_type":building_type,"city":city,
                                "province":province,"climate_zone":climate_zone,"subtype":subtype,"software":software,
                                "model_type":model_type,"phase":phase,"date":str(date.today()),
                                "ef_elec":ef_elec,"ef_gas":ef_gas,"ef_other":ef_other}}
                    st.session_state.step=4; st.rerun()

    elif st.session_state.step == 4 and st.session_state.results:
        r=st.session_state.results; kpis=r["kpis"]; bm=r["bm"]; meta=r["meta"]; flags=r["flags"]; pct=r["percentile"]

        # Build the per-metric comparison rows once, and derive the QA/QC Flags from the
        # same Auto Flag statuses so the flag list and the comparison table always agree.
        comp_rows = build_comparison_rows(kpis, bm) if bm else None
        flags = flags_from_comparison(comp_rows, kpis, bm, meta["building_type"]) if bm else flags

        # ── Project header banner ──
        fc={"pass":0,"warn":0,"fail":0}
        for f in flags: fc[f[0]]=fc.get(f[0],0)+1
        overall       = "Issues Found"    if fc["fail"]>0 else "Review Required" if fc["warn"]>0 else "All Clear"
        overall_icon  = "🔴"              if fc["fail"]>0 else "🟡"              if fc["warn"]>0 else "🟢"
        overall_tone  = "err"             if fc["fail"]>0 else "warn"            if fc["warn"]>0 else "ok"
        overall_accent= UI["err"]         if fc["fail"]>0 else UI["warn"]        if fc["warn"]>0 else UI["ok"]
        overall_ic    = "x"               if fc["fail"]>0 else "alert"           if fc["warn"]>0 else "check"
        proj_name     = meta["project_name"] or "Project Results"
        subtype_disp  = f" &middot; {meta.get('subtype','')}" if meta.get('subtype','') not in ("","General","All") else ""
        phase_disp    = f" &middot; {meta['phase']}" if meta.get('phase') else ""
        city_disp     = f"{meta['city']} &middot; " if meta.get('city') else ""
        scope_disp    = "" if meta.get('city') else " &middot; zone average"
        info_line     = (f"{meta['building_type']} &middot; {city_disp}"
                         f"Climate Zone {meta['climate_zone']}{scope_disp}{subtype_disp} &middot; {meta['model_type']}"
                         f"{phase_disp} &middot; {meta['software']}")
        pct_block = ""
        if pct:
            pct_block = (f'<div style="text-align:right;padding-left:22px;border-left:1px solid {UI["bd"]}">'
                         f'<div style="font-family:{FONT_SERIF};font-size:46px;font-weight:700;'
                         f'color:{UI["tx"]};line-height:1;letter-spacing:-.02em">'
                         f'{pct}<span style="font-size:19px;color:{UI["tx3"]}">th</span></div>'
                         f'<div style="font-size:10.5px;color:{UI["tx3"]};text-transform:uppercase;'
                         f'letter-spacing:.11em;font-weight:700;margin-top:5px">Percentile &middot; lower is better</div></div>')

        # Emitted as one unbroken string on purpose. When `pct_block` is empty a
        # multi-line template leaves a whitespace-only line, which closes the HTML
        # block for the markdown parser — the indented lines after it are then read
        # as an indented code block and the closing </div> tags render as text.
        _pills = (pill(f"{fc['pass']} Pass", "ok") + pill(f"{fc['warn']} Review", "warn")
                  + pill(f"{fc['fail']} Fail", "err"))
        st.markdown(
            f'<div class="ei-card rise" style="border:1px solid {UI["bd"]};'
            f'border-left:5px solid {overall_accent};background:{BRAND["mist"]};'
            f'padding:22px 26px;margin-bottom:6px">'
            f'<div style="display:flex;align-items:center;justify-content:space-between;'
            f'gap:22px;flex-wrap:wrap">'
            f'<div style="min-width:260px">'
            f'<div style="display:flex;align-items:center;gap:11px;flex-wrap:wrap">'
            f'<span style="font-family:{FONT_SERIF};font-size:28px;font-weight:700;'
            f'color:{UI["tx"]};letter-spacing:-.01em">{proj_name}</span>'
            f'{pill(overall, overall_tone)}</div>'
            f'<div style="font-size:13px;color:{UI["tx3"]};margin-top:7px">{info_line}</div>'
            f'<div style="display:flex;gap:9px;margin-top:13px;flex-wrap:wrap">{_pills}</div>'
            f'</div>{pct_block}</div></div>',
            unsafe_allow_html=True)

        # ── Section 1: Executive Summary ──
        section("Executive Summary", "Headline performance vs the benchmark portfolio", "gauge")

        _tr = (lambda v, b: round((v - b) / b * 100) if b else None)

        r1c1, r1c2, r1c3, r1c4, r1c5 = st.columns(5)
        # Total EUI
        _eui_tone = ("ok" if bm and kpis["total_eui"] <= bm["good_eui"]
                     else "err" if bm and kpis["total_eui"] > bm["high_eui"] else "warn") if bm else "cyan"
        kpi(r1c1, "Total EUI", kpis["total_eui"], "kWh/m\u00b2\u00b7yr", "zap", _eui_tone,
            compare=(f"Median {bm['median_eui']}" if bm else "No benchmark"),
            trend=(_tr(kpis["total_eui"], bm["median_eui"]) if bm else None),
            spark=(sorted(bm["pct_data"])[:6] if bm and bm.get("pct_data") else None))
        # TEDI
        _has_tedi_bm = bool(bm and bm.get("median_tedi"))
        _tedi = kpis.get("tedi", 0)
        if _has_tedi_bm and _tedi > 0:
            _t_tone = "ok" if _tedi <= bm["good_tedi"] else "err" if _tedi > bm["high_tedi"] else "warn"
            _t_cmp  = f"Median {bm['median_tedi']}"
            _t_trend = _tr(_tedi, bm["median_tedi"])
        else:
            _t_tone, _t_trend = "idle", None
            _t_cmp = "No TEDI benchmark" if not _has_tedi_bm else "Not provided"
        kpi(r1c2, "TEDI", (_tedi if _tedi > 0 else "\u2014"), "kWh/m\u00b2\u00b7yr", "flame", _t_tone,
            compare=_t_cmp, trend=_t_trend)
        # GHGI
        _ghgi = kpis.get("ghgi")
        _g_cmp = (f"Median {bm['median_ghgi']}" if bm else "No benchmark") if _ghgi is not None \
                 else "Set emission factors"
        kpi(r1c3, "GHGI", (_ghgi if _ghgi is not None else "\u2014"), "kgCO\u2082e/m\u00b2", "leaf",
            ("ok" if (_ghgi is not None and bm and _ghgi <= bm["median_ghgi"]) else
             "err" if (_ghgi is not None and bm and _ghgi > bm["median_ghgi"] * 1.15) else
             "warn" if _ghgi is not None else "idle"),
            compare=_g_cmp,
            trend=(_tr(_ghgi, bm["median_ghgi"]) if (_ghgi is not None and bm) else None))
        # Electricity / Gas
        kpi(r1c4, "Electricity EUI", kpis["elec_eui"], "kWh/m\u00b2\u00b7yr", "zap", "cyan",
            compare=f"Floor area {round(kpis['area']):,} m\u00b2")
        kpi(r1c5, "Gas EUI", kpis["gas_eui"], "kWh/m\u00b2\u00b7yr", "flame", "warn",
            compare=f"Total {round(kpis['total_energy']/1000):,} MWh/yr")

        st.markdown("")

        # Row 2 — secondary metrics + score gauges
        r2c1, r2c2, gcol1, gcol2, gcol3 = st.columns([1, 1, 1, 1, 1])
        kpi(r2c1, "Other Fuel / Biomass", kpis.get("other_fuel_eui", 0), "kWh/m\u00b2\u00b7yr",
            "leaf", "ok", compare="Counted in total energy", small=True)
        kpi(r2c2, "Unmet Hours", f"{kpis.get('unmet_total',0):,}", "hrs/yr", "clock",
            ("ok" if kpis.get("unmet_total", 0) <= 300 else "warn" if kpis.get("unmet_total", 0) <= 1000 else "err"),
            compare=f"Htg {kpis.get('unmet_heating',0):,} \u00b7 Clg {kpis.get('unmet_cooling',0):,}", small=True)

        # Scores derived from the same QA results shown below (no new logic).
        _tot_f = max(fc["pass"] + fc["warn"] + fc["fail"], 1)
        _qa_score = round(100 * (fc["pass"] + 0.5 * fc["warn"]) / _tot_f)
        _perf_score = round(max(0, min(100, 100 - (kpis["total_eui"] - bm["median_eui"]) / bm["median_eui"] * 100))) \
                      if bm and bm.get("median_eui") else 0
        _comp_score = round(100 * fc["pass"] / _tot_f)
        with gcol1:
            st.plotly_chart(make_gauge(_perf_score, "PERFORMANCE SCORE"), use_container_width=True,
                            config={"displayModeBar": False})
        with gcol2:
            st.plotly_chart(make_gauge(_qa_score, "QA SCORE"), use_container_width=True,
                            config={"displayModeBar": False})
        with gcol3:
            st.plotly_chart(make_gauge(_comp_score, "COMPLIANCE"), use_container_width=True,
                            config={"displayModeBar": False})

        st.divider()

        # ── Section 2: Charts ──
        section("Energy Analytics", "End-use split, portfolio ranking and benchmark deltas", "bar")
        cc1,cc2 = st.columns(2)
        with cc1:
            eu_l=["Heating","Cooling","Fans","Lighting","DHW","Receptacle","Pumps"]
            eu_v=[kpis["heat_eui"],kpis["cool_eui"],kpis["fan_eui"],kpis["ltg_eui"],kpis["dhw_eui"],kpis.get("recept_eui",0),kpis["pumps_eui"]]
            pl=[l for l,v in zip(eu_l,eu_v) if v>0]; pv=[v for v in eu_v if v>0]
            st.plotly_chart(make_pie(pl,pv,"End-Use Energy Split"), use_container_width=True)
        with cc2:
            if bm:
                sorted_pcts=sorted(bm["pct_data"])
                bar_colors=[UI["ok"] if v<=bm["good_eui"] else BRAND["marine"] if v<=bm["median_eui"] else UI["warn"] if v<=bm["high_eui"] else UI["err"] for v in sorted_pcts]
                fig_pct=go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)],y=sorted_pcts,marker_color=bar_colors,
                    marker_cornerradius=2, opacity=.85,
                    hovertemplate="<b>%{x} percentile</b><br>EUI: %{y} kWh/m\u00b2\u00b7yr<extra></extra>"))
                fig_pct.add_hline(y=kpis["total_eui"], line_dash="dash", line_color=BRAND["black"], line_width=2,
                    annotation_text=f"  Your model \u00b7 {kpis['total_eui']}",
                    annotation_font=dict(color=BRAND["black"], size=12), annotation_position="top left")
                st.plotly_chart(style_fig(fig_pct, 340, "Portfolio Ranking \u00b7 lower is better", legend=False,
                                          xtitle="Percentile rank of similar buildings", ytitle="EUI (kWh/m\u00b2\u00b7yr)"),
                                use_container_width=True, config={"displayModeBar": False})

        if bm:
            eu_l2=["Heating","Cooling","Central Fan","Local Fan","Exhaust Fan","Lighting","DHW","Receptacle","Pumps"]
            your_eu=[kpis["heat_eui"],kpis["cool_eui"],kpis.get("central_fan_eui",0),kpis.get("local_fan_eui",0),
                     kpis.get("exhaust_fan_eui",0),kpis["ltg_eui"],kpis["dhw_eui"],kpis.get("recept_eui",0),kpis["pumps_eui"]]
            fan_split = bm["fan_pct"] / 3  # split benchmark fan % evenly across central/local/exhaust
            bm_good=[round(bm["good_eui"]*p/100,1)  for p in [bm["heat_pct"],bm["cool_pct"],fan_split,fan_split,fan_split,bm["ltg_pct"],bm["dhw_pct"],bm["recept_pct"],bm["pumps_pct"]]]
            bm_med =[round(bm["median_eui"]*p/100,1) for p in [bm["heat_pct"],bm["cool_pct"],fan_split,fan_split,fan_split,bm["ltg_pct"],bm["dhw_pct"],bm["recept_pct"],bm["pumps_pct"]]]
            fig_cmp = go.Figure()
            # Benchmark median bars with ±15% error bars
            fig_cmp.add_trace(go.Bar(
                name="Benchmark Median", x=eu_l2, y=bm_med,
                marker_color=BRAND["fossil"], marker_cornerradius=2,
                hovertemplate="<b>%{x}</b><br>Benchmark %{y:.1f}<extra></extra>",
                error_y=dict(type="data", symmetric=True,
                             array=[round(v*0.15,1) for v in bm_med],
                             color=BRAND["black"], thickness=1.2, width=5),
            ))
            # Your model bars, labelled with % difference vs benchmark median
            pct_labels = [f"{(y-m)/m*100:+.0f}%" if m else "" for y, m in zip(your_eu, bm_med)]
            fig_cmp.add_trace(go.Bar(
                name="Your Model", x=eu_l2, y=your_eu,
                marker_color=BRAND["orange"], marker_cornerradius=2,
                text=pct_labels, textposition="outside",
                textfont=dict(size=11, color=UI["tx2"], family=FONT_SANS),
                hovertemplate="<b>%{x}</b><br>Your model %{y:.1f}<extra></extra>",
                cliponaxis=False,
            ))
            fig_cmp.update_layout(barmode="group")
            st.plotly_chart(style_fig(fig_cmp, 380,
                                      "Your Model vs Benchmark Median \u00b7 labels = % vs median, error bars = \u00b115%",
                                      ytitle="EUI (kWh/m\u00b2\u00b7yr)"),
                            use_container_width=True, config={"displayModeBar": False})

        st.divider()

        # ── Section 3: Benchmark Table ──
        comparison_edited = None
        if bm:
            section("Benchmark Comparison", "Reviewer-editable QA decisions and audit trail", "columns")
            st.caption("QA/QC thresholds: 🟢 ≤ −15% of median &nbsp; 🟡 Within ±15% of median &nbsp; 🔴 > +15% of median")

            rows = comp_rows   # built once at the top of Step 4 (same statuses drive the QA/QC Flags)

            # Editable comparison: reviewer can override the QA Status and add a Comment.
            # "Auto Flag" preserves the tool's original 4-level computed status for audit;
            # the editable QA Status is a binary pass/fail decision (🟢 / 🔴).
            comp_df = pd.DataFrame(rows).rename(columns={"QA Status": "Auto Flag"})
            comp_df["QA Status"] = comp_df["Auto Flag"].apply(lambda s: "🔴" if s == "🔴" else "🟢")  # default: red only for fails
            comp_df["Comment"]   = ""
            comp_df = comp_df[["End Use","Your Model","Benchmark Median","Difference","Auto Flag","QA Status","Comment"]]

            st.caption("Set the **QA Status** of any row to 🟢 Pass or 🔴 Fail (e.g. pass a flagged item after review) and add a **Comment**. The **Auto Flag** column keeps the tool's original result for audit, and your overrides + comments are saved into the exported PDF.")
            comparison_edited = st.data_editor(
                comp_df, use_container_width=True, hide_index=True, key="bm_review",
                column_config={
                    "End Use":          st.column_config.TextColumn(disabled=True),
                    "Your Model":       st.column_config.NumberColumn(disabled=True),
                    "Benchmark Median": st.column_config.NumberColumn(disabled=True),
                    "Difference":       st.column_config.TextColumn("Difference (%)", disabled=True,
                                            help="Percent difference of your model vs the benchmark median."),
                    "Auto Flag":        st.column_config.TextColumn(
                                            "Auto Flag", disabled=True,
                                            help="Status the tool calculated automatically — kept for audit."),
                    "QA Status":        st.column_config.SelectboxColumn(
                                            "QA Status", options=["🟢","🔴"], required=True,
                                            help="Reviewer decision — 🟢 Pass or 🔴 Fail."),
                    "Comment":          st.column_config.TextColumn(
                                            "Comment", help="Reviewer notes / justification for any override."),
                },
            )
            binary_default = comp_df["QA Status"]
            n_over = int((comparison_edited["QA Status"].values != binary_default.values).sum())
            if n_over:
                st.caption(f"✏️ {n_over} status value(s) overridden by reviewer.")

        st.divider()

        # ── NECB Savings QA (only when a Reference model was also uploaded) ──
        ref_vals = st.session_state.get("ref_vals")
        if ref_vals is not None:
            code = st.session_state.get("compliance_code", "NECB 2020")
            necb_all = load_necb_savings()
            sheet_savings = necb_savings_for(necb_all, meta.get("building_type"), meta.get("city"), code)
            from_sheet = bool(sheet_savings)
            # Use the sheet's targets exactly (blank cells → no target). Defaults only when no row matches.
            savings_map = sheet_savings if from_sheet else dict(DEFAULT_NECB_SAVINGS)
            necb_rows = build_necb_rows(st.session_state.vals, ref_vals, savings_map)

            section(f"Code Compliance Check \u00b7 {code}", "Proposed vs reference savings per end use", "target")
            if from_sheet:
                src_note = f"targets loaded from the **{NECB_SHEET_NAME}** sheet for {meta.get('building_type','?')} · {meta.get('city','?')}"
            else:
                src_note = f"using built-in defaults (no matching row found on the **{NECB_SHEET_NAME}** tab)"
            st.caption(
                f"Proposed vs Reference savings per end use against NECB targets ({src_note}). "
                f"**Savings = (Reference − Proposed) / Reference × 100.** "
                f"Auto Flag: 🟢 savings within {NECB_TOL} pts of (or above) target · 🔴 below target · ⚪ n/a (reference = 0). "
                f"Override **QA Status** and add a **Comment** as needed."
            )
            necb_df = pd.DataFrame(necb_rows)
            necb_df["QA Status"] = necb_df["Auto Flag"]
            necb_df["Comment"]   = ""
            necb_df = necb_df[["End Use","Proposed","Reference Model","Savings","Benchmark","Auto Flag","QA Status","Comment"]]
            necb_edited = st.data_editor(
                necb_df, use_container_width=True, hide_index=True, key="necb_review",
                column_config={
                    "End Use":         st.column_config.TextColumn(disabled=True),
                    "Proposed":        st.column_config.NumberColumn("Proposed (kWh)",        disabled=True, format="%.1f"),
                    "Reference Model": st.column_config.NumberColumn("Reference Model (kWh)", disabled=True, format="%.1f"),
                    "Savings":         st.column_config.TextColumn(disabled=True),
                    "Benchmark":       st.column_config.TextColumn(disabled=True, help="NECB target savings for this end use."),
                    "Auto Flag":       st.column_config.TextColumn(disabled=True, help="Computed automatically; kept for audit."),
                    "QA Status":       st.column_config.SelectboxColumn(options=["🟢","🔴","⚪"], required=True, help="Reviewer override (🟢 pass / 🔴 fail / ⚪ n/a)."),
                    "Comment":         st.column_config.TextColumn(help="Justification or notes for any override."),
                },
            )
            n_g  = int((necb_edited["QA Status"]=="🟢").sum())
            n_r  = int((necb_edited["QA Status"]=="🔴").sum())
            n_na = int((necb_edited["QA Status"]=="⚪").sum())
            n_ov = int((necb_edited["QA Status"].values != necb_df["QA Status"].values).sum())
            st.caption(f"🟢 {n_g} meet target · 🔴 {n_r} below target · ⚪ {n_na} n/a"
                       + (f"  ·  ✏️ {n_ov} overridden by reviewer" if n_ov else ""))
            st.divider()

        # ── Section 4: Validation Results ──
        section("Validation Results", "Automated QA/QC checks against the benchmark median", "shield")

        _vt = {"pass": ("PASSED", UI["ok"]), "warn": ("WARNING", UI["warn"]),
               "fail": ("FAILED", UI["err"]), "info": ("INFO", UI["info"])}
        st.markdown(f'''<div class="vsum">
          <div class="vsum-c"><div class="vsum-n" style="color:{UI['ok']}">{fc['pass']}</div>
            <div class="vsum-l">Passed</div></div>
          <div class="vsum-c"><div class="vsum-n" style="color:{UI['warn']}">{fc['warn']}</div>
            <div class="vsum-l">Warning</div></div>
          <div class="vsum-c"><div class="vsum-n" style="color:{UI['err']}">{fc['fail']}</div>
            <div class="vsum-l">Failed</div></div>
          <div class="vsum-c"><div class="vsum-n" style="color:{UI['cyan']}">{_qa_score}<span
            style="font-size:15px;color:{UI['tx3']}">%</span></div>
            <div class="vsum-l">QA Score</div></div>
        </div>''', unsafe_allow_html=True)

        # Category is inferred from the metric name already present in each message.
        def _cat_of(m):
            for key in ("Total EUI", "TEDI", "Heating", "Cooling", "Fan", "Lighting",
                        "DHW", "Receptacle", "Pumps", "GHGI"):
                if key.lower() in m.lower():
                    return key
            return "General"

        _ICO = {"pass": "check", "warn": "alert", "fail": "x", "info": "info"}
        for level, _flag_icon, msg in flags:
            _label, _col = _vt.get(level, _vt["info"])
            st.markdown(
                f'<div class="vf {level}"><div class="vf-b"></div>'
                f'<div class="vf-i">{icon(_ICO.get(level, "info"), 13)}</div>'
                f'<div class="vf-body"><div class="vf-h">'
                f'<span class="pill {"ok" if level=="pass" else "warn" if level=="warn" else "err" if level=="fail" else "info"}">'
                f'<i></i>{_label}</span>'
                f'<span class="vf-cat">{_cat_of(msg)}</span></div>'
                f'<div class="vf-m">{msg}</div></div></div>',
                unsafe_allow_html=True)

        st.divider()

        # ── Export ──
        section("Export & Share", "Presentation-ready QA/QC documentation", "download")
        pdf_bytes = build_pdf_report(meta, kpis, bm, flags, pct, comparison_edited)

        cd1,cd2=st.columns(2)
        with cd1:
            st.download_button("Download PDF Report", data=pdf_bytes,
                file_name=f"QA_QC_{(meta['project_name'] or 'report').replace(' ','_')}_{meta['date']}.pdf",
                mime="application/pdf", use_container_width=True)
        with cd2:
            if st.button("Run Another Project", use_container_width=True):
                for k in ["step","vals","results","headers","csv_df","mapping","ref_csv_df","ref_vals","compliance_code"]:
                    if k in st.session_state: del st.session_state[k]
                st.session_state.step=1; st.rerun()

        st.divider()

        # ── Add to Benchmark ──
        section("Contribute to Benchmark Database", "Add this validated model to the portfolio", "database")
        # This section reads from the QA Status column. With every row overridden to 🟢
        # there are no fails and no reviews, so neither the error nor the warning shows.
        if comparison_edited is not None:
            fail_count = int((comparison_edited["QA Status"] == "🔴").sum())
            warn_count = int(comparison_edited["QA Status"].isin(["🟡", "🟠"]).sum())
            pass_count = int((comparison_edited["QA Status"] == "🟢").sum())
        else:
            fail_count = sum(1 for f in flags if f[0]=="fail")
            warn_count = sum(1 for f in flags if f[0]=="warn")
            pass_count = sum(1 for f in flags if f[0]=="pass")

        if fail_count > 0:
            st.error(f"❌ This model has **{fail_count} unresolved QA/QC fail(s)** (🔴) in the comparison table above. Override them to 🟢 after review — or fix the model — before adding to the benchmark database.")
        else:
            if warn_count > 0:
                st.warning(f"⚠️ This model has **{warn_count} review flag(s)**. You can still add it to the database, but review the flags first.")

            with st.expander("What happens after you submit?", expanded=False):
                st.markdown("""
Submitting sends this model to the review queue. It does **not** enter the
benchmark pool straight away.

**What gets recorded:**
- Your name and email, so the reviewer knows who to come back to
- Project name, building type, city, climate zone, subtype
- Total EUI, all end-use EUIs, GHGI
- Floor area, model type, phase, date, software

**The review step:**
- The row lands on the **Submissions** tab marked *Pending*, and the reviewer is emailed
- A senior reviewer approves or rejects it; approved rows are copied to
  **Project Results** with the reviewer's name and the date
- You get an email either way, with the reviewer's comment if there is one
- Median EUI and end-use percentages in the Benchmarks sheet are still **not
  changed automatically** — a reviewer recalculates those periodically from the
  accumulated approved projects
                """)

            with st.form("add_to_benchmark_form"):
                st.markdown("**Confirm submission details:**")
                fc1, fc2, fc3 = st.columns(3)
                with fc1:
                    confirm_name    = st.text_input("Project Name",    value=meta["project_name"])
                    # Remembered for the session so a modeller submitting several
                    # projects in a row types this once.
                    confirm_modeller = st.text_input(
                        "Modeller *", value=st.session_state.get("modeller_name", ""),
                        placeholder="First Last",
                        help="Who built this model — recorded with the submission.")
                    confirm_email = st.text_input(
                        "Modeller email *", value=st.session_state.get("modeller_email", ""),
                        placeholder="first.last@stantec.com",
                        help="Where the approval or rejection notice is sent.")
                with fc2:
                    confirm_subtype = st.text_input("Benchmark Subtype", value=meta.get("subtype","General"),
                                                    help="e.g. Boiler + VAV · Medium, Heat Pump + DOAS")
                    confirm_notes   = st.text_area("Notes for the reviewer (optional)", placeholder="e.g. NECB 2020 proposed model, 100% design stage", height=80)
                with fc3:
                    st.markdown("**Summary of values being submitted:**")
                    st.markdown(f"""
| Field | Value |
|---|---|
| Building Type | {meta['building_type']} |
| City | {meta['city']} · Zone {meta['climate_zone']} |
| Total EUI | **{kpis['total_eui']} kWh/m²·yr** |
| GHGI | {kpis['ghgi'] if kpis.get('ghgi') is not None else '—'} kgCO₂e/m²·yr |
| Floor Area | {round(kpis['area']):,} m² |
| QA/QC | ✅ {pass_count} pass · ⚠️ {warn_count} review |
                    """)

                submitted = st.form_submit_button("Submit for Review", type="primary", use_container_width=True)

                if submitted:
                    problems = []
                    if not confirm_name.strip():
                        problems.append("a project name")
                    if not confirm_modeller.strip():
                        problems.append("the modeller's name")
                    if not _valid_email(confirm_email):
                        problems.append("a valid modeller email")
                    if problems:
                        st.error("Please enter " + ", ".join(problems) + ".")
                    else:
                        # Keep the modeller's details for the rest of the session.
                        st.session_state.modeller_name = confirm_modeller.strip()
                        st.session_state.modeller_email = confirm_email.strip()
                        # Build the row for the "Submissions" sheet. Submission ID,
                        # timestamp, status and reviewer columns are filled in
                        # server-side by the Apps Script — never by this client.
                        project_row = [
                            confirm_modeller.strip(),
                            confirm_email.strip(),
                            confirm_name.strip(),
                            meta["building_type"],
                            meta["city"],
                            meta["climate_zone"],
                            confirm_subtype.strip(),
                            meta["software"],
                            meta["model_type"],
                            meta["phase"],
                            meta["date"],
                            round(kpis["area"]),
                            kpis["total_eui"],
                            kpis.get("tedi", 0),
                            kpis["elec_eui"],
                            kpis["gas_eui"],
                            kpis["heat_eui"],
                            kpis["cool_eui"],
                            kpis["fan_eui"],
                            kpis["ltg_eui"],
                            kpis["dhw_eui"],
                            kpis.get("recept_eui", 0),
                            kpis["pumps_eui"],
                            kpis.get("ghgi") if kpis.get("ghgi") is not None else "",
                            f"{pct}th" if pct else "N/A",
                            f"{pass_count} pass · {warn_count} review · {fail_count} fail",
                            confirm_notes.strip(),
                        ]

                        with st.spinner("Sending for review…"):
                            ok, msg = append_to_google_sheet(
                                SUBMISSIONS_SHEET_NAME, project_row, SUBMISSION_HEADERS,
                                pending_workflow=True)
                        if ok:
                            _ref = f" Reference **{msg}**." if msg and msg != "Success" else ""
                            st.success(f"✅ **{confirm_name}** has been submitted for review.{_ref} "
                                       f"The reviewer has been notified, and "
                                       f"{confirm_email.strip()} will get an email once it is "
                                       f"approved or rejected.")
                            st.info("💡 Approved projects are copied to the **Project Results** "
                                    "tab. A senior reviewer periodically recalculates the "
                                    "Benchmarks medians and percentiles from those.")
                        else:
                            # Fallback — show the data so user can manually add it
                            st.warning(f"⚠️ Could not write to Google Sheets automatically: {msg}")
                            _sa_email = sheets_account_email()
                            with st.expander("How to switch the automatic write on"):
                                if _sa_email:
                                    st.markdown(
                                        f"A service account **is** configured as "
                                        f"`{_sa_email}`. If the message above mentions "
                                        f"access, open the spreadsheet, press **Share**, "
                                        f"and give that address **Editor** rights.")
                                else:
                                    st.markdown(
                                        "No credentials are configured yet. Add **one** of "
                                        "these to `.streamlit/secrets.toml` (App settings → "
                                        "Secrets on Streamlit Community Cloud):\n\n"
                                        "- `SHEETS_WEBHOOK_URL` — an Apps Script web app "
                                        "bound to this spreadsheet. No Google Cloud project "
                                        "needed.\n"
                                        "- `GOOGLE_SERVICE_ACCOUNT` — the service-account "
                                        "JSON key, with the sheet shared to its "
                                        "`client_email` as Editor.\n\n"
                                        "`GOOGLE_SHEETS_SETUP.md` in the repo has both "
                                        "procedures step by step.")
                            st.markdown(f"**In the meantime, copy this row into the "
                                        f"'{SUBMISSIONS_SHEET_NAME}' sheet and set its "
                                        f"Status to Pending:**")
                            st.dataframe(pd.DataFrame([dict(zip(SUBMISSION_HEADERS, project_row))]),
                                        use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════════════════════
footer(len(BENCHMARKS))
