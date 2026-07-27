"""
Energy Model Benchmarking & QA/QC Platform
===========================================
Run:  streamlit run app.py
Deps: pip install streamlit pandas plotly openpyxl reportlab

IMPORTANT: Keep benchmarks.xlsx in the same folder as app.py
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
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root{
  --bg-0:#0E1117; --bg-1:#161B22; --card:#1E2530; --side:#101827;
  --blue:#0078D4; --cyan:#18B6F6; --ok:#32D583; --warn:#F79009; --err:#F04438;
  --tx:#FFFFFF; --tx2:#B8C2CC; --tx3:#7D8894;
  --bd:rgba(255,255,255,0.08); --bd2:rgba(255,255,255,0.14); --hov:rgba(255,255,255,0.05);
  --r:14px;
}

/* ── Canvas: blueprint grid + radial lighting ───────────────────────── */
.stApp{
  background-color:var(--bg-0);
  background-image:
    radial-gradient(900px 480px at 78% -8%, rgba(0,120,212,0.16), transparent 60%),
    radial-gradient(700px 420px at 8% 4%, rgba(24,182,246,0.09), transparent 60%),
    linear-gradient(rgba(255,255,255,0.022) 1px, transparent 1px),
    linear-gradient(90deg, rgba(255,255,255,0.022) 1px, transparent 1px);
  background-size:100% 100%,100% 100%,44px 44px,44px 44px;
  background-attachment:fixed;
}
.main .block-container{padding-top:1.1rem; padding-bottom:3rem; max-width:1500px;}
html,body,[class*="css"],.stApp,p,li,span,label,div,input,select,textarea,button{
  font-family:'Inter',-apple-system,'Segoe UI',system-ui,sans-serif;
}
::-webkit-scrollbar{width:10px;height:10px}
::-webkit-scrollbar-track{background:var(--bg-0)}
::-webkit-scrollbar-thumb{background:#2A3341;border-radius:8px;border:2px solid var(--bg-0)}
::-webkit-scrollbar-thumb:hover{background:#38455A}

/* ── Typography ─────────────────────────────────────────────────────── */
h1,h2,h3,h4,h5{color:var(--tx)!important;letter-spacing:-0.02em;font-family:'Inter',sans-serif!important}
h1{font-size:32px!important;font-weight:800!important;margin-bottom:.15rem!important}
h2{font-size:24px!important;font-weight:700!important}
h3{font-size:18px!important;font-weight:700!important}
h4{font-size:15px!important;font-weight:600!important;color:var(--tx2)!important}
p,li,label,.stMarkdown{color:var(--tx2);font-size:15px;line-height:1.62}
.stCaption,[data-testid="stCaptionContainer"],[data-testid="stCaptionContainer"] p{color:var(--tx3)!important;font-size:13px!important}
small{color:var(--tx3)}
strong,b{color:var(--tx)}
code{background:#0B0E14!important;color:var(--cyan)!important;border:1px solid var(--bd);
     border-radius:6px;padding:1px 6px;font-size:13px}
a{color:var(--cyan)!important;text-decoration:none}
a:hover{text-decoration:underline}
hr{border:none!important;border-top:1px solid var(--bd)!important;margin:1.5rem 0!important}
[data-testid="stMarkdownContainer"] table{border:1px solid var(--bd);border-radius:10px;overflow:hidden;
     border-collapse:separate;border-spacing:0;font-size:13px}
[data-testid="stMarkdownContainer"] th{background:#131A24!important;color:var(--tx2)!important;
     font-weight:600!important;border-bottom:1px solid var(--bd)!important;padding:8px 12px!important}
[data-testid="stMarkdownContainer"] td{color:var(--tx2)!important;border-bottom:1px solid var(--bd)!important;padding:8px 12px!important}

/* ── Motion ─────────────────────────────────────────────────────────── */
@keyframes riseIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes pulseDot{0%,100%{box-shadow:0 0 0 0 rgba(50,213,131,.5)}70%{box-shadow:0 0 0 7px rgba(50,213,131,0)}}
@keyframes flowLine{to{background-position:200% 0}}
@keyframes shimmer{0%{background-position:-500px 0}100%{background-position:500px 0}}
.rise{animation:riseIn .45s cubic-bezier(.22,1,.36,1) both}
.rise-1{animation-delay:.04s}.rise-2{animation-delay:.08s}.rise-3{animation-delay:.12s}
.rise-4{animation-delay:.16s}.rise-5{animation-delay:.20s}.rise-6{animation-delay:.24s}

/* ── Top header ─────────────────────────────────────────────────────── */
.ei-header{
  display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;
  background:linear-gradient(135deg,rgba(30,37,48,.94),rgba(16,24,39,.94));
  border:1px solid var(--bd);border-radius:16px;padding:14px 22px;margin-bottom:18px;
  backdrop-filter:blur(18px);box-shadow:0 8px 30px rgba(0,0,0,.42);
  animation:riseIn .5s cubic-bezier(.22,1,.36,1) both;
}
.ei-brand{display:flex;align-items:center;gap:13px}
.ei-mark{width:40px;height:40px;flex-shrink:0;filter:drop-shadow(0 3px 10px rgba(0,120,212,.5))}
.ei-name{font-size:19px;font-weight:800;color:#fff;letter-spacing:-.025em;line-height:1.15}
.ei-tag{font-size:11px;color:var(--tx3);letter-spacing:.14em;text-transform:uppercase;margin-top:2px;font-weight:600}
.ei-hactions{display:flex;align-items:center;gap:9px;flex-wrap:wrap}
.ei-pill{
  display:inline-flex;align-items:center;gap:7px;padding:7px 13px;border-radius:999px;
  border:1px solid var(--bd2);background:rgba(255,255,255,.045);
  font-size:11.5px;font-weight:700;color:var(--tx2);letter-spacing:.05em;text-transform:uppercase;
  transition:all .22s ease;white-space:nowrap;
}
.ei-pill:hover{background:var(--hov);border-color:rgba(255,255,255,.24);transform:translateY(-1px)}
.ei-pill.live{color:#B4F0D2;border-color:rgba(50,213,131,.42);background:rgba(50,213,131,.10)}
.ei-pill.ver{color:var(--cyan);border-color:rgba(24,182,246,.34);background:rgba(24,182,246,.09)}
.ei-dot{width:7px;height:7px;border-radius:50%;background:var(--ok);animation:pulseDot 2.4s infinite}
.ei-ico{width:34px;height:34px;border-radius:10px;display:inline-flex;align-items:center;justify-content:center;
  border:1px solid var(--bd);background:rgba(255,255,255,.035);color:var(--tx2);transition:all .22s ease;cursor:default}
.ei-ico:hover{background:var(--hov);color:#fff;border-color:var(--bd2);transform:translateY(-1px)}
.ei-avatar{width:34px;height:34px;border-radius:10px;display:inline-flex;align-items:center;justify-content:center;
  background:linear-gradient(135deg,var(--blue),var(--cyan));color:#fff;font-weight:800;font-size:12.5px;
  box-shadow:0 3px 12px rgba(0,120,212,.45)}

/* ── Page title block ───────────────────────────────────────────────── */
.ei-page{margin:2px 0 18px 0;animation:riseIn .5s cubic-bezier(.22,1,.36,1) both}
.ei-crumb{font-size:11.5px;color:var(--tx3);letter-spacing:.1em;text-transform:uppercase;font-weight:600;margin-bottom:7px}
.ei-crumb span{color:var(--cyan)}
.ei-h1{font-size:34px;font-weight:800;color:#fff;letter-spacing:-.03em;line-height:1.14}
.ei-sub{font-size:15px;color:var(--tx2);margin-top:7px;max-width:820px;line-height:1.6}

/* ── Section header ─────────────────────────────────────────────────── */
.ei-sec{display:flex;align-items:center;gap:11px;margin:26px 0 13px 0}
.ei-sec-bar{width:3px;height:20px;border-radius:3px;background:linear-gradient(180deg,var(--cyan),var(--blue))}
.ei-sec-t{font-size:19px;font-weight:700;color:#fff;letter-spacing:-.015em}
.ei-sec-d{font-size:13px;color:var(--tx3);margin-left:2px}

/* ── Cards ──────────────────────────────────────────────────────────── */
.ei-card{
  background:linear-gradient(160deg,rgba(30,37,48,.92),rgba(22,27,34,.92));
  border:1px solid var(--bd);border-radius:var(--r);padding:18px 20px;
  box-shadow:0 4px 20px rgba(0,0,0,.30);transition:all .26s cubic-bezier(.22,1,.36,1);height:100%;
}
.ei-card:hover{border-color:rgba(24,182,246,.34);transform:translateY(-2px);box-shadow:0 12px 34px rgba(0,0,0,.46)}

/* KPI card (Power BI style) */
.kpi{
  position:relative;overflow:hidden;
  background:linear-gradient(160deg,rgba(30,37,48,.94),rgba(22,27,34,.94));
  border:1px solid var(--bd);border-radius:var(--r);padding:16px 18px 14px 18px;
  box-shadow:0 4px 20px rgba(0,0,0,.30);transition:all .26s cubic-bezier(.22,1,.36,1);height:100%;
}
.kpi:hover{transform:translateY(-3px);border-color:rgba(24,182,246,.36);box-shadow:0 14px 36px rgba(0,0,0,.5)}
.kpi::after{content:'';position:absolute;top:0;left:0;right:0;height:2px;
  background:linear-gradient(90deg,var(--accent-c,var(--cyan)),transparent 78%);opacity:.85}
.kpi-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:9px}
.kpi-l{font-size:11px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.1em;line-height:1.35}
.kpi-i{width:30px;height:30px;border-radius:9px;display:flex;align-items:center;justify-content:center;flex-shrink:0;
  background:var(--accent-bg,rgba(24,182,246,.13));color:var(--accent-c,var(--cyan));border:1px solid var(--accent-bd,rgba(24,182,246,.24))}
.kpi-v{font-size:42px;font-weight:800;color:#fff;line-height:1;letter-spacing:-.035em;font-variant-numeric:tabular-nums}
.kpi-v.sm{font-size:32px}
.kpi-u{font-size:12px;color:var(--tx3);font-weight:600;margin-left:6px;letter-spacing:.01em}
.kpi-foot{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:10px;
  padding-top:9px;border-top:1px solid var(--bd)}
.kpi-cmp{font-size:11.5px;color:var(--tx3);line-height:1.4}
.kpi-trend{font-size:11.5px;font-weight:700;display:inline-flex;align-items:center;gap:4px;white-space:nowrap}

/* ── Status pills ───────────────────────────────────────────────────── */
.pill{display:inline-flex;align-items:center;gap:6px;padding:4px 11px;border-radius:999px;
  font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;white-space:nowrap}
.pill i{width:6px;height:6px;border-radius:50%;background:currentColor;display:inline-block}
.pill.ok{color:#7BE7B0;background:rgba(50,213,131,.13);border:1px solid rgba(50,213,131,.32)}
.pill.info{color:#8FD4FA;background:rgba(24,182,246,.13);border:1px solid rgba(24,182,246,.32)}
.pill.warn{color:#FBC66B;background:rgba(247,144,9,.13);border:1px solid rgba(247,144,9,.32)}
.pill.err{color:#FB9A93;background:rgba(240,68,56,.13);border:1px solid rgba(240,68,56,.32)}
.pill.idle{color:var(--tx3);background:rgba(255,255,255,.05);border:1px solid var(--bd)}

/* ── Workflow stepper ───────────────────────────────────────────────── */
.wf{display:flex;align-items:stretch;gap:0;margin:6px 0 20px 0;flex-wrap:nowrap;overflow-x:auto;padding-bottom:4px}
.wf-s{flex:1 1 0;min-width:132px;background:linear-gradient(160deg,rgba(30,37,48,.9),rgba(22,27,34,.9));
  border:1px solid var(--bd);border-radius:12px;padding:13px 14px;transition:all .28s cubic-bezier(.22,1,.36,1)}
.wf-s:hover{transform:translateY(-2px);border-color:var(--bd2)}
.wf-s.done{border-color:rgba(50,213,131,.30);background:linear-gradient(160deg,rgba(50,213,131,.07),rgba(22,27,34,.9))}
.wf-s.act{border-color:rgba(0,120,212,.55);background:linear-gradient(160deg,rgba(0,120,212,.15),rgba(22,27,34,.94));
  box-shadow:0 0 0 1px rgba(0,120,212,.28),0 10px 30px rgba(0,120,212,.20)}
.wf-n{width:25px;height:25px;border-radius:8px;display:flex;align-items:center;justify-content:center;
  font-size:11.5px;font-weight:800;margin-bottom:9px;border:1px solid var(--bd);
  background:rgba(255,255,255,.05);color:var(--tx3)}
.wf-s.done .wf-n{background:rgba(50,213,131,.16);color:#7BE7B0;border-color:rgba(50,213,131,.36)}
.wf-s.act .wf-n{background:linear-gradient(135deg,var(--blue),var(--cyan));color:#fff;border-color:transparent;
  box-shadow:0 3px 12px rgba(0,120,212,.5)}
.wf-t{font-size:12.5px;font-weight:700;color:var(--tx2);letter-spacing:-.01em;line-height:1.3}
.wf-s.act .wf-t{color:#fff}
.wf-d{font-size:10.5px;color:var(--tx3);margin-top:3px;line-height:1.35}
.wf-c{flex:0 0 26px;display:flex;align-items:center;justify-content:center;min-width:26px}
.wf-c i{display:block;width:100%;height:2px;border-radius:2px;background:var(--bd2)}
.wf-c.on i{background:linear-gradient(90deg,var(--blue),var(--cyan),var(--blue));
  background-size:200% 100%;animation:flowLine 2.2s linear infinite}

/* ── Validation feed (CI style) ─────────────────────────────────────── */
.vf{border:1px solid var(--bd);border-radius:11px;padding:12px 15px;margin-bottom:8px;
  display:flex;gap:12px;align-items:flex-start;transition:all .22s ease;
  background:linear-gradient(160deg,rgba(30,37,48,.86),rgba(22,27,34,.86))}
.vf:hover{border-color:var(--bd2);transform:translateX(2px)}
.vf-b{width:3px;border-radius:3px;align-self:stretch;flex-shrink:0}
.vf-i{width:22px;height:22px;border-radius:7px;display:flex;align-items:center;justify-content:center;
  font-size:12px;font-weight:800;flex-shrink:0;margin-top:1px}
.vf-body{flex:1;min-width:0}
.vf-h{display:flex;align-items:center;gap:9px;flex-wrap:wrap;margin-bottom:3px}
.vf-cat{font-size:10.5px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.09em}
.vf-m{font-size:13.5px;color:var(--tx2);line-height:1.55}
.vf.pass{border-left:none}
.vf.pass .vf-b{background:var(--ok)} .vf.pass .vf-i{background:rgba(50,213,131,.15);color:#7BE7B0}
.vf.warn .vf-b{background:var(--warn)} .vf.warn .vf-i{background:rgba(247,144,9,.15);color:#FBC66B}
.vf.fail .vf-b{background:var(--err)} .vf.fail .vf-i{background:rgba(240,68,56,.15);color:#FB9A93}
.vf.info .vf-b{background:var(--cyan)} .vf.info .vf-i{background:rgba(24,182,246,.15);color:#8FD4FA}

/* summary counters */
.vsum{display:flex;gap:11px;flex-wrap:wrap;margin-bottom:14px}
.vsum-c{flex:1;min-width:118px;border:1px solid var(--bd);border-radius:12px;padding:13px 16px;
  background:linear-gradient(160deg,rgba(30,37,48,.9),rgba(22,27,34,.9));transition:all .24s ease}
.vsum-c:hover{transform:translateY(-2px);border-color:var(--bd2)}
.vsum-n{font-size:30px;font-weight:800;line-height:1;letter-spacing:-.03em;font-variant-numeric:tabular-nums}
.vsum-l{font-size:10.5px;font-weight:700;color:var(--tx3);text-transform:uppercase;letter-spacing:.1em;margin-top:6px}

/* ── Upload dropzone ────────────────────────────────────────────────── */
.dz{border:1.5px dashed rgba(24,182,246,.34);border-radius:16px;padding:30px 24px;text-align:center;
  background:radial-gradient(560px 200px at 50% 0%,rgba(0,120,212,.12),transparent 70%),rgba(22,27,34,.62);
  transition:all .3s cubic-bezier(.22,1,.36,1);margin-bottom:10px}
.dz:hover{border-color:rgba(24,182,246,.65);background:radial-gradient(560px 200px at 50% 0%,rgba(0,120,212,.20),transparent 70%),rgba(22,27,34,.8);
  box-shadow:0 0 0 4px rgba(0,120,212,.10),0 16px 44px rgba(0,0,0,.4);transform:translateY(-2px)}
.dz-t{font-size:16.5px;font-weight:700;color:#fff;margin-top:11px;letter-spacing:-.01em}
.dz-d{font-size:13px;color:var(--tx3);margin-top:5px}
.dz-f{display:flex;gap:7px;justify-content:center;flex-wrap:wrap;margin-top:15px}
.dz-f span{font-size:10.5px;font-weight:700;color:var(--tx2);padding:5px 11px;border-radius:7px;
  border:1px solid var(--bd);background:rgba(255,255,255,.04);letter-spacing:.05em;transition:all .2s ease}
.dz-f span:hover{border-color:rgba(24,182,246,.4);color:var(--cyan)}

/* ── Empty state ────────────────────────────────────────────────────── */
.es{text-align:center;padding:42px 24px;border:1px dashed var(--bd2);border-radius:16px;background:rgba(22,27,34,.5)}
.es-t{font-size:16px;font-weight:700;color:var(--tx2);margin-top:13px}
.es-d{font-size:13.5px;color:var(--tx3);margin-top:6px;max-width:430px;margin-left:auto;margin-right:auto;line-height:1.6}

/* ── Footer ─────────────────────────────────────────────────────────── */
.ei-foot{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;
  margin-top:32px;padding:15px 20px;border-top:1px solid var(--bd);border-radius:12px;
  background:rgba(16,24,39,.5);font-size:11.5px;color:var(--tx3)}
.ei-foot b{color:var(--tx2);font-weight:600}
.ei-foot-r{display:flex;gap:16px;flex-wrap:wrap;align-items:center}

/* ── Sidebar ────────────────────────────────────────────────────────── */
section[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#101827 0%,#0B111C 100%);
  border-right:1px solid var(--bd);
}
section[data-testid="stSidebar"] > div{padding-top:1.1rem}
section[data-testid="stSidebar"] *{color:var(--tx2)}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3{color:#fff!important}
section[data-testid="stSidebar"] hr{border-top:1px solid var(--bd)!important;margin:.9rem 0!important}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p{color:var(--tx3)!important}

/* nav label = enterprise nav item */
section[data-testid="stSidebar"] div[role="radiogroup"]{gap:5px!important;display:flex;flex-direction:column}
section[data-testid="stSidebar"] div[role="radiogroup"] > label{
  display:flex!important;align-items:center;gap:11px;padding:10px 12px;border-radius:11px;
  border:1px solid transparent;cursor:pointer;transition:all .22s cubic-bezier(.22,1,.36,1);
  margin:0!important;position:relative;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{background:var(--hov);border-color:var(--bd)}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:last-child{
  font-size:13.5px!important;font-weight:600!important;color:var(--tx2)!important;letter-spacing:-.01em}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){
  background:linear-gradient(100deg,rgba(0,120,212,.20),rgba(24,182,246,.06));
  border-color:rgba(0,120,212,.44);box-shadow:0 0 22px rgba(0,120,212,.16),inset 0 0 0 1px rgba(255,255,255,.03);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked)::before{
  content:'';position:absolute;left:0;top:19%;height:62%;width:3px;border-radius:0 3px 3px 0;
  background:linear-gradient(180deg,var(--cyan),var(--blue));box-shadow:0 0 12px rgba(24,182,246,.85);
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) > div:last-child{color:#fff!important;font-weight:700!important}
section[data-testid="stSidebar"] div[role="radiogroup"] input{accent-color:var(--blue)}

section[data-testid="stSidebar"] button{
  background:rgba(255,255,255,.05)!important;border:1px solid var(--bd2)!important;
  color:var(--tx2)!important;font-weight:600!important;border-radius:10px!important;font-size:13px!important;
  transition:all .22s ease!important}
section[data-testid="stSidebar"] button:hover{
  background:rgba(0,120,212,.24)!important;border-color:rgba(0,120,212,.6)!important;
  color:#fff!important;transform:translateY(-1px)}

/* ── Inputs ─────────────────────────────────────────────────────────── */
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div{
  background:rgba(22,27,34,.9)!important;border:1px solid var(--bd)!important;border-radius:10px!important;
  color:var(--tx)!important;transition:all .2s ease!important;min-height:40px}
.stSelectbox div[data-baseweb="select"] > div:hover{border-color:rgba(24,182,246,.44)!important;background:rgba(30,37,48,.95)!important}
div[data-baseweb="select"] svg{color:var(--tx3)!important}
div[data-baseweb="popover"] li{background:var(--card)!important;color:var(--tx2)!important;font-size:13.5px!important}
div[data-baseweb="popover"] li:hover{background:rgba(0,120,212,.24)!important;color:#fff!important}
div[data-baseweb="popover"] ul{background:var(--card)!important;border:1px solid var(--bd2)!important;border-radius:10px!important}

.stTextInput input,.stNumberInput input,.stTextArea textarea{
  background:rgba(22,27,34,.9)!important;border:1px solid var(--bd)!important;border-radius:10px!important;
  color:var(--tx)!important;font-size:14px!important;transition:all .2s ease!important}
.stTextInput input:focus,.stNumberInput input:focus,.stTextArea textarea:focus{
  border-color:var(--blue)!important;box-shadow:0 0 0 3px rgba(0,120,212,.18)!important}
.stTextInput input::placeholder,.stTextArea textarea::placeholder{color:#5C6672!important}
.stNumberInput button{background:rgba(255,255,255,.05)!important;border-color:var(--bd)!important;color:var(--tx2)!important}
[data-testid="stWidgetLabel"] p,[data-testid="stWidgetLabel"] label{
  font-size:12px!important;font-weight:600!important;color:var(--tx2)!important;
  letter-spacing:.04em!important;text-transform:uppercase!important}

/* ── Buttons ────────────────────────────────────────────────────────── */
.stButton button,.stDownloadButton button,.stFormSubmitButton button{
  border-radius:10px!important;font-weight:600!important;font-size:13.5px!important;
  transition:all .24s cubic-bezier(.22,1,.36,1)!important;border:1px solid var(--bd2)!important;
  background:rgba(255,255,255,.05)!important;color:var(--tx)!important;padding:.5rem 1.1rem!important}
.stButton button:hover,.stDownloadButton button:hover,.stFormSubmitButton button:hover{
  background:var(--hov)!important;border-color:rgba(255,255,255,.26)!important;transform:translateY(-1px)}
.stButton button[kind="primary"],.stFormSubmitButton button[kind="primary"],
.stButton button[data-testid="baseButton-primary"]{
  background:linear-gradient(135deg,var(--blue),#0A94F0)!important;border-color:transparent!important;
  color:#fff!important;font-weight:700!important;box-shadow:0 4px 16px rgba(0,120,212,.42)!important}
.stButton button[kind="primary"]:hover,.stFormSubmitButton button[kind="primary"]:hover,
.stButton button[data-testid="baseButton-primary"]:hover{
  box-shadow:0 8px 26px rgba(0,120,212,.58)!important;transform:translateY(-2px)!important;
  background:linear-gradient(135deg,#0A88E8,var(--cyan))!important}
.stDownloadButton button{background:rgba(24,182,246,.13)!important;border-color:rgba(24,182,246,.36)!important;color:#8FD4FA!important}
.stDownloadButton button:hover{background:rgba(24,182,246,.24)!important;color:#fff!important}

/* ── Tabs ───────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"]{gap:5px;background:transparent;border-bottom:1px solid var(--bd);padding-bottom:0}
.stTabs [data-baseweb="tab"]{
  background:transparent!important;border-radius:10px 10px 0 0!important;padding:9px 17px!important;
  font-weight:600!important;font-size:13.5px!important;color:var(--tx3)!important;transition:all .22s ease!important}
.stTabs [data-baseweb="tab"]:hover{background:var(--hov)!important;color:var(--tx2)!important}
.stTabs [aria-selected="true"]{color:#fff!important;background:rgba(0,120,212,.14)!important}
.stTabs [data-baseweb="tab-highlight"]{background:linear-gradient(90deg,var(--blue),var(--cyan))!important;height:2px!important}
.stTabs [data-baseweb="tab-border"]{background:transparent!important}

/* ── Data grid ──────────────────────────────────────────────────────── */
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{
  border:1px solid var(--bd)!important;border-radius:12px!important;overflow:hidden!important;
  box-shadow:0 4px 20px rgba(0,0,0,.28)!important;background:rgba(22,27,34,.6)!important}
[data-testid="stDataFrame"] div[role="columnheader"],[data-testid="stDataEditor"] div[role="columnheader"]{
  background:#131A24!important;color:var(--tx2)!important;font-weight:700!important;
  text-transform:uppercase;letter-spacing:.06em;font-size:11px!important}

/* ── Alerts ─────────────────────────────────────────────────────────── */
div[data-testid="stAlert"]{border-radius:12px!important;border:1px solid var(--bd)!important;
  background:rgba(30,37,48,.85)!important;backdrop-filter:blur(10px)}
div[data-testid="stAlert"] p{color:var(--tx2)!important;font-size:13.5px!important}

/* ── Expander / file uploader / metric / progress ───────────────────── */
[data-testid="stExpander"]{border:1px solid var(--bd)!important;border-radius:12px!important;
  background:rgba(22,27,34,.72)!important;overflow:hidden}
[data-testid="stExpander"] summary{font-size:13.5px!important;font-weight:600!important;color:var(--tx2)!important}
[data-testid="stExpander"] summary:hover{background:var(--hov)!important}
[data-testid="stFileUploader"] section{
  background:rgba(22,27,34,.65)!important;border:1.5px dashed rgba(24,182,246,.30)!important;
  border-radius:14px!important;transition:all .28s cubic-bezier(.22,1,.36,1)!important;padding:18px!important}
[data-testid="stFileUploader"] section:hover{
  border-color:rgba(24,182,246,.62)!important;background:rgba(0,120,212,.10)!important;
  box-shadow:0 0 0 4px rgba(0,120,212,.10),0 14px 40px rgba(0,0,0,.42)!important;transform:translateY(-2px)}
[data-testid="stFileUploader"] section small,[data-testid="stFileUploader"] section span{color:var(--tx3)!important}
[data-testid="stFileUploader"] button{background:rgba(0,120,212,.18)!important;border-color:rgba(0,120,212,.42)!important;color:#8FD4FA!important}
[data-testid="stMetric"]{background:linear-gradient(160deg,rgba(30,37,48,.92),rgba(22,27,34,.92));
  border:1px solid var(--bd);border-radius:var(--r);padding:15px 17px;transition:all .26s ease}
[data-testid="stMetric"]:hover{transform:translateY(-2px);border-color:rgba(24,182,246,.34)}
[data-testid="stMetricLabel"] p{font-size:11px!important;font-weight:700!important;color:var(--tx3)!important;
  text-transform:uppercase!important;letter-spacing:.1em!important}
[data-testid="stMetricValue"]{font-size:27px!important;font-weight:800!important;color:#fff!important;letter-spacing:-.03em}
.stProgress > div > div > div{background:linear-gradient(90deg,var(--blue),var(--cyan))!important}
.stSpinner > div{border-top-color:var(--cyan)!important}
[data-testid="stForm"]{border:1px solid var(--bd)!important;border-radius:14px!important;
  background:rgba(22,27,34,.6)!important;padding:20px!important}
.js-plotly-plot .plotly .modebar{background:transparent!important}
.js-plotly-plot{border-radius:12px;overflow:hidden}
#MainMenu,footer,header[data-testid="stHeader"]{visibility:hidden;height:0}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  DESIGN SYSTEM — tokens, Lucide icons and reusable UI components
#  (Presentation layer only — no analytical behaviour lives here.)
# ══════════════════════════════════════════════════════════════════════════════
UI = {
    "bg": "#0E1117", "bg2": "#161B22", "card": "#1E2530", "side": "#101827",
    "blue": "#0078D4", "cyan": "#18B6F6", "ok": "#32D583", "warn": "#F79009",
    "err": "#F04438", "tx": "#FFFFFF", "tx2": "#B8C2CC", "tx3": "#7D8894",
    "grid": "rgba(255,255,255,0.07)",
}
APP_VERSION = "v2.0"

# Navigation destinations (each maps to a real, implemented page).
PAGE_QA       = "QA/QC Validation"
PAGE_EXPLORER = "Benchmark Explorer"
PAGE_DB       = "Benchmark Database"

# Chart palette — consistent across every figure in the app.
CHART_SEQ = ["#18B6F6", "#0078D4", "#7C5CFF", "#32D583", "#F79009",
             "#F04438", "#00C2A8", "#E879F9", "#8FA3B8"]

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

def _svg_uri(name, stroke="%23B8C2CC"):
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

def logo_mark(size=40):
    """Hexagon brand mark with blue/cyan gradient core and an energy bolt."""
    return f'''<svg class="ei-mark" width="{size}" height="{size}" viewBox="0 0 48 48" fill="none"
        xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="eiG" x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
          <stop stop-color="#18B6F6"/><stop offset=".55" stop-color="#0078D4"/><stop offset="1" stop-color="#F79009"/>
        </linearGradient>
        <linearGradient id="eiG2" x1="24" y1="6" x2="24" y2="42" gradientUnits="userSpaceOnUse">
          <stop stop-color="#FFFFFF" stop-opacity=".95"/><stop offset="1" stop-color="#FFFFFF" stop-opacity=".55"/>
        </linearGradient>
      </defs>
      <path d="M24 2.6 43.2 13.3v21.4L24 45.4 4.8 34.7V13.3L24 2.6z" fill="url(#eiG)" fill-opacity=".18"/>
      <path d="M24 2.6 43.2 13.3v21.4L24 45.4 4.8 34.7V13.3L24 2.6z" stroke="url(#eiG)" stroke-width="2.1"/>
      <path d="M24 9.6 37 16.9v14.2L24 38.4 11 31.1V16.9L24 9.6z" stroke="url(#eiG)" stroke-width="1" stroke-opacity=".45"/>
      <path d="M25.6 14.5 17.2 25.4h5.6l-1.4 8.6 8.4-11.4h-5.6l1.4-8.1z" fill="url(#eiG2)"/>
    </svg>'''

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

def sparkline(values, color="#18B6F6", w=104, h=26):
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

_TONE = {
    "ok":   ("#32D583", "rgba(50,213,131,.13)",  "rgba(50,213,131,.26)"),
    "warn": ("#F79009", "rgba(247,144,9,.13)",   "rgba(247,144,9,.26)"),
    "err":  ("#F04438", "rgba(240,68,56,.13)",   "rgba(240,68,56,.26)"),
    "blue": ("#0078D4", "rgba(0,120,212,.14)",   "rgba(0,120,212,.28)"),
    "cyan": ("#18B6F6", "rgba(24,182,246,.13)",  "rgba(24,182,246,.26)"),
    "idle": ("#7D8894", "rgba(255,255,255,.05)", "rgba(255,255,255,.10)"),
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
    """Apply the dark enterprise chart theme to any Plotly figure."""
    fig.update_layout(
        template="plotly_dark",
        height=height,
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=UI["tx2"], size=12),
        title=(dict(text=title, font=dict(size=14, color=UI["tx"], family="Inter, sans-serif"),
                    x=0, xanchor="left", y=.97) if title else None),
        margin=dict(t=48 if title else 18, b=12, l=10, r=12),
        showlegend=legend,
        legend=dict(orientation="h", y=-0.20, x=0, font=dict(size=11, color=UI["tx3"]),
                    bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="#1E2530", bordercolor="rgba(255,255,255,.18)",
                        font=dict(color="#FFFFFF", size=12, family="Inter, sans-serif")),
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
    col = UI["ok"] if lvl >= 0.75 else UI["cyan"] if lvl >= 0.5 else UI["warn"] if lvl >= 0.25 else UI["err"]
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=round(v),
        number=dict(font=dict(size=38, color="#FFFFFF", family="Inter, sans-serif"), suffix=""),
        gauge=dict(
            axis=dict(range=[0, vmax], tickwidth=1, tickcolor=UI["grid"],
                      tickfont=dict(color=UI["tx3"], size=10)),
            bar=dict(color=col, thickness=0.30),
            bgcolor="rgba(255,255,255,0.03)", borderwidth=0,
            steps=[dict(range=[0, vmax * .25], color="rgba(240,68,56,.10)"),
                   dict(range=[vmax * .25, vmax * .5], color="rgba(247,144,9,.10)"),
                   dict(range=[vmax * .5, vmax * .75], color="rgba(24,182,246,.10)"),
                   dict(range=[vmax * .75, vmax], color="rgba(50,213,131,.10)")],
            threshold=dict(line=dict(color="rgba(255,255,255,.55)", width=2),
                           thickness=0.8, value=v)),
    ))
    fig.update_layout(
        height=190, margin=dict(t=42, b=6, l=22, r=22),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=UI["tx2"]),
        title=dict(text=title, font=dict(size=12.5, color=UI["tx3"]), x=.5, xanchor="center", y=.97),
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


def append_to_google_sheet(sheet_name: str, row: list):
    """
    Append a row to a Google Sheet using the Sheets API via requests.
    Requires SHEET_ID and GOOGLE_SERVICE_ACCOUNT in st.secrets.
    Falls back gracefully if credentials are not configured.
    """
    try:
        import json, requests
        creds_raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT", None)
        if not creds_raw:
            return False, "No Google service account configured in secrets."

        creds = json.loads(creds_raw) if isinstance(creds_raw, str) else dict(creds_raw)

        # Get OAuth2 token using service account
        import time, base64, hashlib
        from urllib.parse import urlencode

        # Build JWT
        header = base64.urlsafe_b64encode(
            json.dumps({"alg":"RS256","typ":"JWT"}).encode()
        ).rstrip(b"=").decode()

        now = int(time.time())
        claim = base64.urlsafe_b64encode(json.dumps({
            "iss": creds["client_email"],
            "scope": "https://www.googleapis.com/auth/spreadsheets",
            "aud": "https://oauth2.googleapis.com/token",
            "iat": now, "exp": now + 3600
        }).encode()).rstrip(b"=").decode()

        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import padding

        private_key = serialization.load_pem_private_key(
            creds["private_key"].encode(), password=None
        )
        sig_input = f"{header}.{claim}".encode()
        signature = base64.urlsafe_b64encode(
            private_key.sign(sig_input, padding.PKCS1v15(), hashes.SHA256())
        ).rstrip(b"=").decode()

        jwt = f"{header}.{claim}.{signature}"

        token_resp = requests.post("https://oauth2.googleapis.com/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": jwt,
        })
        access_token = token_resp.json().get("access_token")
        if not access_token:
            return False, f"Could not get access token: {token_resp.text}"

        # Append row to sheet
        url = (f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
               f"/values/{sheet_name}!A1:append?valueInputOption=USER_ENTERED"
               f"&insertDataOption=INSERT_ROWS")
        resp = requests.post(url,
            headers={"Authorization": f"Bearer {access_token}"},
            json={"values": [row]}
        )
        if resp.status_code == 200:
            return True, "Success"
        else:
            return False, f"Sheets API error: {resp.text}"
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
    """Dark donut chart with the aggregate total rendered in the centre."""
    total = sum(values) if values else 0
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.62, sort=False,
        marker=dict(colors=CHART_SEQ[:len(labels)],
                    line=dict(color="rgba(14,17,23,.85)", width=2.5)),
        textinfo="percent", textposition="outside",
        textfont=dict(size=11.5, color=UI["tx2"], family="Inter, sans-serif"),
        hovertemplate="<b>%{label}</b><br>%{value:.1f} kWh/m\u00b2\u00b7yr &middot; %{percent}<extra></extra>",
    ))
    fig.add_annotation(text=f"<b>{total:,.0f}</b>", showarrow=False,
                       font=dict(size=27, color="#FFFFFF", family="Inter, sans-serif"), y=.545)
    fig.add_annotation(text="kWh/m\u00b2\u00b7yr", showarrow=False,
                       font=dict(size=10.5, color=UI["tx3"], family="Inter, sans-serif"), y=.40)
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
    h1    = ParagraphStyle("h1", parent=styles["Heading1"], textColor=colors.HexColor("#0E1117"), fontSize=17, spaceAfter=2)
    h2    = ParagraphStyle("h2", parent=styles["Heading2"], textColor=colors.HexColor("#0E1117"), fontSize=12, spaceBefore=8, spaceAfter=4)
    body  = styles["BodyText"]
    small = ParagraphStyle("small", parent=body, fontSize=9, textColor=colors.HexColor("#64748b"))
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
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0078D4")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0078D4")),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#e2e8f0")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
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
    f'content:"";width:30px;height:30px;border-radius:9px;order:-1;flex-shrink:0;'
    f'border:1px solid rgba(255,255,255,.08);background-color:rgba(255,255,255,.04);'
    f'background-image:url("data:image/svg+xml;utf8,{_svg_uri(ic)}");'
    f'background-repeat:no-repeat;background-position:center;background-size:15px 15px;'
    f'transition:all .22s ease}}'
    f'section[data-testid="stSidebar"] div[role="radiogroup"] > label:nth-of-type({i}):has(input:checked)::after{{'
    f'background-image:url("data:image/svg+xml;utf8,{_svg_uri(ic, "%2318B6F6")}");'
    f'border-color:rgba(24,182,246,.42);background-color:rgba(24,182,246,.14);'
    f'box-shadow:0 0 14px rgba(24,182,246,.30)}}'
    for i, (_lbl, ic, _d) in enumerate(NAV, 1)
)
st.markdown(f"<style>{_NAV_ICONS}</style>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown(f'''<div style="display:flex;align-items:center;gap:11px;padding:2px 2px 16px 2px">
      {logo_mark(36)}
      <div><div style="font-size:16.5px;font-weight:800;color:#fff;letter-spacing:-.025em;line-height:1.15">Energy Intelligence</div>
      <div style="font-size:9.5px;color:{UI['tx3']};letter-spacing:.14em;text-transform:uppercase;margin-top:2px;font-weight:700">Stantec Buildings</div></div>
    </div>''', unsafe_allow_html=True)

    st.markdown(f'''<div style="font-size:10px;font-weight:700;color:{UI['tx3']};letter-spacing:.14em;
        text-transform:uppercase;margin:2px 0 8px 2px">Workspace</div>
      <div style="display:flex;align-items:center;gap:9px;padding:9px 11px;border-radius:10px;
        border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.035);margin-bottom:14px">
        <span style="width:24px;height:24px;border-radius:7px;background:linear-gradient(135deg,{UI['blue']},{UI['cyan']});
          display:flex;align-items:center;justify-content:center;color:#fff;font-weight:800;font-size:10.5px">BP</span>
        <div style="line-height:1.25"><div style="font-size:12.5px;font-weight:700;color:#fff">Buildings Practice</div>
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
                _bg, _fg, _tc, _lb, _bd = "rgba(50,213,131,.16)", UI["ok"], UI["tx2"], icon("check", 12), "rgba(50,213,131,.34)"
            elif _i == _cur:
                _bg, _fg, _tc, _lb, _bd = "rgba(0,120,212,.24)", "#fff", "#fff", str(_i), "rgba(0,120,212,.55)"
            else:
                _bg, _fg, _tc, _lb, _bd = "rgba(255,255,255,.04)", UI["tx3"], UI["tx3"], str(_i), "rgba(255,255,255,.08)"
            _h += (f'<div style="display:flex;align-items:center;gap:10px">'
                   f'<span style="width:22px;height:22px;border-radius:7px;background:{_bg};color:{_fg};'
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
        <span style="font-size:12.5px;color:#fff;font-weight:800">{len(BENCHMARKS)}</span></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:11px;color:{UI['tx3']};font-weight:600">Data source</span>
        <span style="font-size:11px;color:{UI['ok']};font-weight:700">&#9679; Synced</span></div>
      <div style="display:flex;align-items:center;justify-content:space-between">
        <span style="font-size:11px;color:{UI['tx3']};font-weight:600">Environment</span>
        <span style="font-size:11px;color:{UI['cyan']};font-weight:700">Production</span></div>
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
                                 marker=dict(color=bar_med_f, colorscale=[[0, "#0A5F9E"], [1, UI["cyan"]]],
                                             line=dict(width=0)),
                                 hovertemplate="<b>%{x}</b><br>%{y:.1f} kWh/m\u00b2\u00b7yr<extra></extra>",
                                 error_y=dict(type="data", symmetric=True,
                                              array=[round(v*0.15,1) for v in bar_med_f],
                                              color="rgba(255,255,255,.34)", thickness=1.4, width=4)))
        fig_bar.update_traces(marker_cornerradius=6)
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
        bar_colors  = [UI["ok"] if v<=bm["good_eui"] else UI["cyan"] if v<=bm["median_eui"] else UI["warn"] if v<=bm["high_eui"] else UI["err"] for v in sorted_pcts]
        fig_pct = go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)], y=sorted_pcts, marker_color=bar_colors,
            marker_cornerradius=6,
            hovertemplate="<b>%{x} percentile</b><br>EUI: %{y} kWh/m\u00b2\u00b7yr<extra></extra>"))
        fig_pct.add_hline(y=bm["good_eui"],   line_dash="dot",  line_color=UI["ok"],   line_width=1.4, opacity=.8)
        fig_pct.add_hline(y=bm["median_eui"], line_dash="dash", line_color=UI["cyan"], line_width=1.8, opacity=.9)
        fig_pct.add_hline(y=bm["high_eui"],   line_dash="dot",  line_color=UI["err"],  line_width=1.4, opacity=.8)
        st.plotly_chart(style_fig(fig_pct, 310, legend=False, xtitle="Percentile", ytitle="EUI (kWh/m\u00b2\u00b7yr)"),
                        use_container_width=True, config={"displayModeBar": False})
        st.caption(f"Good (\u221215%): {bm['good_eui']}  \u00b7  Median: {bm['median_eui']}  \u00b7  High flag (+15%): {bm['high_eui']}  kWh/m\u00b2\u00b7yr")

        if bm.get("median_tedi") and bm.get("tedi_pct_data"):
            section("Percentile Distribution \u00b7 TEDI", "Thermal demand across the portfolio", "flame")
            sorted_tedi = sorted(bm["tedi_pct_data"])
            tedi_colors = [UI["ok"] if v<=bm["good_tedi"] else UI["cyan"] if v<=bm["median_tedi"] else UI["warn"] if v<=bm["high_tedi"] else UI["err"] for v in sorted_tedi]
            fig_tedi = go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)], y=sorted_tedi, marker_color=tedi_colors,
                marker_cornerradius=6,
                hovertemplate="<b>%{x} percentile</b><br>TEDI: %{y} kWh/m\u00b2\u00b7yr<extra></extra>"))
            fig_tedi.add_hline(y=bm["good_tedi"],   line_dash="dot",  line_color=UI["ok"],   line_width=1.4, opacity=.8)
            fig_tedi.add_hline(y=bm["median_tedi"], line_dash="dash", line_color=UI["cyan"], line_width=1.8, opacity=.9)
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
            pct_block = (f'<div style="text-align:right;padding-left:22px;border-left:1px solid rgba(255,255,255,.10)">'
                         f'<div style="font-size:44px;font-weight:800;color:#fff;line-height:1;letter-spacing:-.04em">'
                         f'{pct}<span style="font-size:19px;color:{UI["tx3"]}">th</span></div>'
                         f'<div style="font-size:10.5px;color:{UI["tx3"]};text-transform:uppercase;'
                         f'letter-spacing:.11em;font-weight:700;margin-top:5px">Percentile &middot; lower is better</div></div>')

        st.markdown(f'''<div class="ei-card rise" style="border-color:{overall_accent}44;
            background:linear-gradient(140deg,{overall_accent}14,rgba(22,27,34,.94));padding:22px 26px;margin-bottom:6px">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:22px;flex-wrap:wrap">
            <div style="min-width:260px">
              <div style="display:flex;align-items:center;gap:11px;flex-wrap:wrap">
                <span style="font-size:27px;font-weight:800;color:#fff;letter-spacing:-.03em">{proj_name}</span>
                {pill(overall, overall_tone)}
              </div>
              <div style="font-size:13px;color:{UI['tx3']};margin-top:7px">{info_line}</div>
              <div style="display:flex;gap:9px;margin-top:13px;flex-wrap:wrap">
                {pill(f"{fc['pass']} Pass", "ok")}{pill(f"{fc['warn']} Review", "warn")}{pill(f"{fc['fail']} Fail", "err")}
              </div>
            </div>
            {pct_block}
          </div>
        </div>''', unsafe_allow_html=True)

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
                bar_colors=[UI["ok"] if v<=bm["good_eui"] else UI["cyan"] if v<=bm["median_eui"] else UI["warn"] if v<=bm["high_eui"] else UI["err"] for v in sorted_pcts]
                fig_pct=go.Figure(go.Bar(x=[f"{i*10}th" for i in range(1,11)],y=sorted_pcts,marker_color=bar_colors,
                    marker_cornerradius=6, opacity=.55,
                    hovertemplate="<b>%{x} percentile</b><br>EUI: %{y} kWh/m\u00b2\u00b7yr<extra></extra>"))
                fig_pct.add_hline(y=kpis["total_eui"], line_dash="dash", line_color="#FFFFFF", line_width=2,
                    annotation_text=f"  Your model \u00b7 {kpis['total_eui']}",
                    annotation_font=dict(color="#FFFFFF", size=12), annotation_position="top left")
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
                marker_color="rgba(255,255,255,.16)", marker_cornerradius=5,
                hovertemplate="<b>%{x}</b><br>Benchmark %{y:.1f}<extra></extra>",
                error_y=dict(type="data", symmetric=True,
                             array=[round(v*0.15,1) for v in bm_med],
                             color="rgba(255,255,255,.34)", thickness=1.4, width=5),
            ))
            # Your model bars, labelled with % difference vs benchmark median
            pct_labels = [f"{(y-m)/m*100:+.0f}%" if m else "" for y, m in zip(your_eu, bm_med)]
            fig_cmp.add_trace(go.Bar(
                name="Your Model", x=eu_l2, y=your_eu,
                marker_color=UI["cyan"], marker_cornerradius=5,
                text=pct_labels, textposition="outside",
                textfont=dict(size=11, color=UI["tx2"], family="Inter, sans-serif"),
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
               "fail": ("FAILED", UI["err"]), "info": ("INFO", UI["cyan"])}
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

            with st.expander("What does adding to the benchmark do?", expanded=False):
                st.markdown("""
When you add this model to the benchmark database, it contributes to the pool of real projects used for future comparisons.

**What gets recorded:**
- Project name, building type, city, climate zone, subtype
- Total EUI, all end-use EUIs, GHGI
- Floor area, model type, phase, date, software

**How it updates the benchmark:**
- Your project's EUI is added to the percentile dataset for that building type/city/subtype
- The median EUI and end-use percentages in the Benchmarks sheet are **not changed automatically** — a senior reviewer can periodically recalculate and update those based on accumulated project data
- All submitted projects are visible in the **Project Results** tab of the Google Sheet for full audit trail
                """)

            with st.form("add_to_benchmark_form"):
                st.markdown("**Confirm submission details:**")
                fc1, fc2, fc3 = st.columns(3)
                with fc1:
                    confirm_name    = st.text_input("Project Name",    value=meta["project_name"])
                with fc2:
                    confirm_subtype = st.text_input("Benchmark Subtype", value=meta.get("subtype","General"),
                                                    help="e.g. Boiler + VAV · Medium, Heat Pump + DOAS")
                    confirm_notes   = st.text_area("Notes (optional)", placeholder="e.g. NECB 2020 proposed model, 100% design stage", height=80)
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

                submitted = st.form_submit_button("Add to Benchmark Database", type="primary", use_container_width=True)

                if submitted:
                    if not confirm_name.strip():
                        st.error("Please enter a project name.")
                    else:
                        # Build the row to append to "Project Results" sheet
                        project_row = [
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

                        ok, msg = append_to_google_sheet("Project Results", project_row)
                        if ok:
                            st.success(f"✅ **{confirm_name}** has been added to the benchmark database. Thank you!")
                            st.info("💡 A senior reviewer can periodically update the Benchmarks sheet median EUI and percentile data using the accumulated project results.")
                            load_benchmarks.clear()
                        else:
                            # Fallback — show the data so user can manually add it
                            st.warning(f"⚠️ Could not write to Google Sheets automatically: {msg}")
                            st.markdown("**Please manually copy this row into the 'Project Results' sheet in Google Sheets:**")
                            headers_pr = ["Project Name","Building Type","City","Climate Zone","Subtype",
                                          "Software","Model Type","Phase","Date","Area (m²)","Total EUI","TEDI","Elec EUI",
                                          "Gas EUI","Heating EUI","Cooling EUI","Fan EUI","Lighting EUI","DHW EUI",
                                          "Receptacle EUI","Pumps EUI","GHGI","Percentile","QA Flags","Notes"]
                            st.dataframe(pd.DataFrame([dict(zip(headers_pr, project_row))]),
                                        use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════════════════════
footer(len(BENCHMARKS))
