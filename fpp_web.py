#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FPP Web App - Facility Project Proposal
Web interface for generating IBM FPP Excel reports
"""

import os, sys, io
from datetime import date
import streamlit as st
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FPP Generator – IBM Apleona",
    page_icon="🏢",
    layout="wide",
)

# ── Constants ─────────────────────────────────────────────────────────────────
CONTRACT_NUMBERS = (
    "MSA-CW3790072\n"
    "FM SOW-CW3790289\n"
    "PA FM SOW-CW3796451-IBM Israel Ltd\n"
    "PA FM SOW-CW3796611-IBM Israel S&T Ltd"
)
PREPARED_BY = "Tomer Cohen"
FEE         = 0.06

LABOR_ROLES = [
    {"title": "Small Job Coordinator Sr",         "st_rate": 358.49, "ot_rate": 537.74},
    {"title": "Small Job Coordinator",            "st_rate": 286.79, "ot_rate": 430.19},
    {"title": "Facilities Engineer Sr",           "st_rate": 269.81, "ot_rate": 404.72},
    {"title": "Facilities Engineer",              "st_rate": 248.11, "ot_rate": 372.17},
    {"title": "Facilities Technician",            "st_rate": 162.26, "ot_rate": 243.39},
    {"title": "Facilities Administrator",         "st_rate": 130.19, "ot_rate": 195.29},
    {"title": "Space Administrator",              "st_rate": 216.04, "ot_rate": 324.06},
    {"title": "Handyman/Porter",                  "st_rate": 118.87, "ot_rate": 178.31},
    {"title": "Workplace Experience Manager",     "st_rate": 301.89, "ot_rate": 452.84},
    {"title": "Workplace Experience Coordinator", "st_rate": 301.89, "ot_rate": 452.84},
    {"title": "Concierge",                        "st_rate":  85.85, "ot_rate": 128.78},
]

SITES = ["PTK", "Haifa", "Shahar", "Be'er Sheva"]

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    [data-testid="stAppViewContainer"] { direction: rtl; }
    .main-title {
        background: linear-gradient(135deg, #1F3864, #2E75B6);
        color: white; padding: 28px 32px; border-radius: 12px;
        margin-bottom: 28px; text-align: center;
    }
    .main-title h1 { margin: 0; font-size: 2rem; }
    .main-title p  { margin: 6px 0 0; opacity: .85; font-size: 1rem; }
    .section-card {
        background: #f8fafc; border: 1px solid #dce6f1;
        border-radius: 10px; padding: 22px 26px; margin-bottom: 22px;
    }
    .section-title {
        color: #1F3864; font-size: 1.1rem; font-weight: 700;
        border-bottom: 2px solid #2E75B6; padding-bottom: 8px; margin-bottom: 16px;
    }
    .stButton > button {
        background: linear-gradient(135deg, #1F3864, #2E75B6);
        color: white; border: none; padding: 14px 40px;
        font-size: 1.1rem; border-radius: 8px; width: 100%;
        font-weight: 600; cursor: pointer;
    }
    .stButton > button:hover { opacity: .9; }
    .cost-row { background: #edf4fb; border-radius: 6px; padding: 10px 14px; margin-bottom: 8px; }
    label { font-weight: 600 !important; color: #1F3864 !important; }
</style>
""", unsafe_allow_html=True)

# ── Translation ───────────────────────────────────────────────────────────────
def translate(client, text):
    if not text or not any("\u0590" <= c <= "\u05FF" for c in text):
        return text
    msg = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=1024,
        messages=[{"role": "user", "content": (
            "Translate the following Hebrew text to professional English "
            "suitable for a formal facility management proposal submitted to IBM. "
            "Preserve line breaks. Return ONLY the translation.\n\n" + text
        )}]
    )
    return msg.content[0].text.strip()

# ── Excel helpers ─────────────────────────────────────────────────────────────
def _b():
    t = Side(style="thin", color="B8C4CE")
    return Border(left=t, right=t, top=t, bottom=t)

def _fill(c): return PatternFill("solid", fgColor=c)
def _font(bold=False, color="000000", size=9): return Font(name="Arial", bold=bold, color=color, size=size)
def _align(h="left", v="center", wrap=True): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def sc(cell, value=None, font=None, fill=None, align=None, fmt=None):
    if value is not None: cell.value = value
    if font:  cell.font      = font
    if fill:  cell.fill      = fill
    if align: cell.alignment = align
    if fmt:   cell.number_format = fmt
    cell.border = _b()

def generate_excel(data) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for col, w in zip("ABCDEFG", [46, 15, 12, 12, 17, 8, 17]):
        ws.column_dimensions[col].width = w

    br = _b()

    # Row 1 – title
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:G1")
    sc(ws["A1"], "Facility Project Proposal",
       font=_font(bold=True, color="FFFFFF", size=14),
       fill=_fill("1F3864"), align=_align("center"))

    # Rows 2-7 – header
    lf, lfl = _font(bold=True), _fill("F8F8F8")
    vf, vfl = _font(), _fill("F8F8F8")

    def hrow(row, label, value, height=16, va=None):
        ws.row_dimensions[row].height = height
        ws.merge_cells(f"B{row}:G{row}")
        sc(ws[f"A{row}"], label, font=lf, fill=lfl, align=_align())
        sc(ws[f"B{row}"], value, font=vf, fill=vfl, align=va or _align())

    hrow(2, "Project Name:", data["project_name"])
    hrow(3, "Site:",         data["site"])
    hrow(4, "Contract #:",   CONTRACT_NUMBERS, height=70, va=_align("left","top"))
    hrow(5, "Revision Date:",data["revision_date"])
    ws["B5"].number_format = "DD/MM/YYYY"
    hrow(6, "Fixed Price or Estimate/Not to Exceed (NTE)", "Fixed")
    hrow(7, "Proposal Prepared by:", PREPARED_BY)

    ws.row_dimensions[8].height = 6

    # Row 9 – scope label
    ws.row_dimensions[9].height = 18
    ws.merge_cells("A9:G9")
    sc(ws["A9"], "Scope of Work:", font=_font(bold=True), fill=_fill("CFE2F3"), align=_align())

    # Rows 10-11 – scope content
    ws.merge_cells("A10:G11")
    sc(ws["A10"], data["scope_of_work"], font=vf, fill=vfl, align=_align("left","top"))
    ws.row_dimensions[10].height = 65
    ws.row_dimensions[11].height = 65

    # Row 12 – financial label
    ws.row_dimensions[12].height = 18
    ws.merge_cells("A12:G12")
    sc(ws["A12"], "Financial Proposal:", font=_font(bold=True), fill=_fill("CFE2F3"), align=_align())

    # Row 13 – financial headers
    ws.row_dimensions[13].height = 28
    thf, thfl = _font(bold=True, color="FFFFFF"), _fill("2E75B6")
    for col, hdr in enumerate(["Description","Unit Price","Quantity","UoM","Michlol net ILS","Fee","Apleona net ILS"],1):
        sc(ws.cell(row=13,column=col,value=hdr), font=thf, fill=thfl, align=_align("center"))

    # Rows 14-23 – items
    alt, base = _fill("EBF3FB"), _fill("F8F8F8")
    for i, item in enumerate(data["items"]):
        row = 14 + i
        ws.row_dimensions[row].height = 18
        f = alt if i%2==0 else base
        sc(ws.cell(row=row,column=1), item["description"], font=vf, fill=f, align=_align())
        sc(ws.cell(row=row,column=2), item["unit_price"], font=_font(bold=True), fill=f, align=_align("right",wrap=False), fmt="#,##0.00")
        sc(ws.cell(row=row,column=3), item["quantity"], font=vf, fill=f, align=_align("center"))
        sc(ws.cell(row=row,column=4), item["uom"], font=vf, fill=f, align=_align("center"))
        sc(ws.cell(row=row,column=5,value=f"=B{row}*C{row}*1.05"), font=vf, fill=f, align=_align("right",wrap=False), fmt="#,##0.00")
        sc(ws.cell(row=row,column=6), FEE, font=vf, fill=f, align=_align("center"), fmt="0%")
        sc(ws.cell(row=row,column=7,value=f"=E{row}*(1+F{row})"), font=vf, fill=f, align=_align("right",wrap=False), fmt="#,##0.00")

    # Row 24 – total
    ws.row_dimensions[24].height = 20
    tf, tfl = _font(bold=True), _fill("BDD7EE")
    sc(ws["A24"], "Total net", font=tf, fill=tfl, align=_align())
    for col in range(2,8):
        ws.cell(row=24,column=col).fill=tfl; ws.cell(row=24,column=col).border=br
    sc(ws["E24"],"=SUM(E14:E23)", font=tf, fill=tfl, align=_align("right",wrap=False), fmt="#,##0.00")
    sc(ws["G24"],"=SUM(G14:G23)", font=tf, fill=tfl, align=_align("right",wrap=False), fmt="#,##0.00")

    # Row 25 – clarifications label
    ws.row_dimensions[25].height = 18
    ws.merge_cells("A25:G25")
    sc(ws["A25"],"Clarifications / Assumptions:", font=_font(bold=True), fill=_fill("CFE2F3"), align=_align())
    ws.row_dimensions[26].height = 6

    # Rows 27-30 – clarifications
    ws.merge_cells("A27:G30")
    sc(ws["A27"], data["clarifications"] or " ", font=vf, fill=vfl, align=_align("left","top"))
    for r in range(27,31): ws.row_dimensions[r].height = 20

    # Row 31 – labor label
    ws.row_dimensions[31].height = 18
    ws.merge_cells("A31:G31")
    sc(ws["A31"],"FM Provider Self Performed Labor Breakdown:", font=_font(bold=True), fill=_fill("CFE2F3"), align=_align())

    # Row 32 – labor headers
    ws.row_dimensions[32].height = 18
    for col, hdr in enumerate(["Job Title","ST Hours","ST Rate","ST Total","OT Hours","OT Rate","OT Total"],1):
        sc(ws.cell(row=32,column=col,value=hdr), font=thf, fill=thfl, align=_align("center"))

    # Rows 33-43 – labor
    for i, role in enumerate(data["labor_roles"]):
        row = 33+i
        ws.row_dimensions[row].height = 14
        f = alt if i%2==0 else base
        sc(ws.cell(row=row,column=1), role["title"], font=_font(bold=True), fill=f, align=_align())
        sc(ws.cell(row=row,column=2), role["st_hours"], font=_font(bold=True), fill=f, align=_align("center"))
        sc(ws.cell(row=row,column=3), role["st_rate"], font=vf, fill=f, align=_align("center"), fmt="#,##0.00")
        sc(ws.cell(row=row,column=4,value=f"=B{row}*C{row}"), font=vf, fill=f, align=_align("right",wrap=False), fmt="#,##0.00")
        sc(ws.cell(row=row,column=5), role["ot_hours"], font=_font(bold=True), fill=f, align=_align("center"))
        sc(ws.cell(row=row,column=6), role["ot_rate"], font=vf, fill=f, align=_align("center"), fmt="#,##0.00")
        sc(ws.cell(row=row,column=7,value=f"=E{row}*F{row}"), font=vf, fill=f, align=_align("right",wrap=False), fmt="#,##0.00")

    # Row 44 – labor total
    last = 33+len(data["labor_roles"])-1
    trow = last+1
    ws.row_dimensions[trow].height = 18
    sc(ws.cell(row=trow,column=1), "TOTAL", font=tf, fill=tfl, align=_align())
    for col in range(2,8):
        ws.cell(row=trow,column=col).fill=tfl; ws.cell(row=trow,column=col).border=br
    sc(ws.cell(row=trow,column=4,value=f"=SUM(D33:D{last})"), font=tf, fill=tfl, align=_align("right",wrap=False), fmt="#,##0.00")
    sc(ws.cell(row=trow,column=7,value=f"=SUM(G33:G{last})"), font=tf, fill=tfl, align=_align("right",wrap=False), fmt="#,##0.00")

    ws.print_area = f"A1:G{trow}"
    ws.page_setup.orientation="landscape"; ws.page_setup.fitToPage=True
    ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-title">
    <h1>🏢 Facility Project Proposal</h1>
    <p>Apleona Israel · IBM Account · FPP Generator</p>
</div>
""", unsafe_allow_html=True)

# ── Section 1: Project Details ────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">📋 פרטי הפרויקט</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    project_name = st.text_input("שם הפרויקט *", placeholder="לדוגמה: החלפת מזגן במשרד PTK")
with col2:
    site = st.selectbox("אתר *", SITES)
st.markdown('</div>', unsafe_allow_html=True)

# ── Section 2: Scope of Work ──────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">📝 מהות העבודה</div>', unsafe_allow_html=True)
st.caption("תאר את העבודה בחופשיות בעברית — יתורגם לאנגלית מקצועית אוטומטית")
scope_he = st.text_area("תיאור העבודה *", height=140, placeholder="לדוגמה:\nהחלפת יחידת מיזוג אוויר בקומה 3\nפירוק היחידה הישנה ופינוי לאתר פסולת מאושר\nהתקנת יחידה חדשה כולל בדיקות הרצה")
st.markdown('</div>', unsafe_allow_html=True)

# ── Section 3: Costs ──────────────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">💰 עלויות</div>', unsafe_allow_html=True)
st.caption("הכנס כל פריט עלות — המערכת תוסיף 6% אוטומטית ותחשב סה\"כ")

cost_items = []
for i in range(10):
    cols = st.columns([3, 1.5, 0.3])
    with cols[0]:
        desc = st.text_input(f"תיאור פריט {i+1}", key=f"desc_{i}",
                             placeholder="לדוגמה: אספקת מזגן 24BTU" if i==0 else "")
    with cols[1]:
        price = st.number_input(f"מחיר (₪)", key=f"price_{i}", min_value=0.0, step=100.0, format="%.2f")
    with cols[2]:
        st.markdown("<br>", unsafe_allow_html=True)
    if desc:
        cost_items.append({"description_he": desc, "unit_price": price})

if cost_items:
    subtotal = sum(i["unit_price"] for i in cost_items)
    total_with_fee = subtotal * 1.05 * 1.06
    st.markdown(f"""
    <div style="background:#ddeeff;border-radius:8px;padding:12px 18px;margin-top:12px;text-align:left">
        <b>תת-סה"כ לפני עמלה:</b> ₪{subtotal:,.2f} &nbsp;|&nbsp;
        <b>סה"כ כולל 5% מיקלול + 6% עמלה:</b> <span style="color:#1F3864;font-size:1.1em">₪{total_with_fee:,.2f}</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ── Section 4: Labor ──────────────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">👷 שעות עובד נוסף</div>', unsafe_allow_html=True)
st.caption("הכנס שעות לתפקידים הרלוונטיים בלבד (0 = לא רלוונטי)")

labor_data = []
for role in LABOR_ROLES:
    col1, col2, col3 = st.columns([2.5, 1, 1.8])
    with col1:
        st.markdown(f"<span style='font-weight:600'>{role['title']}</span>", unsafe_allow_html=True)
    with col2:
        hours = st.number_input("שעות", key=f"h_{role['title']}", min_value=0, step=1, label_visibility="collapsed")
    with col3:
        if hours > 0:
            shift = st.radio("משמרת", ["בשעות העבודה (ST)", "מחוץ לשעות (OT)"],
                             key=f"s_{role['title']}", horizontal=True, label_visibility="collapsed")
            st_h = int(hours) if "ST" in shift else 0
            ot_h = int(hours) if "OT" in shift else 0
        else:
            st.markdown("")
            st_h = ot_h = 0
    labor_data.append({**role, "st_hours": st_h, "ot_hours": ot_h})

st.markdown('</div>', unsafe_allow_html=True)

# ── Section 5: Clarifications ─────────────────────────────────────────────────
st.markdown('<div class="section-card"><div class="section-title">📌 הבהרות והנחות</div>', unsafe_allow_html=True)
clarifications_he = st.text_area("הבהרות/הנחות (אופציונלי)", height=100,
                                  placeholder="לדוגמה: המחיר אינו כולל עבודות בנייה אזרחית")
st.markdown('</div>', unsafe_allow_html=True)

# ── Generate ──────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
if st.button("⚡ צור קובץ FPP"):
    if not project_name:
        st.error("אנא הכנס שם פרויקט")
    elif not scope_he:
        st.error("אנא הכנס תיאור עבודה")
    elif not cost_items:
        st.error("אנא הכנס לפחות פריט עלות אחד")
    else:
        with st.spinner("מתרגם לאנגלית ומייצר קובץ Excel..."):
            try:
                api_key = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY"))
                client = anthropic.Anthropic(api_key=api_key)

                project_name_en = translate(client, project_name)
                scope_en        = translate(client, scope_he)
                clari_en        = translate(client, clarifications_he) if clarifications_he else ""

                items_en = []
                for item in cost_items:
                    desc_en = translate(client, item["description_he"])
                    items_en.append({
                        "description": desc_en,
                        "unit_price":  item["unit_price"],
                        "quantity":    1,
                        "uom":         "Lump Sum",
                        "fee":         FEE,
                    })
                while len(items_en) < 10:
                    items_en.append({"description":"","unit_price":0,"quantity":0,"uom":"-","fee":FEE})

                excel_bytes = generate_excel({
                    "project_name":   project_name_en,
                    "site":           site,
                    "revision_date":  date.today(),
                    "scope_of_work":  scope_en,
                    "items":          items_en,
                    "clarifications": clari_en,
                    "labor_roles":    labor_data,
                })

                safe = "".join(c if c.isalnum() or c in " _-" else "_" for c in project_name_en)[:40].strip()
                filename = f"FPP_{safe}_{site}_{date.today().strftime('%Y%m%d')}.xlsx"

                st.success("✅ הקובץ מוכן להורדה!")
                st.download_button(
                    label="📥 הורד קובץ Excel",
                    data=excel_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"שגיאה: {e}")
