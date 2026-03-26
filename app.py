import streamlit as st
import json
import io
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)

W, H = A4
M    = 18 * mm
CW   = W - 2 * M

NAVY  = colors.HexColor("#0b1829")
TEAL  = colors.HexColor("#0099b8")
RED   = colors.HexColor("#dc2626")
AMBER = colors.HexColor("#d97706")
GREEN = colors.HexColor("#059669")
GREY  = colors.HexColor("#64748b")
LGREY = colors.HexColor("#f1f5f9")
MGREY = colors.HexColor("#e2e8f0")
BLACK = colors.HexColor("#0f172a")
WHITE = colors.white
LRED  = colors.HexColor("#fff5f5")
LCREAM= colors.HexColor("#fffbeb")
LGREEN= colors.HexColor("#f0fdf4")

st.set_page_config(page_title="MN Master Report", page_icon="📊", layout="wide")

st.markdown("""
<style>
body,.stApp{background:#f4f6f9;color:#0f172a}
.source-badge{display:inline-block;padding:3px 10px;border-radius:20px;
              font-size:11px;font-weight:600;margin-right:6px}
.badge-parks{background:#e0f2fe;color:#0369a1}
.badge-intel{background:#f0fdf4;color:#166534}
</style>
""", unsafe_allow_html=True)

if "parks_uploads" not in st.session_state:
    st.session_state.parks_uploads = []
if "intel_uploads" not in st.session_state:
    st.session_state.intel_uploads = []


def parse_upload(f):
    try:
        return json.loads(f.read())
    except Exception as e:
        st.error(f"Could not parse {f.name}: {e}")
        return None


def _styles():
    b = getSampleStyleSheet()
    def S(name, **kw):
        return ParagraphStyle(name, parent=b["Normal"], **kw)
    return {
        "h1":    S("h1",  fontName="Helvetica-Bold",  fontSize=20, textColor=NAVY,  leading=26),
        "h2":    S("h2",  fontName="Helvetica-Bold",  fontSize=14, textColor=NAVY,  leading=18),
        "h3":    S("h3",  fontName="Helvetica-Bold",  fontSize=11, textColor=NAVY,  leading=14),
        "body":  S("bd",  fontName="Helvetica",        fontSize=9,  textColor=BLACK, leading=13),
        "small": S("sm",  fontName="Helvetica",        fontSize=8,  textColor=GREY,  leading=11),
        "bold9": S("b9",  fontName="Helvetica-Bold",   fontSize=9,  textColor=BLACK, leading=13),
        "bold11":S("b11", fontName="Helvetica-Bold",   fontSize=11, textColor=BLACK, leading=15),
        "mono":  S("mo",  fontName="Courier",          fontSize=7.5,textColor=GREY,  leading=10),
        "monob": S("mob", fontName="Courier-Bold",     fontSize=7.5,textColor=NAVY,  leading=10),
        "monow": S("mow", fontName="Courier-Bold",     fontSize=7.5,textColor=WHITE, leading=10),
        "teal":  S("te",  fontName="Helvetica-Bold",   fontSize=9,  textColor=TEAL,  leading=12),
        "white": S("wh",  fontName="Helvetica",        fontSize=9,  textColor=WHITE, leading=13),
        "whiteb":S("wb",  fontName="Helvetica-Bold",   fontSize=9,  textColor=WHITE, leading=13),
        "italic":S("it",  fontName="Helvetica-Oblique",fontSize=9,  textColor=GREY,  leading=12),
        "red":   S("re",  fontName="Helvetica-Bold",   fontSize=9,  textColor=RED,   leading=12),
        "green": S("gr",  fontName="Helvetica-Bold",   fontSize=9,  textColor=GREEN, leading=12),
        "amber": S("am",  fontName="Helvetica-Bold",   fontSize=9,  textColor=AMBER, leading=12),
    }


def _bar(title, S):
    t = Table([[Paragraph(title, S["whiteb"])]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), NAVY),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
    ]))
    return t


def _hr():
    return HRFlowable(width=CW, thickness=0.5, color=MGREY, spaceAfter=4, spaceBefore=4)


def _footer(title):
    def _fn(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(GREY)
        canvas.drawString(M, 10*mm,
            f"Modern Networks — {title} — CONFIDENTIAL — INTERNAL USE ONLY")
        canvas.drawRightString(W-M, 10*mm, f"Page {doc.page}")
        canvas.restoreState()
    return _fn


def _score_colour(score):
    try:
        s = int(score)
        return RED if s < 50 else AMBER if s < 70 else GREEN
    except (TypeError, ValueError):
        return GREY


def _build_opportunities(parks_uploads, intel_uploads):
    opps = []

    # From science parks
    for upload in parks_uploads:
        for p in upload.get("parks", []):
            name     = p.get("name", "")
            postcode = p.get("postcode", "")
            sector   = p.get("sector", "")
            tenants  = p.get("tenants", "")
            ofcom    = p.get("ofcom", {}) or {}

            # Try to get gigabit pct from nested or flat ofcom data
            gig = 0
            try:
                gig = float(
                    ofcom.get("gigabit_pct") or
                    (ofcom.get("connectivity") or {}).get("gigabit_pct") or 0
                )
            except (TypeError, ValueError):
                gig = 0

            if gig < 50:
                opps.append({
                    "priority": "High" if gig < 20 else "Medium",
                    "property": name,
                    "postcode": postcode,
                    "type":     "Science Park",
                    "gap":      f"Gigabit coverage {gig:.0f}% — below standard for innovation park",
                    "service":  "Fibre Broadband · Network-as-a-Service · WiredScore AP Services",
                    "reason":   f"{tenants} tenants in {sector} — gigabit infrastructure required to attract and retain premium occupiers",
                })
            else:
                opps.append({
                    "priority": "Medium",
                    "property": name,
                    "postcode": postcode,
                    "type":     "Science Park",
                    "gap":      "WiredScore / SmartScore certification not confirmed",
                    "service":  "WiredScore AP Services · SmartScore AP Services",
                    "reason":   f"Certification differentiates {name} for premium tenant attraction — MN are Accredited Professionals",
                })

            # Large tenant base — managed IT
            try:
                t_num = int(str(tenants).replace("+","").replace(",","").split()[0])
            except (ValueError, IndexError):
                t_num = 0
            if t_num >= 100:
                opps.append({
                    "priority": "Medium",
                    "property": name,
                    "postcode": postcode,
                    "type":     "Science Park",
                    "gap":      f"{tenants} tenant organisations — managed IT and connectivity opportunity",
                    "service":  "Service Guardian · Managed Network · Desktop Support · M365",
                    "reason":   "Large multi-tenant environment with demand for a single managed services partner across all occupiers",
                })

    # From building intelligence
    for upload in intel_uploads:
        for b in upload.get("briefings", []):
            pc      = b.get("postcode", "")
            company = b.get("company", "") or pc
            score   = b.get("score", 50)
            verdict = b.get("verdict", "")
            gaps    = b.get("gaps", [])
            for g in gaps[:2]:
                opps.append({
                    "priority": "High" if g.get("sev") == "critical" else "Medium",
                    "property": company,
                    "postcode": pc,
                    "type":     "Building Assessment",
                    "gap":      g.get("title", ""),
                    "service":  g.get("service", "").split("\n")[0],
                    "reason":   g.get("desc", "")[:120],
                })

    opps.sort(key=lambda x: (0 if x["priority"]=="High" else 1, x["property"]))
    return opps


def generate_master_pdf(parks_uploads, intel_uploads, report_title, prepared_by):
    buf = io.BytesIO()
    S   = _styles()

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=M, rightMargin=M,
        topMargin=M, bottomMargin=20*mm,
        title=report_title,
    )
    story = []

    total_parks    = sum(len(u.get("parks",[])) for u in parks_uploads)
    total_briefings= sum(len(u.get("briefings",[])) for u in intel_uploads)
    opportunities  = _build_opportunities(parks_uploads, intel_uploads)
    high_opps      = sum(1 for o in opportunities if o["priority"]=="High")

    # ── COVER ──────────────────────────────────────────────────────────────
    t = Table([[Paragraph(
        "INTERNAL  ·  MODERN NETWORKS SALES & MARKETING INTELLIGENCE  ·  "
        "NOT FOR EXTERNAL DISTRIBUTION", S["whiteb"]
    )]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), NAVY),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
    ]))
    story.append(t)
    story.append(Spacer(1, 10*mm))

    story.append(Paragraph("Modern Networks  |  Territory Intelligence Report", S["teal"]))
    story.append(Paragraph(
        f"Prepared by: {prepared_by}  ·  {datetime.now().strftime('%d %b %Y')}",
        S["mono"]
    ))
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(report_title or "Territory Intelligence Report", S["h1"]))
    story.append(Spacer(1, 4*mm))

    stats = [
        [
            Paragraph(str(total_parks),
                ParagraphStyle("sn",  fontName="Helvetica-Bold", fontSize=22, textColor=TEAL,  leading=26)),
            Paragraph(str(total_briefings),
                ParagraphStyle("sn2", fontName="Helvetica-Bold", fontSize=22, textColor=GREEN, leading=26)),
            Paragraph(str(len(opportunities)),
                ParagraphStyle("sn3", fontName="Helvetica-Bold", fontSize=22, textColor=AMBER, leading=26)),
            Paragraph(str(high_opps),
                ParagraphStyle("sn4", fontName="Helvetica-Bold", fontSize=22, textColor=RED,   leading=26)),
        ],
        [
            Paragraph("Science Parks",       S["small"]),
            Paragraph("Buildings Assessed",  S["small"]),
            Paragraph("Opportunities",       S["small"]),
            Paragraph("High Priority",       S["small"]),
        ],
    ]
    st_t = Table(stats, colWidths=[CW/4]*4)
    st_t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), LGREY),
        ("BOX",           (0,0),(-1,-1), 0.5, MGREY),
        ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ("TOPPADDING",    (0,0),(-1,-1), 10),
        ("BOTTOMPADDING", (0,0),(-1,-1), 10),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("LINEBEFORE",    (1,0),(3,-1),  0.5, MGREY),
    ]))
    story.append(st_t)
    story.append(Spacer(1, 6*mm))

    # ── EXECUTIVE SUMMARY ──────────────────────────────────────────────────
    story.append(_bar("EXECUTIVE SUMMARY", S))
    story.append(Spacer(1, 4*mm))

    park_areas = list(set(
        u.get("area_label","") for u in parks_uploads if u.get("area_label")
    ))
    area_str = ", ".join(park_areas) if park_areas else "the assessed territory"

    exec_text = (
        f"This report covers digital infrastructure intelligence for {area_str}. "
        f"It includes data for {total_parks} science and innovation park{'s' if total_parks!=1 else ''}"
        f"{' and '+str(total_briefings)+' individual building assessment'+('s' if total_briefings!=1 else '') if total_briefings else ''}. "
        f"The analysis has identified {len(opportunities)} service opportunities for Modern Networks, "
        f"of which {high_opps} are high priority requiring prompt action. "
        "Opportunities span connectivity upgrades, WiredScore and SmartScore certification, "
        "managed IT services, and cybersecurity — all core Modern Networks service lines."
    )
    story.append(Paragraph(exec_text, S["body"]))
    story.append(PageBreak())

    # ── SCIENCE PARKS ──────────────────────────────────────────────────────
    if parks_uploads:
        for upload in parks_uploads:
            area_label = upload.get("area_label","") or "Science Parks"
            parks      = upload.get("parks", [])
            exported   = upload.get("exported_at","")

            story.append(_bar(f"SCIENCE PARKS — {area_label.upper()}", S))
            story.append(Spacer(1, 3*mm))
            story.append(Paragraph(
                f"{len(parks)} parks  ·  Exported: {exported}",
                S["mono"]
            ))
            story.append(Spacer(1, 4*mm))

            # Summary table
            hdr = [
                Paragraph("PARK",        S["monow"]),
                Paragraph("POSTCODE",    S["monow"]),
                Paragraph("SECTOR",      S["monow"]),
                Paragraph("TENANTS",     S["monow"]),
                Paragraph("OPERATOR",    S["monow"]),
            ]
            rows = [hdr]
            for p in parks:
                rows.append([
                    Paragraph(p.get("name","")[:35],    S["bold9"]),
                    Paragraph(p.get("postcode",""),      S["body"]),
                    Paragraph(p.get("sector","")[:28],   S["small"]),
                    Paragraph(str(p.get("tenants","")),  S["body"]),
                    Paragraph(p.get("operator","")[:28], S["small"]),
                ])
            cws = [55*mm, 22*mm, 43*mm, 18*mm, CW-138*mm]
            pt = Table(rows, colWidths=cws)
            pt.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,0),  NAVY),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),  [WHITE, LGREY]),
                ("LINEBELOW",     (0,0),(-1,-1),  0.3, MGREY),
                ("TOPPADDING",    (0,0),(-1,-1),  6),
                ("BOTTOMPADDING", (0,0),(-1,-1),  6),
                ("LEFTPADDING",   (0,0),(-1,-1),  5),
                ("VALIGN",        (0,0),(-1,-1),  "TOP"),
            ]))
            story.append(pt)
            story.append(Spacer(1, 6*mm))

            # Park notes
            notes_added = False
            for p in parks:
                if p.get("notes"):
                    if not notes_added:
                        story.append(Paragraph("Park Notes", S["bold9"]))
                        story.append(Spacer(1, 2*mm))
                        notes_added = True
                    story.append(Paragraph(
                        f"{p.get('name','')} ({p.get('postcode','')}) — {p.get('notes','')[:200]}",
                        S["small"]
                    ))
                    story.append(Spacer(1, 2*mm))

            story.append(PageBreak())

    # ── INDIVIDUAL BUILDING ASSESSMENTS ────────────────────────────────────
    if intel_uploads:
        story.append(_bar("INDIVIDUAL BUILDING ASSESSMENTS", S))
        story.append(Spacer(1, 3*mm))
        story.append(Paragraph(
            "Individual building assessments from the Modern Networks Building Intelligence Platform. "
            "Each building has been assessed against Ofcom connectivity data, EPC Register, "
            "Companies House, Environment Agency flood risk, and crime data.",
            S["italic"]
        ))
        story.append(Spacer(1, 4*mm))

        for upload in intel_uploads:
            briefings = upload.get("briefings", [])
            # Sort by score ascending (worst first)
            briefings_sorted = sorted(briefings, key=lambda b: b.get("score", 100))

            for b in briefings_sorted:
                pc      = b.get("postcode","")
                company = b.get("company","")
                score   = b.get("score", 0)
                verdict = b.get("verdict","")
                label   = b.get("scoreLabel","")
                gaps    = b.get("gaps",[])
                pos     = b.get("positives",[])
                ws      = b.get("wiredScore",{})
                sc_col  = _score_colour(score)

                title_str = f"{company} — {pc}" if company else f"Building Assessment — {pc}"

                # Header row
                hdr_t = Table(
                    [[
                        Paragraph(title_str, S["bold11"]),
                        Paragraph(
                            f'<font color="{sc_col.hexval()}"><b>{score}/100</b></font>',
                            ParagraphStyle("bsc",fontName="Helvetica-Bold",fontSize=14,
                                           textColor=BLACK,leading=18)
                        ),
                    ]],
                    colWidths=[CW-30*mm, 30*mm]
                )
                hdr_t.setStyle(TableStyle([
                    ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
                    ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
                    ("TOPPADDING",    (0,0),(-1,-1), 4),
                    ("BOTTOMPADDING", (0,0),(-1,-1), 4),
                    ("LINEBELOW",     (0,0),(-1,-1), 1, NAVY),
                ]))
                story.append(hdr_t)
                story.append(Spacer(1, 2*mm))

                if verdict or label:
                    story.append(Paragraph(
                        f'{verdict or label}  ·  Saved {b.get("savedAt","")}',
                        ParagraphStyle("vd",fontName="Helvetica-Bold",fontSize=8,
                                       textColor=sc_col,leading=11)
                    ))
                    story.append(Spacer(1, 2*mm))

                # WiredScore
                ws_status = ws.get("status","unconfirmed")
                ws_icon   = "✓" if ws_status=="certified" else "✕" if ws_status=="not-certified" else "?"
                story.append(Paragraph(
                    f"WiredScore: {ws_icon} {ws_status.title()}"
                    f"{' — '+ws.get('scheme','')+' '+ws.get('level','') if ws_status=='certified' else ''}",
                    S["small"]
                ))
                story.append(Spacer(1, 3*mm))

                # Gaps
                if gaps:
                    story.append(Paragraph("Gaps & Opportunities", S["bold9"]))
                    story.append(Spacer(1, 2*mm))
                    for g in gaps:
                        lc     = RED if g.get("sev")=="critical" else AMBER if g.get("sev")=="advisory" else TEAL
                        bg_col = LRED if g.get("sev")=="critical" else LCREAM
                        sev    = g.get("sev","").upper()
                        g_t = Table(
                            [[
                                Paragraph(sev, ParagraphStyle(
                                    "gs",fontName="Courier-Bold",fontSize=7,
                                    textColor=lc,leading=10)),
                                [
                                    Paragraph(f"{g.get('icon','')} {g.get('title','')}", S["bold9"]),
                                    Paragraph(g.get("desc","")[:200], S["small"]),
                                ],
                                Paragraph(
                                    g.get("service","").replace("\n","  ·  ")[:60],
                                    S["teal"]
                                ),
                            ]],
                            colWidths=[14*mm, CW-80*mm, 66*mm]
                        )
                        g_t.setStyle(TableStyle([
                            ("BACKGROUND",    (0,0),(-1,-1), bg_col),
                            ("LINEBEFORE",    (0,0),(0,-1),  3, lc),
                            ("BOX",           (0,0),(-1,-1), 0.5, MGREY),
                            ("VALIGN",        (0,0),(-1,-1), "TOP"),
                            ("TOPPADDING",    (0,0),(-1,-1), 6),
                            ("BOTTOMPADDING", (0,0),(-1,-1), 6),
                            ("LEFTPADDING",   (0,0),(-1,-1), 6),
                            ("RIGHTPADDING",  (0,0),(-1,-1), 6),
                        ]))
                        story.append(KeepTogether([g_t, Spacer(1, 3*mm)]))

                # Strengths
                if pos:
                    story.append(Spacer(1, 2*mm))
                    story.append(Paragraph("Confirmed Strengths", S["bold9"]))
                    story.append(Spacer(1, 2*mm))
                    for p_item in pos[:3]:
                        story.append(Paragraph(
                            f"✓ {p_item.get('icon','')} {p_item.get('title','')} — "
                            f"{p_item.get('desc','')[:120]}",
                            ParagraphStyle("str",fontName="Helvetica",fontSize=8,
                                           textColor=GREEN,leading=11)
                        ))
                        story.append(Spacer(1, 2*mm))

                story.append(Spacer(1, 4*mm))
                story.append(_hr())

        story.append(PageBreak())

    # ── PRIORITY ACTION LIST ────────────────────────────────────────────────
    story.append(_bar("PRIORITY ACTION LIST", S))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "All identified opportunities ranked by priority. "
        "High priority items should be actioned within 30 days.",
        S["italic"]
    ))
    story.append(Spacer(1, 5*mm))

    if opportunities:
        for o in opportunities[:30]:
            pc_col = RED if o["priority"]=="High" else AMBER
            bg_col = LRED if o["priority"]=="High" else LCREAM
            type_badge = o.get("type","")

            opp_t = Table(
                [[
                    Paragraph(o["priority"], ParagraphStyle(
                        "pr",fontName="Helvetica-Bold",fontSize=8,
                        textColor=pc_col,leading=11)),
                    [
                        Paragraph(
                            f'{o["property"]}'
                            f'{"  ·  "+o["postcode"] if o["postcode"] else ""}'
                            f'{"  ·  "+type_badge if type_badge else ""}',
                            S["bold9"]
                        ),
                        Paragraph(o["gap"], S["body"]),
                        Paragraph(f"Why now: {o['reason']}", S["italic"]),
                    ],
                    Paragraph(o["service"], S["teal"]),
                ]],
                colWidths=[14*mm, CW-82*mm, 68*mm]
            )
            opp_t.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,-1), bg_col),
                ("LINEBEFORE",    (0,0),(0,-1),  3, pc_col),
                ("BOX",           (0,0),(-1,-1), 0.5, MGREY),
                ("VALIGN",        (0,0),(-1,-1), "TOP"),
                ("TOPPADDING",    (0,0),(-1,-1), 8),
                ("BOTTOMPADDING", (0,0),(-1,-1), 8),
                ("LEFTPADDING",   (0,0),(-1,-1), 8),
                ("RIGHTPADDING",  (0,0),(-1,-1), 8),
            ]))
            story.append(KeepTogether([opp_t, Spacer(1, 4*mm)]))
    else:
        story.append(Paragraph("No opportunities identified from uploaded data.", S["italic"]))

    story.append(PageBreak())

    # ── APPENDIX ───────────────────────────────────────────────────────────
    story.append(_bar("APPENDIX — DATA SOURCES & METHODOLOGY", S))
    story.append(Spacer(1, 4*mm))

    for name, desc in [
        ("Ofcom Connected Nations",
         "Postcode-level fixed broadband coverage data including gigabit, ultrafast, and superfast availability. Updated quarterly by Ofcom."),
        ("EPC Register",
         "Non-domestic Energy Performance Certificates by postcode. Source: MHCLG Get Energy Performance Data API."),
        ("Companies House",
         "Active company registrations by postcode. Source: Companies House public API."),
        ("Environment Agency",
         "Flood zone classification by postcode. Source: EA Postcodes Risk Assessment dataset (data.gov.uk)."),
        ("Police API",
         "Street-level crime data aggregated by location. Source: data.police.uk."),
        ("OS Names API",
         "Postcode to coordinate resolution for mapping and area lookups. Source: OS Data Hub."),
        ("WiredScore / SmartScore",
         "Building certification status. Manually verified by MN staff via wiredscore.com/map."),
        ("Science Parks Data",
         "UK science and innovation park profiles including operator, tenant count, sector, and postcode. Source: MN Science Parks Intelligence Platform."),
    ]:
        story.append(Paragraph(f"<b>{name}</b> — {desc}", S["body"]))
        story.append(Spacer(1, 3*mm))

    fn = _footer(report_title or "Master Report")
    doc.build(story, onFirstPage=fn, onLaterPages=fn)
    return buf.getvalue()


# ── MAIN UI ────────────────────────────────────────────────────────────────────

st.markdown("""
<div style="background:#0b1829;padding:20px 28px;border-radius:10px;margin-bottom:24px">
<div style="font-size:10px;letter-spacing:2px;color:#f59e0b;font-family:monospace;
            margin-bottom:6px">INTERNAL USE ONLY</div>
<div style="font-size:24px;font-weight:800;color:#fff;margin-bottom:4px">Modern Networks</div>
<div style="font-size:14px;color:#64748b">Territory Intelligence Report Generator</div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    "Combine Science Parks data and Building Intelligence assessments "
    "into a single territory intelligence report."
)

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### Report Settings")
    report_title = st.text_input("Report title",
                                  placeholder="e.g. Oxford-Cambridge Arc Q2 2026")
    prepared_by  = st.text_input("Prepared by", placeholder="Your name")

    st.divider()

    # Science Parks uploads
    st.markdown("### 🔬 Science Parks Data")
    st.caption("Upload JSON export from the Science Parks Intelligence app.")
    parks_files = st.file_uploader(
        "Science Parks JSON",
        type=["json"],
        accept_multiple_files=True,
        key="parks_uploader",
        label_visibility="collapsed"
    )
    if parks_files:
        for f in parks_files:
            data = parse_upload(f)
            if data and data.get("source_app") == "science_parks":
                already = any(
                    u.get("exported_at") == data.get("exported_at") and
                    u.get("area_label")  == data.get("area_label")
                    for u in st.session_state.parks_uploads
                )
                if not already:
                    st.session_state.parks_uploads.append(data)
                    st.success(f"Loaded: {f.name}")
            elif data:
                st.warning(f"{f.name} is not a Science Parks export.")

    if st.session_state.parks_uploads:
        for i, u in enumerate(st.session_state.parks_uploads):
            c1, c2 = st.columns([4,1])
            with c1:
                st.markdown(
                    f"<span class='source-badge badge-parks'>Science Parks</span> "
                    f"{u.get('area_label','')} — {len(u.get('parks',[]))} parks",
                    unsafe_allow_html=True
                )
            with c2:
                if st.button("✕", key=f"dp_{i}"):
                    st.session_state.parks_uploads.pop(i)
                    st.rerun()

    st.divider()

    # Building Intelligence uploads
    st.markdown("### 🏢 Building Intelligence Data")
    st.caption("Upload JSON export from the Building Intelligence Platform (Saved Briefings tab).")
    intel_files = st.file_uploader(
        "Building Intelligence JSON",
        type=["json"],
        accept_multiple_files=True,
        key="intel_uploader",
        label_visibility="collapsed"
    )
    if intel_files:
        for f in intel_files:
            data = parse_upload(f)
            if data and data.get("source_app") == "building_intelligence":
                already = any(
                    u.get("exported_at") == data.get("exported_at")
                    for u in st.session_state.intel_uploads
                )
                if not already:
                    st.session_state.intel_uploads.append(data)
                    st.success(f"Loaded: {f.name}")
            elif data:
                st.warning(f"{f.name} is not a Building Intelligence export.")

    if st.session_state.intel_uploads:
        for i, u in enumerate(st.session_state.intel_uploads):
            c1, c2 = st.columns([4,1])
            with c1:
                n = len(u.get("briefings",[]))
                st.markdown(
                    f"<span class='source-badge badge-intel'>Building Intelligence</span> "
                    f"{n} briefing{'s' if n!=1 else ''}",
                    unsafe_allow_html=True
                )
            with c2:
                if st.button("✕", key=f"di_{i}"):
                    st.session_state.intel_uploads.pop(i)
                    st.rerun()


with col2:
    st.markdown("### Report Preview")

    has_data = st.session_state.parks_uploads or st.session_state.intel_uploads

    if not has_data:
        st.markdown("""
        <div style="background:#fff;border:2px dashed #e2e8f0;border-radius:10px;
                    padding:48px;text-align:center;color:#94a3b8">
            <div style="font-size:44px;margin-bottom:14px">📊</div>
            <div style="font-size:15px;font-weight:600;color:#64748b;margin-bottom:8px">
                No data loaded yet</div>
            <div style="font-size:13px;line-height:1.9">
                Upload a JSON export from the <strong>Science Parks app</strong><br>
                and/or the <strong>Building Intelligence Platform</strong><br>
                using the panels on the left.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("#### This report will contain:")

        sections = ["Cover page with territory summary statistics", "Executive summary"]

        for u in st.session_state.parks_uploads:
            label = u.get("area_label","")
            parks = u.get("parks",[])
            sections.append(f"Science Parks — {label} ({len(parks)} parks)")
            for p in parks[:5]:
                st.caption(f"  · {p.get('name','')} ({p.get('postcode','')})")
            if len(parks) > 5:
                st.caption(f"  · +{len(parks)-5} more parks")

        for u in st.session_state.intel_uploads:
            n = len(u.get("briefings",[]))
            sections.append(f"Individual building assessments ({n} properties)")
            for b in u.get("briefings",[])[:5]:
                company = b.get("company","") or b.get("postcode","")
                score   = b.get("score",0)
                sc_text = f"Score {score}/100"
                st.caption(f"  · {company} — {sc_text}")
            if n > 5:
                st.caption(f"  · +{n-5} more buildings")

        opps = _build_opportunities(
            st.session_state.parks_uploads,
            st.session_state.intel_uploads
        )
        high = sum(1 for o in opps if o["priority"]=="High")
        sections.append(f"Priority action list ({len(opps)} opportunities, {high} high priority)")
        sections.append("Data sources appendix")

        for i, s in enumerate(sections, 1):
            st.markdown(f"{i}. {s}")

        st.divider()

        if st.button("⬇ Generate Territory Intelligence Report",
                     type="primary", use_container_width=True):
            if not report_title:
                st.warning("Please enter a report title first.")
            else:
                with st.spinner("Building report…"):
                    pdf_bytes = generate_master_pdf(
                        st.session_state.parks_uploads,
                        st.session_state.intel_uploads,
                        report_title,
                        prepared_by or "MN Staff"
                    )
                safe = report_title.replace(" ","-").replace("/","-")
                st.download_button(
                    "⬇ Download Territory Intelligence Report",
                    data=pdf_bytes,
                    file_name=f"MN-Territory-{safe}-{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
                st.success("Report generated.")