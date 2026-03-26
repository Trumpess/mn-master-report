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

st.set_page_config(page_title="MN Master Report", page_icon="📊", layout="wide")

st.markdown("""
<style>
body,.stApp{background:#f4f6f9;color:#0f172a}
.source-badge{display:inline-block;padding:3px 10px;border-radius:20px;
              font-size:11px;font-weight:600;margin-right:6px}
.badge-parks{background:#e0f2fe;color:#0369a1}
.badge-vuln{background:#fef3c7;color:#92400e}
.badge-intel{background:#f0fdf4;color:#166534}
</style>
""", unsafe_allow_html=True)

if "uploads" not in st.session_state:
    st.session_state.uploads = []
if "postcodes" not in st.session_state:
    st.session_state.postcodes = []


def parse_upload(uploaded_file):
    try:
        data = json.loads(uploaded_file.read())
        return data
    except Exception as e:
        st.error(f"Could not parse {uploaded_file.name}: {e}")
        return None


def source_label(source_app):
    labels = {
        "science_parks":        ("Science Parks",         "badge-parks"),
        "vulnerability_scanner":("Vulnerability Scanner", "badge-vuln"),
        "building_intelligence":("Building Intelligence", "badge-intel"),
    }
    return labels.get(source_app, (source_app, "badge-intel"))


def _styles():
    b = getSampleStyleSheet()
    def S(name, **kw):
        return ParagraphStyle(name, parent=b["Normal"], **kw)
    return {
        "h1":    S("h1",  fontName="Helvetica-Bold",  fontSize=20, textColor=NAVY,  leading=26),
        "h2":    S("h2",  fontName="Helvetica-Bold",  fontSize=14, textColor=NAVY,  leading=18),
        "body":  S("body",fontName="Helvetica",        fontSize=9,  textColor=BLACK, leading=13),
        "small": S("sml", fontName="Helvetica",        fontSize=8,  textColor=GREY,  leading=11),
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


def _section_bar(title, S):
    t = Table([[Paragraph(title, S["whiteb"])]], colWidths=[CW])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), NAVY),
        ("TOPPADDING",    (0,0),(-1,-1), 7),
        ("BOTTOMPADDING", (0,0),(-1,-1), 7),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
    ]))
    return t


def _divider():
    return HRFlowable(width=CW, thickness=0.5, color=MGREY, spaceAfter=4, spaceBefore=4)


def _footer_fn(title):
    def _fn(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(GREY)
        canvas.drawString(M, 10*mm,
            f"Modern Networks — {title} — CONFIDENTIAL — INTERNAL USE ONLY")
        canvas.drawRightString(W-M, 10*mm, f"Page {doc.page}")
        canvas.restoreState()
    return _fn


def _build_opportunities(uploads, postcodes_data):
    """Build a meaningful opportunity list from uploaded data."""
    opps = []

    for upload in uploads:
        source = upload.get("source_app","")

        if source == "science_parks":
            for p in upload.get("parks", []):
                name     = p.get("name","")
                postcode = p.get("postcode","")
                sector   = p.get("sector","")
                tenants  = p.get("tenants","")
                ofcom    = p.get("ofcom",{}) or {}
                notes    = p.get("notes","")

                # Connectivity gap
                gigabit_pct = ofcom.get("gigabit_pct", 0) or ofcom.get("connectivity",{}).get("gigabit_pct",0)
                try:
                    gig = float(gigabit_pct)
                except (TypeError, ValueError):
                    gig = 0

                if gig < 50:
                    opps.append({
                        "park":    name,
                        "postcode":postcode,
                        "gap":     f"Gigabit coverage {gig:.0f}% — below standard for innovation park",
                        "service": "Fibre Broadband · Network-as-a-Service",
                        "reason":  f"{tenants} tenants in {sector} sector requiring high-speed reliable connectivity",
                        "priority":"High" if gig < 20 else "Medium",
                    })
                else:
                    opps.append({
                        "park":    name,
                        "postcode":postcode,
                        "gap":     "WiredScore / SmartScore certification not confirmed",
                        "service": "WiredScore AP Services · SmartScore AP Services",
                        "reason":  f"Certification would differentiate {name} and support premium tenant attraction",
                        "priority":"Medium",
                    })

                # EPC
                epc_rating = ofcom.get("epc_rating","")
                if epc_rating and epc_rating not in ("A","B","C",""):
                    opps.append({
                        "park":    name,
                        "postcode":postcode,
                        "gap":     f"EPC {epc_rating} — below proposed 2027 minimum (C)",
                        "service": "Smart Building Technology · Azure Managed Services",
                        "reason":  "2027 MEES regulations risk lettability — digital infrastructure upgrades contribute to EPC improvement",
                        "priority":"High" if epc_rating in ("E","F","G") else "Medium",
                    })

                # Multi-tenant managed IT
                try:
                    t_num = int(str(tenants).replace("+","").replace(",","").split()[0])
                except (ValueError, IndexError):
                    t_num = 0
                if t_num >= 50:
                    opps.append({
                        "park":    name,
                        "postcode":postcode,
                        "gap":     f"{tenants} tenant organisations — managed IT opportunity",
                        "service": "Service Guardian · Desktop Support · M365 Managed Services",
                        "reason":  "Large multi-tenant environment with demand for managed connectivity, IT support and security",
                        "priority":"Medium",
                    })

        elif source == "vulnerability_scanner":
            for area in upload.get("areas", [upload]):
                a_label = area.get("area_label","") or area.get("search_term","")
                crit    = area.get("critical_count",0)
                devices = area.get("device_count",0)
                score   = str(area.get("score","")).lower()
                if crit > 0 or score in ("red","critical"):
                    opps.append({
                        "park":    a_label,
                        "postcode":"",
                        "gap":     f"{crit} critical vulnerabilities · {devices} exposed devices",
                        "service": "Cybersecurity Services · Managed Firewall · ISO 27001",
                        "reason":  "Exposed devices and critical vulnerabilities represent immediate security and insurance risk",
                        "priority":"High",
                    })

    for bld in postcodes_data:
        pc      = bld.get("postcode","")
        company = bld.get("company","") or pc
        score   = bld.get("score",0)
        gaps    = bld.get("gaps",[])
        for g in gaps[:2]:
            opps.append({
                "park":    company,
                "postcode":pc,
                "gap":     g.get("title",""),
                "service": g.get("service","").split("\n")[0],
                "reason":  g.get("desc","")[:100],
                "priority":"High" if g.get("sev")=="critical" else "Medium",
            })

    opps.sort(key=lambda x: 0 if x["priority"]=="High" else 1)
    return opps


def generate_master_pdf(uploads, postcodes_data, report_title, prepared_by):
    buf = io.BytesIO()
    S   = _styles()

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=M, rightMargin=M,
        topMargin=M, bottomMargin=20*mm,
        title=report_title,
    )
    story = []

    total_parks     = sum(len(u.get("parks",[])) for u in uploads
                         if u.get("source_app")=="science_parks")
    total_areas     = sum(len(u.get("areas",[u])) for u in uploads
                         if u.get("source_app")=="vulnerability_scanner")
    total_postcodes = len(postcodes_data)
    opportunities   = _build_opportunities(uploads, postcodes_data)

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

    story.append(Paragraph("Modern Networks  |  Master Intelligence Report", S["teal"]))
    story.append(Paragraph(
        f"Prepared by: {prepared_by}  ·  {datetime.now().strftime('%d %b %Y')}",
        S["mono"]
    ))
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(report_title or "Master Intelligence Report", S["h1"]))
    story.append(Spacer(1, 4*mm))

    stats_data = [
        [
            Paragraph(str(len(uploads)),
                ParagraphStyle("sn",  fontName="Helvetica-Bold", fontSize=22,
                               textColor=NAVY,  leading=26)),
            Paragraph(str(total_areas + total_parks),
                ParagraphStyle("sn2", fontName="Helvetica-Bold", fontSize=22,
                               textColor=TEAL,  leading=26)),
            Paragraph(str(len(opportunities)),
                ParagraphStyle("sn3", fontName="Helvetica-Bold", fontSize=22,
                               textColor=AMBER, leading=26)),
            Paragraph(str(sum(1 for o in opportunities if o["priority"]=="High")),
                ParagraphStyle("sn4", fontName="Helvetica-Bold", fontSize=22,
                               textColor=RED,   leading=26)),
        ],
        [
            Paragraph("Data Sources",      S["small"]),
            Paragraph("Areas / Parks",     S["small"]),
            Paragraph("Opportunities",     S["small"]),
            Paragraph("High Priority",     S["small"]),
        ],
    ]
    st_t = Table(stats_data, colWidths=[CW/4]*4)
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

    story.append(_section_bar("EXECUTIVE SUMMARY", S))
    story.append(Spacer(1, 4*mm))

    sources     = list(set(u.get("source_app","") for u in uploads))
    source_desc = " and ".join([{
        "science_parks":        "science and innovation park data",
        "vulnerability_scanner":"commercial area vulnerability intelligence",
        "building_intelligence":"individual building assessments",
    }.get(s, s) for s in sources])

    high_count = sum(1 for o in opportunities if o["priority"]=="High")
    exec_text = (
        f"This report consolidates {source_desc} across {len(uploads)} data "
        f"source{'s' if len(uploads)!=1 else ''}. "
        f"It covers"
        f"{' '+str(total_areas)+' assessed areas' if total_areas else ''}"
        f"{' and' if total_areas and total_parks else ''}"
        f"{' '+str(total_parks)+' science and innovation parks' if total_parks else ''}"
        f"{', and '+str(total_postcodes)+' individual building assessments' if total_postcodes else ''}. "
        f"The analysis has identified {len(opportunities)} service opportunities "
        f"for Modern Networks, of which {high_count} are high priority. "
        "The report is intended for internal use by Modern Networks sales and marketing teams "
        "to support territory planning, prospect identification, and meeting preparation."
    )
    story.append(Paragraph(exec_text, S["body"]))
    story.append(PageBreak())

    # ── TERRITORY DATA ─────────────────────────────────────────────────────
    for upload in uploads:
        source   = upload.get("source_app","unknown")
        exported = upload.get("exported_at","")
        label    = upload.get("area_label","") or upload.get("search_term","")

        source_names = {
            "science_parks":        "Science Parks Intelligence",
            "vulnerability_scanner":"Commercial Area Vulnerability Intelligence",
            "building_intelligence":"Building Intelligence",
        }
        story.append(_section_bar(
            f"{source_names.get(source,'Intelligence')} — {label}", S
        ))
        story.append(Spacer(1, 3*mm))
        story.append(Paragraph(
            f"Exported: {exported}",
            S["mono"]
        ))
        story.append(Spacer(1, 4*mm))

        if source == "vulnerability_scanner":
            areas = upload.get("areas", [upload])
            for area in areas:
                a_label = area.get("area_label","") or area.get("search_term","")
                score   = area.get("score","")
                rating  = area.get("rating","")
                devices = area.get("device_count",0)
                crit    = area.get("critical_count",0)
                vulns   = area.get("vuln_count",0)
                sc_col  = (RED   if str(score).lower() in ("red","critical") else
                           AMBER if str(score).lower() in ("amber","warning") else
                           GREEN)

                rows = [
                    [Paragraph("AREA",     S["mono"]), Paragraph(str(a_label), S["bold9"])],
                    [Paragraph("SCORE",    S["mono"]), Paragraph(str(score),
                        ParagraphStyle("asc",fontName="Helvetica-Bold",fontSize=11,
                                       textColor=sc_col,leading=14))],
                    [Paragraph("RATING",   S["mono"]), Paragraph(str(rating),  S["bold9"])],
                    [Paragraph("DEVICES",  S["mono"]), Paragraph(str(devices), S["body"])],
                    [Paragraph("CRITICAL", S["mono"]),
                     Paragraph(str(crit), S["red"] if crit > 0 else S["body"])],
                    [Paragraph("VULNS",    S["mono"]),
                     Paragraph(str(vulns), S["amber"] if vulns > 0 else S["body"])],
                ]
                at = Table(rows, colWidths=[30*mm, CW-30*mm])
                at.setStyle(TableStyle([
                    ("BACKGROUND",    (0,0),(-1,-1), LGREY),
                    ("LINEBELOW",     (0,0),(-1,-1), 0.3, MGREY),
                    ("TOPPADDING",    (0,0),(-1,-1), 5),
                    ("BOTTOMPADDING", (0,0),(-1,-1), 5),
                    ("LEFTPADDING",   (0,0),(-1,-1), 8),
                ]))
                story.append(at)
                story.append(Spacer(1, 3*mm))

                ctx = area.get("area_context",{})
                if ctx:
                    ctx_items = []
                    for k, v in list(ctx.items())[:8]:
                        if v:
                            ctx_items.append([
                                Paragraph(str(k).replace("_"," ").title(), S["mono"]),
                                Paragraph(str(v)[:80], S["body"]),
                            ])
                    if ctx_items:
                        ct = Table(ctx_items, colWidths=[40*mm, CW-40*mm])
                        ct.setStyle(TableStyle([
                            ("LINEBELOW",     (0,0),(-1,-1), 0.2, MGREY),
                            ("TOPPADDING",    (0,0),(-1,-1), 4),
                            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
                            ("LEFTPADDING",   (0,0),(-1,-1), 8),
                        ]))
                        story.append(ct)
                story.append(_divider())

        elif source == "science_parks":
            parks = upload.get("parks",[])
            story.append(Paragraph(f"{len(parks)} parks in this dataset", S["small"]))
            story.append(Spacer(1, 3*mm))

            # Table with explicit white header text
            hdr_row = [
                Paragraph("PARK",     S["monow"]),
                Paragraph("POSTCODE", S["monow"]),
                Paragraph("SECTOR",   S["monow"]),
                Paragraph("TENANTS",  S["monow"]),
                Paragraph("OPERATOR", S["monow"]),
            ]
            rows = [hdr_row]
            for p in parks:
                rows.append([
                    Paragraph(p.get("name","")[:35],    S["bold9"]),
                    Paragraph(p.get("postcode",""),      S["body"]),
                    Paragraph(p.get("sector","")[:30],   S["small"]),
                    Paragraph(str(p.get("tenants","")),  S["body"]),
                    Paragraph(p.get("operator","")[:30], S["small"]),
                ])
            cws = [55*mm, 22*mm, 45*mm, 18*mm, CW-140*mm]
            pt = Table(rows, colWidths=cws)
            pt.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,0),  NAVY),
                ("BACKGROUND",    (0,1),(-1,-1),  WHITE),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),  [WHITE, LGREY]),
                ("LINEBELOW",     (0,0),(-1,-1),  0.3, MGREY),
                ("TOPPADDING",    (0,0),(-1,-1),  6),
                ("BOTTOMPADDING", (0,0),(-1,-1),  6),
                ("LEFTPADDING",   (0,0),(-1,-1),  5),
                ("VALIGN",        (0,0),(-1,-1),  "TOP"),
            ]))
            story.append(pt)
            story.append(Spacer(1, 4*mm))

            for p in parks:
                if p.get("notes"):
                    story.append(Paragraph(
                        f"{p.get('name','')} — {p.get('notes','')[:200]}",
                        S["small"]
                    ))
                    story.append(Spacer(1, 2*mm))

        story.append(PageBreak())

    # ── INDIVIDUAL BUILDING ASSESSMENTS ────────────────────────────────────
    if postcodes_data:
        story.append(_section_bar("INDIVIDUAL BUILDING ASSESSMENTS", S))
        story.append(Spacer(1, 3*mm))
        story.append(Paragraph(
            f"{len(postcodes_data)} individual building postcodes included. "
            "For full building intelligence run assessments in the Building Intelligence Platform.",
            S["italic"]
        ))
        story.append(Spacer(1, 4*mm))

        for bld in postcodes_data:
            pc      = bld.get("postcode","")
            company = bld.get("company","")
            score   = bld.get("score",0)
            verdict = bld.get("verdict","Pending assessment")
            gaps    = bld.get("gaps",[])
            sc_col  = RED if score < 50 else AMBER if score < 70 else GREEN

            title_str = f"{company} — {pc}" if company else f"Building — {pc}"
            story.append(Paragraph(title_str, S["bold11"]))
            if score:
                story.append(Paragraph(
                    f"Score: {score}/100  ·  {verdict}",
                    ParagraphStyle("vs",fontName="Helvetica-Bold",fontSize=9,
                                   textColor=sc_col,leading=12)
                ))
            for g in gaps[:3]:
                lc = RED if g.get("sev")=="critical" else AMBER
                story.append(Paragraph(
                    f"● {g.get('icon','')} {g.get('title','')}",
                    ParagraphStyle("gi",fontName="Helvetica",fontSize=8,
                                   textColor=lc,leading=11)
                ))
            story.append(Spacer(1, 4*mm))
            story.append(_divider())

    # ── PRIORITISED OPPORTUNITY LIST ────────────────────────────────────────
    story.append(PageBreak())
    story.append(_section_bar("PRIORITISED OPPORTUNITY LIST", S))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "Service opportunities identified across all data sources. "
        "High priority items should be actioned first.",
        S["italic"]
    ))
    story.append(Spacer(1, 5*mm))

    if opportunities:
        for o in opportunities[:25]:
            pc_col = RED if o["priority"]=="High" else AMBER
            bg_col = colors.HexColor("#fff5f5") if o["priority"]=="High" else colors.HexColor("#fffbeb")

            opp_t = Table([
                [
                    Paragraph(o["priority"], ParagraphStyle(
                        "pr",fontName="Helvetica-Bold",fontSize=8,
                        textColor=pc_col,leading=11)),
                    [
                        Paragraph(o["park"], S["bold9"]),
                        Paragraph(o["gap"],  S["body"]),
                        Paragraph(f"Why now: {o['reason']}", S["italic"]),
                    ],
                    Paragraph(o["service"], S["teal"]),
                ]
            ], colWidths=[14*mm, CW-80*mm, 66*mm])
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

    story.append(PageBreak())

    # ── APPENDIX ───────────────────────────────────────────────────────────
    story.append(_section_bar("APPENDIX — DATA SOURCES", S))
    story.append(Spacer(1, 4*mm))

    sources_list = [
        ("Ofcom Connected Nations",
         "Postcode-level fixed broadband and mobile coverage data. Updated quarterly."),
        ("EPC Register",
         "Non-domestic Energy Performance Certificates. Source: MHCLG."),
        ("Companies House",
         "Active company registrations by postcode. Source: Companies House API."),
        ("Environment Agency",
         "Flood zone classification by postcode. Source: EA Postcodes Risk Assessment."),
        ("Police API",
         "Street-level crime data by location. Source: data.police.uk."),
        ("OS Names API",
         "Postcode to coordinate resolution. Source: OS Data Hub."),
        ("Shodan",
         "Internet-exposed device and vulnerability intelligence. Source: Shodan.io."),
        ("WiredScore",
         "Building certification status. Manually verified via wiredscore.com/map."),
    ]
    for name, desc in sources_list:
        story.append(Paragraph(f"<b>{name}</b> — {desc}", S["body"]))
        story.append(Spacer(1, 3*mm))

    fn = _footer_fn(report_title or "Master Report")
    doc.build(story, onFirstPage=fn, onLaterPages=fn)
    return buf.getvalue()


# ── MAIN UI ────────────────────────────────────────────────────────────────────

st.markdown("""
<div style="background:#0b1829;padding:20px 28px;border-radius:10px;margin-bottom:24px">
<div style="font-size:10px;letter-spacing:2px;color:#f59e0b;font-family:monospace;
            margin-bottom:6px">INTERNAL USE ONLY</div>
<div style="font-size:24px;font-weight:800;color:#fff;margin-bottom:4px">
    Modern Networks</div>
<div style="font-size:14px;color:#64748b">Master Intelligence Report Generator</div>
</div>
""", unsafe_allow_html=True)

st.markdown(
    "Combine data from the Science Parks app and Vulnerability Scanner "
    "into a single master PDF for sales and marketing use."
)

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### Report Settings")
    report_title = st.text_input("Report title",
                                  placeholder="e.g. Oxford-Cambridge Arc Q2 2026")
    prepared_by  = st.text_input("Prepared by", placeholder="Your name")

    st.divider()
    st.markdown("### Upload Intelligence Data")
    st.caption("Upload JSON export files from the Science Parks app and/or Vulnerability Scanner.")

    uploaded_files = st.file_uploader(
        "Upload JSON files",
        type=["json"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if uploaded_files:
        for f in uploaded_files:
            data = parse_upload(f)
            if data:
                already = any(
                    u.get("exported_at") == data.get("exported_at") and
                    u.get("area_label")  == data.get("area_label")
                    for u in st.session_state.uploads
                )
                if not already:
                    st.session_state.uploads.append(data)
                    st.success(f"Loaded: {f.name}")

    if st.session_state.uploads:
        st.markdown(f"**{len(st.session_state.uploads)} file(s) loaded:**")
        for i, u in enumerate(st.session_state.uploads):
            lbl, badge = source_label(u.get("source_app",""))
            parks_n = len(u.get("parks",[]))
            areas_n = len(u.get("areas",[]))
            desc = (f"{parks_n} parks" if parks_n else
                    f"{areas_n} areas" if areas_n else "1 area")
            c1, c2 = st.columns([4,1])
            with c1:
                st.markdown(
                    f"<span class='source-badge {badge}'>{lbl}</span> "
                    f"{u.get('area_label','') or u.get('search_term','')} — {desc}",
                    unsafe_allow_html=True
                )
            with c2:
                if st.button("✕", key=f"del_{i}"):
                    st.session_state.uploads.pop(i)
                    st.rerun()
        if st.button("Clear All Uploads", use_container_width=True):
            st.session_state.uploads = []
            st.rerun()

    st.divider()
    st.markdown("### Individual Building Postcodes")
    st.caption("Optionally add postcodes for building-level entries.")

    pc_input = st.text_area(
        "Postcodes",
        height=100,
        placeholder="EC3V 1AB\nSW1A 1AA\nCB4 0WS",
        label_visibility="collapsed"
    )
    if st.button("Add Postcodes", use_container_width=True):
        pcs = [p.strip().upper() for p in pc_input.strip().split("\n") if p.strip()]
        st.session_state.postcodes = list(dict.fromkeys(
            st.session_state.postcodes + pcs
        ))
        st.success(f"{len(pcs)} postcode(s) added.")

    if st.session_state.postcodes:
        st.caption(
            f"{len(st.session_state.postcodes)} postcodes: " +
            ", ".join(st.session_state.postcodes[:8]) +
            (f" +{len(st.session_state.postcodes)-8} more"
             if len(st.session_state.postcodes) > 8 else "")
        )
        if st.button("Clear Postcodes", use_container_width=True):
            st.session_state.postcodes = []
            st.rerun()


with col2:
    st.markdown("### Report Preview")

    if not st.session_state.uploads and not st.session_state.postcodes:
        st.markdown("""
        <div style="background:#fff;border:2px dashed #e2e8f0;border-radius:10px;
                    padding:48px;text-align:center;color:#94a3b8">
            <div style="font-size:44px;margin-bottom:14px">📊</div>
            <div style="font-size:15px;font-weight:600;color:#64748b;margin-bottom:8px">
                No data loaded yet</div>
            <div style="font-size:13px;line-height:1.8">
                Upload JSON export files from the Science Parks app<br>
                or Vulnerability Scanner using the panel on the left.
            </div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("#### Contents of this report:")
        sections = ["Cover page with summary statistics", "Executive summary"]

        for u in st.session_state.uploads:
            lbl   = source_label(u.get("source_app",""))[0]
            label = u.get("area_label","") or u.get("search_term","")
            parks = u.get("parks",[])
            areas = u.get("areas",[])

            if parks:
                sections.append(f"{lbl} — {label} ({len(parks)} parks)")
                for p in parks[:4]:
                    st.caption(f"  · {p.get('name','')} ({p.get('postcode','')})")
                if len(parks) > 4:
                    st.caption(f"  · +{len(parks)-4} more parks")
            elif areas:
                for a in areas:
                    sections.append(
                        f"{lbl} — {a.get('area_label','') or a.get('search_term','')} "
                        f"(Score: {a.get('score','')} · "
                        f"{a.get('critical_count',0)} critical)"
                    )
            else:
                sections.append(f"{lbl} — {label}")

        if st.session_state.postcodes:
            sections.append(
                f"Individual buildings — {len(st.session_state.postcodes)} postcodes"
            )

        sections.append("Prioritised opportunity list")
        sections.append("Data sources appendix")

        for i, s in enumerate(sections, 1):
            st.markdown(f"{i}. {s}")

        st.divider()

        postcodes_data = [
            {"postcode": pc, "score": 0, "verdict": "Pending",
             "gaps": [], "company": ""}
            for pc in st.session_state.postcodes
        ]

        if st.button("⬇ Generate Master Report PDF", type="primary",
                     use_container_width=True):
            if not report_title:
                st.warning("Please enter a report title first.")
            else:
                with st.spinner("Building master report…"):
                    pdf_bytes = generate_master_pdf(
                        st.session_state.uploads,
                        postcodes_data,
                        report_title,
                        prepared_by or "MN Staff"
                    )
                safe = report_title.replace(" ","-").replace("/","-")
                st.download_button(
                    "⬇ Download Master Report PDF",
                    data=pdf_bytes,
                    file_name=f"MN-Master-{safe}-{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
                st.success("Report generated successfully.")