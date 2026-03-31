import streamlit as st
import json
import io
import requests

# ── AI NARRATIVE ───────────────────────────────────────────────────────────────
def _ai_narrative(prompt, fallback):
    """
    Call the Claude API to generate a narrative section.
    Returns the AI text if successful, or fallback text if anything goes wrong.
    """
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 1000,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=30,
        )
        if r.status_code == 200:
            blocks = r.json().get("content", [])
            text = " ".join(b.get("text","") for b in blocks if b.get("type")=="text").strip()
            if text:
                return text
    except Exception:
        pass
    return fallback


def _build_ai_exec_prompt(items, mode, area_str, opportunities, enriched):
    """Build the prompt for the AI executive summary."""
    AUDIENCE = {
        "retail": {
            "audience":       "shopping centre and retail park management teams, managing agents, and landlords",
            "decision_makers":"centre directors, estate managers, managing agents (CBRE, JLL, Savills, Cushman & Wakefield), and asset managers",
            "context":        "Modern Networks provides managed IT and connectivity infrastructure to major UK retail and leisure destinations including Manchester Arndale, Centre MK, and Watford. Modern Networks holds WiredScore and SmartScore Accredited Professional status.",
            "services":       "managed connectivity, full fibre, guest WiFi, Network-as-a-Service, EPOS connectivity, WiredScore/SmartScore AP services, managed IT, and network resilience",
            "hooks":          "anchor tenant digital requirements, footfall and dwell time analytics, F&B and leisure connectivity, click-and-collect infrastructure, the 2027 EPC commercial minimum, repositioning projects requiring new infrastructure from scratch, and WiredScore/SmartScore certification as a differentiator for tenant attraction",
        },
        "parks": {
            "audience":       "science and innovation park directors, estates managers, and park operators",
            "decision_makers":"park directors, estates and facilities managers, and parent university or institutional landlords",
            "context":        "Modern Networks provides managed IT and connectivity infrastructure to UK science and innovation parks. Modern Networks holds WiredScore and SmartScore Accredited Professional status.",
            "services":       "research-grade managed connectivity, full fibre, gigabit infrastructure, Network-as-a-Service, private 5G, WiredScore/SmartScore AP services, managed IT, and network resilience",
            "hooks":          "research-grade bandwidth requirements, tenant density and company growth, sector-specific needs (life sciences, AI, deep tech, genomics), the PSTN copper switch-off affecting legacy connections, 2027 EPC commercial minimum, IoT and smart campus applications, and WiredScore certification as a premium tenant differentiator",
        },
        "intel": {
            "audience":       "building owners, managing agents, and building managers of commercial office properties",
            "decision_makers":"building managers, managing agents (CBRE, JLL, Savills, Cushman & Wakefield), and asset managers",
            "context":        "Modern Networks provides managed IT and connectivity infrastructure to UK commercial office buildings. Modern Networks holds WiredScore and SmartScore Accredited Professional status.",
            "services":       "managed connectivity, full fibre, Network-as-a-Service, WiredScore/SmartScore AP services, managed IT, cybersecurity, and network resilience",
            "hooks":          "tenant attraction and retention in a flight-to-quality market, WiredScore and SmartScore certification as a competitive differentiator, EPC pressure ahead of the 2027 minimum, flood risk and network resilience, and post-pandemic office repurposing requiring new infrastructure",
        },
    }
    a    = AUDIENCE.get(mode, AUDIENCE["parks"])
    noun = "retail assets" if mode=="retail" else "science and innovation parks"
    high_opps   = sum(1 for o in opportunities if o["priority"]=="High")
    flags       = _prospect_flags(items, mode) if items else []
    ofcom_items = [p for p in items if _get_ofcom_flat(p).get("gigabit_pct") is not None]
    avg_gig  = round(sum(_get_ofcom_flat(p)["gigabit_pct"] for p in ofcom_items)/len(ofcom_items)) if ofcom_items else None
    low_gig  = sum(1 for p in ofcom_items if _get_ofcom_flat(p)["gigabit_pct"] < 50) if ofcom_items else 0
    low_ff   = sum(1 for p in ofcom_items if _get_ofcom_flat(p)["full_fibre_pct"] < 60) if ofcom_items else 0
    epc_items= [p for p in items if p.get("epc")]
    poor_epc = sum(1 for p in epc_items if (p.get("epc") or {}).get("most_common","") in ("D","E","F","G"))
    good_epc = sum(1 for p in epc_items if (p.get("epc") or {}).get("most_common","") in ("A","B","C"))
    flood_risk=sum(1 for p in items if p.get("flood_risk","") in ("Zone 3 (High)","Zone 2 (Medium)"))
    top_prospects = "\n".join(
        f"- {p['name']} ({p['postcode']}): opportunity score {p['opp_score']}/100 — {p['rationale']}"
        for p in flags[:5]
    ) if flags else "No prospect scoring data available"
    asset_detail = []
    for p in items[:15]:
        o   = _get_ofcom_flat(p)
        epc = (p.get("epc") or {}).get("most_common","unknown")
        fl  = p.get("flood_risk","unknown")
        cos = sum(1 for c in (p.get("companies") or []) if (c.get("company_status") or "").lower()=="active")
        if mode == "retail":
            detail = (f"{p.get('name','')} ({p.get('type','')}, {p.get('gla_sqft',0):,} sq ft, "
                      f"landlord: {p.get('landlord','')}): gigabit {o['gigabit_pct']:.0f}%, "
                      f"full fibre {o['full_fibre_pct']:.0f}%, EPC {epc}, flood {fl}, "
                      f"{cos} active companies"
                      + (", repositioning underway" if p.get("repositioning") else ""))
        else:
            detail = (f"{p.get('name','')} ({p.get('sector','')}, {p.get('tenants','')} tenants): "
                      f"gigabit {o['gigabit_pct']:.0f}%, full fibre {o['full_fibre_pct']:.0f}%, "
                      f"EPC {epc}, flood {fl}, {cos} active companies")
        asset_detail.append(detail)

    return f"""You are a senior sales analyst at Modern Networks, a UK managed IT and connectivity provider. You are writing an executive summary for an internal sales intelligence report.

ABOUT MODERN NETWORKS:
{a['context']}
Services offered: {a['services']}

REPORT PURPOSE:
This is an internal sales tool. The sales team will use it to identify and prioritise outreach to {a['audience']}. Key decision-makers are {a['decision_makers']}. The report should help the team understand which assets to target first, what conversation to open with, and why now is the right time.

TERRITORY: {area_str}
TOTAL {noun.upper()} PROFILED: {len(items)}

KEY DATA:
- Average gigabit coverage: {avg_gig if avg_gig is not None else 'no data'}%
- Assets below 50% gigabit: {low_gig} of {len(ofcom_items)}
- Assets with full fibre below 60%: {low_ff} of {len(ofcom_items)}
- EPC data available for: {len(epc_items)} assets ({poor_epc} rated D or below, {good_epc} rated C or above)
- Assets in EA Flood Zone 2 or 3: {flood_risk}
- Sales opportunities identified: {len(opportunities)} ({high_opps} high priority)

TOP PROSPECTS:
{top_prospects}

ASSET DETAIL:
{chr(10).join(asset_detail)}

RELEVANT SALES HOOKS:
{a['hooks']}

Write an executive summary of 4-5 substantial paragraphs covering:
1. The overall territory picture — what kind of assets, what is the digital infrastructure story, what does it mean commercially
2. The most significant connectivity gaps, naming specific assets, explaining what the gap means for operators and occupiers
3. EPC and flood risk findings where the data supports it — connect to specific service conversations
4. Top 2-3 priority prospects by name and exactly why each is a priority right now
5. The most effective sales approach for this territory — opening angle, value proposition, timing

Write in UK English, active voice, short sentences. Use specific numbers and names. Be commercially direct. No bullet points. No headings, preamble, or sign-off."""


def _build_ai_gap_prompt(items, mode, area_str):
    pass


def _build_ai_gap_prompt(items, mode, area_str):
    """Build the prompt for the AI gap analysis."""
    AUDIENCE = {
        "retail": {
            "decision_makers":"centre directors, managing agents (CBRE, JLL, Savills, Cushman & Wakefield), and asset managers",
            "context":        "Modern Networks provides managed IT and connectivity infrastructure to major UK retail and leisure destinations. Modern Networks holds WiredScore and SmartScore Accredited Professional status.",
            "services":       "managed connectivity, full fibre, guest WiFi, Network-as-a-Service, EPOS connectivity, WiredScore/SmartScore AP services, managed IT, and network resilience",
            "hooks":          "anchor tenant digital requirements, F&B and leisure connectivity needs, click-and-collect infrastructure, 2027 EPC commercial minimum, repositioning projects, and WiredScore/SmartScore certification",
        },
        "parks": {
            "decision_makers":"park directors, estates managers, and parent university or institutional landlords",
            "context":        "Modern Networks provides managed IT and connectivity infrastructure to UK science and innovation parks. Modern Networks holds WiredScore and SmartScore Accredited Professional status.",
            "services":       "research-grade managed connectivity, full fibre, gigabit infrastructure, Network-as-a-Service, private 5G, WiredScore/SmartScore AP services, managed IT, and network resilience",
            "hooks":          "research-grade bandwidth, tenant density, sector requirements (life sciences, AI, deep tech), PSTN switch-off, 2027 EPC minimum, smart campus, and WiredScore certification",
        },
        "intel": {
            "decision_makers":"building managers, managing agents, and asset managers",
            "context":        "Modern Networks provides managed IT and connectivity to UK commercial office buildings. Modern Networks holds WiredScore and SmartScore Accredited Professional status.",
            "services":       "managed connectivity, full fibre, WiredScore/SmartScore AP services, managed IT, cybersecurity, and network resilience",
            "hooks":          "flight-to-quality, WiredScore/SmartScore certification, 2027 EPC minimum, flood resilience, and office repurposing",
        },
    }
    a    = AUDIENCE.get(mode, AUDIENCE["parks"])
    noun = "retail assets" if mode=="retail" else "science and innovation parks"

    asset_detail = []
    for p in items[:20]:
        o    = _get_ofcom_flat(p)
        epc  = (p.get("epc") or {}).get("most_common","unknown")
        fl   = p.get("flood_risk","unknown")
        cos  = sum(1 for c in (p.get("companies") or []) if (c.get("company_status") or "").lower()=="active")
        if mode == "retail":
            anchors = ", ".join((p.get("anchor_tenants") or [])[:3])
            detail  = (f"{p.get('name','')} ({p.get('type','')}, {p.get('gla_sqft',0):,} sq ft, "
                       f"landlord: {p.get('landlord','')}): "
                       f"gigabit {o['gigabit_pct']:.0f}%, full fibre {o['full_fibre_pct']:.0f}%, "
                       f"5G {o['outdoor_5g_pct']:.0f}%, EPC {epc}, flood {fl}, "
                       f"{cos} active companies at postcode, anchors: {anchors or 'not specified'}"
                       + (", REPOSITIONING UNDERWAY" if p.get("repositioning") else ""))
        else:
            detail = (f"{p.get('name','')} (sector: {p.get('sector','')}, {p.get('tenants','')} tenants, "
                      f"operator: {p.get('operator','')}): "
                      f"gigabit {o['gigabit_pct']:.0f}%, full fibre {o['full_fibre_pct']:.0f}%, "
                      f"5G {o['outdoor_5g_pct']:.0f}%, EPC {epc}, flood {fl}, "
                      f"{cos} active companies at postcode")
        asset_detail.append(detail)

    return f"""You are a senior sales analyst at Modern Networks, a UK managed IT and connectivity provider. You are writing the gap analysis section of an internal sales intelligence report.

ABOUT MODERN NETWORKS:
{a['context']}
Services: {a['services']}

TERRITORY: {area_str}
TARGET AUDIENCE: {a['decision_makers']}

ASSET DATA:
{chr(10).join(asset_detail)}

RELEVANT SALES HOOKS:
{a['hooks']}

Write a gap analysis of 5-7 substantial paragraphs covering:

1. CONNECTIVITY GAPS: Which assets have the most critical shortfalls? Group assets with similar profiles. Explain what the gaps mean operationally for the asset operators and their tenants — not just the percentage figures.

2. EPC AND ENERGY: Where the data shows D or below ratings, identify which assets face the most pressure ahead of the 2027 commercial EPC minimum. Explain how a connectivity upgrade conversation fits naturally alongside an energy efficiency conversation at these assets.

3. FLOOD RISK AND RESILIENCE: Where assets are in Flood Zone 2 or 3, name them and explain the specific network resilience and business continuity argument that is relevant to their operators.

4. COMBINED-RISK PRIORITIES: Identify assets where multiple risk factors combine — connectivity gap plus poor EPC plus flood risk, or large scale plus repositioning. These represent the strongest sales cases. Explain specifically why.

5. SALES AND MARKETING APPROACH: Recommend specific tactics for engaging the decision-makers at the priority assets. Should the approach be research-led (presenting the report findings as a value-add), event-led (UKSPA conference, Revo, BCSC), or relationship-led (approaching through their managing agent)? What is the strongest opening line for this specific territory? Are there timing factors — upcoming lease events, repositioning timelines, EPC deadlines — that create urgency?

Write in UK English, active voice. Name specific assets. Be analytically sharp — find patterns, not just facts. No bullet points. No headings or preamble."""



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
.badge-retail{background:#fef3c7;color:#92400e}
</style>
""", unsafe_allow_html=True)

# ── SESSION STATE ──────────────────────────────────────────────────────────────
for key in ("uploads", "report_mode"):
    if key not in st.session_state:
        st.session_state[key] = [] if key == "uploads" else None

# ── HELPERS ────────────────────────────────────────────────────────────────────
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
        canvas.drawString(M, 10*mm, f"Modern Networks — {title} — CONFIDENTIAL — INTERNAL USE ONLY")
        canvas.drawRightString(W-M, 10*mm, f"Page {doc.page}")
        canvas.restoreState()
    return _fn

def _score_colour(score):
    try:
        s = int(score)
        return RED if s < 50 else AMBER if s < 70 else GREEN
    except (TypeError, ValueError):
        return GREY

# ── DATA HELPERS ───────────────────────────────────────────────────────────────
def _get_ofcom_flat(p):
    ofcom = p.get("ofcom") or {}
    if "connectivity" in ofcom:
        conn = ofcom.get("connectivity") or {}
        mob  = ofcom.get("mobile") or {}
        return {
            "gigabit_pct":    conn.get("gigabit_pct", 0) or 0,
            "full_fibre_pct": conn.get("full_fibre_pct", 0) or 0,
            "superfast_pct":  conn.get("superfast_pct", 0) or 0,
            "no_decent_pct":  conn.get("no_decent_pct", 0) or 0,
            "indoor_4g_pct":  mob.get("indoor_4g_all_operators_pct", 0) or 0,
            "outdoor_5g_pct": mob.get("outdoor_5g_all_operators_pct", 0) or 0,
        }
    return {k: float(ofcom.get(k, 0) or 0) for k in
            ["gigabit_pct","full_fibre_pct","superfast_pct","no_decent_pct","indoor_4g_pct","outdoor_5g_pct"]}

def _conn_score(p):
    o = _get_ofcom_flat(p)
    ff, gig, sup, nd = o["full_fibre_pct"], o["gigabit_pct"], o["superfast_pct"], o["no_decent_pct"]
    return round(min(40, ff*0.4) + min(20, gig*0.3) + min(20, sup*0.2) + max(0, 20 - nd*2))

def _rag(score):
    return "Green" if score >= 70 else "Amber" if score >= 40 else "Red"

def _opp_score(p, mode):
    """Combined opportunity score. Higher = better MN prospect."""
    o   = _get_ofcom_flat(p)
    gig = o["gigabit_pct"]
    conn_opp = 40 if gig < 20 else 30 if gig < 50 else 20 if gig < 75 else 10
    epc_mc   = (p.get("epc") or {}).get("most_common","")
    epc_opp  = {"A":0,"B":2,"C":5,"D":12,"E":18,"F":20,"G":20}.get(epc_mc, 8)
    flood    = p.get("flood_risk","")
    flood_opp= {"Zone 3 (High)":10,"Zone 2 (Medium)":6,"Zone 1 (Low)":0}.get(flood, 3)
    companies= p.get("companies") or []
    active   = sum(1 for c in companies if (c.get("company_status") or "").lower()=="active")
    co_opp   = min(20, active * 2)
    # Scale proxy
    if mode == "retail":
        gla = p.get("gla_sqft", 0) or 0
        scale_opp = min(10, gla // 100000)
    else:
        tenants = str(p.get("tenants","") or "")
        try:
            t_num = int("".join(filter(str.isdigit, tenants.split("+")[0].split(",")[0])))
        except Exception:
            t_num = 0
        scale_opp = min(10, t_num // 10)
    return min(100, conn_opp + epc_opp + flood_opp + co_opp + scale_opp)

def _get_items(uploads):
    """Flatten all parks/assets from all uploads."""
    items = []
    for u in uploads:
        area = u.get("area_label","")
        for p in u.get("parks", []):
            p2 = dict(p)
            p2["_area"] = area
            items.append(p2)
    return items

# ── NARRATIVE GENERATORS ───────────────────────────────────────────────────────
def _gap_narrative(items, mode):
    """Return list of plain-English paragraph strings for the territory."""
    if not items:
        return ["No data available for gap analysis."]

    total   = len(items)
    noun    = "assets" if mode == "retail" else "parks"
    paras   = []

    ofcom_items = [p for p in items if _get_ofcom_flat(p).get("gigabit_pct") is not None]
    if ofcom_items:
        low_gig = [p for p in ofcom_items if _get_ofcom_flat(p)["gigabit_pct"] < 50]
        low_ff  = [p for p in ofcom_items if _get_ofcom_flat(p)["full_fibre_pct"] < 60]
        low_5g  = [p for p in ofcom_items if _get_ofcom_flat(p)["outdoor_5g_pct"] < 40]

        if low_gig:
            pct   = round(len(low_gig)/total*100)
            names = ", ".join(p.get("name","") for p in low_gig[:3])
            extra = f" — including {names}" + (f" and {len(low_gig)-3} others" if len(low_gig)>3 else "")
            paras.append(
                f"{len(low_gig)} of {total} {noun} ({pct}%) are in local authority areas where gigabit "
                f"broadband coverage is below 50%{extra}. This represents a direct connectivity upgrade "
                f"opportunity and should be the primary conversation opener at these locations."
            )
        if low_ff and len(low_ff) != len(low_gig):
            paras.append(
                f"A further {len(low_ff)} {noun} have full fibre availability below 60%, indicating that "
                f"legacy broadband infrastructure may be limiting digital capability. "
                f"A managed connectivity assessment is recommended at each site."
            )
        if low_5g:
            paras.append(
                f"{len(low_5g)} {noun} have outdoor 5G coverage below 40%, limiting mobile-first "
                f"applications{', smart campus operations,' if mode=='parks' else ', delivery management,'} "
                f"and IoT use cases. Private 5G and indoor mobile enhancement are relevant service conversations."
            )

    epc_items  = [p for p in items if p.get("epc")]
    poor_epc   = [p for p in epc_items if (p.get("epc") or {}).get("most_common","") in ("D","E","F","G")]
    if poor_epc:
        paras.append(
            f"{len(poor_epc)} of {len(epc_items)} {noun} with EPC data show a most common rating of D or below. "
            f"With the proposed 2027 commercial EPC minimum of C, these assets face significant upgrade pressure. "
            f"Connectivity modernisation conversations can be paired with energy efficiency messaging — "
            f"both point to the same infrastructure investment cycle."
        )
    elif epc_items:
        good_epc = [p for p in epc_items if (p.get("epc") or {}).get("most_common","") in ("A","B","C")]
        if good_epc:
            paras.append(
                f"{len(good_epc)} of {len(epc_items)} {noun} show EPC ratings of C or above, indicating "
                f"well-maintained building stock. WiredScore and SmartScore certification is the natural "
                f"next conversation — the infrastructure investment case is already present."
            )

    flood_high = [p for p in items if p.get("flood_risk","") == "Zone 3 (High)"]
    flood_med  = [p for p in items if p.get("flood_risk","") == "Zone 2 (Medium)"]
    if flood_high:
        names = ", ".join(p.get("name","") for p in flood_high[:2])
        paras.append(
            f"{len(flood_high)} {'asset' if len(flood_high)==1 else noun} {'is' if len(flood_high)==1 else 'are'} "
            f"in EA Flood Zone 3 ({names}{'...' if len(flood_high)>2 else ''}). "
            f"Network resilience, dual-path routing, and business continuity planning are high-relevance "
            f"service lines at these locations."
        )
    elif flood_med:
        paras.append(
            f"{len(flood_med)} {'asset' if len(flood_med)==1 else noun} "
            f"{'sits' if len(flood_med)==1 else 'sit'} in EA Flood Zone 2. "
            f"Infrastructure resilience is a relevant conversation, particularly for "
            f"{'research-intensive occupiers' if mode!='retail' else 'anchor tenants'} with continuity obligations."
        )

    cos_items = [p for p in items if p.get("companies")]
    if cos_items:
        total_active = sum(
            sum(1 for c in (p.get("companies") or []) if (c.get("company_status") or "").lower()=="active")
            for p in cos_items
        )
        paras.append(
            f"Across {len(cos_items)} {noun} with Companies House data, "
            f"{total_active} active registered companies were identified at postcodes. "
            f"This represents the addressable occupier base for estate-wide managed connectivity, "
            f"IT support, and M365 services."
        )

    if not paras:
        paras.append(
            f"Connectivity and infrastructure data has been collected for the {noun} in this territory. "
            f"Run the Full Intelligence option in the source app to enrich the export "
            f"with EPC, Companies House, and flood risk data for deeper gap analysis."
        )
    return paras

def _prospect_flags(items, mode):
    scored = []
    for p in items:
        opp  = _opp_score(p, mode)
        conn = _conn_score(p)
        o    = _get_ofcom_flat(p)
        epc  = p.get("epc") or {}
        flood= p.get("flood_risk","")
        cos  = sum(1 for c in (p.get("companies") or []) if (c.get("company_status") or "").lower()=="active")

        reasons = []
        gig = o["gigabit_pct"]
        if gig < 30:   reasons.append(f"gigabit coverage only {gig:.0f}% — direct fibre opportunity")
        elif gig < 60: reasons.append(f"gigabit coverage {gig:.0f}% — upgrade conversation")

        mc = (epc.get("most_common") or "").upper()
        if mc in ("E","F","G"): reasons.append(f"EPC {mc} — below 2027 minimum, energy+connectivity messaging")
        elif mc == "D":         reasons.append("EPC D — approaching 2027 threshold, upgrade pressure building")

        if flood == "Zone 3 (High)":   reasons.append("EA Flood Zone 3 — resilience and continuity planning")
        elif flood == "Zone 2 (Medium)": reasons.append("EA Flood Zone 2 — resilience conversation relevant")

        if cos >= 15:  reasons.append(f"{cos} active companies at postcode — managed services opportunity")
        elif cos >= 5: reasons.append(f"{cos} active companies — connectivity addressable market")

        if mode == "retail":
            gla = p.get("gla_sqft", 0) or 0
            if gla >= 1000000: reasons.append(f"{gla:,} sq ft — estate-scale managed network opportunity")
            anchors = [a.lower() for a in (p.get("anchor_tenants") or [])]
            if any(a in anchors for a in ["vue cinema","cineworld","odeon"]): reasons.append("cinema anchor — high-bandwidth digital projection and ticketing")
        else:
            sector = (p.get("sector") or "").lower()
            if any(x in sector for x in ["life science","biomedical","genomic"]): reasons.append("life sciences — high bandwidth + compliance requirements")
            elif any(x in sector for x in ["ai","deep tech","gpu","hpc"]): reasons.append("AI/deep tech — 10Gbps+ for GPU/HPC workloads")

        if not reasons: reasons.append("connectivity and certification assessment recommended")

        scored.append({
            "name":      p.get("name",""),
            "postcode":  p.get("postcode",""),
            "area":      p.get("_area",""),
            "type":      p.get("type","") if mode=="retail" else p.get("sector",""),
            "opp_score": opp,
            "conn_score":conn,
            "epc":       mc or "—",
            "flood":     flood or "—",
            "rationale": "  ·  ".join(reasons[:3]),
        })
    scored.sort(key=lambda x: -x["opp_score"])
    return scored

# ── OPPORTUNITY BUILDERS ───────────────────────────────────────────────────────
def _build_opps_parks(uploads):
    opps = []
    for upload in uploads:
        for p in upload.get("parks", []):
            name, pc = p.get("name",""), p.get("postcode","")
            sector, tenants = p.get("sector",""), p.get("tenants","")
            gig = _get_ofcom_flat(p).get("gigabit_pct",0)
            if gig < 50:
                opps.append({"priority":"High" if gig<20 else "Medium","property":name,"postcode":pc,
                    "type":"Science Park","gap":f"Gigabit coverage {gig:.0f}% — below standard for innovation park",
                    "service":"Fibre Broadband · Network-as-a-Service · WiredScore AP Services",
                    "reason":f"{tenants} tenants in {sector} — gigabit infrastructure required to attract and retain premium occupiers"})
            else:
                opps.append({"priority":"Medium","property":name,"postcode":pc,
                    "type":"Science Park","gap":"WiredScore / SmartScore certification not confirmed",
                    "service":"WiredScore AP Services · SmartScore AP Services",
                    "reason":f"Certification differentiates {name} for premium tenant attraction — MN are Accredited Professionals"})
            try:
                t_num = int(str(tenants).replace("+","").replace(",","").split()[0])
            except Exception:
                t_num = 0
            if t_num >= 100:
                opps.append({"priority":"Medium","property":name,"postcode":pc,"type":"Science Park",
                    "gap":f"{tenants} tenant organisations — managed IT and connectivity opportunity",
                    "service":"Service Guardian · Managed Network · Desktop Support · M365",
                    "reason":"Large multi-tenant environment with demand for a single managed services partner"})
    opps.sort(key=lambda x: (0 if x["priority"]=="High" else 1, x["property"]))
    return opps

def _build_opps_retail(uploads):
    opps = []
    for upload in uploads:
        for p in upload.get("parks",[]):
            name, pc = p.get("name",""), p.get("postcode","")
            gig  = _get_ofcom_flat(p).get("gigabit_pct",0)
            ff   = _get_ofcom_flat(p).get("full_fibre_pct",0)
            gla  = p.get("gla_sqft",0) or 0
            reposition = p.get("repositioning",False)
            anchors = [a.lower() for a in (p.get("anchor_tenants") or [])]
            notes = (p.get("notes") or "").lower()
            atype = (p.get("type") or "").lower()

            if gig < 50:
                opps.append({"priority":"High" if gig<20 else "Medium","property":name,"postcode":pc,
                    "type":p.get("type",""),"gap":f"Gigabit coverage {gig:.0f}% — below threshold for modern retail operations",
                    "service":"Fibre Broadband · Network-as-a-Service","reason":f"Major retail asset in area with {gig:.0f}% gigabit coverage — connectivity upgrade required for retailer systems"})
            if ff < 60:
                opps.append({"priority":"Medium","property":name,"postcode":pc,"type":p.get("type",""),
                    "gap":"Full fibre availability gap — legacy broadband limiting EPOS and cloud point-of-sale",
                    "service":"Managed Connectivity · Full Fibre Upgrade","reason":"Below 60% full fibre availability in LA area — retailer digital operations at risk"})
            if gla >= 1000000:
                opps.append({"priority":"High","property":name,"postcode":pc,"type":p.get("type",""),
                    "gap":f"Estate-wide managed network — {gla:,} sq ft requires enterprise-grade multi-tenant infrastructure",
                    "service":"Network-as-a-Service · Managed Network · Service Guardian","reason":f"Major regional asset at scale requiring centralised network management across all retailer units"})
            elif gla >= 500000:
                opps.append({"priority":"Medium","property":name,"postcode":pc,"type":p.get("type",""),
                    "gap":f"Multi-tenant managed connectivity — {gla:,} sq ft with multiple anchor tenants",
                    "service":"Managed Connectivity · Network-as-a-Service","reason":"Large asset benefits from single managed services provider across all retail units"})
            if reposition:
                opps.append({"priority":"High","property":name,"postcode":pc,"type":p.get("type",""),
                    "gap":"Repositioning / redevelopment — live infrastructure brief",
                    "service":"Network Design · Full Fibre · WiredScore AP Services · SmartScore AP Services",
                    "reason":"Asset undergoing significant change — opportunity to specify modern network infrastructure from the outset"})
            if any(x in notes for x in ["food court","food hall","restaurants","dining","leisure"]):
                opps.append({"priority":"Medium","property":name,"postcode":pc,"type":p.get("type",""),
                    "gap":"F&B and leisure managed connectivity — reservation systems, EPOS, and guest WiFi",
                    "service":"Guest WiFi · Managed Connectivity · EPOS Connectivity","reason":"Food and leisure operators require reliable high-bandwidth connectivity for customer-facing systems"})
            if "regional" in atype or "sub-regional" in atype:
                opps.append({"priority":"Medium","property":name,"postcode":pc,"type":p.get("type",""),
                    "gap":"WiredScore / SmartScore certification gap",
                    "service":"WiredScore AP Services · SmartScore AP Services","reason":f"Major retail scheme — certification differentiates {name} for premium occupier attraction; MN are Accredited Professionals"})
    opps.sort(key=lambda x: (0 if x["priority"]=="High" else 1, x["property"]))
    return opps

def _build_opps_intel(intel_uploads):
    opps = []
    for upload in intel_uploads:
        for b in upload.get("briefings",[]):
            pc      = b.get("postcode","")
            company = b.get("company","") or pc
            gaps    = b.get("gaps",[])
            for g in gaps[:2]:
                opps.append({"priority":"High" if g.get("sev")=="critical" else "Medium",
                    "property":company,"postcode":pc,"type":"Building Assessment",
                    "gap":g.get("title",""),"service":g.get("service","").split("\n")[0],
                    "reason":g.get("desc","")[:120]})
    opps.sort(key=lambda x: (0 if x["priority"]=="High" else 1, x["property"]))
    return opps

# ── PDF BUILDERS ───────────────────────────────────────────────────────────────
def _pdf_cover(story, S, report_title, prepared_by, mode, stats_data):
    t = Table([[Paragraph(
        "INTERNAL  ·  MODERN NETWORKS SALES & MARKETING INTELLIGENCE  ·  NOT FOR EXTERNAL DISTRIBUTION",
        S["whiteb"])]], colWidths=[CW])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),("TOPPADDING",(0,0),(-1,-1),6),
                            ("BOTTOMPADDING",(0,0),(-1,-1),6),("LEFTPADDING",(0,0),(-1,-1),10)]))
    story.append(t)
    story.append(Spacer(1, 10*mm))

    mode_labels = {"parks":"Science & Innovation Parks","retail":"Retail Property","intel":"Building Intelligence"}
    story.append(Paragraph(f"Modern Networks  |  {mode_labels.get(mode,'Territory')} Intelligence Report", S["teal"]))
    story.append(Paragraph(f"Prepared by: {prepared_by}  ·  {datetime.now().strftime('%d %b %Y')}", S["mono"]))
    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(report_title or "Territory Intelligence Report", S["h1"]))
    story.append(Spacer(1, 4*mm))

    nums  = [Paragraph(str(v), ParagraphStyle(f"sn{i}", fontName="Helvetica-Bold", fontSize=22,
             textColor=[TEAL,GREEN,AMBER,RED][i], leading=26)) for i, (v,_) in enumerate(stats_data)]
    labels= [Paragraph(lbl, S["small"]) for _,lbl in stats_data]
    st_t  = Table([nums, labels], colWidths=[CW/len(stats_data)]*len(stats_data))
    st_t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),LGREY),("BOX",(0,0),(-1,-1),0.5,MGREY),
                               ("VALIGN",(0,0),(-1,-1),"TOP"),("TOPPADDING",(0,0),(-1,-1),10),
                               ("BOTTOMPADDING",(0,0),(-1,-1),10),("LEFTPADDING",(0,0),(-1,-1),10),
                               ("LINEBEFORE",(1,0),(len(stats_data)-1,-1),0.5,MGREY)]))
    story.append(st_t)
    story.append(Spacer(1, 6*mm))

def _pdf_exec_summary(story, S, items, mode, opportunities, area_str, use_ai=False):
    story.append(_bar("EXECUTIVE SUMMARY", S))
    story.append(Spacer(1, 4*mm))

    noun     = "assets" if mode=="retail" else "parks"
    enriched = any(p.get("epc") or p.get("companies") or p.get("flood_risk") for p in items)
    high_opps= sum(1 for o in opportunities if o["priority"]=="High")

    # Build template fallback first
    flags    = _prospect_flags(items, mode)
    ofcom_items = [p for p in items if _get_ofcom_flat(p).get("gigabit_pct") is not None]
    fallback_lines = []
    fallback_lines.append(
        f"This report covers digital infrastructure intelligence for {area_str}, "
        f"profiling {len(items)} {'retail asset' if mode=='retail' else 'science and innovation park'}{'s' if len(items)!=1 else ''}."
    )
    if ofcom_items:
        avg_gig = round(sum(_get_ofcom_flat(p)["gigabit_pct"] for p in ofcom_items)/len(ofcom_items))
        low_gig = sum(1 for p in ofcom_items if _get_ofcom_flat(p)["gigabit_pct"] < 50)
        fallback_lines.append(
            f"Connectivity analysis shows an average gigabit coverage of {avg_gig}% across the territory. "
            f"{low_gig} of {len(ofcom_items)} {noun} are in areas below the 50% gigabit threshold."
        )
    if flags:
        top3 = ", ".join(p["name"] for p in flags[:3])
        fallback_lines.append(
            f"Opportunity scoring identifies {top3} as the highest-priority prospects in this territory."
        )
    fallback_lines.append(
        f"The analysis has identified {len(opportunities)} service opportunities for Modern Networks, "
        f"of which {high_opps} are high priority."
    )
    fallback_text = " ".join(fallback_lines)

    if use_ai and items and enriched:
        prompt = _build_ai_exec_prompt(items, mode, area_str, opportunities, enriched)
        text   = _ai_narrative(prompt, fallback_text)
        # Split on double newlines or ". " boundaries into paragraphs
        paras  = [p.strip() for p in text.split("\n\n") if p.strip()]
        if len(paras) == 1:
            # Try splitting into ~3 chunks by sentence
            sentences = text.split(". ")
            third = max(1, len(sentences)//3)
            paras = [
                ". ".join(sentences[:third]) + ".",
                ". ".join(sentences[third:2*third]) + ".",
                ". ".join(sentences[2*third:]),
            ]
            paras = [p.strip() for p in paras if p.strip()]
        for para in paras:
            story.append(Paragraph(para, S["body"]))
            story.append(Spacer(1, 3*mm))
        story.append(Paragraph(
            "AI-assisted narrative — Modern Networks Building Intelligence Platform",
            S["small"]))
    else:
        for line in fallback_lines:
            story.append(Paragraph(line, S["body"]))
            story.append(Spacer(1, 3*mm))

    story.append(PageBreak())

def _pdf_asset_table(story, S, upload, mode):
    area_label = upload.get("area_label","") or ("Retail Assets" if mode=="retail" else "Science Parks")
    parks      = upload.get("parks",[])
    exported   = upload.get("exported_at","")
    noun       = "assets" if mode=="retail" else "parks"
    bar_label  = "RETAIL ASSETS" if mode=="retail" else "SCIENCE & INNOVATION PARKS"

    story.append(_bar(f"{bar_label} — {area_label.upper()}", S))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"{len(parks)} {noun}  ·  Exported: {exported}", S["mono"]))
    story.append(Spacer(1, 4*mm))

    enriched = any(p.get("epc") or p.get("companies") or p.get("flood_risk") for p in parks)

    if enriched:
        hdr = [Paragraph(h, S["monow"]) for h in
               (["ASSET","POSTCODE","CONN","EPC","FLOOD","COS","GLA"] if mode=="retail"
                else ["PARK","POSTCODE","CONN","EPC","FLOOD","COS","TENANTS"])]
        rows = [hdr]
        for p in sorted(parks, key=lambda x: -_opp_score(x, mode)):
            cs  = _conn_score(p)
            rag = _rag(cs)
            rc  = {"Green":"●","Amber":"◑","Red":"○"}.get(rag,"")
            epc = (p.get("epc") or {}).get("most_common","—") or "—"
            fl  = p.get("flood_risk","—") or "—"
            fls = {"Zone 3 (High)":"Z3 ⚠","Zone 2 (Medium)":"Z2","Zone 1 (Low)":"Z1"}.get(fl,"—")
            cos = sum(1 for c in (p.get("companies") or []) if (c.get("company_status") or "").lower()=="active")
            scl = f"{p.get('gla_sqft',0):,} sf" if mode=="retail" else str(p.get("tenants",""))
            rows.append([Paragraph(p.get("name","")[:28],S["bold9"]),Paragraph(p.get("postcode",""),S["body"]),
                         Paragraph(f"{rc} {cs}/100",S["body"]),Paragraph(epc,S["body"]),
                         Paragraph(fls,S["body"]),Paragraph(str(cos) if cos else "—",S["body"]),
                         Paragraph(scl,S["body"])])
        cws = [46*mm,20*mm,20*mm,12*mm,16*mm,12*mm,CW-126*mm]
    else:
        if mode == "retail":
            hdr  = [Paragraph(h,S["monow"]) for h in ["ASSET","POSTCODE","TYPE","GLA","LANDLORD"]]
            rows = [hdr]
            for p in parks:
                rows.append([Paragraph(p.get("name","")[:35],S["bold9"]),Paragraph(p.get("postcode",""),S["body"]),
                             Paragraph(p.get("type","")[:28],S["small"]),Paragraph(f"{p.get('gla_sqft',0):,}",S["body"]),
                             Paragraph(p.get("landlord","")[:28],S["small"])])
            cws = [55*mm,22*mm,40*mm,18*mm,CW-135*mm]
        else:
            hdr  = [Paragraph(h,S["monow"]) for h in ["PARK","POSTCODE","SECTOR","TENANTS","OPERATOR"]]
            rows = [hdr]
            for p in parks:
                rows.append([Paragraph(p.get("name","")[:35],S["bold9"]),Paragraph(p.get("postcode",""),S["body"]),
                             Paragraph(p.get("sector","")[:28],S["small"]),Paragraph(str(p.get("tenants","")),S["body"]),
                             Paragraph(p.get("operator","")[:28],S["small"])])
            cws = [55*mm,22*mm,43*mm,18*mm,CW-138*mm]

    pt = Table(rows, colWidths=cws)
    pt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,LGREY]),
                             ("LINEBELOW",(0,0),(-1,-1),0.3,MGREY),("TOPPADDING",(0,0),(-1,-1),5),
                             ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),5),
                             ("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(pt)
    if enriched:
        story.append(Paragraph(
            "CONN = connectivity score  ·  EPC = most common non-domestic rating  ·  "
            "FLOOD = EA zone  ·  COS = active Companies House registrations  ·  sorted by opportunity score",
            S["small"]))
    story.append(Spacer(1, 5*mm))

    for p in parks:
        if p.get("notes"):
            story.append(Paragraph(f"{p.get('name','')} ({p.get('postcode','')}) — {p.get('notes','')[:200]}", S["small"]))
            story.append(Spacer(1, 2*mm))
    story.append(PageBreak())

def _pdf_rankings(story, S, items, mode, flags):
    story.append(_bar("TERRITORY RANKINGS — OPPORTUNITY SCORE", S))
    story.append(Spacer(1, 3*mm))
    noun = "assets" if mode=="retail" else "parks"
    story.append(Paragraph(
        f"{noun.capitalize()} ranked by combined opportunity score across connectivity, EPC, flood risk, "
        f"and occupier density. Higher score = stronger case for Modern Networks engagement.", S["italic"]))
    story.append(Spacer(1, 4*mm))

    hdr = [Paragraph(h,S["monow"]) for h in ["#","NAME","AREA","OPP","CONN","EPC","FLOOD",
                                               "TYPE" if mode=="retail" else "SECTOR"]]
    rows = [hdr]
    for i, p in enumerate(flags[:20], 1):
        opp_col = RED if p["opp_score"]>=60 else AMBER if p["opp_score"]>=35 else GREEN
        fls = {"Zone 3 (High)":"Z3","Zone 2 (Medium)":"Z2","Zone 1 (Low)":"Z1"}.get(p["flood"],"—")
        rows.append([
            Paragraph(str(i),S["body"]),
            Paragraph(p["name"][:28],S["bold9"]),
            Paragraph((p["area"] or "")[:20],S["small"]),
            Paragraph(str(p["opp_score"]),ParagraphStyle("os",fontName="Helvetica-Bold",fontSize=9,textColor=opp_col,leading=12)),
            Paragraph(f"{p['conn_score']}/100",S["body"]),
            Paragraph(p["epc"],S["body"]),
            Paragraph(fls,S["body"]),
            Paragraph((p["type"] or "")[:22],S["small"]),
        ])
    t = Table(rows, colWidths=[10*mm,46*mm,30*mm,15*mm,18*mm,12*mm,14*mm,CW-145*mm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),NAVY),("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,LGREY]),
                            ("LINEBELOW",(0,0),(-1,-1),0.3,MGREY),("TOPPADDING",(0,0),(-1,-1),5),
                            ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),5),
                            ("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(t)
    story.append(PageBreak())

def _pdf_gap_analysis(story, S, items, mode, area_str="", use_ai=False):
    story.append(_bar("TERRITORY GAP ANALYSIS", S))
    story.append(Spacer(1, 4*mm))

    if use_ai and items:
        fallback_paras = _gap_narrative(items, mode)
        fallback_text  = " ".join(fallback_paras)
        prompt = _build_ai_gap_prompt(items, mode, area_str)
        text   = _ai_narrative(prompt, fallback_text)
        paras  = [p.strip() for p in text.split("\n\n") if p.strip()]
        if not paras:
            paras = [text]
        for para in paras:
            story.append(Paragraph(para, S["body"]))
            story.append(Spacer(1, 4*mm))
        story.append(Paragraph(
            "AI-assisted analysis — Modern Networks Building Intelligence Platform",
            S["small"]))
    else:
        for para in _gap_narrative(items, mode):
            story.append(Paragraph(para, S["body"]))
            story.append(Spacer(1, 4*mm))

    story.append(PageBreak())

def _pdf_prospect_flags(story, S, flags):
    story.append(_bar("PRIORITY PROSPECTS", S))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "Top prospects by opportunity score with rationale to support sales conversation preparation.", S["italic"]))
    story.append(Spacer(1, 4*mm))
    for p in flags[:10]:
        opp_col = RED if p["opp_score"]>=60 else AMBER if p["opp_score"]>=35 else GREEN
        bg_col  = LRED if p["opp_score"]>=60 else LCREAM if p["opp_score"]>=35 else LGREEN
        fls = {"Zone 3 (High)":"Zone 3","Zone 2 (Medium)":"Zone 2","Zone 1 (Low)":"Zone 1"}.get(p["flood"],p["flood"])
        ft = Table([[
            Paragraph(str(p["opp_score"]),ParagraphStyle("os2",fontName="Helvetica-Bold",fontSize=14,textColor=opp_col,leading=18)),
            [Paragraph(f'{p["name"]}  ·  {p["postcode"]}{"  ·  "+p["area"] if p["area"] else ""}',S["bold9"]),
             Paragraph(p["rationale"],S["body"]),
             Paragraph(f'Type: {p["type"] or "—"}  ·  Connectivity: {p["conn_score"]}/100  ·  EPC: {p["epc"]}  ·  Flood: {fls}',S["small"])],
        ]], colWidths=[18*mm,CW-18*mm])
        ft.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),bg_col),("LINEBEFORE",(0,0),(0,-1),3,opp_col),
                                 ("BOX",(0,0),(-1,-1),0.5,MGREY),("VALIGN",(0,0),(-1,-1),"TOP"),
                                 ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
                                 ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))
        story.append(KeepTogether([ft, Spacer(1, 4*mm)]))
    story.append(PageBreak())

def _pdf_action_list(story, S, opportunities):
    story.append(_bar("PRIORITY ACTION LIST", S))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph("All identified opportunities ranked by priority. High priority items should be actioned within 30 days.", S["italic"]))
    story.append(Spacer(1, 5*mm))
    if not opportunities:
        story.append(Paragraph("No opportunities identified from uploaded data.", S["italic"]))
        return
    for o in opportunities[:30]:
        pc_col = RED if o["priority"]=="High" else AMBER
        bg_col = LRED if o["priority"]=="High" else LCREAM
        ot = Table([[
            Paragraph(o["priority"],ParagraphStyle("pr",fontName="Helvetica-Bold",fontSize=8,textColor=pc_col,leading=11)),
            [Paragraph(f'{o["property"]}{"  ·  "+o["postcode"] if o["postcode"] else ""}{"  ·  "+o["type"] if o.get("type") else ""}',S["bold9"]),
             Paragraph(o["gap"],S["body"]),
             Paragraph(f'Why now: {o["reason"]}',S["italic"])],
            Paragraph(o["service"],S["teal"]),
        ]], colWidths=[14*mm,CW-82*mm,68*mm])
        ot.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),bg_col),("LINEBEFORE",(0,0),(0,-1),3,pc_col),
                                 ("BOX",(0,0),(-1,-1),0.5,MGREY),("VALIGN",(0,0),(-1,-1),"TOP"),
                                 ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
                                 ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8)]))
        story.append(KeepTogether([ot, Spacer(1,4*mm)]))

def _pdf_building_intel(story, S, intel_uploads):
    story.append(_bar("INDIVIDUAL BUILDING ASSESSMENTS", S))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(
        "Individual building assessments from the Modern Networks Building Intelligence Platform. "
        "Each building has been assessed against Ofcom, EPC Register, Companies House, EA flood risk, and crime data.",
        S["italic"]))
    story.append(Spacer(1, 4*mm))
    for upload in intel_uploads:
        for b in sorted(upload.get("briefings",[]), key=lambda b: b.get("score",100)):
            pc, company = b.get("postcode",""), b.get("company","")
            score, verdict, label = b.get("score",0), b.get("verdict",""), b.get("scoreLabel","")
            gaps, pos, ws = b.get("gaps",[]), b.get("positives",[]), b.get("wiredScore",{})
            sc_col = _score_colour(score)
            title_str = f"{company} — {pc}" if company else f"Building Assessment — {pc}"
            ht = Table([[Paragraph(title_str,S["bold11"]),
                         Paragraph(f'<font color="{sc_col.hexval()}"><b>{score}/100</b></font>',
                                   ParagraphStyle("bsc",fontName="Helvetica-Bold",fontSize=14,textColor=BLACK,leading=18))]],
                       colWidths=[CW-30*mm,30*mm])
            ht.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(1,0),(1,-1),"RIGHT"),
                                     ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                                     ("LINEBELOW",(0,0),(-1,-1),1,NAVY)]))
            story.append(ht)
            story.append(Spacer(1,2*mm))
            if verdict or label:
                story.append(Paragraph(f'{verdict or label}  ·  Saved {b.get("savedAt","")}',
                                        ParagraphStyle("vd",fontName="Helvetica-Bold",fontSize=8,textColor=sc_col,leading=11)))
                story.append(Spacer(1,2*mm))
            ws_status = ws.get("status","unconfirmed")
            ws_icon   = "✓" if ws_status=="certified" else "✕" if ws_status=="not-certified" else "?"
            story.append(Paragraph(
                f"WiredScore: {ws_icon} {ws_status.title()}"
                f"{' — '+ws.get('scheme','')+' '+ws.get('level','') if ws_status=='certified' else ''}",
                S["small"]))
            story.append(Spacer(1,3*mm))
            if gaps:
                story.append(Paragraph("Gaps & Opportunities",S["bold9"]))
                story.append(Spacer(1,2*mm))
                for g in gaps:
                    lc = RED if g.get("sev")=="critical" else AMBER if g.get("sev")=="advisory" else TEAL
                    bg = LRED if g.get("sev")=="critical" else LCREAM
                    gt = Table([[
                        Paragraph(g.get("sev","").upper(),ParagraphStyle("gs",fontName="Courier-Bold",fontSize=7,textColor=lc,leading=10)),
                        [Paragraph(f"{g.get('icon','')} {g.get('title','')}",S["bold9"]),Paragraph(g.get("desc","")[:200],S["small"])],
                        Paragraph(g.get("service","").replace("\n","  ·  ")[:60],S["teal"]),
                    ]], colWidths=[14*mm,CW-80*mm,66*mm])
                    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),bg),("LINEBEFORE",(0,0),(0,-1),3,lc),
                                             ("BOX",(0,0),(-1,-1),0.5,MGREY),("VALIGN",(0,0),(-1,-1),"TOP"),
                                             ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
                                             ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6)]))
                    story.append(KeepTogether([gt,Spacer(1,3*mm)]))
            if pos:
                story.append(Spacer(1,2*mm))
                story.append(Paragraph("Confirmed Strengths",S["bold9"]))
                story.append(Spacer(1,2*mm))
                for p_item in pos[:3]:
                    story.append(Paragraph(
                        f"✓ {p_item.get('icon','')} {p_item.get('title','')} — {p_item.get('desc','')[:120]}",
                        ParagraphStyle("str",fontName="Helvetica",fontSize=8,textColor=GREEN,leading=11)))
                    story.append(Spacer(1,2*mm))
            story.append(Spacer(1,4*mm))
            story.append(_hr())
    story.append(PageBreak())

def _pdf_appendix(story, S, mode):
    story.append(_bar("APPENDIX — DATA SOURCES & METHODOLOGY", S))
    story.append(Spacer(1, 4*mm))
    sources = [
        ("Ofcom Connected Nations", "Postcode-level fixed broadband coverage data. Updated quarterly by Ofcom."),
        ("EPC Register", "Non-domestic Energy Performance Certificates by postcode. Source: MHCLG Get Energy Performance Data API."),
        ("Companies House", "Active company registrations by postcode. Source: Companies House public API."),
        ("Environment Agency", "Flood zone classification. Source: EA Postcodes Risk Assessment dataset (data.gov.uk)."),
        ("WiredScore / SmartScore", "Building certification status. Manually verified by MN staff via wiredscore.com/map."),
    ]
    if mode == "parks":
        sources.append(("Science Parks Data","UK science and innovation park profiles. Source: MN Science Parks Intelligence Platform."))
    elif mode == "retail":
        sources.append(("Retail Property Data","Major UK retail asset profiles including type, GLA, landlord, and anchor tenants. Source: MN Retail Property Intelligence Platform."))
    else:
        sources.append(("OS Names API","Postcode to coordinate resolution. Source: OS Data Hub."))
        sources.append(("Police API","Street-level crime data. Source: data.police.uk."))
    for name, desc in sources:
        story.append(Paragraph(f"<b>{name}</b> — {desc}", S["body"]))
        story.append(Spacer(1, 3*mm))

# ── MAIN PDF GENERATOR ─────────────────────────────────────────────────────────
def generate_pdf(uploads, intel_uploads, mode, report_title, prepared_by, use_ai=False):
    buf = io.BytesIO()
    S   = _styles()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=M, rightMargin=M,
                             topMargin=M, bottomMargin=20*mm, title=report_title)
    story = []

    items     = _get_items(uploads)
    all_areas = list(set(u.get("area_label","") for u in uploads if u.get("area_label")))
    area_str  = ", ".join(all_areas) if all_areas else "the assessed territory"
    enriched  = any(p.get("epc") or p.get("companies") or p.get("flood_risk") for p in items)
    flags     = _prospect_flags(items, mode) if enriched else []

    if mode == "parks":
        opps = _build_opps_parks(uploads)
    elif mode == "retail":
        opps = _build_opps_retail(uploads)
    else:
        opps = _build_opps_intel(intel_uploads)

    if mode == "intel":
        total_briefings = sum(len(u.get("briefings",[])) for u in intel_uploads)
        high_opps = sum(1 for o in opps if o["priority"]=="High")
        stats = [(total_briefings,"Buildings Assessed"),(len(opps),"Opportunities"),(high_opps,"High Priority")]
    else:
        high_opps = sum(1 for o in opps if o["priority"]=="High")
        noun_cap  = "Retail Assets" if mode=="retail" else "Science Parks"
        stats = [(len(items),noun_cap),(len(opps),"Opportunities"),(high_opps,"High Priority")]
        if enriched:
            epc_count = sum(1 for p in items if p.get("epc"))
            stats.append((epc_count,"With EPC Data"))

    _pdf_cover(story, S, report_title, prepared_by, mode, stats)

    if mode == "intel":
        story.append(_bar("EXECUTIVE SUMMARY", S))
        story.append(Spacer(1, 4*mm))
        high_opps = sum(1 for o in opps if o["priority"]=="High")
        total_briefings = sum(len(u.get("briefings",[])) for u in intel_uploads)
        story.append(Paragraph(
            f"This report covers {total_briefings} individual building assessment{'s' if total_briefings!=1 else ''} "
            f"from the Modern Networks Building Intelligence Platform. "
            f"The analysis has identified {len(opps)} service opportunities, of which {high_opps} are high priority.",
            S["body"]))
        story.append(PageBreak())
        _pdf_building_intel(story, S, intel_uploads)
    else:
        _pdf_exec_summary(story, S, items, mode, opps, area_str, use_ai=use_ai)
        for upload in uploads:
            _pdf_asset_table(story, S, upload, mode)
        if enriched and flags:
            _pdf_rankings(story, S, items, mode, flags)
            _pdf_gap_analysis(story, S, items, mode, area_str=area_str, use_ai=use_ai)
            _pdf_prospect_flags(story, S, flags)

    _pdf_action_list(story, S, opps)
    story.append(PageBreak())
    _pdf_appendix(story, S, mode)

    fn = _footer(report_title or "Master Report")
    doc.build(story, onFirstPage=fn, onLaterPages=fn)
    return buf.getvalue()

# ── MAIN UI ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:#0b1829;padding:20px 28px;border-radius:10px;margin-bottom:24px">
<div style="font-size:10px;letter-spacing:2px;color:#f59e0b;font-family:monospace;margin-bottom:6px">INTERNAL USE ONLY</div>
<div style="font-size:24px;font-weight:800;color:#fff;margin-bottom:4px">Modern Networks</div>
<div style="font-size:14px;color:#64748b">Territory Intelligence Report Generator</div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### Report Settings")
    report_title = st.text_input("Report title", placeholder="e.g. London Retail Q2 2026")
    prepared_by  = st.text_input("Prepared by", placeholder="Your name")

    st.divider()

    mode = st.radio("Report type", ["🔬 Science & Innovation Parks", "🏬 Retail Property", "🏢 Building Intelligence"],
                    key="report_mode_radio")
    mode_key = "parks" if "Parks" in mode else "retail" if "Retail" in mode else "intel"

    st.divider()

    if mode_key in ("parks", "retail"):
        icon  = "🔬" if mode_key=="parks" else "🏬"
        label = "Science Parks" if mode_key=="parks" else "Retail Property"
        src   = "science_parks" if mode_key=="parks" else "retail_intelligence"
        st.markdown(f"### {icon} {label} Data")
        st.caption(f"Upload JSON export from the {'Science Parks' if mode_key=='parks' else 'Retail Property'} Intelligence app.")
        files = st.file_uploader(f"{label} JSON", type=["json"], accept_multiple_files=True, key="main_uploader", label_visibility="collapsed")
        if files:
            for f in files:
                data = parse_upload(f)
                if data and data.get("source_app") == src:
                    if "assets" in data and "parks" not in data:
                        data["parks"] = data.pop("assets")
                    already = any(u.get("exported_at")==data.get("exported_at") and u.get("area_label")==data.get("area_label")
                                  for u in st.session_state.uploads)
                    if not already:
                        st.session_state.uploads.append(data)
                        st.success(f"Loaded: {f.name}")
                elif data:
                    st.warning(f"{f.name} is not a {label} export.")

        if st.session_state.uploads:
            for i, u in enumerate(st.session_state.uploads):
                c1, c2 = st.columns([4,1])
                noun = "assets" if mode_key=="retail" else "parks"
                c1.markdown(f"<span class='source-badge {'badge-retail' if mode_key=='retail' else 'badge-parks'}'>{label}</span> "
                            f"{u.get('area_label','')} — {len(u.get('parks',[]))} {noun}", unsafe_allow_html=True)
                if c2.button("✕", key=f"del_{i}"):
                    st.session_state.uploads.pop(i)
                    st.rerun()

    else:  # Building Intelligence
        st.markdown("### 🏢 Building Intelligence Data")
        st.caption("Upload JSON export from the Building Intelligence Platform (Saved Briefings tab).")
        intel_files = st.file_uploader("Building Intelligence JSON", type=["json"], accept_multiple_files=True, key="intel_uploader", label_visibility="collapsed")
        if intel_files:
            for f in intel_files:
                data = parse_upload(f)
                if data and data.get("source_app") == "building_intelligence":
                    already = any(u.get("exported_at")==data.get("exported_at") for u in st.session_state.uploads)
                    if not already:
                        st.session_state.uploads.append(data)
                        st.success(f"Loaded: {f.name}")
                elif data:
                    st.warning(f"{f.name} is not a Building Intelligence export.")

        if st.session_state.uploads:
            for i, u in enumerate(st.session_state.uploads):
                c1, c2 = st.columns([4,1])
                n = len(u.get("briefings",[]))
                c1.markdown(f"<span class='source-badge badge-intel'>Building Intelligence</span> {n} briefing{'s' if n!=1 else ''}", unsafe_allow_html=True)
                if c2.button("✕", key=f"del_{i}"):
                    st.session_state.uploads.pop(i)
                    st.rerun()

with col2:
    st.markdown("### Report Preview")

    if not st.session_state.uploads:
        st.markdown("""
        <div style="background:#fff;border:2px dashed #e2e8f0;border-radius:10px;padding:48px;text-align:center;color:#94a3b8">
        <div style="font-size:44px;margin-bottom:14px">📊</div>
        <div style="font-size:15px;font-weight:600;color:#64748b;margin-bottom:8px">No data loaded yet</div>
        <div style="font-size:13px;line-height:1.9">Select a report type and upload a JSON export using the panel on the left.</div>
        </div>""", unsafe_allow_html=True)
    else:
        uploads     = st.session_state.uploads
        intel_ups   = [u for u in uploads if u.get("source_app")=="building_intelligence"]
        asset_ups   = [u for u in uploads if u.get("source_app") != "building_intelligence"]
        items       = _get_items(asset_ups)
        enriched    = any(p.get("epc") or p.get("companies") or p.get("flood_risk") for p in items)

        st.markdown("#### This report will contain:")
        sections = ["Cover page with summary statistics", "Executive summary"]

        if mode_key in ("parks","retail"):
            if enriched:
                sections += ["Territory rankings — opportunity score","Gap analysis narrative","Priority prospects with rationale"]
            for u in asset_ups:
                noun = "assets" if mode_key=="retail" else "parks"
                sections.append(f"{u.get('area_label','')} — {len(u.get('parks',[]))} {noun}")
                for p in u.get("parks",[])[:4]:
                    st.caption(f"  · {p.get('name','')} ({p.get('postcode','')})")
                extra = len(u.get("parks",[])) - 4
                if extra > 0: st.caption(f"  · +{extra} more")
        else:
            for u in intel_ups:
                n = len(u.get("briefings",[]))
                sections.append(f"Building assessments — {n} properties")
                for b in u.get("briefings",[])[:4]:
                    st.caption(f"  · {b.get('company','') or b.get('postcode','')} — {b.get('score',0)}/100")

        sections.append("Priority action list")
        sections.append("Data sources appendix")

        for i, s in enumerate(sections, 1):
            st.markdown(f"{i}. {s}")

        st.divider()

        use_ai = st.toggle("✨ AI-powered narrative (executive summary & gap analysis)",
                           value=True,
                           help="Uses Claude AI to write the executive summary and gap analysis in natural language. Falls back to template text if unavailable.")

        if st.button("⬇ Generate Report", type="primary", use_container_width=True):
            if not report_title:
                st.warning("Please enter a report title first.")
            else:
                spinner_msg = "Building report with AI narrative…" if use_ai else "Building report…"
                with st.spinner(spinner_msg):
                    pdf_bytes = generate_pdf(
                        asset_ups, intel_ups, mode_key,
                        report_title, prepared_by or "MN Staff",
                        use_ai=use_ai
                    )
                safe = report_title.replace(" ","-").replace("/","-")
                st.download_button("⬇ Download Report", data=pdf_bytes,
                    file_name=f"MN-{mode_key.title()}-{safe}-{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf", use_container_width=True)
                st.success("Report generated.")
