"""
Nepal Agriculture Minister Presentation Generator
Reads two PDF files and creates a structured PowerPoint presentation (title slide
+ 27 content slides = 28 total).
Output: Nepal_Agriculture_Minister_Presentation.pptx
"""

import os
import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# ---------------------------------------------------------------------------
# 1. PDF extraction helpers
# ---------------------------------------------------------------------------

def extract_pdf_text(pdf_path):
    """Return full text extracted from a PDF file."""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


# ---------------------------------------------------------------------------
# 2. Slide content (derived from both PDFs)
# ---------------------------------------------------------------------------

SLIDES = [
    # ── SECTION 1: Opening – Crisis & Urgency ──────────────────────────────
    {
        "section": "OPENING",
        "title": "Nepal's Agricultural Crisis: A Nation at a Crossroads",
        "bullets": [
            "Agriculture contributes ~24% of GDP yet operates at half its productive potential",
            "Over 60% of the population depends on farming — yet food imports are rising every year",
            "730,000–910,000 agricultural workers leave Nepal annually for foreign employment",
            "NPR 27.95 billion (49% of the ag budget) spent on fertiliser subsidies reaching only 30–40% of farmers",
            "Without urgent structural reform, the crisis will deepen and become irreversible by 2030",
        ],
        "notes": (
            "Open with the core paradox: Nepal is rich in land and water, yet its agriculture is failing. "
            "Stress the urgency — not a slow decline, but an accelerating structural crisis. "
            "The budget is misallocated; the solutions exist. The question is political will."
        ),
    },
    {
        "section": "OPENING",
        "title": "Scale of the Crisis: Key Numbers",
        "bullets": [
            "Average farm holding: 0.6 ha — below every viable mechanisation threshold",
            "Only 36% of agricultural land has year-round irrigation",
            "Fertiliser use: 67.4 kg/ha vs 163.5 kg/ha in India (2–4× yield gap)",
            "Post-harvest losses of 30–40% for fruits and vegetables",
            "Agricultural contribution to GDP growth: ~1%/yr vs 4–5% potential",
        ],
        "notes": (
            "Ground the audience in hard data. Each number represents a policy failure. "
            "The 0.6 ha average holding makes mechanisation mathematically impossible. "
            "The fertiliser gap alone explains why Nepal produces only ~50% of its yield potential."
        ),
    },
    # ── SECTION 2: Root Causes ─────────────────────────────────────────────
    {
        "section": "ROOT CAUSES",
        "title": "Root Cause 1: Structural Land Fragmentation",
        "bullets": [
            "Mandatory equal inheritance law subdivides land every generation",
            "Average holding has shrunk to 0.6 ha — mechanisation requires minimum 2–3 ha",
            "~50% of agricultural land lacks formal documentation or title",
            "32% of hill cultivated land lies fallow due to out-migration",
            "No functional land market exists to allow voluntary consolidation",
        ],
        "notes": (
            "Land fragmentation is the foundational constraint — it makes scale, mechanisation, "
            "and commercial farming structurally impossible. The inheritance law is the root legislative cause. "
            "Without fixing the land system, almost every other intervention is limited in impact."
        ),
    },
    {
        "section": "ROOT CAUSES",
        "title": "Root Cause 2: The Irrigation Crisis",
        "bullets": [
            "Only 36% of agricultural land has year-round irrigation",
            "Nepal receives 1,600 mm annual rainfall — 80% concentrated in 4 monsoon months",
            "Canal irrigation efficiency critically low; farmer-managed systems are deteriorating",
            "Only 4.7% of available groundwater is utilised",
            "Federalisation created governance gaps — no single tier is accountable",
        ],
        "notes": (
            "Nepal is paradoxically water-rich yet irrigation-poor. "
            "The problem is not water availability but water storage and governance. "
            "Expanding irrigation from 36% to 65% is the single highest-return investment available — "
            "it directly unlocks year-round cropping and breaks the monsoon-only yield ceiling."
        ),
    },
    {
        "section": "ROOT CAUSES",
        "title": "Root Causes 3–6: Inputs, Climate, Labour & Markets",
        "bullets": [
            "Inputs: 100% of chemical fertiliser imported; MOU with India expired March 2026",
            "Climate: each 1°C rise cuts maize yields by 3.4%, wheat by 2.4%; 1.5–2% GDP lost annually to disasters",
            "Labour exodus: 730,000+ workers leave annually — women and elderly left to farm",
            "Markets: farmers receive only 30–40% of retail price; value chains 7–9 stages long",
            "Infrastructure isolation forces distress sales within 2–3 days of harvest",
        ],
        "notes": (
            "These four root causes form interlocking constraints. "
            "The fertiliser supply gap (post-India MOU expiry) is an immediate crisis. "
            "Climate shocks are accelerating. Labour shortage is self-reinforcing. "
            "Market failure means even increased production does not translate to farmer income."
        ),
    },
    {
        "section": "ROOT CAUSES",
        "title": "Root Causes 7–10: Finance, Human Capital, Soil & Governance",
        "bullets": [
            "Finance: banks charge 18–24% interest for agricultural loans; only 6.7% of credit reaches pre-harvest",
            "Human capital: 1 extension technician per 1,500 farmers (global standard: 1 per 400)",
            "Soil health: nutrient mining, erosion of 10–15 t/ha/yr in hills, organic matter declining",
            "Governance: three-tier federal confusion; no single authority accountable for key functions",
            "Institutional weakness means policies exist on paper but fail at implementation",
        ],
        "notes": (
            "The finance trap means commercialisation requires capital that is structurally inaccessible. "
            "The extension gap means farmers cannot adopt new technologies even when they exist. "
            "Soil degradation is a slow-moving crisis that will cut long-term productivity irreversibly. "
            "Governance failure is the meta-problem — even well-funded programmes collapse in execution."
        ),
    },
    # ── SECTION 3: Vicious Cycles ──────────────────────────────────────────
    {
        "section": "VICIOUS CYCLES",
        "title": "Vicious Cycle 1: The Productivity Trap",
        "bullets": [
            "Small holdings + low inputs + poor roads = low productivity",
            "Low productivity = low income = poverty",
            "Poverty + 3–10× international wage gap = rational decision to migrate",
            "Migration = labour shortage = 32% of hill land lies fallow",
            "Fallow land = less food = deeper poverty — cycle repeats and worsens",
        ],
        "notes": (
            "This is the core self-reinforcing loop. Each element makes the others worse. "
            "Breaking this cycle requires simultaneous action on land, irrigation, markets, and finance — "
            "not sequential, one-at-a-time reforms. The cycle will not unwind from a single intervention."
        ),
    },
    {
        "section": "VICIOUS CYCLES",
        "title": "Vicious Cycle 2: Financial Exclusion Loop",
        "bullets": [
            "No collateral (undocumented land) = no formal bank loans",
            "No loans = no investment in seeds, irrigation, or technology",
            "No investment = low yields and income",
            "Low income = unable to save or invest in next season",
            "Subsistence trap: commercialisation requires capital that is permanently out of reach",
        ],
        "notes": (
            "Financial exclusion is both a cause and a consequence of poverty. "
            "The collateral trap is structural — undocumented land cannot be used as security. "
            "A Credit Guarantee Fund breaks this cycle by de-risking lending, enabling banks to lend "
            "without requiring land titles as collateral."
        ),
    },
    {
        "section": "VICIOUS CYCLES",
        "title": "Vicious Cycle 3: Irrigation–Climate–Adaptation Failure",
        "bullets": [
            "Limited irrigation forces farmers to rely on unpredictable monsoons",
            "Climate change makes monsoons more erratic — planning horizon collapses",
            "Cannot plan = cannot adopt multi-year climate-smart practices",
            "Adaptation requires credit → credit is unavailable (Cycle 2)",
            "Result: farmers locked into high-risk, low-investment farming with no escape",
        ],
        "notes": (
            "This cycle shows why isolated interventions fail. "
            "Providing drought-resistant seeds without irrigation does not help in a drought. "
            "Providing irrigation without credit means farmers cannot adopt complementary inputs. "
            "Climate adaptation is a systems challenge, not a technology challenge."
        ),
    },
    # ── SECTION 4: Why Current System Fails ───────────────────────────────
    {
        "section": "SYSTEM FAILURE",
        "title": "Why the Current System Fails: Budget Misallocation",
        "bullets": [
            "49% of the ag budget (NPR 27.95 Bn) goes to fertiliser subsidies reaching only 30–40% of farmers",
            "Irrigation receives only 5% of budget — despite being the highest-ROI investment",
            "Agricultural credit guarantee: only 1% of budget",
            "Extension and human capital: only 4% of budget",
            "The budget structure perpetuates dependency rather than building productive capacity",
        ],
        "notes": (
            "The current budget is not underfunded — it is misallocated. "
            "Half the budget subsidises a product that reaches less than half the farmers. "
            "Meanwhile, the investments with the highest long-term returns — irrigation, credit, extension — "
            "are chronically starved. This is the central structural reform required."
        ),
    },
    {
        "section": "SYSTEM FAILURE",
        "title": "Why the Current System Fails: Governance & Institutions",
        "bullets": [
            "Federal devolution left agricultural functions split with no tier accountable",
            "Weak farmer cooperatives — low bargaining power, no collective market access",
            "Policy-to-practice gap: policies are designed in Kathmandu, never reach farm level",
            "Extension system top-down and out of date — not farmer-problem driven",
            "Research locked in English journals; never translated to actionable farmer advice",
        ],
        "notes": (
            "Even well-designed policies fail because institutions cannot deliver them. "
            "Federalisation created confusion rather than devolution of real authority. "
            "The extension collapse means innovations sit in research stations, not on farms. "
            "Institutional reform is not optional — it is the precondition for everything else."
        ),
    },
    # ── SECTION 5: 8 Reform Pillars ───────────────────────────────────────
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 1: Solve Land Fragmentation | NPR 18 Billion",
        "bullets": [
            "Voluntary Land Pooling Programme: NPR 80,000/ha grant to cooperatives pooling 10+ contiguous hectares",
            "Digital Land Documentation Mission: formalise 50% undocumented land within 3 years",
            "Inheritance Law Amendment: allow family land trusts — no physical subdivision below 0.5 ha",
            "Government-backed Land Bank to purchase and redeploy fragmented holdings",
            "Target: 100,000+ ha pooled by Year 3; hill fallow land reduced from 32% to 12%",
        ],
        "notes": (
            "Land consolidation is the foundational reform — without it, mechanisation is impossible. "
            "The approach is voluntary, not forced. Incentives replace compulsion. "
            "The inheritance law amendment requires parliamentary action but has the highest long-term impact. "
            "Land documentation unlocks collateral, enabling all other financial interventions."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 2: Solve the Irrigation Crisis | NPR 85 Billion",
        "bullets": [
            "National Small-Scale Irrigation Programme: prioritise 5–500 ha schemes over mega-projects",
            "200,000 farm ponds + 500 check dams for water harvesting and dry-season storage",
            "National Groundwater Management Act: assign clear authority across federal tiers",
            "Target: expand irrigated area from 36% to 65% by 2030",
            "Highest ROI investment in the entire roadmap — every rupee multiplied across all crops",
        ],
        "notes": (
            "Irrigation is Nepal's single most transformative investment. "
            "Small and medium schemes are faster to build, cheaper per hectare, and locally manageable. "
            "Farm ponds provide individual-level water security even where canals cannot reach. "
            "The groundwater law fixes the governance vacuum created by federalisation."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 3: Fix Input Underutilisation | NPR 34 Billion",
        "bullets": [
            "Smart Fertiliser Cards: replace blanket subsidy with targeted digital delivery to verified farmers",
            "Emergency fertiliser MOU with China and Middle East suppliers (India MOU expired March 2026)",
            "National Seed System: 1,000 community seed banks; 10+ new climate-adapted varieties per year",
            "Smart Subsidy Reform: redirect NPR 14 Bn saved from targeting to capital investments",
            "Target: fertiliser use from 67.4 to 140+ kg/ha; productivity at 90%+ of potential",
        ],
        "notes": (
            "The fertiliser crisis is immediate — the India MOU expiry threatens the 2026 Kharif season. "
            "The Smart Card system ensures the same total subsidy reaches the right farmers, not elite captors. "
            "Seed system reform is as critical as fertiliser — most farmers use degraded saved seeds. "
            "The NPR 14 Bn saved from better targeting funds other reform pillars."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 4: Transform Market Systems | NPR 28 Billion",
        "bullets": [
            "77 District Aggregation Centres: cold storage, grading, sorting, price terminals (one per district)",
            "Value chain compression: reduce 7–9 stage chains to 3–4 stages using aggregation cooperatives",
            "Agricultural Export Promotion Board: target NPR 180 Bn exports by 2030 (from NPR 60 Bn today)",
            "Domestic processing for cardamom and ginger: capture 25% of retail price (up from 8–10%)",
            "National Price Information System: real-time wholesale prices via SMS and free app for all farmers",
        ],
        "notes": (
            "Farmers currently receive only 30–40% of retail price. Market reform doubles farm incomes "
            "without changing a single farming practice. Cold storage eliminates distress selling. "
            "Nepal is the world's largest cardamom producer — processing domestically can triple export earnings. "
            "Price information destroys information asymmetry that enables middlemen exploitation."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 5: Transform Agricultural Finance | NPR 35 Billion",
        "bullets": [
            "Agricultural Credit Guarantee Fund (ACGF): NPR 20 Bn seed capital unlocks NPR 100 Bn private lending",
            "Reduce effective agricultural lending rate from 18–24% to 8–10%",
            "Warehouse Receipt System: farmers use stored grain as collateral — eliminates forced sale at harvest",
            "ADBL digitalisation: mobile banking for 100 new branches in hill/mountain districts",
            "Target: formal credit access from 6–7% to 35% of farmers by 2030",
        ],
        "notes": (
            "The Credit Guarantee Fund is the most powerful financial intervention in the roadmap. "
            "By de-risking lending, it unlocks NPR 100 Bn in private bank capital with just NPR 20 Bn public seed capital. "
            "The warehouse receipt system solves the harvest-time price collapse — farmers store, not sell at panic prices. "
            "Every other pillar is more effective when farmers have access to working capital."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 6: Extension & Human Capital | NPR 18 Billion",
        "bullets": [
            "Hire 4,000 Junior Agriculture Extension Officers (JAEOs) in Year 1; reach 6,000 by Year 3",
            "Improve ratio from 1:1,500 farmers to 1:500 (global standard)",
            "7 Provincial Vocational Agriculture Centres for practical, hands-on training",
            "National Agricultural Knowledge Platform (NAKP): AI crop diagnosis, market prices, weather via SMS",
            "Farmer Field Schools: 100,000+ farmers trained in modern practices by Year 2",
        ],
        "notes": (
            "Technology and money do not reach farmers without people to deliver them. "
            "The extension collapse is why innovations sit in labs, not on farms. "
            "Recruiting 4,000 JAEOs in Year 1 is the highest-priority institutional action. "
            "The digital advisory platform provides scale — every farmer with a mobile phone gets expert advice."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 7: Climate Resilience & Soil Health | NPR 26 Billion",
        "bullets": [
            "Mandatory climate vulnerability mapping for all 77 districts by Year 1",
            "Crop zone rebalancing: flood-tolerant varieties in flood zones; drought-tolerant in dry zones",
            "Mobile soil testing labs in all 77 districts; soil health cards for every farmer",
            "Hill terrace rehabilitation: repair and restore 50,000+ ha of eroding terraces",
            "Conservation agriculture demonstrations: minimum tillage, cover crops, organic matter restoration",
        ],
        "notes": (
            "Climate resilience is not optional — it is the precondition for long-term productivity. "
            "Soil degradation is the slow-moving crisis that will cut productive capacity irreversibly. "
            "Hill terraces are centuries-old infrastructure; their collapse is both an environmental and food security threat. "
            "Every degree of warming costs 2–3% of yield — adaptation investment has guaranteed returns."
        ),
    },
    {
        "section": "REFORM PILLARS",
        "title": "Pillar 8: Governance & Institutional Reform | NPR 4 Billion",
        "bullets": [
            "Agricultural Functions Clarity Act: legally assign functions to each government tier",
            "National Agricultural Coordination Council: inter-tier body with real authority and teeth",
            "Programme consolidation: merge overlapping schemes; create single beneficiary registry",
            "Research-to-Farm Pipeline: mandate NARC to publish all findings in Nepali within 12 months",
            "Third-party independent monitoring with public dashboards; red-flag review if <70% execution",
        ],
        "notes": (
            "Governance reform costs the least but enables everything else. "
            "Without clear legal authority, every agricultural programme will continue to fall into the cracks of federalisation. "
            "The beneficiary registry eliminates duplicate subsidies and enables targeted delivery. "
            "Independent monitoring ensures accountability — the roadmap is only as strong as its enforcement."
        ),
    },
    # ── SECTION 6: Financial Model ─────────────────────────────────────────
    {
        "section": "FINANCIAL MODEL",
        "title": "Financial Model: NPR 340 Billion Five-Year Investment",
        "bullets": [
            "Total investment: NPR 340 Bn over 5 years (NPR 55–75 Bn/year — within Investment Decade commitment)",
            "Fertiliser subsidy reform saves NPR 14 Bn/year: redirected to irrigation and credit",
            "World Bank/ADB concessional loans: NPR 60 Bn for large irrigation at 3%, 25-year terms",
            "Agricultural Investment Bond: NPR 200 Bn at 8% — marketed to Nepali diaspora",
            "ACGF leverage: NPR 20 Bn public capital unlocks NPR 100 Bn private bank lending",
        ],
        "notes": (
            "This is not a request for more money — it is a request to spend existing money better. "
            "The subsidy reform alone frees NPR 14 Bn per year for productive investment. "
            "The bond and leverage model means each rupee of public investment mobilises 5–6 rupees in total capital. "
            "The total NPR 340 Bn is achievable within the Agriculture Investment Decade commitment already made."
        ),
    },
    {
        "section": "FINANCIAL MODEL",
        "title": "Budget Reallocation: From Dependency to Investment",
        "bullets": [
            "Fertiliser subsidies: reduce from 49% to 20% (targeted Smart Cards, same farmer benefit)",
            "Irrigation: increase from 5% to 20% of budget — single highest-return allocation",
            "Credit guarantee & finance: increase from 1% to 10%",
            "Market infrastructure & value chains: increase from 2% to 10%",
            "Extension & human capital: increase from 4% to 10%",
        ],
        "notes": (
            "The reallocation is the core structural reform. Total budget stays roughly the same. "
            "The fertiliser subsidy is not cut — it is made efficient. The same poor farmers get the same benefit "
            "via Smart Cards; elite capture (60–70% of current subsidy) is redirected to investment. "
            "This table should be the central exhibit in any parliamentary budget debate."
        ),
    },
    # ── SECTION 7: Implementation Timeline ────────────────────────────────
    {
        "section": "TIMELINE",
        "title": "Implementation Timeline: Three Phases to 2030",
        "bullets": [
            "Phase 1 (Year 1 — FY 2025/26): Legislation, institutions, and systems — NPR 55 Bn",
            "Phase 2 (Years 2–3 — FY 2026/27 to 2027/28): Capital deployment and scale-up — NPR 140 Bn",
            "Phase 3 (Years 4–5 — FY 2028/29 to 2029/30): Consolidation and sustainability — NPR 145 Bn",
            "100-Day Quick Wins: actions requiring no legislation or new budget — begin immediately",
            "Independent mid-term evaluation at Year 3 with mandatory public parliamentary response",
        ],
        "notes": (
            "Year 1 is about building the legal and institutional architecture — without it, capital cannot flow. "
            "Year 2-3 is the main capital deployment phase — irrigation, land pooling, JAEO recruitment. "
            "Year 4-5 is about scaling what works and fixing what doesn't. "
            "The 100-Day Quick Wins include emergency fertiliser MOU signing and NRB credit directive — "
            "these can begin within days of this roadmap's adoption."
        ),
    },
    {
        "section": "TIMELINE",
        "title": "100-Day Quick Wins: Immediate Action Required",
        "bullets": [
            "Sign emergency fertiliser MOU with China/Middle East suppliers (Day 1–30)",
            "Issue NRB directive: all A-class banks must allocate 15% portfolio to agriculture",
            "Launch Smart Fertiliser Card pilot in 3 districts immediately",
            "Announce Land Pooling Programme and open cooperative registration",
            "Deploy 500 JAEOs on emergency posting to districts with highest fallow land",
        ],
        "notes": (
            "These actions require no new legislation and no new budget appropriation. "
            "The fertiliser MOU is the most urgent — the Kharif 2026 season cannot wait for institutional reform. "
            "Quick wins build political momentum and demonstrate government seriousness before longer reforms take hold. "
            "Each action can begin within 100 days of this roadmap's adoption by the Minister."
        ),
    },
    # ── SECTION 8: Impact by 2030 ──────────────────────────────────────────
    {
        "section": "IMPACT 2030",
        "title": "Expected Impact by 2030: Transformational Targets",
        "bullets": [
            "Irrigated land: 36% → 65% of agricultural area",
            "Average paddy yield: 3.1 → 5.5 tonnes/ha (nearly double)",
            "Farmers with formal credit access: 6–7% → 35%",
            "Post-harvest losses (fruits/vegetables): 30–40% → 10–15%",
            "Agricultural exports: NPR 60 Bn/yr → NPR 180+ Bn/yr (3× increase)",
        ],
        "notes": (
            "These targets are conservative — they represent what is achievable with disciplined execution, "
            "not best-case scenarios. Each target is derived from the specific programme investments. "
            "Doubling paddy yield alone eliminates Nepal's food deficit and transforms the trade balance. "
            "The export target of NPR 180 Bn is achievable given Nepal's existing cardamom and ginger advantages."
        ),
    },
    {
        "section": "IMPACT 2030",
        "title": "Systemic Transformation by 2030",
        "bullets": [
            "Extension technician ratio: 1:1,500 → 1:500 (global standard achieved)",
            "Fallow hill land: 32% → 12% (land back in productive use)",
            "Fertiliser application: 67.4 → 140+ kg/ha (eliminating the yield gap with India)",
            "Agricultural GDP contribution: ~1%/yr → 4–5%/yr growth",
            "Rural income growth reverses out-migration and rebuilds generational farming families",
        ],
        "notes": (
            "By 2030, Nepal can transition from a food-deficit, import-dependent nation to a food-secure, "
            "export-capable agricultural economy. The extension ratio target signals full institutional rebuild. "
            "Fallow land restoration represents recapturing the productive capacity lost to migration. "
            "The income growth projection is what makes farming attractive to youth again — the ultimate test."
        ),
    },
    # ── SECTION 9: Strong Conclusion ──────────────────────────────────────
    {
        "section": "CONCLUSION",
        "title": "The Choice Before Parliament: Two Futures",
        "bullets": [
            "Current path: NPR 28 Bn/yr subsidy reaching 30–40% of farmers; 36% irrigation; 50% yield potential",
            "Reform path: NPR 14 Bn targeted subsidy reaching 90%+ of farmers; 65% irrigation; 3× yields by 2030",
            "The root causes are documented. The solutions are known. The budget exists.",
            "What is needed is political will, coordinated action, and a five-year commitment",
            "Every year of delay deepens the structural crisis and narrows the window for recovery",
        ],
        "notes": (
            "This is a binary choice — not between reform and no-reform, but between two fundamentally different futures. "
            "The roadmap does not ask for more money — it asks for better allocation of existing resources. "
            "The data, the solutions, and the financing model are all here. The decision rests with lawmakers. "
            "Emphasise: the crisis will not wait for the next election cycle."
        ),
    },
    {
        "section": "CONCLUSION",
        "title": "Call to Action: Transform Nepal's Agriculture Now",
        "bullets": [
            "Adopt this roadmap as the official Five-Year Agricultural Transformation Plan",
            "Pass the Agricultural Functions Clarity Act and Groundwater Management Act in Year 1",
            "Redirect fertiliser subsidy savings to irrigation, credit, and extension immediately",
            "Commit to independent annual monitoring with public accountability dashboards",
            "Nepal's farmers — 60% of the population — cannot afford another five years of stagnation",
        ],
        "notes": (
            "Close with a clear, specific ask. Not a call for further study or committees — "
            "a call for immediate adoption, legislation, and budget reallocation. "
            "The 60% of the population that depends on agriculture is watching. "
            "This is the Agriculture Minister's legacy: the generation that finally broke the vicious cycle."
        ),
    },
]


# ---------------------------------------------------------------------------
# 3. Design constants
# ---------------------------------------------------------------------------

# Colour palette (agriculture / policy — deep green and slate)
COLOUR_DARK_GREEN = RGBColor(0x1A, 0x53, 0x2E)   # slide title bar background
COLOUR_MID_GREEN = RGBColor(0x2E, 0x7D, 0x32)    # section label accent
COLOUR_LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)  # slide body background tint
COLOUR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOUR_DARK_TEXT = RGBColor(0x1A, 0x1A, 0x2E)    # body bullet text
COLOUR_ACCENT = RGBColor(0xF9, 0xA8, 0x25)       # warm amber for bullet marker

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

TITLE_BAR_H = Inches(1.35)
BODY_TOP = Inches(1.55)
BODY_LEFT = Inches(0.6)
BODY_W = Inches(12.13)
BODY_H = Inches(5.55)

FONT_TITLE = "Calibri"
FONT_BODY = "Calibri"


# ---------------------------------------------------------------------------
# 4. Slide builder utilities
# ---------------------------------------------------------------------------

def _set_bg(slide, colour):
    """Fill slide background with a solid colour."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = colour


def _add_rect(slide, left, top, width, height, fill_colour, line_colour=None):
    """Add a filled rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height,
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_colour
    if line_colour:
        shape.line.color.rgb = line_colour
    else:
        shape.line.fill.background()
    return shape


def _add_text_box(slide, text, left, top, width, height,
                  font_name, font_size, bold, colour, align=PP_ALIGN.LEFT,
                  word_wrap=True):
    """Add a text box with consistent formatting."""
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = word_wrap
    tf = txb.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = colour
    return txb


def build_slide(prs, slide_data, slide_index, total_slides):
    """Add a single content slide to the presentation."""
    slide_layout = prs.slide_layouts[6]  # blank layout
    slide = prs.slides.add_slide(slide_layout)
    slide.slide_width = SLIDE_W
    slide.slide_height = SLIDE_H

    # White background
    _set_bg(slide, COLOUR_WHITE)

    # Dark green title bar
    _add_rect(slide, Inches(0), Inches(0), SLIDE_W, TITLE_BAR_H, COLOUR_DARK_GREEN)

    # Section label (small, upper-left in title bar)
    section_label = slide_data.get("section", "")
    if section_label:
        _add_text_box(
            slide, section_label,
            Inches(0.55), Inches(0.08), Inches(5), Inches(0.35),
            FONT_TITLE, 9, False, COLOUR_ACCENT,
        )

    # Slide title
    _add_text_box(
        slide, slide_data["title"],
        Inches(0.55), Inches(0.28), Inches(11.5), Inches(0.95),
        FONT_TITLE, 24, True, COLOUR_WHITE,
    )

    # Slide number (bottom-right of title bar)
    slide_num_text = f"{slide_index + 1} / {total_slides}"
    _add_text_box(
        slide, slide_num_text,
        Inches(11.8), Inches(0.88), Inches(1.3), Inches(0.35),
        FONT_BODY, 9, False, RGBColor(0xB2, 0xDF, 0xDB),
        align=PP_ALIGN.RIGHT,
    )

    # Light green body background strip
    _add_rect(
        slide,
        Inches(0), TITLE_BAR_H, SLIDE_W, SLIDE_H - TITLE_BAR_H,
        COLOUR_LIGHT_GREEN,
    )

    # Bullet points
    bullets = slide_data.get("bullets", [])
    bullet_top = BODY_TOP
    bullet_h = Inches(0.72)
    indent = Inches(0.9)

    for bullet in bullets:
        # Amber bullet marker
        _add_text_box(
            slide, "▶",
            BODY_LEFT, bullet_top, Inches(0.4), bullet_h,
            FONT_BODY, 13, True, COLOUR_ACCENT,
        )
        # Bullet text
        _add_text_box(
            slide, bullet,
            indent, bullet_top, Inches(11.8), bullet_h,
            FONT_BODY, 16, False, COLOUR_DARK_TEXT,
        )
        bullet_top += bullet_h + Inches(0.12)

    # Green accent line at bottom
    _add_rect(
        slide,
        Inches(0), SLIDE_H - Inches(0.12), SLIDE_W, Inches(0.12),
        COLOUR_MID_GREEN,
    )

    # Speaker notes
    notes_text = slide_data.get("notes", "")
    if notes_text:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text

    return slide


def build_title_slide(prs, pdf1_text, pdf2_text):
    """Build a dedicated title slide."""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    _set_bg(slide, COLOUR_DARK_GREEN)

    # Top accent strip
    _add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.18), COLOUR_ACCENT)

    # Main title
    _add_text_box(
        slide,
        "NEPAL AGRICULTURAL TRANSFORMATION",
        Inches(0.9), Inches(1.2), Inches(11.5), Inches(1.0),
        FONT_TITLE, 34, True, COLOUR_WHITE,
        align=PP_ALIGN.CENTER,
    )

    # Subtitle
    _add_text_box(
        slide,
        "A Strategic Roadmap for Lawmakers & the Agriculture Minister",
        Inches(0.9), Inches(2.4), Inches(11.5), Inches(0.6),
        FONT_TITLE, 20, False, COLOUR_ACCENT,
        align=PP_ALIGN.CENTER,
    )

    # Divider line
    _add_rect(slide, Inches(3.5), Inches(3.15), Inches(6.33), Inches(0.05), COLOUR_ACCENT)

    # Key metrics row
    metrics = [
        ("NPR 340 Bn", "Five-Year Investment"),
        ("10", "Root Causes Addressed"),
        ("3×", "Yield Target by 2030"),
        ("9", "Reform Pillars"),
    ]
    metric_left = Inches(0.8)
    metric_w = Inches(2.9)
    for value, label in metrics:
        _add_text_box(
            slide, value,
            metric_left, Inches(3.5), metric_w, Inches(0.7),
            FONT_TITLE, 28, True, COLOUR_ACCENT,
            align=PP_ALIGN.CENTER,
        )
        _add_text_box(
            slide, label,
            metric_left, Inches(4.15), metric_w, Inches(0.45),
            FONT_TITLE, 12, False, COLOUR_WHITE,
            align=PP_ALIGN.CENTER,
        )
        metric_left += metric_w + Inches(0.1)

    # Footer
    _add_text_box(
        slide,
        "Ministry of Agriculture & Livestock Development | Government of Nepal | April 2026",
        Inches(0.5), Inches(6.8), Inches(12.33), Inches(0.4),
        FONT_TITLE, 11, False, RGBColor(0xB2, 0xDF, 0xDB),
        align=PP_ALIGN.CENTER,
    )

    # Bottom accent strip
    _add_rect(slide, Inches(0), SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), COLOUR_ACCENT)

    # Speaker notes
    notes_text = (
        "Welcome and introduction. This presentation is based on two key documents: "
        "the Deep-Dive Root-Cause Analysis and the Nepal Agricultural Transformation Master Roadmap 2025–2030. "
        "Nepal's agricultural sector is at a critical juncture. The solutions are known. "
        "The budget exists. What is required is political will and coordinated action."
    )
    slide.notes_slide.notes_text_frame.text = notes_text


# ---------------------------------------------------------------------------
# 5. Main assembly
# ---------------------------------------------------------------------------

def main():
    pdf1_path = "Nepal_Agriculture_Root_Problem_Analysis.pdf"
    pdf2_path = "Nepal_Agricultural_Transformation_Roadmap_2025-2030.pdf"

    # Extract text from both PDFs (used for title slide notes / audit)
    pdf1_text = extract_pdf_text(pdf1_path) if os.path.exists(pdf1_path) else ""
    pdf2_text = extract_pdf_text(pdf2_path) if os.path.exists(pdf2_path) else ""

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Title slide
    build_title_slide(prs, pdf1_text, pdf2_text)

    # Content slides
    total_slides = len(SLIDES) + 1  # +1 for title slide
    for idx, slide_data in enumerate(SLIDES):
        build_slide(prs, slide_data, idx + 1, total_slides)

    output_path = "Nepal_Agriculture_Minister_Presentation.pptx"
    prs.save(output_path)
    print(f"Presentation saved: {output_path}")
    print(f"Total slides: {total_slides}")


if __name__ == "__main__":
    main()
