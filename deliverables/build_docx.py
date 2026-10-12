import json
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

R = json.load(open("results.json"))
TEAL = RGBColor(0x0E, 0x5A, 0x57)
GREY = RGBColor(0x5B, 0x69, 0x63)
doc = Document()

sec = doc.sections[0]
sec.left_margin = sec.right_margin = Inches(1)
sec.top_margin = sec.bottom_margin = Inches(1)

st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(11)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(4)
st.paragraph_format.line_spacing = 1.08
for name, size in [("Heading 1", 16), ("Heading 2", 13), ("Heading 3", 11.5)]:
    h = doc.styles[name]
    h.font.name = "Calibri"; h.font.size = Pt(size); h.font.color.rgb = TEAL; h.font.bold = True
    h.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    h.paragraph_format.space_before = Pt(14 if name == "Heading 1" else 10)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True

def add_field(par, instr):
    r = par.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = "1"
    f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end")
    for el in (f1, it, f2, t, f3):
        r._r.append(el)

fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fr = fp.add_run("Caruthers Fair Band Showcase: Proposal and Planning Report  |  Page "); fr.font.size = Pt(8.5); fr.font.color.rgb = GREY
add_field(fp, "PAGE")
hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
hr = hp.add_run("Draft for review. Planning estimates, not quotes."); hr.font.size = Pt(8.5); hr.font.color.rgb = GREY

def H1(t): doc.add_heading(t, level=1)
def H2(t): doc.add_heading(t, level=2)
def H3(t): doc.add_heading(t, level=3)
def P(t, bold=False, italic=False, size=None, color=None, align=None, after=None):
    p = doc.add_paragraph()
    r = p.add_run(t); r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    if align is not None: p.alignment = align
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p
def BL(items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        if isinstance(it, tuple):
            r = p.add_run(it[0] + " "); r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(3)
def NL(items):
    for it in items:
        p = doc.add_paragraph(style="List Number")
        p.add_run(it); p.paragraph_format.space_after = Pt(3)
def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)
def T(header, rows, widths=None, size=9):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""
        r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(size); r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, "0E5A57")
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v)); r.font.size = Pt(size)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t
def FIG(path, caption, width=6.2):
    doc.add_picture(path, width=Inches(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(caption, italic=True, size=9, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
def PB(): doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
def money(x): return f"${x:,.0f}"
def num(x): return f"{x:,.0f}"

sp, cars, ld, sup, cost, rev, net = (R[k] for k in ["spectators", "cars", "lot_demand", "supply", "cost", "revenue", "net"])

# ---------- COVER ----------
for _ in range(6): doc.add_paragraph()
P("CARUTHERS FAIR BAND SHOWCASE", bold=True, size=13, color=GREY)
P("Proposal and Planning Report", bold=True, size=28, color=TEAL, after=4)
P("A free-gate band exhibition at the Caruthers District Fair: format, day plan, venues, traffic and parking, fire and emergency coordination, predictive outcomes, budget, risks, and implementation", size=13, color=GREY)
doc.add_paragraph()
P("Prepared for: Caruthers District Fair Board; Caruthers Unified School District; Caruthers High School band program; Fresno County Fire Protection District (Station 90)", size=10.5)
P("Prepared by: Proposal team, Caruthers Fair Band Showcase working group", size=10.5)
P("Status: Draft for review. Planning figures are estimates from stated assumptions. Items marked VERIFY need confirmation before any commitment.", size=10.5)
P("Companion file: Caruthers_Showcase_Planning_Model.xlsx (live formulas, Monte Carlo sample).", size=10.5)
PB()

# ---------- CONTENTS ----------
H1("Contents")
toc = ["Executive summary", "1. Purpose, scope and method", "2. The Caruthers District Fair: what is known",
       "3. The band circuit and how the showcase fits it", "4. Format: exhibition, preview or invitational",
       "5. Participants: field-show bands, parade bands and middle schools", "6. Day plan and schedule design",
       "7. Venues and facility requirements", "8. Transportation, trailers and parking",
       "9. Road closures, walking route and traffic control", "10. Fire station coordination and noise",
       "11. Safety, medical, heat and weather", "12. Predictive analysis", "13. Budget and funding",
       "14. Governance, roles and responsibilities", "15. Risk register", "16. Alternatives and contingency plans",
       "17. Implementation timeline", "18. Recommendations and decisions required",
       "Appendix A. Evidence register and sources", "Appendix B. Verification checklist and data gaps",
       "Appendix C. Draft band invitation letter", "Appendix D. Model method and workbook guide",
       "Appendix E. Sample schedule for twelve field-show bands"]
for t in toc:
    P(t, after=2)
PB()

# ---------- EXEC SUMMARY ----------
H1("Executive summary")
P("The Caruthers District Fair runs a free-gate, four-day fair each fall, usually Wednesday through Saturday in late September or early October. Its Saturday already carries the fair's parade, its largest crowd, and its evening programming. This report proposes a first annual Caruthers Fair Band Showcase on that Saturday: a no-score exhibition in which 10 to 16 high school field-show bands from the Central Valley and Southern California each receive a morning clinic, an afternoon preview performance with written judge comments, and a field-show slot. Middle school bands may take part in the parade and a short afternoon clinic, and they are not counted toward the field-show total.")
P("The proposal is positioned as a direct competitor to the Washington Union Preview of Champions, an early-season preview in the same pre-season window. The aim is to draw the bands that would otherwise attend that preview to the fair's Saturday, by offering what the preview may not: a free-gate crowd of the fair's size, a clinic, written judge feedback with no placement, an optional parade, and a live audience during the fair. The Western Band Association (WBA) historically opens its season on the first weekend of October and ends with championships before Thanksgiving, so the window is real and contested. The WBA's own policy manual places exhibition-class ensembles outside caption awards and gives them a participation award at every event, which supports the no-placement offer.")
P("The central finding is financial, and it changes the proposal. Under the planning assumptions in this report, the showcase costs about " + money(cost["50"]) + " at the median and generates about " + money(rev["50"]) + " in donations, sponsorship and band fees. The median net result is about " + money(net["50"]) + ". No simulated run broke even. The showcase therefore needs roughly $7,000 to $9,000 of fair-board, sponsor or school support, or a reduction in cost, before it can run without a deficit. The fair board is the logical first partner, because it already owns the parade, the grounds and the crowd.")
P("Parking is the second finding. Vehicles needing a space exceed usable spaces in about " + f"{R['shortfall_prob']*100:.0f}" + "% of simulated runs, with a median shortfall of roughly " + f"{R['shortfall_p50']:.0f}" + " vehicles. Because the football field is used for fair parking, the showcase needs a plan that does not rely on it. This report recommends staging band trailers in a separate lot, with a gate off Clemenceau Avenue, and shuttling visitors from the elementary lot. It also recommends that bands carry instruments only as far as needed, with the host supplying mallets and other shared equipment.")
P("Three further items need decisions before any invitation is sent:")
BL([("Fire station coordination.", "A courtesy request to reduce non-emergency siren use during the field show, with an explicit statement that emergency response takes precedence. The station listed for Caruthers in public records is at 2701 W. Tahoe Ave., not on Clemenceau Avenue, so the station the working group means must be confirmed (VERIFY)."),
    ("Parade and route.", "The parade is 10 a.m. on the fair's route. The fair's official site says Main Street, while the working group describes W. Tahoe Ave. from S. March Ave. to S. Raider Ave. The route must be confirmed by the fair board before any band packet is released."),
    ("Dates and attendance.", "The official fair site confirms the 2026 fair ran September 30 to October 3, 2026. No official attendance count is published. The model uses an organizer estimate of 60,000 to 80,000 visitors over four days from 2021 (VERIFY).")])
P("Recommendation: approve a one-year pilot for 2027 on the fair's Saturday, conditional on (1) a written fair-board co-sponsorship covering at least $8,000 of the budget, (2) confirmed use of the high school field and lights, (3) a parking and trailer plan approved by the district and the fair, and (4) a fire-station meeting. Section 18 sets out the full decision list.")
PB()

# ---------- 1 ----------
H1("1. Purpose, scope and method")
H2("1.1 Purpose")
P("This report sets out a proposal and implementation plan for the Caruthers Fair Band Showcase. It is written for the decision-makers who must approve or host the event: the fair board, the school district and high school administration, the band program, and the fire agency that serves the community. It also gives visiting band directors a clear picture of what they would be asked to commit to.")
H2("1.2 Scope")
BL(["The event day (the fair's Saturday) and the clinic and preview blocks that lead into it.",
    "The high school field and the associated facilities; the elementary school and its open field; the fairgrounds; and the streets between them.",
    "Transportation, parking, traffic control, emergency access and noise.",
    "Budget, funding sources, predictive outcomes and risks.",
    "Out of scope: the fair's own programming and operations, except where they interact with the showcase."])
H2("1.3 Method")
P("The working group combined four sources of evidence:")
NL(["Public sources on the fair (its official website, public news coverage, and organizer statements), and on the Caruthers community.",
    "The WBA 2025 Policy Manual, retrieved as a PDF and read for event procedures, credentials, critique, awards, weather protocols and copyright. It is the most detailed public rulebook we could retrieve for California high school marching band events.",
    "The CSBC 2016 field rules, the only CSBC rules document we could retrieve. It is dated, and the report treats it as indicative only (Section 3.3).",
    "A Monte Carlo model with 20,000 runs (triangular inputs) and a companion Excel workbook. Inputs that could not be sourced are labelled as assumptions. The model is described in Section 12 and Appendix D."])
P("Where a figure is an assumption, the report says so. Where a fact could not be verified, it is marked VERIFY. Appendix A lists every source; Appendix B lists the open verifications.")
H2("1.4 Terms used in this report")
T(["Term", "Meaning in this report"], [
    ["Showcase", "The Caruthers Fair Band Showcase as a whole."],
    ["Field-show band", "A high school band that takes part in the clinic, the preview and the field-show slot. Counted toward the 10 to 16 target."],
    ["Parade band", "Any band that marches in the 10 a.m. parade. Parade participation is optional for field-show bands and expected for middle schools."],
    ["Preview", "A short performance (about 8 minutes or less) followed by written judge comments."],
    ["Field-show slot", "A timed block of about 20 minutes per band, including up to 8 minutes of performance and 12 minutes of changeover and buffer."],
    ["Leeway", "Buffer time built into each band's day so that a late or early band does not affect the next band."],
    ["Exhibition", "A non-competitive performance with no placements. Judges still provide comments."],
    ["Free gate", "The fair does not charge admission. Revenue comes from concessions, carnival tickets, sponsors and donations."]], widths=[1.6, 4.9])
PB()

# ---------- 2 ----------
H1("2. The Caruthers District Fair: what is known")
H2("2.1 Identity and dates")
P("The fair is operated under the name Caruthers District Fair, at 13599 S. Raider Ave., Caruthers, CA 93609. Its official listed phone number is (559) 864-3768, and its postal address is PO Box 511, Caruthers, CA 93609. The fair's official site describes it as \"California's Largest and Longest-Running Free-Gate Fair.\" That is the fair's own claim. We found no independent ranking that confirms it, and this report does not repeat it as fact.")
P("Dates. The official site confirms the 2026 fair as Wednesday, September 30 through Saturday, October 3, 2026. Earlier years show a late-September pattern: September 25 to 28, 2024, and September 24 to 27, 2025, as reported by UC Agriculture and Natural Resources and a Fresno County 4-H listing. A 2018 public report implies September 26 to 29. The fair therefore runs Wednesday through Saturday, and the Saturday has been the closing day in each year we could verify. The 2026 dates moved into early October, so the 2027 dates must be confirmed with the board (VERIFY).")
H2("2.2 Public schedule and programming")
P("The official site, as retrieved for the 2026 fair, lists the following gate hours and programming. These figures are the fair's own and should be re-checked each year.")
T(["Day", "Gates (2026)", "Notable items (2026)"], [
    ["Wed, Sept. 30", "10:00 a.m. to 11:00 p.m.", "Queen coronation and opening ceremony at 6:30 p.m."],
    ["Thu, Oct. 1", "6:00 a.m. to 11:00 p.m.", "Free nightly entertainment on two stages begins (Thursday to Saturday)."],
    ["Fri, Oct. 2", "6:00 a.m. to 11:00 p.m.", "Carnival (Johnston Amusements), junior livestock, exhibits, motorcycle races."],
    ["Sat, Oct. 3", "6:00 a.m. to 12:00 a.m.", "Parade at 10:00 a.m. on Main Street; entertainment; nightly prize drawing."]],
    widths=[1.2, 1.9, 3.4])
P("Other programming on the official site includes horticulture and floriculture exhibits, home arts, a fair queen competition run by high school girls selling tickets over the summer, a nightly prize drawing whose final prize is a new Ford Mustang, and food and commercial booths. The site does not mention a rodeo or a band competition.")
H2("2.3 Parade")
P("The parade is held at 10:00 a.m. on the Saturday. The official 2026 listing describes it as featuring bands, horse groups, vintage cars and community floats, on Main Street. A travel listing (rove.me) describes a route beginning on South Marsh Avenue, continuing along West Tahoe Avenue, and ending at the fairgrounds, which is consistent with the working group's description of a short run from S. March Ave. on W. Tahoe Ave. to S. Raider Ave. The official site and the travel listing disagree on the street names. The route must be confirmed with the fair board before any band packet is issued (VERIFY).")
P("The parade is short in distance and begins promptly at 10:00 a.m. This matters for scheduling: a band that marches needs to be in place around 9:30 a.m., and once the parade ends, it has a limited window to reach the high school field for its clinic or slot.")
H2("2.4 Local context")
P("Caruthers is a small community. The 2020 census count reported on Wikipedia is 2,613, up from 2,497 in 2010. Other sites give different figures for different boundaries; the 2,613 figure is used here. A small population makes the fair a major local event, and it also means that local volunteers and facilities are shared across many activities.")
P("A 2024 news report describes a Caruthers High School parade for three championship teams that began at Henderson Road and Tahoe Avenue and ended at the fairgrounds. This is evidence that the high school already uses the parade route and the fairgrounds for public celebrations. We did not find a record of a band performing in the fair parade in recent years (VERIFY).")
H2("2.5 How the fair is managed")
P("The fair is run by a district fair board. The public sources we found do not describe the board's membership, its staffing, its budget, or its contracting approach. The working group therefore assumes that the fair operates with a volunteer board, a small paid staff, and contracted services for the carnival, sound and security. These are assumptions and must be confirmed with the board. The fair's general phone number and contact form are the entry points for a first conversation.")
PB()

# ---------- 3 ----------
H1("3. The band circuit and how the showcase fits it")
H2("3.1 The Western Band Association (WBA)")
P("The WBA is a nonprofit marching band association that runs competitions in California and neighboring states. Its 2025 Policy Manual (version 10, dated October 2025) sets out the rules the working group used as a benchmark. Key provisions relevant to the showcase are summarized below. Section numbers refer to the manual's own table of contents.")
T(["Topic", "WBA provision (2025 manual)", "Relevance to the showcase"], [
    ["Season timing", "The season historically starts on the first weekend of October and ends with championships the weekend before Thanksgiving.", "A late-September exhibition sits just before the circuit. An October-first date would be at the season's start."],
    ["Event schedule", "Standard events end with the last ensemble around 9:00 p.m., with awards around 9:30 p.m.", "Sets an industry norm for a full day. Our slots (Section 6) end near 8:30 p.m., leaving room for recognition."],
    ["Bands' day", "Arrive 2 to 3 hours before gate time; warm up 1 to 1.5 hours before gate; gate 15 minutes before performance.", "Used directly in the day plan: arrival for clinic, warm-up, then the slot."],
    ["Packets", "Event packets and schedules are finalized two weeks before each event; late updates posted by Wednesday or directly within 48 hours.", "Set the showcase's packet deadline for bands."],
    ["Credentials", "One director, one staff per 10 performers, one media, and one field pass per 10 performers.", "Field passes for each band are needed for the field and staff areas. A band of 100 needs about 10 field passes."],
    ["Critique", "Four-minute sessions with the visual and music panels; two ensembles at a time; registration required by September 1 for WBA events.", "Judge feedback for the preview can follow the same pattern: short, two-at-a-time sessions."],
    ["Awards", "Exhibition and festival-class ensembles receive a participation award at all events and do not qualify for caption awards.", "Supports the recognition model in Section 4.3: participation awards with no placement."],
    ["Recordings and cameras", "Cameras are not allowed on the field; WBA does not provide recordings; copyright must be cleared by the ensemble.", "The showcase must post the same rule and ask bands to clear music rights."],
    ["Weather", "Host checks the field for flooding and the warm-up and parking zones; WBA issues Plan A, B and C conditions for championships.", "A heat plan is needed for Central Valley September and October (Section 11)."],
    ["Ticketing", "Ticket sales are at the discretion of the host and WBA; a cashless ticketing service is available to standard hosts.", "The showcase is free-gate, so ticketing applies only to any seated or premium area the host chooses to charge for."],
    ["Host agreement", "A host agreement governs merchandise and other host responsibilities.", "The showcase needs a written host agreement with the fair and the school."]],
    widths=[1.2, 2.9, 2.4], size=8.5)
H2("3.2 Regional championships and the Kingsburg precedent")
P("The WBA runs regional championships in Northern, Central and Southern California, each held as a one-day event. The Kingsburg Viking Classic is listed as a WBA Regional Championship in the 2021 season, and a Kingsburg-hosted WBA event produced a Class A winner in November 2019. This is a Central Valley precedent for a small community hosting a WBA-related event on its own field. It is evidence that a community school can host a band event at a scale comparable to the showcase (VERIFY the current status of the Kingsburg event).")
H2("3.3 CSBC")
P("The California State Band Championships (CSBC) is a program of the Southwest Music Educators Association. Its fall circuit runs in the Central Valley and Southern California. Our search found a CSBC field rules document from 2016. It states that adjudication is by the Southwest Judges Network, that a one-point penalty applies for every 30 seconds or fraction thereof that a unit is over or under the required time, and that all music must be free of copyright infringement. The 2016 document is dated, and the working group did not retrieve the current CSBC rulebook or its stadium specifications. The CSBC website was not readable from our tools. Current CSBC rules must be obtained before the showcase's time and penalty rules are finalized (VERIFY).")
H2("3.4 The Washington Union Preview of Champions as the competitor")
P("The working group treats the Washington Union Preview of Champions as the event the showcase must out-draw. We could not verify its date, host, format, or fees, so the competitive analysis below is provisional. The name suggests a preview-style event in the same pre-season window, which is exactly where the showcase sits. Before any invitation is sent, the date must be confirmed, because a showcase on the same weekend would share the same bands and directors (VERIFY with Washington Union High School and the WBA).")
H2("3.5 Competitive positioning")
P("The showcase competes for the same bands and directors. Its advantages must be stated plainly in the invitation: the fair's large free-gate crowd, a parade and a fair Saturday for families, a clinic before the show, judges' written feedback, a participation plaque for every band, and no placement pressure. Its weaknesses are also plain: the Washington Union event may have an older track record, a sanctioned status, and an established judge panel. The showcase should not claim sanction or qualifier status it does not have.")
P("The working group also recognises the relationship cost. Directors who attend both events will notice a deliberate move to draw bands away. We recommend that the showcase director contact the Washington Union program first, state the intent openly, and offer a schedule that does not conflict with its date. A transparent approach protects the fair's standing in the band community and reduces the risk of a formal objection from the WBA.")
PB()

# ---------- 4 ----------
H1("4. Format: exhibition, preview or invitational")
H2("4.1 Three formats")
P("The working group considered three formats. They differ in how much judging and recognition they involve, and therefore in cost and in the commitments asked of bands.")
T(["Format", "What it is", "Strengths", "Weaknesses", "Fit for 2027"], [
    ["Exhibition only", "Bands perform for the crowd; no judging; no feedback.", "Lowest cost; simplest logistics.", "Little value to bands who want evaluation.", "Too thin for the first year."],
    ["Clinic and preview (recommended)", "Morning clinic, afternoon preview, field-show slot; written judge comments; participation recognition.", "Gives bands work with a clinician and feedback; no placement pressure; a clear identity.", "Needs judges and clinicians; moderate cost.", "Best fit."],
    ["Invitational", "Scored competition with judges' sheets, rankings and evening awards.", "Attracts more bands; looks like a competition.", "High cost; needs judging panels and sanction questions; invites a placement contest the early season does not need.", "Year two, if year one proves demand."]],
    widths=[1.2, 1.7, 1.5, 1.6, 1.4], size=8.5)
H2("4.2 Recommendation")
P("Adopt the clinic-and-preview format in 2027, with a field-show slot as the centrepiece. Judges provide written comments on the preview and the field show. The event is an exhibition, so no band is ranked. Revisit an invitational format in 2028 if at least 12 bands return and at least three new bands ask for placements.")
H2("4.3 Recognition")
P("Bands will want something to take home. The recommended approach is recognition without ranking, in line with the WBA treatment of exhibition-class ensembles:")
BL(["A participation plaque for every field-show band, presented at the close of the day.",
    "Written judge comments for each band, delivered the same day.",
    "Optional \"Outstanding\" recognitions chosen by the judges, such as Outstanding Music, Outstanding Visual or Outstanding Showmanship. These are announced at the end of the day without placements or scores.",
    "No public placements, scores or caption awards in year one. If bands ask for placements, the question is revisited for 2028."])
P("The working group considered the risk that bands and parents will want a trophy. The plaque and the recognition are designed to meet that need without creating a losing side. Parents are also likely to value a clear, public recognition of their child's band.")
H2("4.4 Night performance")
P("Night performance under field lights is a separate decision, to be made by the school district and each band director. Sunset in early October in Caruthers is around 7 p.m. (VERIFY against the local almanac). A field show that starts at 4:30 p.m. and runs to 8:30 p.m. will be in the dark for most of the second half, so the field must be lit. If the lights are not available, the schedule must end before sunset, which means starting no later than about 3:30 p.m. and shortening the slots.")
PB()

# ---------- 5 ----------
H1("5. Participants: field-show bands, parade bands and middle schools")
H2("5.1 Field-show bands")
P("The target is 10 to 16 high school field-show bands. The working group recommends 10 to 12 in 2027. Sixteen bands on one day push the field-show block past 9:30 p.m. (Section 6.3), which conflicts with lighting and with the WBA norm of ending near 9 p.m. The target is therefore a range that the schedule can absorb, not a fixed number.")
P("Bands should be invited across classes. The WBA's two-tier scoring recognizes that small and large bands have different paradigms, and the showcase should similarly invite bands of different sizes. A mix of classes gives each band a comparison group and makes the day more useful for judges and clinicians.")
H2("5.2 Invitation criteria")
BL(["A confirmed high school marching band with a show that can be performed within the slot.",
    "A director who can commit to the clinic, the preview and the field-show slot, and who accepts that the event is an exhibition.",
    "Acceptance of the music copyright and recording rules in Section 3.1.",
    "Confirmed transportation, including whether the band will bring a trailer or bus and how many vehicles it needs.",
    "A named contact for the band's packet and call times."])
H2("5.3 Target regions and recruitment")
P("The working group's recruitment plan targets three areas. The first is the Central Valley, which is closest and cheapest for bands to reach. The second is Southern California, including Los Angeles County, where a band such as Gahr High School competed at the WBA Championships in Fresno in November 2025. The third is Riverside County, where a band such as Murrieta Valley could be invited if its calendar allows (VERIFY availability and current seasons for every band before contact).")
P("The working group's search did not produce verified 2026 season calendars or director contact details for these bands, so no contact has been made. Each invitation must be preceded by a current check of the band's season, its director, and its availability.")
T(["Area", "Candidate bands (VERIFY before contact)", "Basis for inclusion", "Travel"], [
    ["Central Valley (first)", "Clovis North; other Clovis-area programs; Kingsburg; Orange Cove; Visalia and Selma programs (list to be confirmed)", "Local to the fair; Clovis North reports a 4A state championship (year not stated).", "Under 90 minutes by bus."],
    ["Southern California", "Gahr High School (Cerritos); other WBA-affiliated Los Angeles County programs", "Gahr competed at WBA Championships in Fresno in November 2025.", "About 4 to 5 hours. Clinic-only or stipend option."],
    ["Riverside County", "Murrieta Valley and other Riverside County programs", "No verified result found; included as a candidate only.", "About 4 to 5 hours. Clinic-only or stipend option."]],
    widths=[1.2, 2.3, 2.0, 1.0], size=8.5)
H2("5.4 Middle schools")
P("Middle school bands are not part of the 10 to 16 field-show count. They take part in the 10 a.m. parade and in a shorter afternoon clinic after the parade. The clinic is short, about 30 minutes, and is scheduled after the parade so that the morning clinic can be reserved for field-show bands. Middle school bands do not perform a field show.")
H2("5.5 Parade participation for field-show bands")
P("Parade participation is encouraged for field-show bands but is optional. A band that skips the parade still takes the clinic and the field-show slot. A band that marches gets a short walk on the route and a clear time to return for its call. Bands that skip the parade have free time at the fairgrounds before their call time, as described in Section 6.")
PB()

# ---------- 6 ----------
H1("6. Day plan and schedule design")
H2("6.1 Principles")
BL([("Priority for field-show bands.", "The morning clinic block is reserved for field-show bands, because the clinic prepares the show and the show is the event's centrepiece."),
    ("Parade is optional and does not block the field show.", "A parade band returns in time for its call; a band that skips the parade keeps its clinic and slot."),
    ("Leeway is built in.", "Each band's slot has a buffer. Late arrival does not push the next band."),
    ("The heat is planned for.", "The field-show block starts late, around 4:30 to 5 p.m., to avoid the hottest hours."),
    ("The fair stays open.", "Families and bands can use the fair on Saturday before and after their slot.")])
H2("6.2 The day, in order")
T(["Time (approx.)", "Field-show bands", "Parade bands (middle school)", "Public and fair"], [
    ["6:30 to 8:00 a.m.", "Arrival at the elementary field; equipment unloaded to the staging area; warm-up on the field.", "Not yet on site (middle schools arrive by 9:00 a.m.).", "Fairgrounds open at 6:00 a.m. (2026 hours)."],
    ["8:00 to 10:00 a.m.", "Clinic blocks of 30 to 40 minutes per band, in the elementary music room or on the field.", "Assemble on W. Tahoe Ave. for the parade; warm-up at the elementary lot.", "Visitors arrive; parking through the elementary lot."],
    ["10:00 a.m.", "Field-show bands not marching; free time at the fairgrounds, returning by their call time.", "Parade: W. Tahoe Ave., March to Raider. Short, about one hour including assembly.", "Parade."],
    ["10:45 a.m. to 1:00 p.m.", "Bands that marched return to staging; lunch; free time at the fair; return to the high school for call times.", "Middle school bands dismiss to lunch; return for the afternoon clinic.", "Fair open; food and exhibits."],
    ["1:00 to 2:00 p.m.", "Preview and clinic follow-up as scheduled for bands that need it.", "Short afternoon clinic (about 30 minutes per band).", "Carnival, exhibits, entertainment."],
    ["3:30 p.m.", "Call times for the first field-show band; gate at 15 minutes before the slot.", "Middle school bands leave or stay as spectators.", "Fair continues."],
    ["4:30 to 8:30 p.m.", "Field-show slots, 20 minutes each, including up to 8 minutes of performance (preview) and buffer. Ten to 12 bands.", "Watch from the bleachers if they choose.", "Fair continues; entertainment on stages."],
    ["8:30 to 9:00 p.m.", "Program ends. Judges' comments delivered; plaques presented to bands before departure where possible.", "Dismissal.", "Fair open until its closing hour."]],
    widths=[1.1, 2.4, 1.8, 1.3], size=8)
P("Times are a planning sketch to be tested with the band directors, the high school and the fair board. They are not a confirmed schedule.", italic=True, size=9)
H2("6.3 Slot arithmetic")
P("Each field-show band receives a 20-minute slot: up to 8 minutes of performance and 12 minutes of changeover and buffer. The field-show block starts at 4:30 p.m. Ten bands occupy 200 minutes (3 hours 20 minutes). Twelve bands occupy 240 minutes (4 hours). Sixteen bands occupy 320 minutes (5 hours 20 minutes) and run past 9:30 p.m. The working group therefore recommends 10 to 12 bands in 2027, and it keeps 16 as the upper limit for a later year, when the lighting and the schedule have been tested.")
T(["Field-show bands", "Slot block (minutes)", "Start 4:30 p.m. ends", "Start 5:00 p.m. ends", "Within 9:00 p.m. (WBA norm)?"], [
    ["10", "200", "7:50 p.m.", "8:20 p.m.", "Yes, both"],
    ["12", "240", "8:30 p.m.", "9:00 p.m.", "Yes if start is 4:30 p.m."],
    ["14", "280", "9:10 p.m.", "9:40 p.m.", "No"],
    ["16", "320", "9:50 p.m.", "10:20 p.m.", "No"]],
    widths=[1.2, 1.3, 1.4, 1.4, 1.5], size=8.5)
H2("6.4 Leeway")
P("Leeway is the buffer built into each band's day. It works in three places. First, each band's slot includes 12 minutes of changeover and buffer. Second, the clinic and the preview each have a 10-minute gap, so a band that runs long does not delay the next clinician. Third, the 30-minute buffer at the end of the block allows late bands to perform without pushing the program past 9 p.m. A band that uses its buffer for its own changeover does not reduce another band's time.")
H2("6.5 Bands that stay")
P("Bands may remain after their slot to watch others. The working group recommends that bands sit together in the designated band seating so that the audience can see the performance. Their call time and departure time are set in advance, and they are free to leave after their slot once their equipment is loaded.")
PB()

# ---------- 7 ----------
H1("7. Venues and facility requirements")
H2("7.1 Sites")
T(["Site", "Use", "What is known", "What must be confirmed (VERIFY)"], [
    ["Caruthers Elementary and the open field beside it", "Warm-up, load-in, clinic blocks.", "Elementary school on the route between the fair and the high school. An open community field sits close by (location and ownership to be confirmed).", "Use hours on the Saturday; ownership; restrooms; power; parking capacity; insurance."],
    ["Caruthers High School field", "Preview and field-show slots.", "High school site adjacent to the fairgrounds area; located near the parade finish on S. Raider Ave.", "Lights; bleachers; press box and judge positions; field condition; restrooms; power; use fee; availability."],
    ["Fairgrounds, 13599 S. Raider Ave.", "Parade finish; family gathering; fair programming.", "Official address and contact; gate hours; carnival and entertainment on the grounds.", "Board approval; Saturday use; staging; shared restrooms; parking."],
    ["Football field across from the fairgrounds", "Used for fair parking on the Saturday (per the working group).", "Used for parking during the fair (per the working group).", "Availability for showcase parking; number of spaces; cost; agreement with the fair."]],
    widths=[1.5, 1.4, 2.0, 1.6], size=8.5)
H2("7.2 Field and stadium requirements")
P("The following requirements are a planning checklist. They are drawn from the WBA's credential, event-procedure and weather provisions and from the working group's operating needs. The CSBC's current stadium specifications could not be retrieved, so the checklist must be checked against that document once it is obtained (VERIFY).")
T(["Requirement", "Basis", "Status"], [
    ["Field lights for the period after about 7 p.m.", "Sunset in early October; WBA events run to about 9 p.m.", "Unknown (VERIFY)"],
    ["Bleacher seating; seated capacity counted for the model", "Model input; WBA hosts provide staff viewing areas.", "Unknown (VERIFY)"],
    ["Designated judges' and staff positions, with a clear view and no interference with spectators", "WBA credentials: staff viewing areas must not be affected by media or band activity.", "To plan"],
    ["Field-pass system: one pass per 10 performers, plus director and staff passes", "WBA credential rules.", "To plan"],
    ["Public address and a program announcer", "WBA standard event schedule; spiel sheet template.", "To source (rental)"],
    ["Restrooms, portable and permanent", "Modelled at 4 to 6 portable units.", "To source"],
    ["Water and shade for performers and audience", "Heat planning (Section 11).", "To plan"],
    ["First aid and medical standby", "Section 11.", "To plan"],
    ["Field condition check and flooding protocol", "WBA host checklist for field, warm-up and parking zones.", "To plan"],
    ["Power for sound, judges' equipment and the scoring laptop", "Standard event need.", "To confirm"]],
    widths=[2.6, 2.8, 1.1], size=8.5)
H2("7.3 Facility use agreement")
P("The school district must approve use of the field, lights and restrooms in writing. The agreement should set the use hours, the custodial responsibilities, the insurance requirement, the fee, the cleanup standard, and the procedure if the field is damaged. A draft agreement is an early deliverable in the implementation plan.")
PB()

# ---------- 8 ----------
H1("8. Transportation, trailers and parking")
H2("8.1 The problem")
P("The football field across from the fairgrounds is used for fair parking on the Saturday. That is the natural lot for spectators and it is also the closest to the high school. The working group recommends that the showcase not depend on it. Bands bring trailers, buses and cars. Visitors bring cars. The field-show audience is largely local, so it is likely to arrive by car, and it may be large.")
H2("8.2 Predicted demand and supply")
P("The Monte Carlo model estimates the number of vehicles that need a space at the high school and its lots, and compares it with usable spaces. The main results are below. Section 12 describes the model.")
T(["Measure", "P10", "Median (P50)", "P90"], [
    ["Vehicles needing space (spectator cars plus band vehicles)", num(ld["10"]), num(ld["50"]), num(ld["90"])],
    ["Usable parking spaces (football field net of the fair's share, plus two school lots)", num(sup["10"]), num(sup["50"]), num(sup["90"])],
    ["Spectator cars", num(cars["10"]), num(cars["50"]), num(cars["90"])]],
    widths=[3.7, 0.9, 1.0, 0.9], size=9)
P(f"Vehicles needing space exceed usable spaces in about {R['shortfall_prob']*100:.0f}% of simulated runs. The median shortfall is about {R['shortfall_p50']:.0f} vehicles, and the 90th-percentile shortfall is about {R['shortfall_p90']:.0f}. These figures are sensitive to the assumed fair share of the football field and to the unverified lot sizes, so they are a basis for planning, not a forecast.")
FIG("figures/fig_parking.png", "Figure 1. Simulated vehicles needing space against usable spaces (20,000 runs).")
H2("8.3 Options")
P("The working group compared five ways to handle the demand. Each option is evaluated on its feasibility, cost, risk and the alternative it leaves.")
T(["Option", "Description", "Advantages", "Disadvantages and risks", "Assessment"], [
    ["A. Share the football field", "Keep the fair's parking; showcase visitors use the field as well.", "No new arrangement.", "Not enough spaces in most runs; fair demand is high.", "Not recommended as the only plan."],
    ["B. Elementary lot and shuttle", "Visitors park at the elementary lot and walk or are shuttled to the high school.", "Keeps the football field for the fair; uses a school lot.", "Requires a shuttle and a walking route; school lot may be needed for staff.", "Recommended for visitors."],
    ["C. Trailer staging off Clemenceau Avenue", "A dedicated gate between the fire station area and Clemenceau Avenue, for band trailers and buses only.", "Keeps trailers off the fair's parking and away from the high school's frontage; gives a single staging point.", "Needs district and fair approval; needs a gate, signage, fencing and a traffic plan; fire lanes must stay clear.", "Recommended (subject to VERIFY of the location and the fire station, Section 10)."],
    ["D. Fairgrounds overflow", "Use the fairgrounds' own overflow parking with the fair's permission.", "Central location.", "Fair demand; cost to be agreed.", "Backup option."],
    ["E. Bands carry instruments only as far as needed", "Trailers stay at staging; the clinic uses the elementary room; mallets and shared equipment are supplied by the host.", "Less carrying; trailers stay put.", "Drums and heavy equipment still need to move; requires a secure storage room at the elementary school.", "Recommended, as a supporting measure."]],
    widths=[1.2, 1.6, 1.4, 1.6, 1.0], size=8)
H2("8.4 Recommendation")
P("Combine B, C and E. Visitors park at the elementary lot and walk or are shuttled to the high school. Band trailers and buses stage at a gate off Clemenceau Avenue, with an operated gate to keep traffic controlled. Bands carry instruments only from their trailers to the clinic or the field, and the host provides mallets and shared equipment. Option D is the backup. The football field remains the fair's parking.")
P("Before the plan is adopted, the working group must confirm the Clemenceau gate's location, ownership, width and turning radius for trailers, and must show that the fire lane stays clear (VERIFY). The gate off Clemenceau is best located between the fire station area and the street, but its exact position must be confirmed with the fire district, because the fire station that serves Caruthers is listed at 2701 W. Tahoe Ave. (Section 10).")
H2("8.5 Shuttle and walking")
P("A shuttle from the elementary lot to the fairgrounds and high school is the simplest way to handle visitors if the lot is used. The shuttle should run continuously from 2:00 p.m. to 9:00 p.m., with a stop at the high school field. If the shuttle is not available, the walking distance from the elementary lot to the high school should be measured and signed (VERIFY). Mobility-impaired visitors need a dedicated drop-off close to the field.")
PB()

# ---------- 9 ----------
H1("9. Road closures, walking route and traffic control")
H2("9.1 Routes")
P("The working group identified three routes. The first is the parade route, which runs 10 a.m. on W. Tahoe Ave. from S. March Ave. to S. Raider Ave. (Section 2.3, VERIFY). The second is the walking route for field-show bands from the elementary field to the high school, which the working group proposes runs along W. Sandy Ave. and S. Raider Ave. (VERIFY). The third is the route for trailers and buses from Clemenceau Avenue to the staging gate.")
P("The earlier red line drawn on the aerial map that the working group reviewed was not the parade route. It marked the school and fair locations and was used only to orient the locations. Street names and exact segments for the walking route must be confirmed by a field walk.")
H2("9.2 Proposed closures and controls")
T(["Location", "Proposed control", "When", "Notes"], [
    ["S. Raider Ave. between W. Sandy Ave. and W. Tahoe Ave.", "Closure to through traffic; local access by flagger.", "Walking window and the field-show block (about 4:00 to 9:00 p.m.).", "Full closure requires a county permit (VERIFY)."],
    ["W. Sandy Ave. (one block around the school frontage)", "Partial closure or lane control for the walking route.", "Walking windows only.", "Use flaggers at each crossing."],
    ["Crossings on the walking route", "Flaggers and cones; bands walk with an escort.", "As bands move.", "Crossing control is required whatever the closure."],
    ["Clemenceau Avenue gate", "Controlled gate for trailers and buses only.", "Arrival and departure windows.", "Fire lane stays clear; signage required."],
    ["Parade route, W. Tahoe Ave.", "Existing parade control by the fair.", "10 a.m. parade.", "The showcase relies on the fair's parade control."]],
    widths=[1.7, 1.9, 1.5, 1.4], size=8.5)
H2("9.3 Permits")
P("A street closure needs a permit from the agency that controls the road. The public sources we found did not include Fresno County's special event or encroachment rules. The City of Fresno requires a permit when an event closes streets, sidewalks or public spaces for any period of time, and requires notice to the police at least 48 hours ahead for parades of more than 50 people. These are the City of Fresno's rules, not the county's, and they are cited only as a model. The working group must confirm with Fresno County Public Works and Planning whether an encroachment permit or special event permit is required for each county road on the route, and must confirm the Sheriff's office's role (VERIFY).")
H2("9.4 Traffic control staffing")
P("The working group estimates four to six flaggers for about six hours on the day (Section 13). This covers the walking route crossings, the Raider Ave. control, and the Clemenceau gate. Flaggers must be trained and equipped. A traffic control plan should be prepared by a qualified provider and approved before the day.")
H2("9.5 Noise")
P("The football field across from the fairgrounds is used for parking and will generate vehicle noise during the field show. The working group recommends a sheriff's or traffic officer's presence during the field-show block, and recommends that horns and idling be addressed by signage and the flaggers. This is a request to the Sheriff's office and the fair, not a guarantee.")
PB()

# ---------- 10 ----------
H1("10. Fire station coordination and noise")
H2("10.1 The station")
P("Public records we retrieved list Fresno County Fire Station 90 at 2701 W. Tahoe Ave., Caruthers, CA 93609, with the public phone line 559-864-3211. The station building dates from 1993 and replaces an earlier station near the fairgrounds that opened in 1947, according to Fresno County Fire Battalion 15's public page. We did not find a Caruthers fire station on Clemenceau Avenue. The working group's proposal refers to a station between the fire station area and Clemenceau Avenue. That must be confirmed: either the station is on W. Tahoe Ave. (which is on the parade route's neighbourhood), or the working group means a different site (VERIFY).")
H2("10.2 The request")
P("The working group would like to ask the station whether it can keep non-emergency sirens quiet during the field-show block. The request should be framed as a courtesy, with a clear statement at the start of the request that emergency calls take precedence, and that the station decides what it can do. The request has three parts:")
NL(["For emergency responses, the station uses its normal procedures. Sirens may sound when the station is responding to an emergency, and the showcase does not ask the station to change this.",
    "For non-emergency movements, such as drills, training runs or moves between stations, the station is asked to use lights without sirens during the field-show block where it can.",
    "The station is asked to tell the showcase's coordinator of any planned non-emergency movement during the block, so that the coordinator can warn the audience and the flaggers."])
P("The working group should not promise that sirens will be silent. The fire agency's emergency protocols are its own decision, and California law gives authorized emergency vehicles the privilege to use sirens when responding (the Vehicle Code provision should be confirmed by the fire district, VERIFY).")
H2("10.3 Access and emergency routes")
P("Fire access to the high school field, the fairgrounds, the elementary school and the Clemenceau gate must stay clear at all times. The traffic control plan must show fire lanes, and the flaggers must be trained to clear them. A fire engine or ambulance should be able to reach the field without passing through the walking route.")
H2("10.4 Good-neighbour measures")
BL(["A named contact for the fire station and the showcase, with mobile numbers.",
    "A pre-event walk-through of the Clemenceau gate and the fire lane with the station captain.",
    "An invitation to the fire station to attend the event, with an identified seating area.",
    "A thank-you letter after the event, with the field-show schedule so the station knows when the show is."])
PB()

# ---------- 11 ----------
H1("11. Safety, medical, heat and weather")
H2("11.1 Heat")
P("Central Valley weather in late September and early October can be hot. A field-show block that starts at 4:30 p.m. avoids the hottest part of the day, but performers still wear full uniforms and carry instruments. The working group recommends the following: shade and water at staging and the field; a heat-illness plan with a trained first-aid lead; scheduled hydration breaks for bands; and a rule that bands may take a break between clinic and slot.")
H2("11.2 Medical")
P("A first-aid station should be staffed on the field and at staging. The working group recommends a licensed EMT or certified first-aid provider for the field-show block, and a plan to call for an ambulance with the nearest address (VERIFY the cost and the provider). The medical plan should name the nearest hospital and the route to it.")
H2("11.3 Weather")
P("The WBA's weather procedures distinguish between conditions that leave the event on schedule, conditions that require adjustments, and conditions that make the day unsuitable. The showcase adopts the same three categories. For rain or flooding, the host checks the field, the warm-up area and the parking zones the morning before and the day of, and reports to the WBA where it is involved. The showcase's director of operations decides whether the field show is delayed, shortened, moved to the next day, or cancelled, in consultation with the band directors. A cancellation decision must be made by the day before where possible.")
H2("11.4 Security and crowd")
P("The fair's own security and the Sheriff's office should be consulted. The working group recommends a visible presence at the high school field during the field-show block, a bag policy consistent with the school's, and a lost-child procedure. Crowd estimates for the seated area are in Section 12, and the grass areas beside the field should be monitored.")
H2("11.5 Recording and copyright")
P("Following WBA practice, the showcase does not provide recordings. Cameras are not allowed on the field, and still cameras are allowed in the stands only if the host permits. Bands must clear music rights for their arrangements. The working group recommends a notice in the band packet and on the programme.")
H2("11.6 Emergency action plan")
P("An emergency action plan names the incident commander, the first-aid lead, the Sheriff's contact, the fire station contact, the school's contact and the evacuation routes. It is drafted by the working group, reviewed by the district and the fire station, and tested with the flaggers before the day.")
PB()

# ---------- 12 ----------
H1("12. Predictive analysis")
H2("12.1 Why a model")
P("No official attendance count for the Caruthers District Fair is published, and no attendance or parking data exists for a band exhibition in Caruthers. The working group therefore built a predictive model. It does not claim to forecast the real number of visitors. It shows how the outcome changes as the assumptions change, and it identifies the assumptions that matter most. Every assumption is listed in the companion workbook with its basis.")
H2("12.2 Method")
P("The model simulates 20,000 runs. Each run draws a value for each input from a triangular distribution with a low, a most-likely (mode) and a high value. The triangular distribution is used because the planning team could state a plausible range and a most-likely value for each input, but had no data to fit a distribution. For each run, the model computes:")
BL(["Field-show spectators: band guests (bands times guests per band) plus walk-up spectators (a share of Saturday fair attendance), capped at seated capacity.",
    "Cars: spectators times the car share, divided by people per car.",
    "Vehicles needing space: cars plus band vehicles.",
    "Usable spaces: the football field's spaces less the fair's share on the Saturday, plus the elementary and high school lots.",
    "Cost: the sum of the budget lines (each drawn from its range), plus contingency at 15%.",
    "Revenue: donations and programme sales per spectator, plus a sponsor or fair-board contribution, plus band fees.",
    "Net result: revenue minus cost."])
H2("12.3 Inputs and their basis")
rows = []
for k, v in R["inputs"].items():
    rows.append([k.replace("_", " "), f"{v['low']:,}", f"{v['mode']:,}", f"{v['high']:,}", v["basis"]])
T(["Input", "Low", "Mode", "High", "Basis"], rows, widths=[1.4, 0.7, 0.7, 0.7, 3.0], size=8)
H2("12.4 Results: spectators")
P(f"The median predicted seated spectator count is about {num(sp['50'])}, with an 80% range (10th to 90th percentile) of about {num(sp['10'])} to {num(sp['90'])}. The model caps this at seated capacity, which is itself an unverified assumption of 800 to 2,500 seats. Without the cap, the median demand is about 3,100, because band guests average about {num(R['band_guests_mean'])} and walk-up spectators add roughly 1,000 more. The gap between demand and seated capacity is the point of interest for the high school: spectators who cannot sit will stand on grass beside the field, which is safe only if it is planned.")
FIG("figures/fig_spectators.png", "Figure 2. Predicted seated spectators (capped at stand capacity), 20,000 runs.")
H2("12.5 Results: parking")
P(f"The median vehicles needing a space is about {num(ld['50'])}, against median usable spaces of about {num(sup['50'])}. A shortfall occurs in about {R['shortfall_prob']*100:.0f}% of runs. The shortfall is the main reason for Option B and Option C in Section 8.")
H2("12.6 Results: budget and net result")
T(["Measure", "P10", "Median (P50)", "P90"], [
    ["Cost (with 15% contingency)", money(cost["10"]), money(cost["50"]), money(cost["90"])],
    ["Revenue (donations, sponsor, band fees)", money(rev["10"]), money(rev["50"]), money(rev["90"])],
    ["Net result (revenue minus cost)", money(net["10"]), money(net["50"]), money(net["90"])]],
    widths=[3.1, 1.1, 1.2, 1.1], size=9)
P(f"No simulated run broke even: the probability that the net result is zero or better is {R['prob_net_nonneg']*100:.0f}%. Under the planning assumptions, the showcase needs about {money(-net['50'])} of outside support at the median, and about {money(-net['90'])} at the 90th percentile, to avoid a deficit. Put another way, a band fee of about {money(-net['50']/12)} per band (for 12 bands) would be needed to break even at the median, which is not realistic for a first-year exhibition.")
FIG("figures/fig_net.png", "Figure 3. Simulated net result. Runs left of zero need outside funding.")
H2("12.7 Sensitivity")
P("The sensitivity analysis changes one input at a time from its mode, holding the others at their most-likely values. The largest effect on the net result comes from the sponsor or fair-board contribution, followed by band fees, donations per spectator and the number of bands. Changes in attendance drivers have no effect at the base case, because seated capacity binds first. That means attendance above capacity does not add revenue unless the seating or the donation model changes. This is a key finding: if the fair board wants a larger audience, the showcase needs more seating or a model that counts standing spectators.")
FIG("figures/fig_tornado.png", "Figure 4. One-at-a-time sensitivity of the net result (base case about " + money(R['tornado_base']) + ").")
H2("12.8 Scenarios")
T(["Scenario", "Bands", "Seated spectators", "Vehicles needing space", "Usable spaces", "Revenue", "Cost", "Net"], [
    ["Low inputs", "10", "800", "265", "335", "$80", "$11,615", "-$11,535"],
    ["Base case", "12", "1,500", "479", "415", "$3,075", "$11,615", "-$8,540"],
    ["High inputs", "16", "2,500", "790", "410", "$10,500", "$11,615", "-$1,115"]],
    widths=[1.0, 0.6, 0.9, 1.0, 0.8, 0.8, 0.8, 0.8], size=8)
P("Even the high-input scenario, with optimistic sponsor and fee assumptions, produces a deficit. The workbook (Scenarios sheet) recalculates these scenarios when inputs change. The scenarios use the budget midpoint for cost.")
H2("12.9 Interpretation and limits")
BL(["The model is only as good as its inputs. The attendance and parking inputs are assumptions; the budget ranges are planning estimates; the sponsor figure is unknown.",
    "The result is robust in direction: under any input set we tested, the showcase needs outside funding unless its cost falls substantially.",
    "The model does not count revenue from the fair itself. If the fair's concessions, carnival or sponsors gain from the showcase, that is a benefit to the fair and a reason for the fair to share costs.",
    "Attendance above seated capacity is not counted as revenue in the base case. This should be revisited once the seating is known."])
PB()

# ---------- 13 ----------
H1("13. Budget and funding")
H2("13.1 Budget lines")
T(["Budget line", "Low", "High", "Mid (base)", "Basis"], [
    ["Field and facility use, lights, custodial", "$500", "$1,500", "$1,000", "Assumed district use fee; to be confirmed."],
    ["Liability insurance (one-day special event)", "$500", "$1,500", "$1,000", "Assumed; to be quoted."],
    ["Traffic and crossing control (4 to 6 flaggers, about 6 hours)", "$1,000", "$2,500", "$1,750", "Assumed hourly rates; to be quoted."],
    ["Portable restrooms (4 to 6 units)", "$600", "$1,200", "$900", "Assumed rental rate; to be quoted."],
    ["Sound and staging for the field show", "$600", "$1,500", "$1,050", "Assumed; to be quoted."],
    ["Judges (3 to 4)", "$600", "$1,600", "$1,100", "Assumed honoraria and travel."],
    ["Clinicians (2)", "$800", "$1,600", "$1,200", "Assumed honoraria and travel."],
    ["Plaques, programmes, signage", "$500", "$1,200", "$850", "Assumed."],
    ["Fire and EMS standby (added)", "$0", "$1,000", "$500", "Depends on the fire district's position; may be zero."],
    ["Trailer staging, fencing and signage (added)", "$300", "$1,200", "$750", "Assumed; depends on the Clemenceau gate."],
    ["Subtotal", "$5,400", "$14,800", "$10,100", ""],
    ["Contingency (15%)", "$810", "$2,220", "$1,515", ""],
    ["Total", "$6,210", "$17,020", "$11,615", ""]],
    widths=[2.5, 0.7, 0.7, 0.8, 2.0], size=8.5)
P("All lines are planning estimates from assumed rates. None is a quote. The working group's earlier estimate of $5,900 to $13,900 is superseded by the workbook, which adds fire standby and trailer staging.", italic=True, size=9)
H2("13.2 Revenue and funding sources")
BL([("Fair board.", "Co-sponsorship is the most important source. The working group asks for a contribution of at least $8,000 toward the budget, or in-kind support for the parade, staging, restrooms and signage."),
    ("Sponsors.", "Local businesses, the Caruthers Community Services District and agricultural businesses are plausible sponsors. None is confirmed."),
    ("Donations and programmes.", "A donation box and a printed programme. At the base assumption of $0.25 per seated spectator, this is about $375 for 1,500 spectators."),
    ("Band fees.", "A participation fee of up to $250 per band is modelled. Free participation is the recommended first-year policy if the fair board covers costs, because it lowers the barrier for bands."),
    ("Concessions.", "Food and drink sold by the fair's booths or a booster group. These are the fair's revenue, not the showcase's, unless an arrangement is made.")])
H2("13.3 Cost reduction options")
P("If the fair board cannot contribute the full amount, the following reductions are possible, each with its risk: fewer judges (three instead of four); a clinic-only format for distant bands; restrooms shared with the fair; sound and staging supplied by the school; fire standby removed if the fire district agrees; and plaques replaced by certificates.")
PB()

# ---------- 14 ----------
H1("14. Governance, roles and responsibilities")
H2("14.1 Structure")
P("The working group recommends a three-tier structure: a steering group of the fair board, the district and the band program; an operating committee that runs the day; and task teams for transportation, facilities, judging and clinics, safety and communications.")
T(["Role", "Responsibility", "Lead (proposed)"], [
    ["Steering group", "Approves budget, format and key agreements; signs the host agreement.", "Fair board president; district representative; band program lead."],
    ["Showcase director", "Runs the day; final decisions on timing and weather; liaison to bands and the fair.", "A named coordinator, to be appointed."],
    ["Facilities lead", "Field and lights agreement; restrooms; power; field condition.", "High school athletics or facilities office."],
    ["Transportation and parking lead", "Shuttle; elementary lot; Clemenceau gate; trailer plan.", "Fair board operations or a designee."],
    ["Traffic and road lead", "Permits; traffic control plan; flaggers; Sheriff's liaison.", "Traffic control contractor; Sheriff's liaison."],
    ["Fire and emergency lead", "Fire station contact; emergency action plan; first aid.", "Designated safety officer; fire station liaison."],
    ["Judging and clinics lead", "Judges and clinicians; critique and comments; WBA-style procedures.", "Band program lead with invited educators."],
    ["Communications lead", "Band packet; invitations; community promotion; programme.", "Fair board communications or a designee."],
    ["Finance lead", "Budget; contracts; sponsor and donation accounts; reporting.", "Fair board treasurer or district business office."]],
    widths=[1.6, 3.2, 1.7], size=8.5)
H2("14.2 Decision rights")
P("The steering group decides the format, budget and major agreements. The showcase director decides operational matters on the day, within the approved plan and after consulting the fire and emergency lead. Any decision to cancel or delay the field show is made by the director in consultation with the band directors and the fire and emergency lead, and is recorded.")
H2("14.3 Agreements to be signed")
BL(["Host agreement between the fair and the school district.",
    "Facility use agreement for the field, lights and restrooms.",
    "Parking and trailer agreement, including the Clemenceau gate.",
    "Traffic control plan and any county permit.",
    "Insurance certificate and indemnity.",
    "Clinician and judge agreements.",
    "Band participation agreement, including music rights and the schedule."])
PB()

# ---------- 15 ----------
H1("15. Risk register")
T(["Risk", "Likelihood", "Impact", "Mitigation", "Alternative if the mitigation fails"], [
    ["Fair board declines to co-sponsor", "Medium", "High", "Early meeting; written proposal; cost-reduction options.", "Scale to clinic and preview only; no field show in 2027."],
    ["High school field or lights unavailable", "Medium", "High", "Early agreement with the district; confirm lights before invitations.", "Move the field-show block to an earlier start and end before sunset, or move the venue."],
    ["Parking shortfall", "High (about 80% of runs)", "Medium", "Elementary lot and shuttle; Clemenceau gate for trailers.", "Fairgrounds overflow (Option D); fewer band vehicles."],
    ["Clemenceau gate not approved", "Medium", "Medium", "Early walk-through with the district and the fire station; fire lane plan.", "Trailers park at the elementary lot; shuttle the band's instruments."],
    ["Fire station cannot quiet sirens", "Medium", "Low", "Courtesy request, framed with emergency precedence; coordinator notice.", "Schedule field-show slots to avoid known non-emergency movements; warn the audience."],
    ["Road closure refused", "Medium", "Medium", "Early permit application; cone and flagger plan instead of full closure.", "Walking route with crossing control only; bands walk on sidewalks."],
    ["Heat illness", "Medium", "High", "Late start; shade and water; first aid; hydration breaks.", "Shorten slots; move the field show earlier or later, if the field allows."],
    ["Weather (rain, storms, wind)", "Low to medium", "High", "WBA-style weather categories; call-by-day-before decision.", "Move to the Sunday, if the fair allows, or cancel the field show and keep clinic and parade."],
    ["Bands withdraw", "Medium", "Medium", "Invite more bands than needed (14 to 18 for 10 to 12 commitments); confirm early.", "Reduce the slot count; keep the clinic."],
    ["Band or parent expectation of placements", "High", "Low to medium", "Clear message at invitation; recognition plan; judge comments.", "Add a year-two invitational track."],
    ["Music copyright violation", "Low", "Medium", "Written rights clearance requirement in the band packet.", "Remove the piece from the programme; notify the director."],
    ["Injury or medical emergency", "Low", "High", "Trained first-aid lead; ambulance access; emergency action plan.", "Stop the affected band's slot; route the injured person under the plan."],
    ["Traffic accident on the walking route", "Low", "High", "Flaggers; escort; cones; clear routes.", "Halt the walk; move bands by vehicle under the plan."],
    ["Overlapping event", "Medium", "Medium", "Check the WBA calendar and the Washington Union preview date before the date is set.", "Shift to the other Saturday, if the fair dates allow."],
    ["Washington Union program responds (date move, lobbying, or objection through the WBA)", "Medium", "High", "Direct, transparent contact before invitations; a schedule that avoids its date; a clear statement of the showcase's own offer.", "Accept that some bands attend both; keep the showcase on a separate date; limit the invitation to bands not committed elsewhere."],
    ["Directors see the showcase as poaching and decline future invitations", "Medium", "Medium", "Invite only bands free on the Saturday; no pressure on directors already committed to another event.", "Shift invitations toward Southern California and Riverside bands that do not attend the Washington Union event."],
    ["Budget overrun", "Medium", "Medium", "Quotes before commitment; contingency at 15%.", "Cut the cost-reduction items in Section 13.3."]],
    widths=[1.4, 0.9, 0.7, 1.9, 1.7], size=7.5)
PB()

# ---------- 16 ----------
H1("16. Alternatives and contingency plans")
P("The working group identified the alternatives below for the points most likely to fail. Each alternative is a planned fallback, not an improvised response.")
T(["If this fails", "Then", "Trigger", "Owner"], [
    ["Fair board will not co-sponsor", "Run the clinic and preview only; defer the field show to 2028.", "No written commitment by February 2027.", "Steering group"],
    ["Field or lights unavailable", "Move the field-show block to an earlier start (about 2:00 p.m.) and end before sunset; or move the venue to a lit field in the area.", "No district agreement by April 2027.", "Facilities lead"],
    ["Clemenceau gate refused", "Stage trailers at the elementary lot; shuttle instruments as needed.", "No approval by June 2027.", "Transportation lead"],
    ["Parking still short", "Use the fairgrounds overflow; reduce band vehicles through a carpool-and-shuttle plan.", "Shortfall confirmed in the August walk-through.", "Transportation lead"],
    ["Road closure refused", "Use crossing control and sidewalks; no full closure.", "Written refusal from the road authority.", "Traffic lead"],
    ["Fire station cannot quiet sirens", "Keep the courtesy request; the coordinator warns the audience and performers before any movement.", "Station cannot agree in writing by June 2027.", "Fire and emergency lead"],
    ["Heat warning", "Shorten slots by two minutes each; add shade and water; start one hour earlier if the field is lit.", "National Weather Service heat advisory for Fresno County on the day or the day before.", "Showcase director"],
    ["Rain or flooding", "Apply the WBA-style Plan B (covered warm-up where possible) or move the field show.", "Forecast or field check the day before.", "Showcase director"],
    ["Bands withdraw below 10", "Keep the day with the bands that remain and shorten the block; or cancel the field show and keep the clinic.", "Fewer than 10 confirmed by August 2027.", "Band program lead"],
    ["Medical emergency", "Follow the emergency action plan; stop the affected slot.", "Any incident.", "First-aid lead"]],
    widths=[1.5, 2.6, 1.7, 1.0], size=8)
PB()

# ---------- 17 ----------
H1("17. Implementation timeline")
T(["When", "Milestone", "Owner"], [
    ["Now to December 2026", "Confirm the fair's 2027 dates; confirm the parade route, the fire station and the Washington Union preview date; first meeting with the fair board.", "Working group"],
    ["January to February 2027", "Written fair-board position; steering group formed; format and recognition policy agreed; district meeting.", "Steering group"],
    ["February to April 2027", "Recruit clinicians and judges; invite 14 to 18 bands; first band commitments; facility use request to the district.", "Band program lead; facilities lead"],
    ["April to May 2027", "Fire station meeting; road and permit applications; Clemenceau gate walk-through; insurance quotes.", "Fire lead; traffic lead; finance lead"],
    ["May to July 2027", "Facility and parking agreements signed; traffic control plan approved; quotes for restrooms, sound and flaggers; judges and clinicians confirmed.", "Steering group; finance lead"],
    ["July to August 2027", "Band commitments confirmed (10 to 12); band packets issued two weeks before the event; programme drafted; volunteer roster.", "Showcase director; communications lead"],
    ["Late August 2027", "Walk-through of the walking route and the Clemenceau gate with the fire station, the Sheriff's office and the flaggers; emergency action plan tested.", "Showcase director"],
    ["Final week", "Final schedule to bands; weather check; trailer and parking confirmation; staging set-up.", "Showcase director"],
    ["Fair Saturday", "Event day as set out in Section 6.", "Showcase director"],
    ["Within two weeks after", "Debrief with bands, judges, fire station, district and fair board; survey of families; budget close-out; report to steering group.", "Steering group"]],
    widths=[1.5, 4.0, 1.0], size=8.5)
PB()

# ---------- 18 ----------
H1("18. Recommendations and decisions required")
H2("18.1 Recommendations")
NL(["Adopt a one-year pilot, the Caruthers Fair Band Showcase, on the fair's Saturday in 2027, using the clinic-and-preview format with a field-show slot.",
    "Invite 10 to 12 field-show bands in 2027, from the Central Valley first, with a clinic-only or stipend option for distant bands.",
    "Treat middle school bands as parade and afternoon clinic participants, outside the field-show count.",
    "Give recognition without placement: a participation plaque, written comments and optional Outstanding recognitions.",
    "Start the field-show block at 4:30 to 5 p.m. and end before 9 p.m., using the lit field.",
    "Use the elementary lot and shuttle for visitors, and a Clemenceau gate for band trailers and buses, subject to approval.",
    "Ask the fire station for a courtesy quiet-siren arrangement, with emergency precedence stated clearly.",
    "Seek at least $8,000 of fair-board co-sponsorship, and sponsors for the balance, before any invitation is sent.",
    "Position the showcase as a direct alternative to the Washington Union Preview of Champions. Contact that program directly before inviting its bands, and state the intent openly."])
H2("18.2 Decisions required")
T(["#", "Decision", "Who decides", "Needed by"], [
    ["1", "Approve the pilot and the clinic-and-preview format.", "Fair board and district", "February 2027"],
    ["2", "Commit at least $8,000 of co-sponsorship or in-kind support.", "Fair board", "February 2027"],
    ["3", "Approve use of the high school field and lights on the fair's Saturday.", "District", "April 2027"],
    ["4", "Approve the Clemenceau gate and the trailer plan.", "District; fair board; fire district", "June 2027"],
    ["5", "Confirm the fire station's position on courtesy quiet-siren use.", "Fire district", "May 2027"],
    ["6", "Confirm road permits and Sheriff's role.", "Fresno County; Sheriff", "May 2027"],
    ["7", "Confirm recognition policy (no placements in 2027).", "Steering group; band directors", "February 2027"],
    ["8", "Approve budget and contingency.", "Steering group", "June 2027"],
    ["9", "Authorise direct, transparent contact with the Washington Union program before any band invitation.", "Steering group", "January 2027"]],
    widths=[0.4, 3.4, 1.8, 0.9], size=8.5)
PB()

# ---------- APPENDIX A ----------
H1("Appendix A. Evidence register and sources")
P("Each source is listed with the date it was retrieved or published and what it supports. Items marked VERIFY are not confirmed by any source.")
T(["#", "Source", "Used for", "Reliability note"], [
    ["A1", "Caruthers District Fair official website, caruthersfair.com (home page, retrieved for the 2026 fair)", "2026 dates; gate hours; parade time and location; programming; address; contact details; free-gate description.", "Official, but self-described; parade location differs from other listings."],
    ["A2", "UC Agriculture and Natural Resources, ucanr.edu (node 126516 and print edition)", "2024 and 2025 fair dates, September 25 to 28 (2024).", "Public university extension listing."],
    ["A3", "Fresno County 4-H annual events page (ucanr.edu/sites/4-H-Fresno)", "2025 dates, September 24 to 27.", "Public extension listing."],
    ["A4", "ABC30, \"Caruthers District Fair begins today\" and related pieces", "Fair dates and events in prior years; a 2021 organizer estimate of 60,000 to 80,000 visitors over four days, based on carnival ticket sales.", "News report; attendance is an organizer estimate."],
    ["A5", "KMPH (Fox 26), coverage of the Caruthers fair and free admission", "Free gate; largest-fair claim as the fair's own statement.", "News report; the 'largest' claim is not independently verified."],
    ["A6", "Fresno State Collegian, 2014 coverage of the 85th Caruthers fair", "Organizer estimates of 60,000 to 100,000 visitors a year.", "Student newspaper; dated; estimates only."],
    ["A7", "WBA 2025 Policy Manual, version 10, October 2025 (cms.bluedevils.org)", "Season timing; event schedule; credentials; critique; exhibition awards; weather categories; copyright; host checklist; ticketing; host agreement.", "Primary governing document for WBA events. Read as retrieved; section numbers from its table of contents."],
    ["A8", "WBA 2024 Policies and Procedures (cms.bluedevils.org)", "Comparison with the 2025 manual.", "Dated; cited for context only."],
    ["A9", "Wikipedia, Western Band Association", "Organization history; ownership; circuit description.", "Secondary source."],
    ["A10", "CSBC 2016 field rules (cvhs.com)", "Adjudication by the Southwest Judges Network; time penalty; copyright.", "Dated (2016); indicative only."],
    ["A11", "Gahr High School Marching Gladiators, gahrband.org news", "Gahr competed at WBA Championships in Fresno, November 21 to 23, 2025; Trabuco Hills regional.", "School band website; the placement was not confirmed."],
    ["A12", "Clovis North (cnhs.cusd.com) instrumental music page", "Reports a 4A state championship and a score of 90.45.", "School website; year not stated."],
    ["A13", "FloMarching, 2021 WBA Regional Championships article", "The Kingsburg Viking Classic as a WBA Regional Championship; 2021 season ended in mid-November.", "Trade media."],
    ["A14", "Knights Brigade press release (2019), WBA Championships in Kingsburg", "Kingsburg as a WBA event site in November 2019.", "Band press release."],
    ["A15", "Wikipedia, Caruthers, California", "2020 census count of 2,613; 2010 count of 2,497.", "Secondary source citing the census."],
    ["A16", "NERIS (neris.fsri.org), Fresno County Fire Station 90", "Station address 2701 W. Tahoe Ave.; public phone 559-864-3211.", "Public fire-data listing; verify with the district."],
    ["A17", "Fresno County Fire Battalion 15 public page", "Station built 1993; earlier station near the fairgrounds opened in 1947.", "Public agency page."],
    ["A18", "Fresno LAFCO map and report, Caruthers Community Services District (2021)", "Water, sewer and street lights; not fire protection.", "Public agency document."],
    ["A19", "City of Fresno, citywide special events page and FMC 14-2404 (as reported)", "Model for permit and 48-hour notice practice; not Fresno County rules.", "Used as a model only."],
    ["A20", "ABC30 and KFSN, 2024 Caruthers High School celebration parade", "Parade route from Henderson Road and Tahoe Avenue to the fairgrounds.", "News report."],
    ["A21", "Working group's observations of the aerial map and of local streets", "School, fair and fire station locations; walking route; parking locations.", "Working observation; requires field walk."]],
    widths=[0.4, 2.2, 2.4, 1.5], size=7.5)
H2("A.1 Sources not retrieved")
BL(["The current CSBC rulebook and stadium specifications (calstatebandchamps.org did not return readable content).",
    "The Fresno County special event and encroachment permit rules.",
    "The Caruthers Unified School District band calendar and the high school field's specifications.",
    "The Washington Union Preview of Champions date, host and format.",
    "Fair attendance records for any recent year, and the fair's own parking plan."])
PB()

# ---------- APPENDIX B ----------
H1("Appendix B. Verification checklist and data gaps")
T(["#", "Item to verify", "Why it matters", "Who verifies"], [
    ["B1", "2027 fair dates and the Saturday", "The whole plan depends on the fair's Saturday.", "Fair board"],
    ["B2", "Parade route and street names (Main Street or W. Tahoe Ave. from S. March Ave.)", "The parade and the walking route depend on it.", "Fair board; traffic lead"],
    ["B3", "High school field availability, lights, bleachers, restrooms, power", "The field show cannot happen without them.", "District facilities"],
    ["B4", "Seated capacity of the high school bleachers", "Caps the spectator count in the model.", "District facilities"],
    ["B5", "Parking spaces in the football field, elementary lot and high school lot", "Drives the parking model.", "Transportation lead (field count)"],
    ["B6", "Fire station near Clemenceau Avenue, or the W. Tahoe Ave. station (2701 W. Tahoe)", "Determines the fire coordination point and the gate's location.", "Fire district"],
    ["B7", "Fire district's position on courtesy quiet-siren arrangements", "Affects the noise plan.", "Fire district"],
    ["B8", "Fresno County permit requirements for each road on the route", "Required before any closure.", "Fresno County Public Works and Planning"],
    ["B9", "Sheriff's role during the field-show block", "Noise and crowd plan.", "Sheriff's office"],
    ["B10", "Current CSBC rules and stadium specifications", "Time, penalties and host requirements.", "Working group (CSBC)"],
    ["B11", "Current WBA host and sanction rules, if the showcase seeks WBA involvement", "Determines what can be called a WBA-related event.", "Working group (WBA)"],
    ["B12", "Washington Union Preview of Champions date", "Avoid overlap; confirm the rival date before setting ours.", "Working group"],
    ["B13", "Band seasons and availability for each invited band", "Only bands free on the Saturday can be invited.", "Working group"],
    ["B14", "Sunset time on the chosen day", "Lighting plan.", "Working group"],
    ["B15", "Ambulance and first-aid provider costs", "Budget line.", "Finance lead"],
    ["B16", "Insurance quote for a one-day special event", "Budget line; host agreement.", "Finance lead"],
    ["B17", "Fair attendance data for recent years", "Would replace the model's organizer estimate.", "Fair board"],
    ["B18", "Caruthers High School band's participation in the fair parade", "Shows the school's existing engagement.", "School band program"]],
    widths=[0.4, 2.5, 2.3, 1.3], size=8)
PB()

# ---------- APPENDIX C ----------
H1("Appendix C. Draft band invitation letter")
P("[Date]", after=2)
P("[Band director name]\n[School name]\n[Address]", after=8)
P("Dear [Director name],")
P("On behalf of the Caruthers Fair Band Showcase working group, we invite your band to take part in the first annual Caruthers Fair Band Showcase, on Saturday, [date] 2027, at the Caruthers District Fair in Caruthers, California. The fair is a free-gate event, so families and community members attend at no charge.")
P("The showcase is an exhibition. It is not a competition, and bands are not ranked. Each field-show band will take part in a clinic with a guest music educator, a short preview performance, and a field-show slot. Judges will give your band written comments on the day, and each band will receive a participation plaque. Parade participation is optional for field-show bands.")
P("We are inviting 10 to 12 field-show bands, and we would like to hear from you by [date]. Please reply with your band's availability, the number of students and vehicles you expect to bring, and whether your band plans to march in the parade. We will send a full band packet two weeks before the event, with the schedule, call times, parking and walking routes, and music rights requirements.")
P("The showcase gives your students a live audience before the fall circuit, with a clinic, judges' written feedback, and a participation plaque, and no placement. We would be glad to answer any questions.")
P("With thanks,")
P("[Name]\n[Role], Caruthers Fair Band Showcase working group\n[Contact details]")
P("Note: This letter is a draft. Each invitation should be sent only after the fair board, the district and the fire district have confirmed their positions, and after the band's season and availability have been checked.", italic=True, size=9)
PB()

# ---------- APPENDIX D ----------
H1("Appendix D. Model method and workbook guide")
H2("D.1 Files")
BL(["Caruthers_Showcase_Planning_Model.xlsx: the workbook. Yellow cells on the Inputs sheet are the assumptions. The Scenarios, Budget and Summary sheets recalculate from them. The MC_Samples sheet holds 2,000 of the 20,000 simulated runs. Summary percentiles use that sample.",
    "model.py: the Monte Carlo model (20,000 runs, triangular inputs, fixed random seed for reproducibility). It writes results.json, the samples file and the figures.",
    "build_xlsx.py and build_docx.py: scripts that build the workbook and this report from the same results."])
H2("D.2 Reproducing the results")
P("Run model.py to regenerate the results, then build_xlsx.py and build_docx.py. The seed is fixed, so the results are reproducible. Changing an input in the workbook changes the scenario results but not the Monte Carlo summary, which is based on the Python inputs. To change the Monte Carlo inputs, edit the INPUTS table in model.py.")
H2("D.3 Modelling choices")
BL(["Triangular distributions are used because the planning team could state a range and a most-likely value, but had no data to fit a distribution.",
    "Seated spectators are capped at stand capacity. Demand above capacity is not counted, which understates the audience if spectators stand.",
    "The fair's share of the football field is modelled as a random share of its spaces on the Saturday, to reflect uncertainty in fair demand.",
    "Cost lines are drawn independently. In reality, some costs (such as flaggers and restrooms) may move together; correlation is not modelled.",
    "Revenue from concessions, carnival and fair programming is excluded, because it belongs to the fair."])
H2("D.4 Limits")
P("The model is a planning tool. It cannot forecast real attendance or parking. Its value is in showing the direction and the size of the main effects, and in identifying which assumptions to verify first.")
PB()

# ---------- APPENDIX E ----------
H1("Appendix E. Sample schedule for twelve field-show bands")
P("The sample below assumes a 4:30 p.m. start, twelve field-show bands, 20-minute slots, and a 30-minute buffer at the end. Band names are placeholders.")
rows = []
t0 = 16 * 60 + 30
for i in range(12):
    s = t0 + i * 20
    e = s + 8
    rows.append([f"{s//60}:{s%60:02d}", f"{e//60}:{e%60:02d}", f"Band {i+1}", "Preview (up to 8 min), then changeover buffer", "Gate 15 minutes before slot"])
rows.append(["8:30", "9:00", "Buffer and recognition", "Participation plaques; judges' comments", "Program closes by 9:00 p.m."])
T(["Start", "Performance end", "Band (placeholder)", "Segment", "Band call"], rows, widths=[0.7, 1.1, 1.3, 2.4, 1.6], size=8.5)
P("The schedule is a template. Each band's slot is confirmed in the band packet. A band's full field show is performed in the field-show block; the preview and the field-show slot are shown here as a single segment for clarity.", italic=True, size=9)

doc.save("Caruthers_Fair_Band_Showcase_Report.docx")
print("saved docx")
