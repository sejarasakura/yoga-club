import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "/Users/susielim/Documents/ChatGPT/YOGA Hosting";
const SKILL_DIR = "/Users/susielim/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations";
const RUNTIME_PYTHON = "/Users/susielim/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";
const mediaDir = path.join(workspaceDir, ".codex-build/exco-deck/reference-media");
const stagingDir = path.join(workspaceDir, ".codex-build/exco-deck/finalizer-detailed");
const FINAL_PPTX = path.join(workspaceDir, ".docs/03 Meetings/Presentations/Yoga Club EXCO Meeting – 11 September 2026 (Detailed Version).pptx");
const referencePath = path.join(workspaceDir, "1nd EXCO proposal meeting 2026.pptx");

const { makeNativeBulletParagraphs, finalizePresentation } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href,
);

const W = 1920;
const H = 1080;
const FONT = "Poppins";
const FONT_BOLD = "Poppins Bold";
const C = {
  ink: "#433831",
  dark: "#5B4F47",
  brown: "#745E4D",
  tan: "#BD9479",
  sand: "#DDC1A7",
  pale: "#F4EBE4",
  paper: "#FBFAF8",
  white: "#FFFFFF",
  muted: "#766B64",
  gold: "#D99D48",
};

const media = {};
async function asset(name) {
  if (!media[name]) media[name] = new Uint8Array(await fs.readFile(path.join(mediaDir, name)));
  return media[name];
}

const presentation = Presentation.create({ slideSize: { width: W, height: H } });
presentation.theme.colorScheme = {
  name: "UTAR Yoga Warm Neutral",
  themeColors: {
    accent1: C.dark,
    accent2: C.tan,
    accent3: C.sand,
    accent4: C.gold,
    accent5: C.brown,
    accent6: C.pale,
    bg1: C.paper,
    bg2: C.pale,
    tx1: C.ink,
    tx2: C.muted,
    dk1: "#000000",
    dk2: C.dark,
    lt1: C.white,
    lt2: C.pale,
    hlink: C.brown,
    folHlink: C.tan,
  },
};

function addText(slide, text, position, style = {}, name = undefined) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: FONT,
    fontSize: 30,
    color: C.ink,
    autoFit: "none",
    ...style,
  };
  return shape;
}

function addRect(slide, position, fill, name = undefined) {
  return slide.shapes.add({
    geometry: "rect",
    name,
    position,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
}

async function addImage(slide, name, position, options = {}) {
  const extension = path.extname(name).toLowerCase();
  const contentType = extension === ".jpeg" || extension === ".jpg" ? "image/jpeg" : "image/png";
  return slide.images.add({
    blob: await asset(name),
    contentType,
    alt: options.alt ?? "Visual from the supplied reference presentation",
    fit: options.fit ?? "cover",
    position,
    geometry: options.geometry ?? "rect",
    ...(options.borderRadius ? { borderRadius: options.borderRadius } : {}),
    ...(options.crop ? { crop: options.crop } : {}),
  });
}

async function addBase(slide, section, title, slideNo) {
  slide.background.fill = C.paper;
  await addImage(slide, "image1.jpeg", { left: 0, top: 0, width: W, height: H }, {
    fit: "cover",
    alt: "Subtle white geometric background from the supplied reference presentation",
  });
  await addImage(slide, "image2.png", { left: 1600, top: 760, width: 320, height: 320 }, {
    fit: "contain",
    alt: "Warm dotted texture from the supplied reference presentation",
  });
  await addImage(slide, "image4.png", { left: 45, top: 25, width: 78, height: 78 }, {
    fit: "contain",
    alt: "UTAR Yoga Society logo from the supplied reference presentation",
  });
  addRect(slide, { left: 137, top: 41, width: 255, height: 42 }, C.dark, "club-label-band");
  addText(slide, "UTAR SL YOGA CLUB", { left: 157, top: 45, width: 220, height: 30 }, {
    typeface: FONT_BOLD, fontSize: 19, bold: true, color: C.white,
  }, "club-label");
  addText(slide, String(section).padStart(2, "0"), { left: 1650, top: 46, width: 90, height: 32 }, {
    typeface: FONT_BOLD, fontSize: 20, bold: true, color: C.tan,
  }, "section-number");
  addText(slide, title, { left: 110, top: 120, width: 1560, height: 105 }, {
    typeface: FONT_BOLD, fontSize: 52, bold: true, color: C.dark,
  }, "slide-title");
  addRect(slide, { left: 110, top: 222, width: 118, height: 8 }, C.tan, "title-rule");
  addText(slide, `EXCO MEETING, 11 SEPTEMBER 2026`, { left: 110, top: 1012, width: 620, height: 28 }, {
    fontSize: 18, color: C.muted,
  }, "footer-source");
  addText(slide, String(slideNo), { left: 1774, top: 1010, width: 52, height: 28 }, {
    typeface: FONT_BOLD, fontSize: 18, bold: true, color: C.dark,
  }, "folio");
}

function addHeading(slide, text, position, color = C.brown) {
  addText(slide, text.toUpperCase(), position, {
    typeface: FONT_BOLD, fontSize: 23, bold: true, color,
  });
}

function addBullets(slide, items, position, options = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = makeNativeBulletParagraphs(items, {
    marginLeftPoints: options.marginLeftPoints ?? 20,
    hangingPoints: options.hangingPoints ?? 10,
    spaceAfterPoints: options.spaceAfterPoints ?? 10,
  });
  box.text.style = {
    typeface: FONT,
    fontSize: options.fontSize ?? 29,
    color: options.color ?? C.ink,
    autoFit: "none",
  };
  return box;
}

function addKeyLine(slide, lead, body, top, width = 1180) {
  addText(slide, lead.toUpperCase(), { left: 110, top, width: 310, height: 38 }, {
    typeface: FONT_BOLD, fontSize: 23, bold: true, color: C.tan,
  });
  addText(slide, body, { left: 420, top: top - 2, width, height: 62 }, {
    fontSize: 28, color: C.ink,
  });
}

const detailedNotes = {
  "meeting header": [
    "Purpose: open the meeting, confirm attendance and establish the decision-making scope.",
    "Facilitator guidance: confirm whether the meeting has the required EXCO representation. State that this meeting covers programme planning, recruitment, governance, finance, certification, staffing and merchandise. Ask members to distinguish confirmed decisions from proposals requiring later approval.",
    "Record: attendance, apologies, meeting start time and any declared conflicts of interest.",
  ],
  "overview of Items 1–12": [
    "Purpose: align members on the decisions that must leave the meeting with a named owner and completion date.",
    "Discussion: confirm the activity portfolio first because proposals, publicity, budgets and recruitment depend on it. Resolve the October Recruitment Drive offer before publicity begins. Confirm which certification model requires further consultation. Treat merchandise and Members’ Night funding as conditional until the Treasurer completes the budget review.",
    "Decision required: agree the sequence of decisions and identify any item that cannot be concluded without Ms Belinda, Ms Heng, a Core Advisor or DSA approval.",
    "Record: decision, owner, due date, approval dependency and supporting document for every agenda item.",
  ],
  "Item 1": [
    "Purpose: define the next-semester activity portfolio and distinguish compulsory work from optional pilots.",
    "Compulsory activities: Recruitment Drive, weekly yoga classes and free yoga classes. For a long semester, consider six to eight practical sessions in addition to regular classes. These sessions may extend certification hours and create supervised training opportunities for student coaches. Use common student free hours when scheduling.",
    "Options: assess Members’ Night as a reward for regular attendance. Decide whether meditation runs once, twice or weekly and whether participants pay a fee. Decide whether Natural Yoga begins next semester or moves to February 2027.",
    "Decision required: approve the activity list, provisional frequency, semester timing and department responsible for developing each proposal.",
    "Record: confirmed activities, deferred activities, responsible department and proposal due date.",
  ],
  "Item 2": [
    "Purpose: agree the October Recruitment Drive offer and ensure publicity communicates only confirmed benefits.",
    "Promotion: merchandise, lucky draw opportunities and weekly-class certification. If Members’ Night proceeds, promote it after the October free class. Possible benefits include food, health-oriented gifts, participant feedback activities and information about upcoming classes.",
    "Participant offer: review whether first-time participants receive one free class. Consider a fee waiver for members who bring new participants to selected classes. The Treasurer should confirm any cost or revenue effect before Publicity publishes the offer.",
    "Recruitment: include the vacant Vice Treasurer role in Recruitment Drive materials. Conduct a follow-up survey after the activity.",
    "Decision required: approve the offer, benefit wording, survey owner and Vice Treasurer recruitment message.",
  ],
  "Item 3; obvious source typo ‘Office Barriers’ normalised to ‘Office Bearers’": [
    "Purpose: establish a complete and traceable secretarial control system for proposals and semester reports.",
    "Proposal standard: create one proposal for each activity where required. Each proposal should state objectives, programme details, schedule, venue, budget, publicity arrangements, registration process and required approvals. Submit the proposal before the event and within the institutional timeline.",
    "Reporting standard: complete reports by Week 14 for weekly classes, the Recruitment Drive, EXCO Meeting 1, EXCO Meeting 2, the free yoga class, Office Bearers and any other reportable activity or matter.",
    "Control mechanism: maintain a live checklist showing document status, owner, submission date and approval status. Record student coach service hours and attendance in a form that supports accountability and certificate issuance.",
    "Decision required: approve the checklist format, report ownership and review frequency.",
  ],
  "Item 4": [
    "Purpose: define how financial decisions move from the Treasurer into proposals, publicity and operational records.",
    "Treasury responsibilities: prepare event-specific policy and documentation, confirm prices or member benefits, retain receipts and supporting evidence, and maintain records for weekly classes and events.",
    "Information hand-off: once the Treasurer confirms the financial arrangement, the Secretary incorporates it into proposals and reports. Publicity then prepares the pricing announcement and payment instructions. Any subsequent change requires renewed verification before publication.",
    "Digital support: fee tracking and unified receipt uploading already exist as proposed website functions. Treat them as useful but not urgent unless current record-keeping creates material gaps.",
    "Decision required: confirm the minimum documents for each event and who verifies financial information before release.",
  ],
  "Item 5": [
    "Purpose: agree entry conditions, benefits and signing authority for the Student Instructor Programme.",
    "Eligibility: more than two semester-based Certificates of Completion, or one yearly Certificate of Completion if the club adopts the yearly model. Semester certificates apply only to long semesters of about fourteen weeks.",
    "Proposed benefit: consider a fee waiver for student instructors who maintain attendance above 80 per cent. The Treasurer should assess the financial effect and the committee should clarify whether the waiver applies prospectively or retrospectively.",
    "Signing authority: consider Ms Chew, the Club Auditor, the Club Chairperson, and Ms Belinda, Ms Heng or a Core Advisor. Confirm the required combination rather than assuming every person signs.",
    "Decision required: select the eligibility model, attendance evidence, benefit rule and authorised signatories.",
  ],
  "Item 6": [
    "Purpose: create a controlled publicity process that prevents unconfirmed information from reaching members.",
    "Scope: Recruitment Drive, free classes, weekly classes, Members’ Night, lucky draw activities, merchandise, meditation and other approved programmes.",
    "Verification: Publicity should confirm event dates, venues, prices, member benefits, registration links, payment instructions and certification requirements with the Secretary and Treasurer before release. When arrangements change, update every affected poster, caption, announcement and link.",
    "Records: retain the final published version and enough context to identify the event, publication date and approving source. Avoid keeping multiple files without a clear final status.",
    "Decision required: approve the verification route, publication authority and archive location.",
  ],
  "Item 7": [
    "Purpose: finalise a certification policy that the club can explain consistently and verify through attendance records.",
    "Semester model: at least 80 per cent attendance across approximately fourteen hours or fourteen class days during a long semester. Attendance from a short semester may accumulate towards the long-semester requirement.",
    "Yearly alternative: approximately thirty-five hours, thirty-five days or thirty-five yoga practice classes. The committee must choose one measurable unit and explain how mixed class durations count.",
    "Signatories and logos: consider Ms Chew, Student Coach HOD Ms Xinle, the Club Chairperson, and Ms Belinda, Ms Heng or a Core Advisor. Confirm whether Yoga Club, DSA and UTAR logos are permitted and whether separate approval applies to each.",
    "Decision required: select the certification model, measurement unit, evidence source, signatories and permitted logos.",
  ],
  "Item 8": [
    "Purpose: fill leadership and publicity capacity gaps before the next semester begins.",
    "Priority roles: recruit one additional Publicity team member and identify a suitable current committee member with a strong network for Vice Chairperson. The committee should also consider additional higher committee members where workload or succession risk justifies recruitment.",
    "Recruitment routes: committee referrals, direct invitations and EA appointments for experience-based roles connected with student instruction, committee participation, events or publicity support.",
    "Assessment: consider availability, reliability, role understanding, relevant experience and willingness to complete documentation. Personal networks may identify candidates, but the role requirements should govern selection.",
    "Decision required: confirm vacancies, selection criteria, recruiter and target completion date.",
  ],
  "Item 9": [
    "Purpose: reduce the merchandise list to a feasible shortlist before design and supplier work begins.",
    "Choices: higher-end items include a logo cape or cloak, yoga mat and yoga bricks. Middle-range items include a yoga brick, T-shirt, hat and stretching strap. Lower-cost options include a keychain, lanyard, rubber wristband and other suitable small items.",
    "Polling: every committee member should participate and select at least three preferred items from the limited choices. Record the number of votes rather than only the winning items.",
    "Dependencies: Publicity prepares designs, the Secretary includes selected merchandise in the proposal, and the Treasurer confirms supplier pricing, minimum order quantities and pre-funding requirements before commitment.",
    "Decision required: agree the polling deadline, shortlist size and budget ceiling for supplier enquiries.",
  ],
  "Items 2.3 and 10": [
    "Purpose: determine whether Members’ Night and a lucky draw provide a proportionate and affordable reward for member participation.",
    "Audience: regular participants from the June weekly classes and new members joining the club. Clarify whether attendance history, current membership or event registration determines eligibility.",
    "Possible benefits: food, health-oriented gifts, selected merchandise, participant feedback and promotion of future classes. Publicity should avoid promising benefits before budget confirmation.",
    "Financial condition: any use of club funds for the lucky draw requires budget review and approval. The Treasurer should document expected cost, item source and purchasing evidence.",
    "Decision required: confirm whether the activity proceeds, eligibility, benefits, provisional budget and approval route.",
  ],
  "Item 12": [
    "Purpose: test whether student instructors can deliver meditation or additional yoga practice safely and consistently.",
    "Pilot design: choose meditation, yoga practice or both. Decide whether the pilot runs on one selected day or weekly. Confirm venue, participant fee, instructor eligibility, supervision and attendance recording before publicity begins.",
    "Certification: eligible student instructors may receive certification under the approved framework. Event organisation and publicity personnel may receive separate yearly service certificates that state events supported, total service hours and role.",
    "Accountability: record responsibilities by hours and retain enough evidence to verify service claims.",
    "Decision required: approve the pilot scope, frequency, location owner, supervision arrangement and service-record method.",
  ],
  "Items 3.9–3.12 and 11": [
    "Purpose: resolve cross-cutting matters that affect several departments or require external clarification.",
    "Yoga Day: consult Ms Belinda about institutional requirements, information needs and approvals for a larger-scale event. Do not begin public promotion until the approval route is understood.",
    "Attendance: select a recording method that captures participant identity, date, activity and verified attendance. The method should support certification and coach service-hour records.",
    "Website: decide whether the proposed club website should hold attendance, fee status and receipt uploads. Define access rights and the authoritative record before moving information online.",
    "Decision required: identify unresolved matters, owner, consultation target and return date for each issue.",
  ],
  "meeting close and action-recording requirement derived from Items 3, 4, 6 and 11": [
    "Purpose: convert meeting discussion into an accountable record.",
    "Close-out check: read back each decision and confirm that members interpret it in the same way. Name one accountable owner, even when several departments contribute. Add a due date and identify the document or evidence that demonstrates completion.",
    "Dependencies: record any DSA, advisor, instructor, Treasurer or Secretary approval needed before work proceeds. Mark unresolved items clearly and set a date for returning them to the committee.",
    "After the meeting: the Secretary updates the minutes and proposal-report checklist. Department leads verify their actions and notify the Secretary when circumstances change.",
    "Record: final decisions, dissent or reservations where relevant, owners, due dates, approval dependencies and the next review point.",
  ],
};

function notes(slide, item) {
  const detail = detailedNotes[item] ?? [];
  slide.speakerNotes.textFrame.setText([
    `Content source: Yoga Club EXCO Meeting Agenda – 11 September 2026.md, ${item}.`,
    `Visual language and embedded imagery reference: 1nd EXCO proposal meeting 2026.pptx.`,
    "",
    ...detail,
  ].join("\n\n"));
}

// 1. Cover
{
  const slide = presentation.slides.add();
  slide.background.fill = C.paper;
  await addImage(slide, "image1.jpeg", { left: 0, top: 0, width: W, height: H }, { fit: "cover" });
  await addImage(slide, "image2.png", { left: 1380, top: 0, width: 540, height: 540 }, { fit: "contain" });
  addRect(slide, { left: 0, top: 0, width: 46, height: H }, C.dark, "left-accent");
  addRect(slide, { left: 1250, top: 0, width: 670, height: H }, C.pale, "cover-field");
  await addImage(slide, "image4.png", { left: 1390, top: 205, width: 360, height: 360 }, { fit: "contain" });
  addText(slide, "YOGA CLUB SOCIETY", { left: 120, top: 118, width: 830, height: 55 }, {
    typeface: FONT_BOLD, fontSize: 27, bold: true, color: C.tan,
  });
  addText(slide, "EXCO MEETING\nAGENDA", { left: 118, top: 220, width: 1000, height: 235 }, {
    typeface: FONT_BOLD, fontSize: 74, bold: true, color: C.dark,
  }, "cover-title");
  addRect(slide, { left: 120, top: 482, width: 175, height: 9 }, C.tan, "cover-rule");
  addText(slide, "Friday, 11 September 2026\n7:00 p.m., Google Meet", { left: 120, top: 535, width: 890, height: 120 }, {
    fontSize: 31, color: C.ink,
  });
  addText(slide, "SUNGAI LONG CAMPUS", { left: 120, top: 900, width: 520, height: 45 }, {
    typeface: FONT_BOLD, fontSize: 24, bold: true, color: C.brown,
  });
  notes(slide, "meeting header");
}

// 2. Meeting outcomes
{
  const slide = presentation.slides.add();
  await addBase(slide, 1, "Decisions required at this meeting", 2);
  addKeyLine(slide, "Programme", "Confirm compulsory activities and select the proposed trials for the next semester.", 290);
  addKeyLine(slide, "October RD", "Approve the promotion offer, Members’ Night link and Vice Treasurer recruitment message.", 405);
  addKeyLine(slide, "Governance", "Agree proposal, reporting, finance, attendance and publicity controls.", 520);
  addKeyLine(slide, "People", "Confirm certification rules, signatories and committee recruitment priorities.", 635);
  addKeyLine(slide, "Resources", "Shortlist merchandise and determine which activities require budget approval.", 750);
  addText(slide, "Meeting discipline", { left: 1180, top: 825, width: 390, height: 40 }, {
    typeface: FONT_BOLD, fontSize: 24, bold: true, color: C.dark,
  });
  addText(slide, "Record every decision with an owner and due date before the meeting closes.", { left: 1180, top: 875, width: 520, height: 82 }, {
    fontSize: 27, color: C.brown,
  });
  notes(slide, "overview of Items 1–12");
}

// 3. Semester activity plan
{
  const slide = presentation.slides.add();
  await addBase(slide, 2, "Next semester activity plan", 3);
  addHeading(slide, "Compulsory programme", { left: 110, top: 280, width: 650, height: 40 });
  addBullets(slide, [
    "Recruitment Drive",
    "Weekly yoga classes",
    "Free yoga classes",
  ], { left: 110, top: 335, width: 690, height: 260 }, { fontSize: 31, spaceAfterPoints: 14 });
  addHeading(slide, "Options for review", { left: 110, top: 645, width: 650, height: 40 });
  addBullets(slide, [
    "Members’ Night as a reward for regular participants",
    "One or two meditation sessions, free or paid",
    "Natural Yoga pilot next semester or from February 2027",
  ], { left: 110, top: 700, width: 920, height: 255 }, { fontSize: 27, spaceAfterPoints: 10 });
  await addImage(slide, "image29.png", { left: 1160, top: 280, width: 620, height: 500 }, {
    fit: "cover", geometry: "ellipse", alt: "Group yoga class image from the supplied reference presentation",
  });
  addText(slide, "6–8", { left: 1205, top: 825, width: 170, height: 75 }, {
    typeface: FONT_BOLD, fontSize: 55, bold: true, color: C.tan,
  });
  addText(slide, "additional practical sessions during a long semester, scheduled around common free hours and certification needs", { left: 1370, top: 824, width: 420, height: 130 }, {
    fontSize: 24, color: C.ink,
  });
  notes(slide, "Item 1");
}

// 4. October recruitment drive
{
  const slide = presentation.slides.add();
  await addBase(slide, 3, "October Recruitment Drive", 4);
  await addImage(slide, "image22.png", { left: 1150, top: 265, width: 650, height: 610 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Yoga recruitment illustration from the supplied reference presentation",
  });
  addHeading(slide, "Promotion content", { left: 110, top: 280, width: 650, height: 40 });
  addBullets(slide, [
    "Club merchandise and lucky draw",
    "Weekly-class certification",
    "Members’ Night after the October free class, if confirmed",
    "Possible benefits include food, health-oriented gifts and participant feedback",
  ], { left: 110, top: 335, width: 890, height: 260 }, { fontSize: 29 });
  addHeading(slide, "Offers and recruitment", { left: 110, top: 640, width: 650, height: 40 });
  addBullets(slide, [
    "Review a free first class for first-time participants",
    "Review fee waivers for members who bring new participants",
    "Recruit the vacant Vice Treasurer during the drive",
    "Collect participant and member feedback after the activity",
  ], { left: 110, top: 695, width: 930, height: 265 }, { fontSize: 26, spaceAfterPoints: 8 });
  notes(slide, "Item 2");
}

// 5. Secretarial governance
{
  const slide = presentation.slides.add();
  await addBase(slide, 4, "Secretarial deliverables and records", 5);
  addHeading(slide, "Before each activity", { left: 110, top: 275, width: 650, height: 40 });
  addBullets(slide, [
    "Prepare one proposal for every activity where required",
    "Cover objectives, schedule, venue, budget, publicity, registration and approvals",
    "Submit before the activity and within the required timeline",
  ], { left: 110, top: 330, width: 820, height: 300 }, { fontSize: 27 });
  addHeading(slide, "By Week 14", { left: 110, top: 665, width: 650, height: 40 });
  addBullets(slide, [
    "Complete reports for weekly classes, the Recruitment Drive, both EXCO meetings and the free yoga class",
    "Include Office Bearers and any other reportable activities or matters",
  ], { left: 110, top: 720, width: 850, height: 220 }, { fontSize: 26 });
  await addImage(slide, "image46.png", { left: 1120, top: 270, width: 650, height: 390 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Certificate records shown in the supplied reference presentation",
  });
  addHeading(slide, "Control record", { left: 1120, top: 705, width: 480, height: 40 });
  addBullets(slide, [
    "Maintain a completed and outstanding checklist",
    "Record student coach service hours and attendance",
    "Review whether the club website can support traceability",
  ], { left: 1120, top: 755, width: 620, height: 210 }, { fontSize: 24, spaceAfterPoints: 7 });
  notes(slide, "Item 3; obvious source typo ‘Office Barriers’ normalised to ‘Office Bearers’");
}

// 6. Treasury
{
  const slide = presentation.slides.add();
  await addBase(slide, 5, "Treasury policy and documentation", 6);
  await addImage(slide, "image24.png", { left: 1160, top: 290, width: 560, height: 560 }, {
    fit: "contain", geometry: "ellipse", alt: "Financial yoga illustration from the supplied reference presentation",
  });
  addBullets(slide, [
    "Establish a treasury policy and supporting documentation for every event",
    "Maintain complete financial records for events and weekly classes",
    "Pass confirmed financial arrangements to the Secretary and Publicity team",
    "Publicity prepares pricing announcements after verification",
    "The Secretary records the arrangements in proposals and reports",
    "Fee tracking and unified receipt upload functions already exist on the proposed website and are not urgent",
  ], { left: 110, top: 290, width: 920, height: 625 }, { fontSize: 27, spaceAfterPoints: 9 });
  notes(slide, "Item 4");
}

// 7. Student Instructor Programme
{
  const slide = presentation.slides.add();
  await addBase(slide, 6, "Student Instructor Programme", 7);
  await addImage(slide, "image26.png", { left: 1070, top: 280, width: 720, height: 460 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Student yoga group illustration from the supplied reference presentation",
  });
  addHeading(slide, "Eligibility", { left: 110, top: 280, width: 450, height: 40 });
  addBullets(slide, [
    "More than two semester-based Certificates of Completion",
    "Or one yearly Certificate of Completion if the yearly system is adopted",
    "Semester certificates apply only to long semesters of about 14 weeks",
  ], { left: 110, top: 335, width: 820, height: 280 }, { fontSize: 27 });
  addHeading(slide, "Benefit under review", { left: 110, top: 655, width: 520, height: 40 });
  addText(slide, "Fee waiver for student instructors who maintain more than 80% class attendance", { left: 110, top: 710, width: 820, height: 100 }, {
    fontSize: 31, color: C.ink,
  });
  addHeading(slide, "Proposed signatories", { left: 1070, top: 780, width: 540, height: 40 });
  addText(slide, "Instructor Ms Chew, Club Auditor, Club Chairperson, and Ms Belinda, Ms Heng or Core Advisor", { left: 1070, top: 835, width: 710, height: 105 }, {
    fontSize: 25, color: C.ink,
  });
  notes(slide, "Item 5");
}

// 8. Publicity
{
  const slide = presentation.slides.add();
  await addBase(slide, 7, "Publicity guidelines and controls", 8);
  await addImage(slide, "image51.jpeg", { left: 1120, top: 285, width: 660, height: 590 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Communications planning photograph from the supplied reference presentation",
  });
  addHeading(slide, "Publication scope", { left: 110, top: 280, width: 550, height: 40 });
  addBullets(slide, [
    "Recruitment Drive and free yoga classes",
    "Weekly classes and certification",
    "Members’ Night, lucky draw and merchandise",
    "Meditation and other approved programmes",
  ], { left: 110, top: 335, width: 850, height: 310 }, { fontSize: 28 });
  addHeading(slide, "Release control", { left: 110, top: 680, width: 550, height: 40 });
  addBullets(slide, [
    "Verify dates, venues, fees, payment instructions and participation requirements with the Secretary and Treasurer",
    "Update materials when confirmed arrangements change",
    "Archive posters, captions, announcements and registration links",
  ], { left: 110, top: 735, width: 900, height: 230 }, { fontSize: 25, spaceAfterPoints: 7 });
  notes(slide, "Item 6");
}

// 9. Certification policy
{
  const slide = presentation.slides.add();
  await addBase(slide, 8, "Weekly class certification policy", 9);
  await addImage(slide, "image53.png", { left: 1210, top: 275, width: 520, height: 420 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Yoga certificates from the supplied reference presentation",
  });
  addText(slide, "80%", { left: 110, top: 292, width: 270, height: 120 }, {
    typeface: FONT_BOLD, fontSize: 78, bold: true, color: C.tan,
  });
  addText(slide, "proposed attendance threshold for a semester-based Certificate of Completion (结业证书)", { left: 375, top: 305, width: 730, height: 110 }, {
    fontSize: 29, color: C.ink,
  });
  addHeading(slide, "Semester model", { left: 110, top: 485, width: 500, height: 40 });
  addBullets(slide, [
    "Approximately 14 hours or 14 class days during a long semester",
    "Short-semester attendance may accumulate towards the long-semester requirement",
  ], { left: 110, top: 540, width: 930, height: 185 }, { fontSize: 27 });
  addHeading(slide, "Alternative yearly model", { left: 110, top: 755, width: 600, height: 40 });
  addText(slide, "Approximately 35 hours, 35 days or 35 yoga practice classes", { left: 110, top: 810, width: 800, height: 70 }, {
    fontSize: 30, color: C.ink,
  });
  addHeading(slide, "Still to confirm", { left: 1200, top: 745, width: 500, height: 40 });
  addText(slide, "Signatories: Ms Chew, Ms Xinle, Club Chairperson, and Ms Belinda, Ms Heng or Core Advisor. Confirm permitted Yoga Club, DSA and UTAR logos.", { left: 1200, top: 800, width: 560, height: 145 }, {
    fontSize: 24, color: C.ink,
  });
  notes(slide, "Item 7");
}

// 10. Committee recruitment
{
  const slide = presentation.slides.add();
  await addBase(slide, 9, "Higher committee recruitment", 10);
  await addImage(slide, "image23.jpeg", { left: 1110, top: 270, width: 650, height: 650 }, {
    fit: "cover", geometry: "ellipse", alt: "Yoga class community photograph from the supplied reference presentation",
  });
  addHeading(slide, "Priority vacancies", { left: 110, top: 285, width: 600, height: 40 });
  addText(slide, "VICE CHAIRPERSON", { left: 110, top: 360, width: 650, height: 68 }, {
    typeface: FONT_BOLD, fontSize: 40, bold: true, color: C.dark,
  });
  addText(slide, "Identify a suitable current committee member with a strong student network.", { left: 110, top: 435, width: 820, height: 90 }, {
    fontSize: 28, color: C.ink,
  });
  addText(slide, "PUBLICITY TEAM", { left: 110, top: 585, width: 650, height: 68 }, {
    typeface: FONT_BOLD, fontSize: 40, bold: true, color: C.dark,
  });
  addText(slide, "Recruit one additional member to support the coming programme.", { left: 110, top: 660, width: 820, height: 90 }, {
    fontSize: 28, color: C.ink,
  });
  addHeading(slide, "Recruitment route", { left: 110, top: 820, width: 500, height: 40 });
  addText(slide, "Committee referrals and suitable EA appointments for experience-based instructor, committee and event-support roles", { left: 110, top: 870, width: 900, height: 85 }, {
    fontSize: 24, color: C.brown,
  });
  notes(slide, "Item 8");
}

// 11. Merchandise
{
  const slide = presentation.slides.add();
  await addBase(slide, 10, "New Yoga Club merchandise", 11);
  await addImage(slide, "image38.png", { left: 1230, top: 250, width: 590, height: 620 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Yoga merchandise concepts from the supplied reference presentation",
  });
  addHeading(slide, "Higher-end", { left: 110, top: 280, width: 420, height: 40 });
  addText(slide, "Logo cape or cloak, yoga mat, yoga bricks", { left: 110, top: 330, width: 970, height: 70 }, {
    fontSize: 29, color: C.ink,
  });
  addHeading(slide, "Middle-range", { left: 110, top: 465, width: 420, height: 40 });
  addText(slide, "Yoga brick, T-shirt, hat, stretching strap", { left: 110, top: 515, width: 970, height: 70 }, {
    fontSize: 29, color: C.ink,
  });
  addHeading(slide, "Lower-cost", { left: 110, top: 650, width: 420, height: 40 });
  addText(slide, "Keychain, lanyard, rubber wristband, other suitable small items", { left: 110, top: 700, width: 980, height: 85 }, {
    fontSize: 28, color: C.ink,
  });
  addText(slide, "Decision", { left: 110, top: 855, width: 220, height: 45 }, {
    typeface: FONT_BOLD, fontSize: 25, bold: true, color: C.tan,
  });
  addText(slide, "Every committee member selects at least three preferred items. Publicity designs the shortlist, the Secretary updates the proposal and the Treasurer confirms supplier pricing and pre-funding.", { left: 310, top: 848, width: 1060, height: 115 }, {
    fontSize: 24, color: C.dark,
  });
  notes(slide, "Item 9");
}

// 12. Members' Night
{
  const slide = presentation.slides.add();
  await addBase(slide, 11, "Members’ Night and lucky draw", 12);
  addText(slide, "PURPOSE", { left: 110, top: 295, width: 300, height: 46 }, {
    typeface: FONT_BOLD, fontSize: 25, bold: true, color: C.tan,
  });
  addText(slide, "Recognise regular June-class participants and welcome new club members.", { left: 110, top: 355, width: 870, height: 110 }, {
    typeface: FONT_BOLD, fontSize: 37, bold: true, color: C.dark,
  });
  addHeading(slide, "Possible member benefits", { left: 110, top: 545, width: 620, height: 40 });
  addBullets(slide, [
    "Free food or health-oriented gifts",
    "Selected club merchandise in the lucky draw pool",
    "Feedback survey and promotion of upcoming yoga classes",
  ], { left: 110, top: 600, width: 860, height: 260 }, { fontSize: 28 });
  addRect(slide, { left: 1130, top: 300, width: 550, height: 550 }, C.pale, "budget-field");
  addText(slide, "BUDGET CONDITION", { left: 1190, top: 380, width: 430, height: 45 }, {
    typeface: FONT_BOLD, fontSize: 24, bold: true, color: C.tan,
  });
  addText(slide, "Club funding for the lucky draw requires budget review and approval before commitment.", { left: 1190, top: 465, width: 430, height: 210 }, {
    typeface: FONT_BOLD, fontSize: 36, bold: true, color: C.dark,
  });
  notes(slide, "Items 2.3 and 10");
}

// 13. Meditation and practice classes
{
  const slide = presentation.slides.add();
  await addBase(slide, 12, "Meditation and practice-class pilot", 13);
  await addImage(slide, "image31.png", { left: 1090, top: 265, width: 700, height: 650 }, {
    fit: "cover", geometry: "roundRect", borderRadius: "rounded-2xl", alt: "Meditation class photograph from the supplied reference presentation",
  });
  addHeading(slide, "Pilot format", { left: 110, top: 280, width: 520, height: 40 });
  addBullets(slide, [
    "Meditation, additional yoga practice, or both",
    "One selected day or a weekly format",
    "Potentially delivered fully by student instructors",
    "Location and participation fee remain to be confirmed",
  ], { left: 110, top: 335, width: 850, height: 330 }, { fontSize: 28 });
  addHeading(slide, "Service certification", { left: 110, top: 720, width: 600, height: 40 });
  addText(slide, "Yearly certificates may record the number of events, total service hours and the individual’s role in instruction, organisation or publicity.", { left: 110, top: 775, width: 880, height: 145 }, {
    fontSize: 28, color: C.ink,
  });
  notes(slide, "Item 12");
}

// 14. Open matters
{
  const slide = presentation.slides.add();
  await addBase(slide, 13, "Open matters and governance checks", 14);
  addKeyLine(slide, "Yoga Day", "Consult Ms Belinda on information and approval requirements for a larger-scale event.", 300);
  addKeyLine(slide, "Attendance", "Agree an improved recording method for weekly classes and other activities.", 430);
  addKeyLine(slide, "Website", "Determine which record, fee and receipt functions should move into the proposed club website.", 560);
  addKeyLine(slide, "Coverage", "Identify any additional activity, meeting or matter requiring a proposal or semester report.", 690);
  addKeyLine(slide, "Feedback", "Invite final suggestions and unresolved departmental concerns.", 820);
  notes(slide, "Items 3.9–3.12 and 11");
}

// 15. Close and action capture
{
  const slide = presentation.slides.add();
  slide.background.fill = C.dark;
  await addImage(slide, "image2.png", { left: 1280, top: 0, width: 640, height: 640 }, {
    fit: "contain", alt: "Warm dotted texture from the supplied reference presentation",
  });
  await addImage(slide, "image4.png", { left: 1390, top: 280, width: 330, height: 330 }, {
    fit: "contain", alt: "UTAR Yoga Society logo from the supplied reference presentation",
  });
  addText(slide, "BEFORE WE CLOSE", { left: 120, top: 145, width: 800, height: 60 }, {
    typeface: FONT_BOLD, fontSize: 29, bold: true, color: C.sand,
  });
  addText(slide, "Confirm the decision,\nowner and due date", { left: 120, top: 250, width: 990, height: 210 }, {
    typeface: FONT_BOLD, fontSize: 67, bold: true, color: C.white,
  });
  addBullets(slide, [
    "Secretary updates the proposal and report checklist",
    "Departments verify their assigned actions",
    "Any approval dependency is recorded explicitly",
    "Meeting minutes capture the final decisions",
  ], { left: 120, top: 565, width: 990, height: 330 }, { fontSize: 29, color: C.white, spaceAfterPoints: 12 });
  addText(slide, "UTAR SL YOGA CLUB, SUNGAI LONG CAMPUS", { left: 120, top: 982, width: 700, height: 34 }, {
    fontSize: 19, color: C.sand,
  });
  notes(slide, "meeting close and action-recording requirement derived from Items 3, 4, 6 and 11");
}

await fs.mkdir(stagingDir, { recursive: true });
await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });

// Private previews support visual verification before delivery.
const previewDir = path.join(workspaceDir, ".codex-build/exco-deck/detailed-draft-render");
await fs.mkdir(previewDir, { recursive: true });
for (let index = 0; index < presentation.slides.count; index += 1) {
  const slide = presentation.slides.getItem(index);
  const preview = await presentation.export({ slide, format: "png", scale: 0.8 });
  await fs.writeFile(path.join(previewDir, `slide-${index + 1}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const candidatePath = path.join(stagingDir, "candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);

const requirements = {
  explicitTotalSlideCount: 15,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
};
const fontPolicy = {
  basis: "reference",
  families: [FONT, FONT_BOLD],
  referencePath,
  referenceSha256: "6777517c2c64f929ec57f8a5a00e5975a4eb5aff379c597c70d6907f4d85c803",
};
const result = await finalizePresentation({
  ...requirements,
  workspaceDir,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: [
    "--expected-slide-size-emu", "18288000,10287000",
    "--validate-bullet-geometry",
    "--validate-heading-fit",
  ],
  requiredNativeTableOwnerSlides: [],
  fontPolicy,
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, `${path.basename(FINAL_PPTX)}.validation.json`),
});

console.log(JSON.stringify({ finalPath: FINAL_PPTX, candidatePath, previewDir, result }, null, 2));
