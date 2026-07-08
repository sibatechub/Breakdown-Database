/*********** IT HEAD DETAILS (GLOBAL) ***********/
const IT_HEAD_NAME   = "Santosh Pradhan";
const IT_HEAD_MOBILE = "8249482016";
const IT_HEAD_EMAIL  = "systemadmin@driems.ac.in";

/*********** TECHNICIAN STATUS FORM (GLOBAL) ***********/
const TECH_FORM_BASE =
  "https://docs.google.com/forms/d/e/1FAIpQLSfBBacWI4YWI8x3Q7RbLZ6ssf1cmknqXYnHWPxuPGFDmkiCwg/viewform?usp=pp_url";

/*********** IMMEDIATE ESCALATION MAIL GROUP ***********/
const IMMEDIATE_TECH_EMAILS = [
  "Satchidanandakar@driems.ac.in",
  "shibaprasadmahanta@driems.ac.in"
  // add more as needed
];
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  if (!e || !e.range) {
    lock.releaseLock();
    return;
  }

  if (e.range.getSheet().getName() !== "Master") {
    lock.releaseLock();
    return;
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Master");

  const row = e.range.getRow();

// ===== SAFE STATIC TICKET =====
let ticketNo = sheet.getRange(row, 2).getValue();

if (!ticketNo) {

  const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "ddMMyyHHmmss");

  const buildingCode = String(sheet.getRange(row,9).getValue()).slice(-4);
  const floorCode    = String(sheet.getRange(row,10).getValue()).slice(-3);
  const assetCode    = String(sheet.getRange(row,12).getValue()).slice(0,4);

  // ===== PERMANENT GLOBAL COUNTER =====
const scriptProps = PropertiesService.getScriptProperties();

let counter = scriptProps.getProperty("TICKET_COUNTER");

if (!counter) {
  counter = 1;
} else {
  counter = parseInt(counter, 10) + 1;
}

// SAVE UPDATED VALUE
scriptProps.setProperty("TICKET_COUNTER", counter);

// FORMAT → 0001, 0002, 0003
const uniquePart = String(counter).padStart(4, "0");

  ticketNo = timestamp + buildingCode + floorCode + assetCode + uniquePart;

  sheet.getRange(row,2).setValue(ticketNo);
}

  // ===== Ticket Details =====
  const priority   = sheet.getRange(row, 4).getValue(); // D
  const staffName  = sheet.getRange(row, 6).getValue(); // F
  const staffEmail = sheet.getRange(row, 8).getValue(); // H
  const mobile     = sheet.getRange(row, 7).getValue(); // G
  const building   = sheet.getRange(row, 9).getValue(); // I
  const floor      = sheet.getRange(row,10).getValue(); // J
  const room       = sheet.getRange(row,11).getValue(); // K
  const assetType  = sheet.getRange(row,12).getValue(); // L
  const assetNo    = sheet.getRange(row,13).getValue(); // M
  const problem    = sheet.getRange(row,14).getValue(); // N

  // ===== Technician Override Check (UNCHANGED) =====
  const overrideSheet = SpreadsheetApp.getActive()
    .getSheetByName("WorkLoad");

  const overrideData = overrideSheet
    .getRange(2, 1, overrideSheet.getLastRow() - 1, 10)
    .getValues();

  const today = new Date();
  let overrideTechName="", overrideTechMobile="", overrideTechEmail="", overrideReason="", originalTechName="";

  for (let i = 0; i < overrideData.length; i++) {
    if (
      overrideData[i][1] === building &&
      overrideData[i][8].toString().toUpperCase() === "YES" &&
      today >= new Date(overrideData[i][5]) &&
      today <= new Date(overrideData[i][6])
    ) {
      originalTechName   = overrideData[i][2]; 
      overrideTechName   = overrideData[i][3];
      overrideTechMobile = overrideData[i][4];
      overrideTechEmail  = overrideData[i][9];
      overrideReason     = overrideData[i][7];
      break;
    }
  }


  // ===== Technician Mapping (UNCHANGED) =====
  const mapSheet = SpreadsheetApp.getActive()
    .getSheetByName("Building_Technician_Map");

  const mapData = mapSheet.getRange(2,1,mapSheet.getLastRow()-1,6).getValues();

  let techName="", techEmail="", techMobile="";
  let allTechEmails=[];

  for (let i = 0; i < mapData.length; i++) {
    if (
      mapData[i][1] === building &&
      mapData[i][5].toString().toUpperCase() === "YES"
    ) {
      allTechEmails.push(mapData[i][3]);

      if (!techEmail) {
        techName   = mapData[i][2];
        techEmail  = mapData[i][3];
        techMobile = mapData[i][4];
      }
    }
  }

  if (overrideTechName) {
    techName   = overrideTechName;
    techMobile = overrideTechMobile;
    techEmail  = overrideTechEmail; 
  }

  // ===== Write Technician Details =====
  sheet.getRange(row,16).setValue(techName);
  sheet.getRange(row,17).setValue(techEmail);
  sheet.getRange(row,18).setValue(techMobile);

  // ===== Auto Status OPEN (UNCHANGED) =====
  sheet.getRange(row,3).setValue("OPEN");

  // ===== Technician Status Form Prefill Link (UNCHANGED STRUCTURE) =====
  const formBaseUrl = "https://docs.google.com/forms/d/e/1FAIpQLSfBBacWI4YWI8x3Q7RbLZ6ssf1cmknqXYnHWPxuPGFDmkiCwg/viewform?usp=pp_url";

  const ticketFieldId      = "entry.327371081";
  const techNameFieldId   = "entry.748529779";
  const techMobileFieldId = "entry.1950349758";
  const techEmailFieldId  = "entry.1640227390";

  const preFilledLink =
    `${formBaseUrl}` +
    `&${ticketFieldId}=${encodeURIComponent(ticketNo)}` +
    `&${techNameFieldId}=${encodeURIComponent(techName)}` +
    `&${techMobileFieldId}=${encodeURIComponent(techMobile)}` +
    `&${techEmailFieldId}=${encodeURIComponent(techEmail)}`;

  // ===== Mail Subject =====
  const subject = `IT Ticket ${ticketNo} | ${building}`;
  const WEBAPP_URL =
"https://script.google.com/a/macros/driems.ac.in/s/AKfycbxZKMgefaveR5xFotc-6ji8XgdumFbiLwJmPyj1J0auKmuhD3lPsxkW2ZTJ-i44uHs9/exec";

const trackingLink =
WEBAPP_URL + "?ticket=" + encodeURIComponent(ticketNo.trim());
// ===== COMPUTER DATABASE LINK =====

const computerNo = assetNo;

const inventorySS = SpreadsheetApp.openById(
  "1w7UhKvXK_zk-M3qXG0cCrc_FVdW3eOQ4Pei7lwiSy5A"
);

const inventorySheet =
  inventorySS.getSheetByName("Master");

const inventoryData =
  inventorySheet.getDataRange().getValues();

let computerLink = trackingLink;

for (let j = 1; j < inventoryData.length; j++) {

  const dbComputer =
    String(inventoryData[j][3]).trim();

  if (dbComputer == String(computerNo).trim()) {

    computerLink =
      inventoryData[j][17];

    break;
  }
}

  // ===== Mail Body (PRESERVED LINE-BY-LINE) =====
// ===== HTML MAIL TEMPLATE =====
const template = HtmlService.createTemplateFromFile("mail");

// PASS VARIABLES
template.staffName = staffName;
template.ticketNo = ticketNo;
template.priority = priority;
template.mobile = mobile;
template.building = building;
template.floor = floor;
template.room = room;
template.assetType = assetType;
template.assetNo = assetNo;
template.problem = problem;
template.techName = techName;
template.techMobile = techMobile;
template.overrideReason = overrideReason;
template.originalTechName = originalTechName;
template.preFilledLink = preFilledLink;
template.trackingLink = trackingLink;
template.computerLink = computerLink;
template.itHeadName = IT_HEAD_NAME;
template.itHeadMobile = IT_HEAD_MOBILE;

// GENERATE HTML
const htmlBody = template.evaluate().getContent();

 // ===== MAIL ROUTING =====
 let toMail = staffEmail;  // Assigned technician 
 let ccMail = techEmail + "," + IT_HEAD_EMAIL;

// 🔥 IMMEDIATE PRIORITY ESCALATION
 if (priority && priority.toUpperCase() === "IMMEDIATE") {
  toMail = staffEmail; 
  ccMail = IMMEDIATE_TECH_EMAILS.join(",") + "," + IT_HEAD_EMAIL + "," + techEmail;
}

  GmailApp.sendEmail(
  toMail,
  subject,
  "Your email client does not support HTML.",
  {
    htmlBody: htmlBody,
    cc: ccMail
  }
);

  // ===== Fetch Thread ID (UNCHANGED) =====
  Utilities.sleep(8000);
  const threads = GmailApp.search(`subject:"IT Ticket ${ticketNo}" from:me`, 0, 1);
  sheet.getRange(row,5).setValue(threads.length ? threads[0].getId() : "THREAD_PENDING");

  lock.releaseLock();
}
function onTechnicianFormSubmit(e) {
  try {
    if (!e || !e.namedValues) return;

    const ticketNo   = e.namedValues["Ticket Number"][0];
    const newStatus  = e.namedValues["Status"][0];
    const comments = e.namedValues["Comments / Notes (Only Use for Closed or Pending)"]?.[0] || "";
    const techName   = e.namedValues["Technician Name"][0];
    const techMobile = e.namedValues["Mobile No"][0];

    const master = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Master");

    const data = master.getRange(
      2, 1,
      master.getLastRow() - 1,
      master.getLastColumn()
    ).getValues();

    for (let i = 0; i < data.length; i++) {

      if (data[i][1] === ticketNo) {

        const row = i + 2;
        const threadId = master.getRange(row,5).getValue();
        const now = new Date();

        // ===== ACCEPTED → NO MAIL =====
        if (newStatus.replace(/\s/g,"").toUpperCase() === "ACCEPTED") {
         master.getRange(row,3).setValue("ACCEPTED");
         master.getRange(row,19).setValue(now);
         return;
        }

        // ===== PENDING → MAIL ONLY IT HEAD (FULL DETAILS) =====
if (newStatus.toUpperCase() === "PENDING") {

  master.getRange(row, 3).setValue("PENDING");
  // ✅ SAVE COMMENT FOR PENDING ONLY
  master.getRange(row, 22).setValue(comments);

  // ===== READ FULL TICKET DETAILS FROM MASTER =====
  const priority   = master.getRange(row, 4).getValue();
  const staffName  = master.getRange(row, 6).getValue();
  const mobile     = master.getRange(row, 7).getValue();
  const staffEmail = master.getRange(row, 8).getValue();
  const building   = master.getRange(row, 9).getValue();
  const floor      = master.getRange(row,10).getValue();
  const room       = master.getRange(row,11).getValue();
  const assetType  = master.getRange(row,12).getValue();
  const assetNo    = master.getRange(row,13).getValue();
  const problem    = master.getRange(row,14).getValue();
  const techName   = master.getRange(row,16).getValue();
  const techMobile = master.getRange(row,18).getValue();

// ===== IT HEAD REASSIGNMENT FORM PREFILL LINK =====
const reassignBaseUrl =
  "https://docs.google.com/forms/d/e/1FAIpQLSfViYbmWTHIajGkhtk-2stXb3Ye1vkA6dPDm3Knw8mRtRNwGg/viewform?usp=pp_url";

const reassignTicketField = "entry.441357556";
const reassignTechField   = "entry.1098786949";

const reassignLink =
  `${reassignBaseUrl}` +
  `&${reassignTicketField}=${encodeURIComponent(ticketNo)}` +
  `&${reassignTechField}=${encodeURIComponent(techName)}`;
  // ===== MAIL BODY =====
  const template = HtmlService.createTemplateFromFile("pending_mail");

template.ticketNo = ticketNo;
template.staffName = staffName;
template.mobile = mobile;
template.building = building;
template.floor = floor;
template.room = room;
template.assetType = assetType;
template.assetNo = assetNo;
template.problem = problem;
template.techName = techName;
template.techMobile = techMobile;
template.comments = comments;
template.reassignLink = reassignLink;

const htmlBody = template.evaluate().getContent();

GmailApp.sendEmail(
  IT_HEAD_EMAIL,
  `PENDING: IT Ticket ${ticketNo}`,
  "HTML required",
  { htmlBody: htmlBody }
);

  return;
}
        

        // ===== CLOSED → DURATION + THREAD REPLY =====
        if (newStatus.replace(/\s/g,"").toUpperCase() === "CLOSED") {
          const startTime = master.getRange(row,1).getValue();
          const diffMs = now - startTime;

          const days  = Math.floor(diffMs / (1000*60*60*24));
          const hours = Math.floor((diffMs / (1000*60*60)) % 24);
          const mins  = Math.floor((diffMs / (1000*60)) % 60);

          const duration = `${days}D:${hours}H:${mins}M`;

          master.getRange(row,3).setValue("CLOSED");
          master.getRange(row,20).setValue(now);
          master.getRange(row,21).setValue(duration);
          master.getRange(row,22).setValue(comments);
          // ===== FEEDBACK FORM PREFILL LINK =====
const feedbackBaseUrl =
  "https://docs.google.com/forms/d/e/1FAIpQLScxKSPaW54d4l0A8s6UUkv20WQSPuoXpZU5gOSDYAJQvCNPHg/viewform?usp=pp_url";

const feedbackTicketField = "entry.1832990317";
const feedbackTechField   = "entry.1945085659";

const feedbackLink =
  `${feedbackBaseUrl}` +
  `&${feedbackTicketField}=${encodeURIComponent(ticketNo)}` +
  `&${feedbackTechField}=${encodeURIComponent(techName)}`;

          // ===== GET STAFF NAME =====
const staffName = master.getRange(row,6).getValue();

// ===== HTML TEMPLATE =====
const template = HtmlService.createTemplateFromFile("closed_mail");

template.staffName = staffName;
template.ticketNo = ticketNo;
template.techName = techName;
template.techMobile = techMobile;
template.duration = duration;
template.comments = comments;
template.feedbackLink = feedbackLink;

const htmlBody = template.evaluate().getContent();

// ===== SEND MAIL IN SAME THREAD =====
if (threadId) {
  GmailApp.getThreadById(threadId).replyAll("", {
    htmlBody: htmlBody
  });
}
          return;
        }
      }
    }
  } catch (err) {
    Logger.log("Technician form submit error: " + err.message);
  }
}
function onFeedbackFormSubmit(e) {
  try {
    if (!e || !e.namedValues) return;

    const getVal = (key) =>
      e.namedValues[key]?.[0] || "";

    const ticketNo   = getVal("Ticket Number");
    const overallSat = getVal("Overall Satisfaction");

    // 🔒 ADD HERE
    if (!ticketNo || !overallSat) return;

    const techBehav  = getVal("Technician Behaviour");
    const resolution = getVal("Resolution Time Satisfaction");
    const addComment = getVal("Additional Comments / Suggestions");
    const feedbackTime = getVal("Timestamp");

    const master = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Master");

    const lastRow = master.getLastRow();
    if (lastRow < 2) return;

    const ticketCol = master.getRange(2, 2, lastRow - 1, 1).getValues();

    for (let i = 0; i < ticketCol.length; i++) {
      if (ticketCol[i][0] === ticketNo) {
        const row = i + 2;

        master.getRange(row, 23).setValue(overallSat); // W
        master.getRange(row, 24).setValue(techBehav);  // X
        master.getRange(row, 25).setValue(resolution); // Y
        master.getRange(row, 26).setValue(addComment); // Z
        master.getRange(row, 31).setValue(feedbackTime);

        break;
      }
    }

  } catch (err) {
    Logger.log("Feedback form submit error: " + err.message);
  }
}
function onReassignFormSubmit(e) {
  try {
    if (!e || !e.namedValues) return;

    const get = k => e.namedValues[k]?.[0] || "";

    // ===== REASSIGN FORM DATA =====
    const ticketNo     = get("Ticket Number");
    const newTechName  = get("New Technician Name");
    const newTechEmail = get("New Technician Email");
    const newTechMob   = get("New Technician Mobile");
    const remark       = get("IT Head Remarks / Instructions");

    if (!ticketNo || !newTechEmail) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName("Master");

    const lastRow = master.getLastRow();
    const tickets = master.getRange(2, 2, lastRow - 1, 1).getValues(); // Column B

    for (let i = 0; i < tickets.length; i++) {
      if (tickets[i][0] === ticketNo) {
        const row = i + 2;

        // ===== READ EXISTING TICKET DETAILS =====
        const staffName = master.getRange(row, 6).getValue();
        const staffMob  = master.getRange(row, 7).getValue();
        const building  = master.getRange(row, 9).getValue();
        const floor     = master.getRange(row,10).getValue();
        const room      = master.getRange(row,11).getValue();
        const assetType = master.getRange(row,12).getValue();
        const assetNo   = master.getRange(row,13).getValue();
        const issue     = master.getRange(row,14).getValue();
        const oldTech   = master.getRange(row,16).getValue();
        const oldMob    = master.getRange(row,18).getValue();

        // ===== UPDATE MASTER =====
        master.getRange(row, 3).setValue("REOPEN");          // Status
        master.getRange(row,27).setValue(newTechName);       // AA
        master.getRange(row,28).setValue(newTechEmail);      // AB
        master.getRange(row,29).setValue(newTechMob);        // AC
        master.getRange(row,30).setValue(remark);            // AD

        // ===== TECHNICIAN STATUS PREFILL LINK =====
        const techPrefillLink =
          `${TECH_FORM_BASE}` +
          `&entry.327371081=${encodeURIComponent(ticketNo)}` +
          `&entry.748529779=${encodeURIComponent(newTechName)}` +
          `&entry.1950349758=${encodeURIComponent(newTechMob)}` +
          `&entry.1640227390=${encodeURIComponent(newTechEmail)}`;

        // ===== MAIL TO NEW TECHNICIAN ONLY =====
        const template = HtmlService.createTemplateFromFile("reassign_mail");

template.ticketNo = ticketNo;
template.newTechName = newTechName;
template.staffName = staffName;
template.staffMob = staffMob;
template.building = building;
template.floor = floor;
template.room = room;
template.assetType = assetType;
template.assetNo = assetNo;
template.issue = issue;
template.oldTech = oldTech;
template.oldMob = oldMob;
template.remark = remark;
template.techPrefillLink = techPrefillLink;

const htmlBody = template.evaluate().getContent();

GmailApp.sendEmail(
  newTechEmail,
  `REOPENED: IT Ticket ${ticketNo}`,
  "HTML required",
  { htmlBody: htmlBody }
);

        break;
      }
    }
  } catch (err) {
    Logger.log("Reassign error: " + err.message);
  }
}
function doGet(e) {
  try {

    const ticketNo = (e && e.parameter && e.parameter.ticket)
      ? e.parameter.ticket.toString().trim()
      : "";

    // ===== MANUAL ENTRY PAGE (FINAL FIXED FOR /dev + /exec) =====
    if (!ticketNo) {

      return HtmlService

        .createHtmlOutputFromFile("index")

        .setSandboxMode(HtmlService.SandboxMode.IFRAME)

        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    }

    // ===== FETCH DATA =====
    const sheet = SpreadsheetApp
    .openById("1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524")
    .getSheetByName("Master");

    const data = sheet.getDataRange().getValues();

    let foundRow = null;

    // ===== STRONG MATCH =====
    for (let i = 1; i < data.length; i++) {

      const sheetTicket = String(data[i][1]).replace(/\s/g, '').toUpperCase();
      const inputTicket = ticketNo.replace(/\s/g, '').toUpperCase();

      if (sheetTicket === inputTicket) {
        foundRow = data[i];
        break;
      }
    }

    // ===== NOT FOUND =====
    if (!foundRow) {
      return HtmlService.createHtmlOutput(`
        <html>
        <body style="font-family:Arial;text-align:center;padding:50px;">
          <h2 style="color:red;">❌ Ticket Not Found</h2>
          <p><b>Entered:</b> ${ticketNo}</p>
          <button onclick="window.history.back()">🔙 Go Back</button>
        </body>
        </html>
      `)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // ===== SUCCESS =====
    return HtmlService
      .createHtmlOutput(generateHTML(foundRow))
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {

    return HtmlService.createHtmlOutput(`
      <html>
      <body style="font-family:Arial;text-align:center;padding:50px;">
        <h2 style="color:red;">⚠️ SYSTEM ERROR</h2>
        <p>${err.message}</p>
      </body>
      </html>
    `);
  }
}
function generateHTML(row) {

  function formatDate(dateStr) {
    if (!dateStr) return "Waiting";
    const d = new Date(dateStr);
    return d.toLocaleString("en-IN", {
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: true
    }).replace(",", " |");
  }

  function getEmoji(rating) {
    if (rating == 1) return "😡";
    if (rating == 2) return "😕";
    if (rating == 3) return "🙂";
    if (rating == 4) return "😃";
    if (rating == 5) return "🤩";
    return "";
  }

  function getColor(rating) {
    if (rating <= 2) return "#ef4444";
    if (rating == 3) return "#f59e0b";
    return "#22c55e";
  }

  const isAccepted = row[18];
  const isClosed   = row[19];
  const isFeedback = row[30];

  const statusText =
    isFeedback ? "COMPLETED" :
    isClosed ? "CLOSED" :
    isAccepted ? "ACCEPTED" :
    "OPEN";

  const statusColor =
    isFeedback ? "#22c55e" :
    isClosed ? "#3b82f6" :
    isAccepted ? "#f59e0b" :
    "#ef4444";

  const currentStep =
    isFeedback ? 4 :
    isClosed ? 3 :
    isAccepted ? 2 :
    1;

  return `
  <html>
  <head>
  <style>

    body {
  font-family: 'Inter', sans-serif;
  background:#ffffff;   /* ✅ full white page */
  display:flex;
  justify-content:center;
  align-items:flex-start; /* keep top spacing */
  padding:40px 20px;
}

    .card {
  width:390px;
  background:#ffffff;
  border-radius:16px;
  padding:20px;

  /* 🔥 THIS creates separation from white page */
  border:1px solid #e5e7eb;
  box-shadow:0 10px 25px rgba(0,0,0,0.12);
}

    .header {
      text-align:center;
      border-bottom:2px solid #E1A835;
      padding-bottom:10px;
      margin-bottom:15px;
    }

    .header h1 {
      margin:0;
      font-size:16px;
      color:#2E2278;
    }

    .header span {
      color:#E1A835;
      font-size:12px;
      font-weight:600;
    }

    .status {
      text-align:center;
      padding:6px;
      border-radius:20px;
      color:white;
      font-weight:600;
      background:${statusColor};
      margin:10px 0;
    }

    .timeline {
      position:relative;
      margin-left:20px;
      padding-left:20px;
    }

    .line {
      position:absolute;
      left:-2px;
      top:0;
      width:4px;
      height:100%;
      background:#ddd;
    }

    .active-line {
      position:absolute;
      left:-2px;
      top:0;
      width:4px;
      background:#22c55e;
      animation:growLine 1.5s ease forwards;
      height:0;
    }

    @keyframes growLine {
      from { height:0; }
      to { height: var(--line-height); }
    }

    .moving-dot {
      position:absolute;
      left:-7px;
      top:0;
      width:14px;
      height:14px;
      background:#E1A835;
      border-radius:50%;
      box-shadow:0 0 10px #E1A835, 0 0 20px rgba(225,168,53,0.6);
      animation:moveDot 1.5s ease forwards;
    }

    @keyframes moveDot {
      from { top:0; }
      to { top: var(--dot-height); }
    }

    .step {
      margin-bottom:25px;
      position:relative;
      animation:fadeUp 0.6s ease;
    }

    @keyframes fadeUp {
      from { opacity:0; transform:translateY(10px);}
      to { opacity:1; transform:translateY(0);}
    }

    .circle {
      position:absolute;
      left:-29px;
      top:2px;
      width:16px;
      height:16px;
      border-radius:50%;
      background:#ddd;
    }

    .active .circle {
      background:#E1A835;
    }

    .current .circle {
      background:#E1A835;
      animation:pulseGlow 1.5s infinite;
    }

    @keyframes pulseGlow {
      0% { box-shadow:0 0 0 0 rgba(225,168,53,0.7); }
      70% { box-shadow:0 0 0 12px rgba(225,168,53,0); }
      100% { box-shadow:0 0 0 0 rgba(225,168,53,0); }
    }

    .label {
      font-weight:600;
      color:#2E2278;
    }

    .time {
      font-size:12px;
      color:#666;
    }

    .tech-box {
      margin-top:8px;
      padding:10px;
      background:#f9fafb;
      border-radius:8px;
      border-left:4px solid #E1A835;
    }

    .call-btn {
      display:inline-block;
      margin-top:6px;
      padding:6px 12px;
      background:#2E2278;
      color:white;
      border-radius:6px;
      text-decoration:none;
      font-size:12px;
    }

    .rating {
      margin-top:15px;
      padding:12px;
      background:#f9fafb;
      border-radius:10px;
      border-left:4px solid #E1A835;
    }

    .rating-top {
      display:flex;
      align-items:center;
      gap:10px;
    }

    .emoji {
      font-size:24px;
      animation:bounce 1s infinite;
    }

    @keyframes bounce {
      0%,100%{transform:translateY(0);}
      50%{transform:translateY(-5px);}
    }

    .stars-container {
      position:relative;
      font-size:18px;
    }

    .stars-bg { color:#ddd; }

    .stars-fill {
      position:absolute;
      top:0;
      left:0;
      width:calc(var(--rating) * 20%);
      color:var(--color);
      overflow:hidden;
      white-space:nowrap;
      animation:fillStars 1s ease;
    }

    @keyframes fillStars {
      from {width:0;}
      to {width:calc(var(--rating) * 20%);}
    }

    .pdf-btn {
      width:100%;
      margin-top:15px;
      padding:10px;
      background:#E1A835;
      border:none;
      color:white;
      font-weight:600;
      border-radius:8px;
      cursor:pointer;
    }

    @media print {
      .pdf-btn { display:none; }
      body { background:white; }
    }

  </style>
  </head>

  <body>
  <div class="card">

<div style="
  background: radial-gradient(
    circle at center,
    rgba(255,255,255,0.06) 0%,
    transparent 40%
  ),
  linear-gradient(
    135deg,
    #1a174f 0%,
    #2E2278 40%,
    #3b2f8f 70%,
    #1f1b5c 100%
  );
  padding:16px 18px;
  border-radius:16px 16px 0 0;
  margin:-20px -20px 15px -20px;

  /* 🔥 THIS FIXES YOUR PROBLEM */
  box-shadow:0 8px 20px rgba(0,0,0,0.25);
  border-bottom:1px solid rgba(255,255,255,0.08);
">

  <table width="100%">
    <tr>

      <!-- LOGO -->
      <td>
        <a href="https://sites.google.com/view/driemsit/home" target="_blank">
          <img src="https://lh3.googleusercontent.com/d/1zMARUFlcRIdEzLKG60unqZRvVmhnODYD=w200"
               style="height:40px; display:block;">
        </a>
      </td>

      <!-- TEXT -->
      <td align="right" style="color:#ffffff;">
        <div style="
          font-size:16px;
          font-weight:700;
          letter-spacing:0.5px;
        ">
          IT Support
        </div>

        <div style="
          font-size:11px;
          opacity:0.85;
        ">
          Ticket Tracking System
        </div>
      </td>

    </tr>
  </table>

  <!-- 🔥 PREMIUM DIVIDER -->
  <div style="
    height:3px;
    background:linear-gradient(to right,#E1A835,#ffd27a,#E1A835);
    margin-top:12px;
    border-radius:2px;
  "></div>

</div>

  <div><b>Ticket ID:</b> ${row[1]}</div>

  <div class="status">${statusText}</div>

  <div class="timeline">
    <div class="line"></div>

    <div class="active-line" style="--line-height:${
      isFeedback ? "100%" :
      isClosed ? "75%" :
      isAccepted ? "50%" :
      "25%"
    }"></div>

    <div class="moving-dot" style="--dot-height:${
      isFeedback ? "100%" :
      isClosed ? "75%" :
      isAccepted ? "50%" :
      "25%"
    }"></div>

    <div class="step ${currentStep>=1?"active":""} ${currentStep==1?"current":""}">
      <div class="circle"></div>
      <div class="label">📝 Created</div>
      <div class="time">${formatDate(row[0])}</div>
    </div>

    <div class="step ${isAccepted?"active":""} ${currentStep==2?"current":""}">
      <div class="circle"></div>
      <div class="label">👨‍🔧 Accepted</div>
      <div class="time">${formatDate(row[18])}</div>

      ${isAccepted ? `
      <div class="tech-box">
        👨‍🔧 ${row[15]}<br>
        <a class="call-btn" href="tel:${row[17]}">📞 Call Technician</a>
      </div>` : ""}
    </div>

    <div class="step ${isClosed?"active":""} ${currentStep==3?"current":""}">
      <div class="circle"></div>
      <div class="label">🛠 Closed</div>
      <div class="time">${formatDate(row[19])}</div>
    </div>

    <div class="step ${isFeedback?"active":""} ${currentStep==4?"current":""}">
      <div class="circle"></div>
      <div class="label">⭐ Feedback</div>
      <div class="time">${formatDate(row[30])}</div>
    </div>

  </div>

  ${row[22] ? `
  <div class="rating">
    <div class="rating-top">
      <div class="emoji">${getEmoji(row[22])}</div>

      <div class="stars-container" style="--rating:${row[22]}; --color:${getColor(row[22])};">
        <div class="stars-bg">★★★★★</div>
        <div class="stars-fill">★★★★★</div>
      </div>

      (${row[22]}/5)
    </div>

    <br>😊 ${row[23]}
    <br>⏱ ${row[24]}
    <br>💬 ${row[25]}
  </div>
  ` : ""}

  <button class="pdf-btn" onclick="window.print()">⬇ Download PDF</button>

  </div>
  </body>
  </html>
  `;
}
