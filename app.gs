/**
 * ============================================================================
 * GOOGLE FORM ADVANCED APPROVAL WORKFLOW
 * ============================================================================
 * * This script transforms a simple Google Form into a multi-stage approval engine.
 * * FEATURES:
 * 1. Dynamic Routing: Routes based on "Department" and "Team" selection.
 * 2. State Management: Stores approval history in a JSON cell (no messy columns).
 * 3. Dashboard: A custom HTML Web App for approvers to view/action requests.
 * 4. Smart Notifications: Skips "Director" level instant emails (weekly digest instead).
 * * SETUP REQUIREMENTS:
 * - A Google Form with questions: "Employee Name", "Department", "Team", "Description", "Cost".
 * - A bound Google Sheet.
 * - Deploy as Web App to get the URL.
 */

// =====================
// 1. GLOBAL CONFIGURATION
// =====================

// While TEST_MODE = true, ALL emails go to TEST_EMAIL to prevent spamming real users during setup.
const TEST_MODE  = true; 
const TEST_EMAIL = "admin@example.com"; // <--- Put your email here for testing

// =====================
// 2. APPROVAL FLOWS
// =====================
// Key format: "Department|Team" 
// These keys MUST exactly match the dropdown options in your Google Form.

const FLOWS = {
  // --- ENGINEERING DEPARTMENT FLOWS ---
  "Engineering|Frontend Development": [
    {
      email: "manager.frontend@example.com",
      name: "Alice Manager",
      title: "Engineering Manager",
    },
    {
      email: "director.tech@example.com",
      name: "Bob Director",
      title: "Director of Technology", // Keywords "Director" triggers silent mode (digest only)
    },
  ],

  "Engineering|Backend Infrastructure": [
    // This team skips the manager and goes straight to the VP/Director
    {
      email: "director.tech@example.com",
      name: "Bob Director",
      title: "Director of Technology",
    },
  ],

  // --- SALES DEPARTMENT FLOWS ---
  "Sales|North America": [
    {
      email: "lead.na@example.com",
      name: "Charlie Lead",
      title: "Team Lead - NA",
    },
    {
      email: "vp.sales@example.com",
      name: "Diana VP",
      title: "VP of Sales",
    },
  ],

  "Sales|Europe": [
    {
      email: "lead.eu@example.com",
      name: "Evan Lead",
      title: "Team Lead - EU",
    },
    {
      email: "vp.sales@example.com",
      name: "Diana VP",
      title: "VP of Sales",
    },
  ],

  // --- FALLBACK ---
  // If the user selects a combo not defined above, it goes here.
  defaultFlow: [
    {
      email: "admin.support@example.com",
      name: "System Admin",
      title: "Routing Error - Please Review",
    },
  ],
};

// ====================================================
// APP CLASS (The Engine)
// ====================================================

function App() {
  this.form = FormApp.getActiveForm();
  
  // 🔴 IMPORTANT: After deploying as a Web App, paste that URL here:
  this.url = "https://script.google.com/macros/s/INSERT_YOUR_DEPLOYMENT_ID_HERE/exec"; 

  this.title = this.form ? this.form.getTitle() : "Approval Workflow";
  
  // Configuration: The exact headers in your Sheet (from Form Questions)
  this.sheetname = "Form Responses 1"; // Default Google Sheet name
  
  this.headers = {
    employee: "Employee Name",
    section: "Department",       // Matches key in FLOWS
    team: "Team",                // Matches key in FLOWS
    desc: "Description",         // Description of request
    cost: "Cost",                // Monetary value
    email: "Email Address"       // Auto-collected email
  };

  this.sectionHeader = this.headers.section;
  this.teamHeader = this.headers.team;

  // System Headers (Script will create these automatically)
  this.uidHeader = "Request ID";
  this.uidPrefix = "REQ-";
  this.uidLength = 5;
  this.statusHeader = "Current Status";
  this.responseIdHeader = "_response_id";
  
  // Status Constants
  this.pending = "Pending";
  this.approved = "Approved";
  this.rejected = "Rejected";
  this.waiting = "Waiting";

  // Role Detection Logic
  this.getRoleFromTitle = function (title) {
    if (!title) return "unknown";
    const t = String(title).toLowerCase();
    if (t.includes("director") || t.includes("vp")) return "director";
    if (t.includes("supervisor") || t.includes("lead")) return "supervisor";
    return "manager";
  };

  // "High Level" roles receive Weekly Digests instead of instant emails
  this.isHighLevel = function (title) {
    if (!title) return false;
    const t = String(title).toLowerCase();
    return t.includes("director") || t.includes("vp");
  };

  // ----------------------------------------------------
  //  DATA ACCESS LAYER
  // ----------------------------------------------------

  // Get the linked responses sheet (or create one if not set)
  this.sheet = (() => {
    try {
      const id = this.form.getDestinationId();
      return SpreadsheetApp.openById(id).getSheetByName(this.sheetname);
    } catch (e) {
      // Fallback if testing directly in Sheet without Form bound
      return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.sheetname);
    }
  })();

  this.parsedValues = () => {
    const values = this.sheet.getDataRange().getDisplayValues();
    return values.map((value) => {
      return value.map((cell) => {
        try {
          return JSON.parse(cell); // Try to parse JSON approval objects
        } catch (e) {
          return cell; // Return plain text if not JSON
        }
      });
    });
  };

  // ----------------------------------------------------
  //  DASHBOARD DATA FETCHER
  // ----------------------------------------------------
  this.getDashboardDataFor = (approverEmail) => {
    const values = this.parsedValues();
    const headers = values[0];
    
    if (!values || values.length < 2) return { pending: [], approved: [], rejected: [], roles: [] };

    const lowerEmail = String(approverEmail || "").toLowerCase();

    // Find Column Indices
    const tsIdx       = headers.indexOf("Timestamp");
    const empIdx      = headers.indexOf(this.headers.employee);
    const sectionIdx  = headers.indexOf(this.headers.section);
    const teamIdx     = headers.indexOf(this.headers.team);
    const descIdx     = headers.indexOf(this.headers.desc);
    const costIdx     = headers.indexOf(this.headers.cost);
    const statusIdx   = headers.indexOf(this.statusHeader);

    const pendingList  = [];
    const approvedList = [];
    const rejectedList = [];
    const foundTitles = new Set(); 

    const makeItem = (row, approverCell) => {
      const tsRaw = approverCell.timestamp || row[tsIdx];
      let updatedAt = 0, lastUpdated = "";

      if (tsRaw) {
        const d = new Date(tsRaw);
        if (!isNaN(d.getTime())) {
          updatedAt = d.getTime();
          lastUpdated = d.toLocaleString();
        }
      }

      return {
        taskId:       approverCell.taskId,
        employeeName: row[empIdx]      || "",
        section:      row[sectionIdx]  || "",
        team:         row[teamIdx]     || "",
        description:  row[descIdx]     || "",
        cost:         row[costIdx]     || "",
        overallStatus: row[statusIdx]  || "",
        actingTitle:  approverCell.title || "Approver", 
        lastUpdated,
        updatedAt
      };
    };

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.length === 0) continue;

      // Scan row for any JSON cell belonging to this approver
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (!cell || typeof cell !== "object") continue;
        if (!cell.email || !cell.status) continue;

        if (String(cell.email).toLowerCase() !== lowerEmail) continue;

        const status = String(cell.status);
        const item = makeItem(row, cell);

        if (status === this.pending) {
          pendingList.push(item);
          if (cell.title) foundTitles.add(cell.title);
        } else if (status === this.approved) {
          approvedList.push(item);
        } else if (status === this.rejected) {
          rejectedList.push(item);
        }
      }
    }

    // Sort by most recent
    const sortByRecent = (arr) => arr.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
    sortByRecent(pendingList);
    sortByRecent(approvedList);
    sortByRecent(rejectedList);

    const roleString = Array.from(foundTitles).join(" / ") || "Approver";

    return {
      pending:  pendingList,
      approved: approvedList.slice(0, 20),
      rejected: rejectedList.slice(0, 20),
      pendingTotal:  pendingList.length,
      approvedTotal: approvedList.length,
      rejectedTotal: rejectedList.length,
      approverTitle: roleString
    };
  };

  // ----------------------------------------------------
  //  CORE HELPERS
  // ----------------------------------------------------

  this.getTaskById = (id) => {
    const values = this.parsedValues();
    const headers = values[0];
    
    // Find row containing the specific taskId
    const recordIdx = values.findIndex((value) => value.some((cell) => cell && cell.taskId === id));
    if (recordIdx === -1) return {}; // Not found

    const record = values[recordIdx];
    const row = recordIdx + 1;

    // Map all standard fields for display
    const task = record.slice(0, headers.indexOf(this.statusHeader) + 1).map((item, i) => {
      let val = item;
      // If it's a file upload, make it a link
      if (headers[i] && headers[i].toLowerCase().includes("document")) {
         val = formatFileLinks(item);
      }
      return { label: headers[i], value: val };
    });
    
    const email = record[headers.indexOf(this.headers.email)];
    const status = record[headers.indexOf(this.statusHeader)];
    const responseId = record[headers.indexOf(this.responseIdHeader)];
    
    const approver = record.find((item) => item && item.taskId === id);
    const column = record.findIndex((item) => item && item.taskId === id) + 1;
    const nextApprover = record[column]; // The cell immediately to the right
    const approvers = record.filter((item) => item && item.taskId); // All approvers in chain

    return { email, status, responseId, task, approver, nextApprover, approvers, row, column, statusColumn: headers.indexOf(this.statusHeader) + 1 };
  };

  this.findRowByResponseId = (responseId) => {
    const values = this.parsedValues();
    const headers = values[0];
    const respIdx = headers.indexOf(this.responseIdHeader);
    for (let r = 1; r < values.length; r++) {
      if (values[r][respIdx] === responseId) return r + 1;
    }
    return -1;
  };

  this.createUid = () => {
    const props = PropertiesService.getDocumentProperties();
    let uid = Number(props.getProperty(this.uidHeader));
    if (!uid) uid = 1;
    props.setProperty(this.uidHeader, uid + 1);
    return this.uidPrefix + (uid + 10 ** this.uidLength).toString().slice(-this.uidLength);
  };

  // ----------------------------------------------------
  //  EMAILS
  // ----------------------------------------------------

  this.sendApproval = ({ task, approver, approvers, isEdit }) => {
    const template = HtmlService.createTemplateFromFile("approval_email.html");
    Object.assign(template, {
      title: this.title, task, approver, approvers, isEdit,
      approveUrl: `${this.url}?taskId=${approver.taskId}&action=approve`,
      rejectUrl: `${this.url}?taskId=${approver.taskId}&action=reject`,
      dashboardUrl: `${this.url}?dashboard=1&email=${encodeURIComponent(approver.email)}`
    });

    const to = TEST_MODE ? TEST_EMAIL : approver.email;
    GmailApp.sendEmail(to, `Approval Required - ${this.title}`, "", { htmlBody: template.evaluate().getContent() });
  };

  this.sendNotification = (taskId, stageMeta) => {
    const data = this.getTaskById(taskId);
    if (!data.email) return;

    const { email, responseId, status, task, approvers } = data;
    const role = stageMeta.role || "initial";
    const decision = stageMeta.decision || "pending";

    const template = HtmlService.createTemplateFromFile("notification_email.html");
    Object.assign(template, {
      title: this.title, task, status, approvers,
      approvalProgressUrl: `${this.url}?responseId=${responseId}`,
      stageRole: role, decision: decision
    });

    const subject = `Update: ${this.title} (${decision})`;
    const to = TEST_MODE ? TEST_EMAIL : email;
    GmailApp.sendEmail(to, subject, "", { htmlBody: template.evaluate().getContent() });
  };

  // ----------------------------------------------------
  //  FORM SUBMISSION HANDLER
  // ----------------------------------------------------
  
  this.onFormSubmit = () => {
    const values = this.parsedValues();
    const headers = values[0];

    // Get Form Response
    const responses = this.form.getResponses();
    const lastResponse = responses[responses.length - 1];
    const responseId = lastResponse.getId();
    const itemResponses = lastResponse.getItemResponses();

    // Check for edits
    const existingRow = this.findRowByResponseId(responseId);
    const isEdit = existingRow !== -1;
    let targetRow = isEdit ? existingRow : values.length;

    // Write Answers to Sheet (in case of Edit, overwrite)
    itemResponses.forEach((ir) => {
      const title = ir.getItem().getTitle();
      const colIndex = headers.indexOf(title) + 1;
      if (colIndex > 0) this.sheet.getRange(targetRow, colIndex).setValue(ir.getResponse());
    });

    // --- BUILD APPROVAL CHAIN ---
    // Look for where to start writing the approval columns
    let startColumn = headers.indexOf(this.uidHeader) + 1;
    if (startColumn === 0) startColumn = headers.length + 1; // Append to end if first run

    const newHeaders = [this.uidHeader, this.statusHeader, this.responseIdHeader];
    
    // Generate or Retrieve UID
    let uidValue = isEdit ? values[targetRow - 1][headers.indexOf(this.uidHeader)] : this.createUid();
    const newValues = [uidValue, this.pending, responseId];

    // Get Department/Team from Sheet
    const section = this.sheet.getRange(targetRow, headers.indexOf(this.sectionHeader) + 1).getValue();
    const team = this.sheet.getRange(targetRow, headers.indexOf(this.teamHeader) + 1).getValue();
    
    // Routing Logic
    const flowKey = `${section}|${team}`;
    let flow = FLOWS[flowKey];
    let isFallback = false;

    if (!flow) {
      flow = FLOWS.defaultFlow;
      isFallback = true;
    }

    let firstTaskId;

    // Create JSON objects for each approver
    flow.forEach((item, i) => {
      newHeaders.push("_approver_" + (i + 1));
      
      const approverObj = {
        email: item.email,
        name: item.name,
        title: item.title,
        taskId: Utilities.base64EncodeWebSafe(Utilities.getUuid()),
        timestamp: new Date(),
        status: (i === 0) ? this.pending : this.waiting,
        hasNext: (i < flow.length - 1),
        comments: (isFallback && i === 0) ? `⚠️ Routed via default: Invalid selection (${flowKey})` : null
      };

      if (i === 0) firstTaskId = approverObj.taskId;
      newValues.push(JSON.stringify(approverObj));
    });

    // Write Headers (Green background)
    this.sheet.getRange(1, startColumn, 1, newHeaders.length)
        .setValues([newHeaders])
        .setBackgroundColor("#34A853")
        .setFontColor("#FFFFFF");

    // Write Data Row
    this.sheet.getRange(targetRow, startColumn, 1, newValues.length).setValues([newValues]);

    // Send Initial Email
    if (isEdit) {
      this.sendNotification(firstTaskId, { role: "initial", decision: "pending" });
    } else {
      this.sendNotification(firstTaskId, { role: "initial", decision: "pending" });
    }

    // Trigger first approver
    const { task, approver, approvers } = this.getTaskById(firstTaskId);
    this.sendApproval({ task, approver, approvers, isEdit });
  };

  // ----------------------------------------------------
  //  ACTIONS (Approve/Reject)
  // ----------------------------------------------------

  this.approve = ({ taskId, comments }) => {
    const data = this.getTaskById(taskId);
    if (!data.approver) return;
    const { task, approver, approvers, nextApprover, row, column, statusColumn } = data;

    // Update Current Approver
    approver.comments = comments;
    approver.status = this.approved;
    approver.timestamp = new Date();
    this.sheet.getRange(row, column).setValue(JSON.stringify(approver));

    const role = this.getRoleFromTitle(approver.title);

    // Is there a next person?
    if (approver.hasNext && nextApprover) {
      // Notify Requester
      this.sendNotification(taskId, { role, decision: "approved" });

      // Activate Next Person
      nextApprover.status = this.pending;
      nextApprover.timestamp = new Date();
      this.sheet.getRange(row, column + 1).setValue(JSON.stringify(nextApprover));

      // SILENT MODE CHECK
      if (!this.isHighLevel(nextApprover.title)) {
        this.sendApproval({ task, approver: nextApprover, approvers });
      } else {
        console.log(`Silent Mode: Queued for digest -> ${nextApprover.email}`);
      }

    } else {
      // Final Approval
      this.sheet.getRange(row, statusColumn).setValue(this.approved);
      this.sendNotification(taskId, { role, decision: "approved" });
    }
  };

  this.reject = ({ taskId, comments }) => {
    const data = this.getTaskById(taskId);
    if (!data.approver) return;
    const { approver, row, column, statusColumn } = data;

    approver.comments = comments;
    approver.status = this.rejected;
    approver.timestamp = new Date();
    
    this.sheet.getRange(row, column).setValue(JSON.stringify(approver));
    this.sheet.getRange(row, statusColumn).setValue(this.rejected);

    const role = this.getRoleFromTitle(approver.title);
    this.sendNotification(taskId, { role, decision: "rejected" });
  };

  // ----------------------------------------------------
  //  WEEKLY DIGEST
  // ----------------------------------------------------
  
  this.sendWeeklyDigest = () => {
    const values = this.parsedValues();
    const pendingCounts = {};
    const pendingNames = {};

    // Scan for pending items for High Level roles
    for (let r = 1; r < values.length; r++) {
      values[r].forEach(cell => {
        try {
          if (cell && typeof cell === 'object' && cell.status === this.pending) {
            if (this.isHighLevel(cell.title)) {
               const email = String(cell.email).toLowerCase().trim();
               if (!pendingCounts[email]) {
                 pendingCounts[email] = 0;
                 pendingNames[email] = cell.name;
               }
               pendingCounts[email]++;
            }
          }
        } catch (e) { }
      });
    }

    // Send Digest Emails
    Object.keys(pendingCounts).forEach(email => {
      const count = pendingCounts[email];
      const name = pendingNames[email] || "Approver";
      const portalLink = `${this.url}?dashboard=1&email=${encodeURIComponent(email)}`;
      
      const htmlBody = `
        <h3>Hello ${name},</h3>
        <p>You have <strong>${count}</strong> pending approval requests.</p>
        <p><a href="${portalLink}">Click here to view your Dashboard</a></p>
      `;

      if (TEST_MODE) {
        GmailApp.sendEmail(TEST_EMAIL, `[TEST] Digest for ${name}`, "", { htmlBody });
      } else {
        GmailApp.sendEmail(email, `Action Required: ${count} Pending Requests`, "", { htmlBody });
      }
    });
  };
}


// ====================================================
// GLOBAL HANDLERS (Web App & Triggers)
// ====================================================

function _onFormSubmit() { new App().onFormSubmit(); }
function createTrigger() {
  const functionName = "_onFormSubmit";
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === functionName)) return;
  ScriptApp.newTrigger(functionName).forForm(FormApp.getActiveForm()).onFormSubmit().create();
}

function doGet(e) {
  const app = new App();
  const { taskId, responseId, dashboard, email, action } = e.parameter || {};
  let template;

  if (taskId && action) {
    // 1. One-click Email Action
    try {
      if (action === 'approve') app.approve({ taskId, comments: "Approved via Email" });
      if (action === 'reject') app.reject({ taskId, comments: "Rejected via Email" });
      return HtmlService.createHtmlOutput(`<h1>Action Recorded: ${action}</h1><p>You can close this window.</p>`);
    } catch (err) {
      return HtmlService.createHtmlOutput(`Error: ${err.message}`);
    }
  } else if (dashboard === "1" && email) {
    // 2. Dashboard View
    template = HtmlService.createTemplateFromFile("dashboard");
    const data = app.getDashboardDataFor(decodeURIComponent(email));
    Object.assign(template, {
      approverEmail: decodeURIComponent(email),
      approverTitle: data.approverTitle,
      pendingItems: data.pending,
      approvedItems: data.approved,
      rejectedItems: data.rejected
    });
  } else if (taskId) {
    // 3. Single Task View
    template = HtmlService.createTemplateFromFile("index");
    const data = app.getTaskById(taskId);
    if(data.task) Object.assign(template, { task: data.task, status: data.status, approver: data.approver, approvers: data.approvers, taskId });
    else template = HtmlService.createTemplateFromFile("404.html");
  } else if (responseId) {
    // 4. Progress View
    // (Logic simplified for brevity - assumes similar structure to getTaskById)
     return HtmlService.createHtmlOutput("Progress View Not Implemented in Demo");
  } else {
    // 5. Portal Login
    template = HtmlService.createTemplateFromFile("portal");
    template.detectedEmail = Session.getActiveUser().getEmail(); 
  }
  
  template.url = app.url;
  return template.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: Format Drive Links for HTML
function formatFileLinks(rawString) {
  if (!rawString || typeof rawString !== 'string') return "";
  // Simple check for Drive ID or URL
  if (rawString.includes("http") || rawString.length > 20) {
    return `<a href="${rawString}" target="_blank">View Document</a>`;
  }
  return rawString;
}
