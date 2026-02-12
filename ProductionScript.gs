  /**
  * PRODUCTION MANAGER SCRIPT - HIGH PERFORMANCE VERSION
  * 
  * Optimized for API limits and UX responsiveness.
  * Features:
  * - Lazy Loading: Permissions and Dropdowns are ONLY created when a task is assigned.
  * - Native Locking: No "Alert/Revert" scripts. Restricted cells are natively greyed out.
  * - Batch Processing: Minimizes API calls during initialization.
  */

  // ==========================================
  // 1. MENU & TRIGGER
  // ==========================================

  function onOpen() {
    const user = Session.getActiveUser().getEmail().toLowerCase();
    const admin = "filip@snowballgames.io";

    if (user === admin) {
      SpreadsheetApp.getUi()
        .createMenu('🚀 Production Manager')
        .addItem('1. Initialize New Sprint/Workstream', 'initSprintSheet')
        .addItem('2. Add New Task', 'showAddTaskDialog')
        .addItem('3. Invite Collaborator to Sheet', 'showInviteCollaboratorDialog')
        .addSeparator()
        .addItem('⚠️ Setup Notifications (Run Once)', 'setupInstallableTrigger')
        .addToUi();
    }
  }

  /**
  * Creates an installable trigger for onEdit.
  * Required for Email Notifications and "Permission" enforcement.
  */
  function setupInstallableTrigger() {
    if (!isAdminUser()) return;
    const ss = SpreadsheetApp.getActive();
    
    const triggers = ScriptApp.getUserTriggers(ss);
    for (const t of triggers) {
      if (t.getHandlerFunction() === 'handleEdit') {
        SpreadsheetApp.getUi().alert("Trigger is already set up!");
        return;
      }
    }
    
    ScriptApp.newTrigger('handleEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
      
    SpreadsheetApp.getUi().alert("✅ Notification & Permission Trigger Set Up Successfully!");
  }

  function isAdminUser() {
    const user = Session.getActiveUser().getEmail().toLowerCase();
    const admin = "filip@snowballgames.io";
    if (user !== admin) {
      SpreadsheetApp.getUi().alert("⛔ Admin Only", "You do not have permission to run this.", SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
    return true;
  }


  // ==========================================
  // 2. SHEET INITIALIZATION (Req 1, 2, 3)
  // ==========================================

  function initSprintSheet() {
    if (!isAdminUser()) return;
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Gather Input
    const nameResponse = ui.prompt('New Workstream', 'Enter name (e.g. "Marketing", "Dev"):', ui.ButtonSet.OK_CANCEL);
    if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
    const rawName = nameResponse.getResponseText().trim();
    const emoji = getSmartEmoji(rawName);
    const sheetName = `${emoji} ${rawName}`;

    if (rawName === "" || ss.getSheetByName(sheetName)) {
      ui.alert('Error', `Sheet "${sheetName}" already exists or name is empty.`, ui.ButtonSet.OK);
      return;
    }

    const teamResponse = ui.prompt('Team Members', 'Enter team emails separated by commas:', ui.ButtonSet.OK_CANCEL);
    if (teamResponse.getSelectedButton() !== ui.Button.OK) return;
    const teamMembers = teamResponse.getResponseText().split(',').map(e => e.trim().toLowerCase()).filter(e => e.includes("@"));

    if (teamMembers.length === 0) {
      ui.alert('Error', 'Please enter at least one valid email.', ui.ButtonSet.OK);
      return;
    }

    // 2. BATCH SHEET CREATION
    const sheet = ss.insertSheet(sheetName);
    const numRows = 200; 
    const dataStartRow = 10;
    
    // Dashboard Setup (Rows 1-7) — single batch write for row 4
    sheet.setRowHeights(1, 7, 30);
    sheet.getRange("B2").setValue(sheetName.toUpperCase() + " DASHBOARD")
        .setFontWeight("bold").setFontSize(18).setFontFamily("Inter");
    
    // Batch: write labels + formulas for row 4 in one call
    const dashRow = sheet.getRange("B4:H4");
    dashRow.setValues([["Total Est. Hours:", '=SUM(F10:F)', "Total Actual Hours:", '=SUM(G10:G)', "", "Overall Progress:", '=SPARKLINE(COUNTIF(D10:D, "Done"), {"charttype","bar";"max",COUNTA(B10:B);"color1","#22c55e";"empty","zero"})']]);
    dashRow.setFontWeight("bold");

    // Headers (Row 9)
    const headers = ["Task Name", "Assignee", "Status", "Deadline", "Est. Hours", "Actual Hours", "Variance"];
    const headerRange = sheet.getRange(9, 2, 1, headers.length);
    headerRange.setValues([headers]).setFontWeight("bold").setBackground("#334155").setFontColor("white").setHorizontalAlignment("center");
    
    // Layout & Formatting
    sheet.setFrozenRows(9);
    
    const dataRange = sheet.getRange(dataStartRow, 2, numRows, headers.length);
    dataRange.setFontFamily("Inter").setVerticalAlignment("middle");
    sheet.getRange(dataStartRow, 5, numRows).setNumberFormat("dd-mm-yyyy");
    
    // Batch Variance Formula
    const formulas = [];
    for (let i = 0; i < numRows; i++) {
      formulas.push([`=IF(ISBLANK(B${dataStartRow + i}), "", F${dataStartRow + i}-G${dataStartRow + i})`]);
    }
    sheet.getRange(dataStartRow, 8, numRows).setFormulas(formulas);
    
    setConditionalFormatting(sheet);

    // 3. MASTER LOCK & BATCH VALIDATION
    // Hide all status dropdowns initially (Batch)
    sheet.getRange(dataStartRow, 4, numRows).setDataValidation(null);

    // Date picker for Deadline column (European DD-MM-YYYY)
    const dateRule = SpreadsheetApp.newDataValidation().requireDate().build();
    sheet.getRange(dataStartRow, 5, numRows).setDataValidation(dateRule);
    
    // Assignee dropdown (Batch)
    addTeamToSheet(ss, sheet, teamMembers, dataStartRow, numRows);

    const me = Session.getEffectiveUser();
    applyAllProtectionsBatch(sheet, dataStartRow, numRows, me.getEmail());

    // Flush all pending writes so the Sheets API resize sees final content
    SpreadsheetApp.flush();

    // Auto-fit columns via Sheets API batchUpdate (more reliable than SpreadsheetApp.autoResizeColumns)
    const resizeSheetId = sheet.getSheetId();
    const resizeRequests = [
      { autoResizeDimensions: { dimensions: { sheetId: resizeSheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 8 } } },
      { updateDimensionProperties: { range: { sheetId: resizeSheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 2 }, properties: { pixelSize: 350 }, fields: "pixelSize" } },
      { updateDimensionProperties: { range: { sheetId: resizeSheetId, dimension: "COLUMNS", startIndex: 2, endIndex: 3 }, properties: { pixelSize: 200 }, fields: "pixelSize" } }
    ];
    Sheets.Spreadsheets.batchUpdate({ requests: resizeRequests }, ss.getId());

    ss.toast("Spreadsheet is ready. Status buttons will appear as you assign tasks.", "✅ Initialization Done", 5);
  }


  // ==========================================
  // 3. LAZY PERMISSIONS & LOGIC
  // ==========================================

  function handleEdit(e) {
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const col = range.getColumn();

    if (row < 10) return; 

    const COL_ASSIGNEE = 3;
    const COL_STATUS = 4;
    const COL_ACTUAL = 7;
    
    // 1. HANDLE ASSIGNMENT (THE "LAZY"permission trigger)
    if (col === COL_ASSIGNEE) {
      const newAssignee = e.value;
      const oldAssignee = e.oldValue;

      if (newAssignee === oldAssignee) return;
      
      // Grant row permissions (fast path)
      updateRowPermissionsAndValidation(sheet, row, newAssignee, oldAssignee);

      // Notification Logic
      if (newAssignee && newAssignee.includes("@")) {
        const ui = SpreadsheetApp.getUi();
        const confirm = ui.alert("Notify User?", `Send an email notification to ${newAssignee}?`, ui.ButtonSet.YES_NO);
        if (confirm === ui.Button.YES) {
          try {
            const rowData = sheet.getRange(row, 2, 1, 6).getValues()[0];
            const tName = rowData[0];
            const dLine = rowData[3];
            const dLineStr = (dLine instanceof Date) ? dLine.toLocaleDateString() : dLine;
            GmailApp.sendEmail(newAssignee, `New Task: ${tName}`, `Hi,\n\nYou have been assigned a task: ${tName}\nDue: ${dLineStr}\n\nPlease update the production tracker.`);
            SpreadsheetApp.getActive().toast(`✅ Email sent to ${newAssignee}`);
          } catch(err) { console.error(err); }
        }
      }
    }

    // 2. STATUS INTEGRITY (Actual Hours must be filled)
    if (col === COL_STATUS && e.value === "Done") {
      const actual = sheet.getRange(row, COL_ACTUAL).getValue();
      if (actual === "" || actual === null || Number(actual) === 0) {
        SpreadsheetApp.getUi().alert(`⚠️ Data Missing\nPlease enter 'Actual Hours' spent before marking the task as Done.`);
        if (e.oldValue === undefined) e.range.setValue("In Progress"); 
        else e.range.setValue(e.oldValue);
      }
    }
  }

  /**
  * PUNCHES HOLES in the Master Lock only when a task is assigned.
  * This keeps the script fast and stays within API limits.
  */
  function updateRowPermissionsAndValidation(sheet, row, assigneeEmail, oldAssignee) {
    const me = Session.getEffectiveUser();
    const statusCell = sheet.getRange(row, 4);
    const actualCell = sheet.getRange(row, 7);
    const statusDesc = `Row${row}-Status`;
    const actualDesc = `Row${row}-Actual`;

    let statusProt;
    let actualProt;
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (const p of protections) {
      const desc = p.getDescription();
      if (desc === statusDesc) statusProt = p;
      if (desc === actualDesc) actualProt = p;
      if (statusProt && actualProt) break;
    }

    if (!statusProt) {
      statusProt = statusCell.protect().setDescription(statusDesc);
      statusProt.addEditor(me);
      if (statusProt.canDomainEdit()) statusProt.setDomainEdit(false);
    }
    if (!actualProt) {
      actualProt = actualCell.protect().setDescription(actualDesc);
      actualProt.addEditor(me);
      if (actualProt.canDomainEdit()) actualProt.setDomainEdit(false);
    }

    if (oldAssignee && oldAssignee !== assigneeEmail) {
      statusProt.removeEditors([oldAssignee]);
      actualProt.removeEditors([oldAssignee]);
    }

    // If we have an assignee, grant them permission + Add Dropdown
    if (assigneeEmail && assigneeEmail.includes("@")) {
      statusProt.addEditor(assigneeEmail);
      actualProt.addEditor(assigneeEmail);
      statusProt.addEditor(me);
      actualProt.addEditor(me);

      // Show the Status Dropdown
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Not Started", "In Progress", "Blocked", "Done"])
        .build();
      statusCell.setDataValidation(statusRule);
      
      // Set initial value if empty
      if (statusCell.getValue() === "") statusCell.setValue("Not Started");
      
    } else {
      // If assignee cleared, keep admin-only protection and hide the dropdown
      const statusEditors = statusProt.getEditors();
      const actualEditors = actualProt.getEditors();
      if (statusEditors.length) statusProt.removeEditors(statusEditors);
      if (actualEditors.length) actualProt.removeEditors(actualEditors);
      statusProt.addEditor(me);
      actualProt.addEditor(me);

      statusCell.setDataValidation(null);
      statusCell.clearContent();
      actualCell.clearContent();
    }
  }

  /**
  * Creates ALL protections (base columns + per-row Status/Actual) in a single
  * Sheets API batchUpdate call. This replaces two separate calls and cuts
  * initialization time roughly in half.
  */
  function applyAllProtectionsBatch(sheet, dataStartRow, numRows, adminEmail) {
    if (typeof Sheets === "undefined") {
      throw new Error("Advanced Sheets Service is required. Enable it in Extensions > Apps Script > Services.");
    }

    const sheetId = sheet.getSheetId();
    const ssId = sheet.getParent().getId();
    const endDataRow = dataStartRow - 1 + numRows;

    // Base protections: dashboard/headers + admin-only columns (B:C, E:F, H)
    const baseRanges = [
      { startRowIndex: 0, endRowIndex: 9, startColumnIndex: 1, endColumnIndex: 8 },
      { startRowIndex: dataStartRow - 1, endRowIndex: endDataRow, startColumnIndex: 1, endColumnIndex: 3 },
      { startRowIndex: dataStartRow - 1, endRowIndex: endDataRow, startColumnIndex: 4, endColumnIndex: 6 },
      { startRowIndex: dataStartRow - 1, endRowIndex: endDataRow, startColumnIndex: 7, endColumnIndex: 8 }
    ];

    const requests = baseRanges.map(r => ({
      addProtectedRange: {
        protectedRange: {
          range: Object.assign({ sheetId }, r),
          description: "Base-Lock",
          editors: { users: [adminEmail] }
        }
      }
    }));

    // Per-row protections: Status (col D) and Actual Hours (col G)
    for (let i = 0; i < numRows; i++) {
      const rowIdx = dataStartRow - 1 + i;
      requests.push({
        addProtectedRange: {
          protectedRange: {
            range: { sheetId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 3, endColumnIndex: 4 },
            description: `Row${dataStartRow + i}-Status`,
            editors: { users: [adminEmail] }
          }
        }
      });
      requests.push({
        addProtectedRange: {
          protectedRange: {
            range: { sheetId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 6, endColumnIndex: 7 },
            description: `Row${dataStartRow + i}-Actual`,
            editors: { users: [adminEmail] }
          }
        }
      });
    }

    Sheets.Spreadsheets.batchUpdate({ requests }, ssId);
  }

  /**
  * Shared: adds team members to a sheet's assignee dropdown + grants file-level access.
  * Merges with any existing assignees already in the validation rule.
  */
  function addTeamToSheet(ss, sheet, newMembers, dataStartRow, numRows) {
    const existingRule = sheet.getRange(dataStartRow, 3).getDataValidation();
    let allMembers = [];
    if (existingRule) {
      const criteria = existingRule.getCriteriaValues();
      if (criteria && criteria.length > 0 && Array.isArray(criteria[0])) {
        allMembers = criteria[0].map(e => e.toLowerCase());
      } else if (criteria && criteria.length > 0) {
        allMembers = criteria.map(e => String(e).toLowerCase());
      }
    }
    for (const m of newMembers) {
      if (!allMembers.includes(m)) allMembers.push(m);
    }
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(allMembers).build();
    sheet.getRange(dataStartRow, 3, numRows).setDataValidation(rule);
    try { ss.addEditors(newMembers); } catch(e) {}
  }

  /**
  * Invite Collaborator: lets admin pick a sheet and add new team members.
  * Reuses addTeamToSheet so the logic is identical to init.
  */
  function showInviteCollaboratorDialog() {
    if (!isAdminUser()) return;
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheets = ss.getSheets().filter(s => s.getFrozenRows() >= 9);
    if (sheets.length === 0) {
      ui.alert("No Sheets", "No initialized sprint sheets found.", ui.ButtonSet.OK);
      return;
    }

    const sheetNames = sheets.map(s => s.getName());
    const listStr = sheetNames.map((n, i) => `${i + 1}. ${n}`).join("\n");
    const pickResp = ui.prompt(
      "Select Sheet",
      `Enter the number of the sheet to invite collaborators to:\n\n${listStr}`,
      ui.ButtonSet.OK_CANCEL
    );
    if (pickResp.getSelectedButton() !== ui.Button.OK) return;

    const idx = parseInt(pickResp.getResponseText().trim(), 10) - 1;
    if (isNaN(idx) || idx < 0 || idx >= sheets.length) {
      ui.alert("Invalid selection.");
      return;
    }
    const targetSheet = sheets[idx];

    const emailResp = ui.prompt(
      "Invite Collaborators",
      `Adding to: ${targetSheet.getName()}\n\nEnter emails separated by commas:`,
      ui.ButtonSet.OK_CANCEL
    );
    if (emailResp.getSelectedButton() !== ui.Button.OK) return;

    const newMembers = emailResp.getResponseText().split(",").map(e => e.trim().toLowerCase()).filter(e => e.includes("@"));
    if (newMembers.length === 0) {
      ui.alert("Error", "No valid emails entered.", ui.ButtonSet.OK);
      return;
    }

    addTeamToSheet(ss, targetSheet, newMembers, 10, 200);
    ss.toast(`Added ${newMembers.length} collaborator(s) to ${targetSheet.getName()}.`, "✅ Done", 5);
  }

  // ==========================================
  // 4. TASK DIALOG
  // ==========================================

  function showAddTaskDialog() {
    if (!isAdminUser()) return;
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    if (sheet.getFrozenRows() < 9) {
      ui.alert("Using Wrong Sheet?", "Please switch to a Sprint Sheet.", ui.ButtonSet.OK);
      return;
    }

    const nameResp = ui.prompt("New Task", "Task Name:", ui.ButtonSet.OK_CANCEL);
    if (nameResp.getSelectedButton() !== ui.Button.OK) return;
    const taskName = nameResp.getResponseText();

    const estResp = ui.prompt("Details", "Estimated Hours spent:", ui.ButtonSet.OK_CANCEL);
    const estHoursValue = estResp.getSelectedButton() === ui.Button.OK ? estResp.getResponseText() : "";

    const deadResp = ui.prompt("Details", "Deadline (DD-MM-YYYY):", ui.ButtonSet.OK_CANCEL);
    const deadlineValue = deadResp.getSelectedButton() === ui.Button.OK ? deadResp.getResponseText() : "";

    // Helper: Find next empty row (Col B)
    const data = sheet.getRange(10, 2, 200).getValues();
    let targetRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === "") { targetRow = 10 + i; break; }
    }
    if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

    sheet.getRange(targetRow, 2, 1, 5).setValues([[taskName, "", "", deadlineValue, estHoursValue]]);
    
    // Flash to Assignee column
    sheet.setActiveRange(sheet.getRange(targetRow, 3));
  }


  // ==========================================
  // 5. STYLING & HELPERS
  // ==========================================

  function setConditionalFormatting(sheet) {
    const range = sheet.getRange("B10:H200");
    const deadlineRange = sheet.getRange("E10:E200");
    const rules = [];

    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D10="Blocked"').setBackground("#fee2e2").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D10="In Progress"').setBackground("#fef9c3").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D10="Done"').setBackground("#dcfce7").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$D10="Not Started"').setBackground("#f1f5f9").setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E10 < TODAY(), $E10 <> "", $D10 <> "Done")').setFontColor("#ef4444").setBold(true).setRanges([deadlineRange]).build());

    sheet.setConditionalFormatRules(rules);
  }

  function getSmartEmoji(name) {
    const n = name.toLowerCase();
    if (n.includes("market") || n.includes("sale")) return "📢";
    if (n.includes("dev") || n.includes("code") || n.includes("tech")) return "💻";
    if (n.includes("art") || n.includes("design")) return "🎨";
    if (n.includes("qa") || n.includes("test") || n.includes("bug")) return "🐛";
    if (n.includes("sound") || n.includes("music")) return "🎵";
    if (n.includes("prod") || n.includes("manage")) return "📅";
    if (n.includes("write") || n.includes("script")) return "📝";
    return "📁";
  }
