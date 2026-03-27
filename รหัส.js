

// ============================================================
// QA_BUTC_System_version_3.0 — Backend Controller (Code.gs)
// Architecture: MVC-like pattern within Google Apps Script
// Roles: Admin | Manager | User
// ============================================================
const SPREADSHEET_NAME = "QA_BUTC_Database_v3";

// เอา ID จาก Log มาใส่ตรงนี้เลยครับ
const SPREADSHEET_ID = "1h5V8lB9vJpYBsaGEKuhhEiyf0iPhlH-NKZccMKd4wGw";
// ─────────────────────────────────────────
// ENTRY POINT
// ─────────────────────────────────────────
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('QA BUTC System v3.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// QA_BUTC_System_version_3.0 — Backend Controller (Code.gs)
// ============================================================

// ─────────────────────────────────────────
// FOLDER MANAGEMENT
// หลักการ: Spreadsheet อยู่ใน Folder เดียวกับ Script เสมอ
// ─────────────────────────────────────────

/**
 * ดึง Folder ที่ Script File นี้อยู่ปัจจุบัน
 * ไม่สร้าง Folder ใหม่ — ใช้ที่อยู่จริงของ Script เลย
 * วิธีนี้ทำงานได้ 100% ไม่ว่า Script จะอยู่ใน Folder ไหนก็ตาม
 */
function getScriptFolder() {
  const scriptId   = ScriptApp.getScriptId();
  const scriptFile = DriveApp.getFileById(scriptId);
  const parents    = scriptFile.getParents();
  if (parents.hasNext()) { return parents.next(); }
  return DriveApp.getRootFolder();
}

/**
 * ค้นหาหรือสร้าง Spreadsheet ใน Folder เดียวกับ Script
 */
/**
 * Patch 1: ปรับแก้การเข้าถึง Spreadsheet ให้เป็น O(1) Direct Access
 */
function getOrCreateSpreadsheet() {
  // 1. ถ้ามีการใส่ ID ไว้แล้ว ให้ดึงข้อมูลจาก ID ทันที (เสถียร 100% ไม่พลาดไฟล์)
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "ใส่_ID_ตรงนี้" && SPREADSHEET_ID !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  // 2. ถ้ายังไม่มี ID (รันครั้งแรก) ให้ใช้ Logic เดิมของคุณ
  const folder = getScriptFolder();
  const files = folder.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  // 3. กรณีต้องสร้างใหม่
  const ss     = SpreadsheetApp.create(SPREADSHEET_NAME);
  const ssFile = DriveApp.getFileById(ss.getId());
  folder.addFile(ssFile);

  try {
    const root = DriveApp.getRootFolder();
    if (root.getId() !== folder.getId()) {
      root.removeFile(ssFile);
    }
  } catch(e) { console.error("Move folder error:", e); }

  return ss;
}

// ─────────────────────────────────────────
// ฟังก์ชัน SETUP
// รันครั้งเดียวหลัง Deploy เพื่อสร้างฐานข้อมูล
// เลือก setupProject ใน Dropdown แล้วกด Run
// ─────────────────────────────────────────
function setupProject() {
  const folder = getScriptFolder();
  const ss     = getOrCreateSpreadsheet();
  const result = initializeDatabase();

  Logger.log("========== SETUP COMPLETE ==========");
  Logger.log("📁 Script อยู่ใน Folder : " + folder.getName());
  Logger.log("🔗 Folder URL           : " + folder.getUrl());
  Logger.log("📊 Spreadsheet          : " + ss.getName());
  Logger.log("🔗 Spreadsheet URL      : " + ss.getUrl());
  Logger.log("🗄️  Database Status      : " + result.message);
  Logger.log("====================================");
}

// ─────────────────────────────────────────
// DATABASE INITIALIZATION
// ─────────────────────────────────────────
function ensureSheetExists(sheetName) {
  const ss    = getOrCreateSpreadsheet();
  let   sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initSheetHeaders(sheet, sheetName);
  }
  return sheet;
}

function initSheetHeaders(sheet, sheetName) {
  const headers = {
    "tb_admins":   ["Email", "Password", "Prefix", "FirstName", "LastName", "Role", "CreatedAt"],
    "tb_managers": ["Email", "Prefix", "FirstName", "LastName", "Phone", "ManagedSections", "CreatedAt"],
    "tb_users":    ["Email", "Auth_Type", "Prefix", "FirstName", "LastName", "Phone", "Positions", "Sections", "CreatedAt"],
    "tb_sections": ["SectionID", "SectionName", "Description", "AssignedFormIDs", "IconClass", "ColorClass", "IsActive", "CreatedAt", "CreatedBy"],
    "tb_forms":    ["FormID", "FormName", "SectionID", "Fields", "FormType", "IsActive", "CreatedAt", "CreatedBy", "Description", "Note"],
    "tb_data":     ["Timestamp", "UserEmail", "Year", "SectionID", "SectionName", "LastUpdate", "DataJSON", "SubmitterName"],
    "tb_settings": ["Year", "StartDate", "EndDate", "Status", "AllowedSections"],
    "tb_audit":    ["Timestamp", "UserEmail", "Action", "Target", "Detail"],
    "tb_theme":    ["Key", "Value", "UpdatedAt", "UpdatedBy"],
    "tb_iqa_mapping": ["PageID", "MappingJSON", "UpdatedAt", "UpdatedBy"], // เพิ่มบรรทัดนี้
    "tb_departments": ["DeptID", "DeptName", "Description", "AssignedSections", "IconClass", "ColorClass", "IsActive", "CreatedAt", "CreatedBy"] // ✅ เพิ่มบรรทัดนี้
  };
  if (headers[sheetName]) {
    sheet.appendRow(headers[sheetName]);
    sheet.getRange(1, 1, 1, headers[sheetName].length)
      .setBackground('#1a237e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

function initializeDatabase() {
  try {
    const sheets = [
      "tb_admins", "tb_managers", "tb_users",
      "tb_sections", "tb_forms", "tb_data",
      "tb_settings", "tb_audit", "tb_theme", "tb_visits", "tb_iqa_mapping", "tb_departments" // ✅ เพิ่ม tb_departments
    ];
    sheets.forEach(s => ensureSheetExists(s));
    return { success: true, message: "ฐานข้อมูลพร้อมใช้งาน" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// AUDIT LOG
// ─────────────────────────────────────────
function writeAuditLog(userEmail, action, target, detail) {
  try {
    const sheet = ensureSheetExists("tb_audit");
    sheet.appendRow([new Date(), userEmail, action, target, detail]);
  } catch(e) { /* silent fail */ }
}

// ─────────────────────────────────────────
// 1. AUTHENTICATION — Google Workspace SSO
// ─────────────────────────────────────────
function verifyGoogleWorkspaceIdentity() {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const scriptUrl   = ScriptApp.getService().getUrl();
    // ✅ เพิ่มบรรทัดนี้: สร้าง URL สำหรับเปลี่ยนบัญชีที่ถูกต้องจากฝั่ง Backend
    const switchUrl   = `https://accounts.google.com/AccountChooser?continue=${encodeURIComponent(scriptUrl)}`;

    if (!activeEmail) {
      return {
        status:  "auth_required",
        message: "ระบบไม่สามารถดึงข้อมูลบัญชีได้ (Third-party Cookies หรือ Incognito Mode)",
        url:     scriptUrl,
        switchUrl: switchUrl // ✅ ส่งไปให้หน้าเว็บ
      };
    }

    if (!activeEmail.endsWith("@butc.ac.th")) {
      return {
        status: "wrong_domain",
        email:  activeEmail,
        url:    switchUrl,
        switchUrl: switchUrl // ✅ ส่งไปให้หน้าเว็บ
      };
    }

    // ตรวจสอบ Admin
    const adminSheet = ensureSheetExists("tb_admins");
    const admins     = adminSheet.getDataRange().getValues();
    for (let i = 1; i < admins.length; i++) {
      if (admins[i][0] === activeEmail) {
        const fullName  = `${admins[i][2]}${admins[i][3]} ${admins[i][4]}`;
        const adminRole = admins[i][5] || "admin";
        writeAuditLog(activeEmail, "LOGIN", "admin", "SSO Login Success");
        return {
          status:    "success",
          role:      "admin",
          adminRole: adminRole,
          email:     activeEmail,
          fullName:  fullName,
          profile: {
            prefix:    admins[i][2] || "",
            fname:     admins[i][3] || "",
            lname:     admins[i][4] || "",
            positions: ["ผู้ดูแลระบบ"],
            sections:  [],
            fullName:  fullName
          }
        };
      }
    }

    // ตรวจสอบ Manager
    const managerSheet = ensureSheetExists("tb_managers");
    const managers     = managerSheet.getDataRange().getValues();
    for (let i = 1; i < managers.length; i++) {
      if (managers[i][0] === activeEmail) {
        const fullName = `${managers[i][1]}${managers[i][2]} ${managers[i][3]}`;
        let   sections = [];
        try { sections = managers[i][5] ? JSON.parse(managers[i][5]) : []; } catch(e) {} const permission = managers[i][7] || "view"; // ✅ ดึงสิทธิ์จากคอลัมน์ที่ 8 (Index 7)
        writeAuditLog(activeEmail, "LOGIN", "manager", "SSO Login Success");
        return {
          status:   "success",
          role:     "manager",
          email:    activeEmail,
          fullName: fullName,
          profile:  {
            prefix:   managers[i][1],
            fname:    managers[i][2],
            lname:    managers[i][3],
            phone:    managers[i][4],
            sections: sections,
            permission: permission, // ✅ ส่งสิทธิ์ไปให้หน้าเว็บ
            fullName: fullName
          }
        };
      }
    }

    // ตรวจสอบ User
    const userSheet = ensureSheetExists("tb_users");
    const users     = userSheet.getDataRange().getValues();
    for (let i = 1; i < users.length; i++) {
      if (users[i][0] === activeEmail) {
        const fullName  = `${users[i][2]}${users[i][3]} ${users[i][4]}`;
        let   positions = [], sections = [];
        try { positions = users[i][6] ? JSON.parse(users[i][6]) : []; } catch(e) {}
        try { sections  = users[i][7] ? JSON.parse(users[i][7]) : []; } catch(e) {}
        writeAuditLog(activeEmail, "LOGIN", "user", "SSO Login Success");
        return {
          status:   "success",
          role:     "user",
          email:    activeEmail,
          fullName: fullName,
          profile:  {
            prefix:    users[i][2],
            fname:     users[i][3],
            lname:     users[i][4],
            phone:     users[i][5],
            positions: positions,
            sections:  sections,
            fullName:  fullName
          }
        };
      }
    }

    // ✅ บรรทัดสุดท้าย (กรณีไม่พบผู้ใช้ในระบบเลย) ให้ส่ง switchUrl กลับไปด้วย
    return { status: "unregistered", email: activeEmail, switchUrl: switchUrl };

  } catch(error) {
    return { status: "error", message: "System Error: " + error.toString() };
  }
}

// ─────────────────────────────────────────
// 2. REGISTRATION (แบบใหม่: แยก User / Manager อัตโนมัติ)
// ─────────────────────────────────────────
function registerUser(formData) {
  try {
    _clearAppCache();
    const email = formData.email.trim();
    if (!email.endsWith("@butc.ac.th")) return { success: false, message: "อนุญาตเฉพาะ @butc.ac.th เท่านั้น" };

    const uSheet = ensureSheetExists("tb_users");
    const mSheet = ensureSheetExists("tb_managers");
    
    // ตรวจสอบอีเมลซ้ำในทั้งสองตาราง
    const uData = uSheet.getDataRange().getValues();
    const mData = mSheet.getDataRange().getValues();
    if (uData.some(r => r[0] === email) || mData.some(r => r[0] === email)) {
      return { success: false, message: "อีเมลนี้ถูกลงทะเบียนแล้ว" };
    }

    const positions = formData.positions || [];
    const isManager = positions.includes("ผู้บริหาร");
    const permission = (isManager && positions.length > 1) ? "edit" : "view";
    const sectionsStr = JSON.stringify(formData.sections || []);
    const positionsStr = JSON.stringify(positions);

    if (isManager) {
      mSheet.appendRow([email, formData.prefix, formData.fname, formData.lname, formData.phone, sectionsStr, new Date(), permission, positionsStr]);
    } else {
      uSheet.appendRow([email, "SSO_GOOGLE", formData.prefix, formData.fname, formData.lname, formData.phone, positionsStr, sectionsStr, new Date()]);
    }

    writeAuditLog(email, "REGISTER", isManager ? "manager" : "user", `Permissions: ${permission}`);
    return { success: true, message: "ลงทะเบียนสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// ระบบลบข้อมูลเก่าก่อนย้ายตาราง (Helper)
// ─────────────────────────────────────────
function _deleteUserFromBothTables(email) {
  ["tb_managers", "tb_users"].forEach(sName => {
    let sheet = ensureSheetExists(sName);
    let data = sheet.getDataRange().getValues();
    for(let i = data.length - 1; i >= 1; i--) {
      if(data[i][0] === email) sheet.deleteRow(i + 1);
    }
  });
}

// ─────────────────────────────────────────
// 2. REGISTRATION
// ─────────────────────────────────────────
/*
function registerUser(formData) {
  try {
    const email = formData.email.trim();
    if (!email.endsWith("@butc.ac.th")) {
      return { success: false, message: "อนุญาตเฉพาะ @butc.ac.th เท่านั้น" };
    }

    const userSheet = ensureSheetExists("tb_users");
    const users     = userSheet.getDataRange().getValues();
    for (let i = 1; i < users.length; i++) {
      if (users[i][0] === email) {
        return { success: false, message: "อีเมลนี้ถูกลงทะเบียนแล้ว" };
      }
    }

    const positions = JSON.stringify(formData.positions || []);
    const sections  = JSON.stringify(formData.sections  || []);
    userSheet.appendRow([
      email, "SSO_GOOGLE",
      formData.prefix, formData.fname, formData.lname,
      formData.phone, positions, sections, new Date()
    ]);

    writeAuditLog(email, "REGISTER", "user", "New user registered");
    return { success: true, message: "ลงทะเบียนสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}*/

// ─────────────────────────────────────────
// 3. SECTION MANAGEMENT (Admin)
// ─────────────────────────────────────────
// ─────────────────────────────────────────
// PATCH: แก้ปัญหา Silent Fail จาก Date Object
// ─────────────────────────────────────────

function getAllSections() {
  try {
    const sheet    = ensureSheetExists("tb_sections");
    const data     = sheet.getDataRange().getValues();
    const sections = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let assignedFormIDs = [];
      try { assignedFormIDs = data[i][3] ? JSON.parse(data[i][3]) : []; } catch(e) {}
      
      // ✅ แปลง Date เป็น String ป้องกันท่อข้อมูลพัง
      let cDate = data[i][7];
      if (cDate instanceof Date) { cDate = cDate.toISOString(); }

      sections.push({
        sectionId:       data[i][0],
        sectionName:     data[i][1],
        description:     data[i][2],
        assignedFormIDs: assignedFormIDs,
        iconClass:       data[i][4] || "bi-folder",
        colorClass:      data[i][5] || "primary",
        isActive:        data[i][6] !== false,
        createdAt:       cDate,
        createdBy:       data[i][8]
      });
    }
    return { success: true, data: sections };
  } catch(e) {
    console.error("🔥 ERROR in getAllSections: " + e.stack);
    return { success: false, data: [], message: e.toString() };
  }
}

function saveSection(sectionData) {
  try {
    _clearAppCache();
    const adminEmail         = Session.getActiveUser().getEmail();
    const sheet              = ensureSheetExists("tb_sections");
    const data               = sheet.getDataRange().getValues();
    const assignedFormIDsStr = JSON.stringify(sectionData.assignedFormIDs || []);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sectionData.sectionId) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          sectionData.sectionId,
          sectionData.sectionName,
          sectionData.description,
          assignedFormIDsStr,
          sectionData.iconClass,
          // รวม customColor เข้า colorClass ถ้าเป็น custom
          (sectionData.colorClass === "custom" && sectionData.customColor
            ? "custom:"+sectionData.customColor
            : sectionData.colorClass),
          sectionData.isActive,
          data[i][7],
          data[i][8]
        ]]);
        writeAuditLog(adminEmail, "UPDATE_SECTION", sectionData.sectionId, sectionData.sectionName);
        return { success: true, message: "อัปเดต Section สำเร็จ" };
      }
    }

    const newId = "SEC_" + Date.now();
    sheet.appendRow([
      newId,
      sectionData.sectionName,
      sectionData.description,
      assignedFormIDsStr,
      sectionData.iconClass,
      sectionData.colorClass,
      true,
      new Date(),
      adminEmail
    ]);
    writeAuditLog(adminEmail, "CREATE_SECTION", newId, sectionData.sectionName);
    return { success: true, message: "สร้าง Section สำเร็จ", sectionId: newId };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteSection(sectionId) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet      = ensureSheetExists("tb_sections");
    const data       = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sectionId) {
        sheet.deleteRow(i + 1);
        writeAuditLog(adminEmail, "DELETE_SECTION", sectionId, data[i][1]);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบ Section" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 4. FORM BUILDER (Admin)
// ─────────────────────────────────────────
function getAllForms() {
  try {
    const sheet = ensureSheetExists("tb_forms");
    const data  = sheet.getDataRange().getValues();
    const forms = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let fields = [];
      try { fields = data[i][3] ? JSON.parse(data[i][3]) : []; } catch(e) {}
      
      // ✅ แปลง Date เป็น String ป้องกันท่อข้อมูลพัง
      let cDate = data[i][6];
      if (cDate instanceof Date) { cDate = cDate.toISOString(); }

      forms.push({
        formId:      data[i][0],
        formName:    data[i][1],
        sectionId:   data[i][2],
        fields:      fields,
        formType:    data[i][4] || "entry",
        isActive:    data[i][5] !== false,
        createdAt:   cDate,
        createdBy:   data[i][7],
        description: data[i][8] || "",
        note:        data[i][9] || ""
      });
    }
    return { success: true, data: forms };
  } catch(e) {
    console.error("🔥 ERROR in getAllForms: " + e.stack);
    return { success: false, data: [], message: e.toString() };
  }
}

function saveForm(formData) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet      = ensureSheetExists("tb_forms");
    const data       = sheet.getDataRange().getValues();
    const fieldsStr  = JSON.stringify(formData.fields || []);
    const desc       = formData.description || "";
    const note       = formData.note        || "";

    if (formData.formId) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === formData.formId) {
          sheet.getRange(i + 1, 1, 1, 10).setValues([[
            formData.formId, formData.formName, formData.sectionId,
            fieldsStr, formData.formType, formData.isActive,
            data[i][6], data[i][7], desc, note
          ]]);
          writeAuditLog(adminEmail, "UPDATE_FORM", formData.formId, formData.formName);
          return { success: true, message: "อัปเดต Form สำเร็จ" };
        }
      }
    }

    const newId = "FORM_" + Date.now();
    sheet.appendRow([
      newId, formData.formName, formData.sectionId,
      fieldsStr, formData.formType, true,
      new Date(), adminEmail, desc, note
    ]);
    writeAuditLog(adminEmail, "CREATE_FORM", newId, formData.formName);
    return { success: true, message: "สร้าง Form สำเร็จ", formId: newId };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteForm(formId) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet      = ensureSheetExists("tb_forms");
    const data       = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === formId) {
        sheet.deleteRow(i + 1);
        writeAuditLog(adminEmail, "DELETE_FORM", formId, data[i][1]);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบ Form" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 5. DATA SUBMISSION (CRUD)
// ─────────────────────────────────────────
function saveFormData(sectionId, sectionName, year, formDataObj, userEmail, userName) {
  try {
    const sheet     = ensureSheetExists("tb_data");
    const timestamp = new Date();
    const data      = sheet.getDataRange().getValues();
    let   rowToUpdate  = -1;
    let   existingData = {};

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(year) && data[i][3] === sectionId) {
        rowToUpdate = i + 1;
        try { existingData = JSON.parse(data[i][6] || "{}"); } catch(e) {}
        break;
      }
    }

    // Merge ข้อมูลเก่ากับใหม่
    const mergedData = { ...existingData, ...formDataObj };
    const dataJson   = JSON.stringify(mergedData);

    if (rowToUpdate > 0) {
      sheet.getRange(rowToUpdate, 2).setValue(userEmail);
      sheet.getRange(rowToUpdate, 6).setValue(timestamp);
      sheet.getRange(rowToUpdate, 7).setValue(dataJson);
      sheet.getRange(rowToUpdate, 8).setValue(userName);
    } else {
      sheet.appendRow([
        timestamp, userEmail, year,
        sectionId, sectionName,
        timestamp, dataJson, userName
      ]);
    }

    writeAuditLog(userEmail, "SAVE_DATA", `${sectionId}:${year}`, userName);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function fetchDashboardData() {
  try {
    const sheet   = ensureSheetExists("tb_data");
    const data    = sheet.getDataRange().getValues();
    const summary = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][2] || !data[i][3]) continue;
      const lUpdate = data[i][5] instanceof Date
        ? data[i][5].toISOString()
        : data[i][5];
      summary.push({
        year:          data[i][2],
        sectionId:     data[i][3],
        sectionName:   data[i][4],
        lastUpdate:    lUpdate,
        submitterName: data[i][7] || data[i][1],
        dataJson:      data[i][6]
      });
    }
    return { success: true, data: summary };
  } catch(e) {
    // [Patch 2]: บันทึก Error ลงระบบเบื้องหลัง
    console.error("🔥 ERROR in fetchDashboardData: " + e.stack);
    return { success: false, data: [], message: e.toString() };
  }
}

function fetchSectionData(sectionId, year) {
  try {
    const sheet = ensureSheetExists("tb_data");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(year) && data[i][3] === sectionId) {
        let parsed = {};
        try { parsed = JSON.parse(data[i][6] || "{}"); } catch(e) {}
        return {
          success:       true,
          data:          parsed,
          submitterName: data[i][7],
          lastUpdate:    data[i][5]
        };
      }
    }
    return { success: true, data: {}, submitterName: "", lastUpdate: null };
  } catch(e) {
    return { success: false, data: {}, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 6. SYSTEM SETTINGS (Admin)
// ─────────────────────────────────────────
function getSystemSettings() {
  try {
    const sheet    = ensureSheetExists("tb_settings");
    const data     = sheet.getDataRange().getValues();
    const settings = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const sDate = data[i][1] instanceof Date ? data[i][1].toISOString() : data[i][1];
      const eDate = data[i][2] instanceof Date ? data[i][2].toISOString() : data[i][2];
      let   allowedSections = [];
      try { allowedSections = data[i][4] ? JSON.parse(data[i][4]) : []; } catch(e) {}
      settings.push({
        year:            data[i][0],
        startDate:       sDate,
        endDate:         eDate,
        isOpen:          data[i][3] === "Open",
        allowedSections: allowedSections
      });
    }
    settings.sort((a, b) => Number(b.year) - Number(a.year));
    return { success: true, data: settings };
  } catch(e) {
    return { success: false, data: [], message: e.toString() };
  }
}

function saveAdminSettings(formData) {
  try {
    _clearAppCache();
    const adminEmail         = Session.getActiveUser().getEmail();
    const sheet              = ensureSheetExists("tb_settings");
    const data               = sheet.getDataRange().getValues();
    const allowedSectionsStr = JSON.stringify(formData.allowedSections || []);
    let   rowToUpdate        = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(formData.year)) {
        rowToUpdate = i + 1;
        break;
      }
    }

    if (rowToUpdate > 0) {
      sheet.getRange(rowToUpdate, 1, 1, 5).setValues([[
        formData.year, formData.startDate, formData.endDate,
        formData.status, allowedSectionsStr
      ]]);
    } else {
      sheet.appendRow([
        formData.year, formData.startDate, formData.endDate,
        formData.status, allowedSectionsStr
      ]);
    }

    writeAuditLog(adminEmail, "SAVE_SETTINGS", formData.year, `Status: ${formData.status}`);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 7. USER MANAGEMENT (Admin)
// ─────────────────────────────────────────
function getAllUsers() {
  try {
    const result = { admins: [], managers: [], users: [] };

    const aSheet = ensureSheetExists("tb_admins");
    const aData  = aSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      if (!aData[i][0]) continue;
      result.admins.push({ email: aData[i][0], prefix: aData[i][2], fname: aData[i][3], lname: aData[i][4], role: aData[i][5] || "admin" });
    }

    const mSheet = ensureSheetExists("tb_managers");
    const mData  = mSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      if (!mData[i][0]) continue;
      let sections = []; try { sections = mData[i][5] ? JSON.parse(mData[i][5]) : []; } catch(e) {}
      result.managers.push({
        email:      mData[i][0],
        prefix:     mData[i][1],
        fname:      mData[i][2],
        lname:      mData[i][3],
        phone:      mData[i][4],
        sections:   sections,
        permission: mData[i][7] || "view"
      });
    }

    const uSheet = ensureSheetExists("tb_users");
    const uData  = uSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      if (!uData[i][0]) continue;
      let positions = [], sections = [];
      try { positions = uData[i][6] ? JSON.parse(uData[i][6]) : []; } catch(e) {}
      try { sections  = uData[i][7] ? JSON.parse(uData[i][7]) : []; } catch(e) {}
      result.users.push({
        email:     uData[i][0],
        prefix:    uData[i][2],
        fname:     uData[i][3],
        lname:     uData[i][4],
        phone:     uData[i][5],
        positions: positions,
        sections:  sections
      });
    }

    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 7. USER MANAGEMENT (Admin)
// ─────────────────────────────────────────
/*function getAllUsers() {
  try {
    const result = { admins: [], managers: [], users: [] };

    const aSheet = ensureSheetExists("tb_admins");
    const aData  = aSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      if (!aData[i][0]) continue;
      result.admins.push({
        email:  aData[i][0],
        prefix: aData[i][2],
        fname:  aData[i][3],
        lname:  aData[i][4],
        role:   aData[i][5] || "admin"
      });
    }

    const mSheet = ensureSheetExists("tb_managers");
    const mData  = mSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      if (!mData[i][0]) continue;
      let sections = [];
      try { sections = mData[i][5] ? JSON.parse(mData[i][5]) : []; } catch(e) {}
      result.managers.push({
        email:    mData[i][0],
        prefix:   mData[i][1],
        fname:    mData[i][2],
        lname:    mData[i][3],
        phone:    mData[i][4],
        sections: sections
      });
    }

    const uSheet = ensureSheetExists("tb_users");
    const uData  = uSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      if (!uData[i][0]) continue;
      let positions = [], sections = [];
      try { positions = uData[i][6] ? JSON.parse(uData[i][6]) : []; } catch(e) {}
      try { sections  = uData[i][7] ? JSON.parse(uData[i][7]) : []; } catch(e) {}
      result.users.push({
        email:     uData[i][0],
        prefix:    uData[i][2],
        fname:     uData[i][3],
        lname:     uData[i][4],
        phone:     uData[i][5],
        positions: positions,
        sections:  sections
      });
    }

    return { success: true, data: result };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}*/

/*
function saveManager(managerData) {
  try {
    const adminEmail  = Session.getActiveUser().getEmail();
    const sheet       = ensureSheetExists("tb_managers");
    const data        = sheet.getDataRange().getValues();
    const email       = managerData.email.trim();
    if (!email.endsWith("@butc.ac.th")) {
      return { success: false, message: "อนุญาตเฉพาะ @butc.ac.th เท่านั้น" };
    }
    const sectionsStr = JSON.stringify(managerData.sections || []);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.getRange(i + 1, 1, 1, 7).setValues([[
          email, managerData.prefix, managerData.fname,
          managerData.lname, managerData.phone, sectionsStr, data[i][6]
        ]]);
        writeAuditLog(adminEmail, "UPDATE_MANAGER", email, "");
        return { success: true, message: "อัปเดต Manager สำเร็จ" };
      }
    }

    sheet.appendRow([
      email, managerData.prefix, managerData.fname,
      managerData.lname, managerData.phone, sectionsStr, new Date()
    ]);
    writeAuditLog(adminEmail, "CREATE_MANAGER", email, "");
    return { success: true, message: "เพิ่ม Manager สำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}*/

// เปลี่ยนจาก saveManager เป็น savePersonnel ให้รองรับการบันทึกได้ทั้ง User และ Manager
function savePersonnel(data) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const email = data.email.trim();
    if (!email.endsWith("@butc.ac.th")) return { success: false, message: "อนุญาตเฉพาะ @butc.ac.th เท่านั้น" };

    const positions = data.positions || [];
    const isManager = positions.includes("ผู้บริหาร");
    const permission = data.permission || "view";
    const sectionsStr = JSON.stringify(data.sections || []);
    const positionsStr = JSON.stringify(positions);

    // ลบข้อมูลจากตารางเดิมก่อน (เผื่อมีการอัปเกรด/ดาวน์เกรดสิทธิ์)
    _deleteUserFromBothTables(email);

    if (isManager) {
      const mSheet = ensureSheetExists("tb_managers");
      mSheet.appendRow([email, data.prefix, data.fname, data.lname, data.phone, sectionsStr, new Date(), permission, positionsStr]);
    } else {
      const uSheet = ensureSheetExists("tb_users");
      uSheet.appendRow([email, "SSO_GOOGLE", data.prefix, data.fname, data.lname, data.phone, positionsStr, sectionsStr, new Date()]);
    }

    writeAuditLog(adminEmail, "SAVE_PERSONNEL", email, isManager ? "Manager" : "User");
    return { success: true, message: "บันทึกข้อมูลบุคลากรสำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteUser(userType, email) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheetMap   = { admin: "tb_admins", manager: "tb_managers", user: "tb_users" };
    const sheet = ensureSheetExists(sheetMap[userType]);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.deleteRow(i + 1);
        writeAuditLog(adminEmail, "DELETE_USER", email, userType);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/*
function saveManager(managerData) {
  try {
    const adminEmail  = Session.getActiveUser().getEmail();
    const sheet       = ensureSheetExists("tb_managers");
    const data        = sheet.getDataRange().getValues();
    const email       = managerData.email.trim();
    if (!email.endsWith("@butc.ac.th")) {
      return { success: false, message: "อนุญาตเฉพาะ @butc.ac.th เท่านั้น" };
    }
    const sectionsStr = JSON.stringify(managerData.sections || []);
    const permission  = managerData.permission || "view"; // ✅ รับค่าสิทธิ์

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        // ✅ อัปเดตข้อมูลและสิทธิ์ (คอลัมน์ที่ 8)
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          email, managerData.prefix, managerData.fname,
          managerData.lname, managerData.phone, sectionsStr, data[i][6], permission
        ]]);
        writeAuditLog(adminEmail, "UPDATE_MANAGER", email, `Permission: ${permission}`);
        return { success: true, message: "อัปเดต Manager สำเร็จ" };
      }
    }

    // ✅ เพิ่มข้อมูลใหม่พร้อมสิทธิ์
    sheet.appendRow([
      email, managerData.prefix, managerData.fname,
      managerData.lname, managerData.phone, sectionsStr, new Date(), permission
    ]);
    writeAuditLog(adminEmail, "CREATE_MANAGER", email, `Permission: ${permission}`);
    return { success: true, message: "เพิ่ม Manager สำเร็จ" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}*/

/*
function deleteUser(userType, email) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const sheetMap   = {
      admin:   "tb_admins",
      manager: "tb_managers",
      user:    "tb_users"
    };
    const sheet = ensureSheetExists(sheetMap[userType]);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.deleteRow(i + 1);
        writeAuditLog(adminEmail, "DELETE_USER", email, userType);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}*/

function updateUserSections(email, sections, role, permission) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const sheetName = role === "manager" ? "tb_managers" : "tb_users";
    const sheet = ensureSheetExists(sheetName);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        if (role === "manager") {
          sheet.getRange(i + 1, 6).setValue(JSON.stringify(sections)); // อัปเดต Section ของ Manager
          sheet.getRange(i + 1, 8).setValue(permission); // อัปเดตสิทธิ์ของ Manager
        } else {
          sheet.getRange(i + 1, 8).setValue(JSON.stringify(sections)); // อัปเดต Section ของ User
        }
        writeAuditLog(adminEmail, "UPDATE_USER_SECTIONS", email, JSON.stringify(sections));
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบ User" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/*
function updateUserSections(email, sections) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet      = ensureSheetExists("tb_users");
    const data       = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.getRange(i + 1, 8).setValue(JSON.stringify(sections));
        writeAuditLog(adminEmail, "UPDATE_USER_SECTIONS", email, JSON.stringify(sections));
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบ User" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}*/


// ─────────────────────────────────────────
// 8. SECURE DATA CLEARING (Admin)
// ─────────────────────────────────────────
function clearSectionData(adminEmail, adminPassword, targetYear, targetSectionId, fieldsToClearJson) {
  try {
    const adminSheet = ensureSheetExists("tb_admins");
    const admins     = adminSheet.getDataRange().getValues();
    let   isAuthenticated = false;
    for (let i = 1; i < admins.length; i++) {
      if (admins[i][0] === adminEmail && String(admins[i][1]) === String(adminPassword)) {
        isAuthenticated = true;
        break;
      }
    }
    if (!isAuthenticated) {
      return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
    }

    const dataSheet = ensureSheetExists("tb_data");
    const data      = dataSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) === String(targetYear) && data[i][3] === targetSectionId) {
        let existingData = {};
        try { existingData = JSON.parse(data[i][6] || "{}"); } catch(e) {}
        const fieldsToClear = JSON.parse(fieldsToClearJson);
        fieldsToClear.forEach(f => delete existingData[f]);
        dataSheet.getRange(i + 1, 7).setValue(JSON.stringify(existingData));
        writeAuditLog(adminEmail, "CLEAR_DATA", `${targetSectionId}:${targetYear}`, `Fields: ${fieldsToClearJson}`);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบข้อมูล" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 9. EXPORT TO GOOGLE SHEETS
// ─────────────────────────────────────────
function exportYearDataToSheet(year) {
  try {
    const adminEmail      = Session.getActiveUser().getEmail();
    const ss              = getOrCreateSpreadsheet();
    const exportSheetName = `Export_${year}_${new Date().getTime()}`;
    const exportSheet     = ss.insertSheet(exportSheetName);

    exportSheet.appendRow([
      "ปีการศึกษา", "Section", "ผู้บันทึกล่าสุด", "วันที่อัปเดต", "ข้อมูล (JSON)"
    ]);
    exportSheet.getRange(1, 1, 1, 5)
      .setBackground('#1a237e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    const dataSheet = ensureSheetExists("tb_data");
    const data      = dataSheet.getDataRange().getValues();
    let   rowCount  = 0;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]) !== String(year)) continue;
      const sectionName = data[i][4] || data[i][3];
      const lastUpdate  = data[i][5] instanceof Date
        ? data[i][5].toLocaleDateString('th-TH')
        : data[i][5];
      exportSheet.appendRow([
        data[i][2], sectionName, data[i][7], lastUpdate, data[i][6]
      ]);
      rowCount++;
    }

    const exportUrl = ss.getUrl() + "#gid=" + exportSheet.getSheetId();
    writeAuditLog(adminEmail, "EXPORT", year, `Rows: ${rowCount}`);
    return { success: true, url: exportUrl, rows: rowCount, sheetName: exportSheetName };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// NEW: THEME & LOGO & COLLEGE SETTINGS
// ─────────────────────────────────────────
/*function saveThemeSettings(themeData) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const sh = ensureSheetExists("tb_theme");
    const data = sh.getDataRange().getValues();
    const keysToSave = ["primary","gradientCss","gradientId","logoUrl","collegeName","logoFileId","iqaMapping"];
    keysToSave.forEach(key => {
      const val = themeData[key] !== undefined ? String(themeData[key]) : "";
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
          sh.getRange(i + 1, 1, 1, 4).setValues([[key, val, new Date(), adminEmail]]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow([key, val, new Date(), adminEmail]);
    });
    writeAuditLog(adminEmail, "SAVE_THEME", "theme", JSON.stringify(Object.keys(themeData)));
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}*/
function saveThemeSettings(themeData) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sh = ensureSheetExists("tb_theme");
    const data = sh.getDataRange().getValues();
    const keysToSave = ["primary","gradientCss","gradientId","logoUrl","collegeName","logoFileId","iqaMapping"];
    
    keysToSave.forEach(key => {
      // ✅ แก้ไข: อัปเดตเฉพาะ key ที่ถูกส่งมาเท่านั้น ป้องกันค่าอื่นๆ โดนลบทิ้งเป็นค่าว่าง
      if (themeData[key] !== undefined) {
        const val = String(themeData[key]);
        let found = false;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === key) {
            sh.getRange(i + 1, 1, 1, 4).setValues([[key, val, new Date(), adminEmail]]);
            found = true; break;
          }
        }
        if (!found) sh.appendRow([key, val, new Date(), adminEmail]);
      }
    });
    
    writeAuditLog(adminEmail, "SAVE_THEME", "theme", "Updated theme settings");
    return { success: true };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}


function getThemeSettings() {
  try {
    const sh = ensureSheetExists("tb_theme");
    const data = sh.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) result[String(data[i][0])] = String(data[i][1] || "");
    }
    return { success: true, data: result };
  } catch(e) { return { success: true, data: {} }; }
}

// ─────────────────────────────────────────
// NEW: LOGO UPLOAD (Base64 → Drive → URL)
// ─────────────────────────────────────────
function uploadLogoImage(base64Data, mimeType, fileName) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const folder = getScriptFolder();
    // ลบโลโก้เก่าถ้ามี
    const themeData = getThemeSettings().data || {};
    if (themeData.logoFileId) {
      try { DriveApp.getFileById(themeData.logoFileId).setTrashed(true); } catch(e) {}
    }
    // สร้างไฟล์ใหม่
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, fileName || "college_logo.png");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId  = file.getId();
    const fileUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w500";
    // บันทึก theme
    saveThemeSettings({ logoUrl: fileUrl, logoFileId: fileId });
    writeAuditLog(adminEmail, "UPLOAD_LOGO", fileId, fileName);
    return { success: true, url: fileUrl, fileId: fileId };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function deleteLogoImage() {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const themeData = getThemeSettings().data || {};
    if (themeData.logoFileId) {
      try { DriveApp.getFileById(themeData.logoFileId).setTrashed(true); } catch(e) {}
    }
    saveThemeSettings({ logoUrl: "", logoFileId: "" });
    writeAuditLog(adminEmail, "DELETE_LOGO", "", "");
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─────────────────────────────────────────
// NEW: COPY DATA FROM PREVIOUS YEAR
// ─────────────────────────────────────────
function getAvailableYearsForSection(sectionId) {
  try {
    const sh = ensureSheetExists("tb_data");
    const data = sh.getDataRange().getValues();
    const years = new Set();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]) === String(sectionId) && data[i][2]) {
        years.add(String(data[i][2]));
      }
    }
    return { success: true, years: Array.from(years).sort((a,b) => Number(b)-Number(a)) };
  } catch(e) { return { success: false, years: [] }; }
}

function getSectionDataByYear(sectionId, year) {
  try {
    const sh = ensureSheetExists("tb_data");
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]) === String(sectionId) && String(data[i][2]) === String(year)) {
        let dataObj = {};
        try { dataObj = JSON.parse(data[i][6] || "{}"); } catch(e) {}
        return {
          success: true,
          data: dataObj,
          submitterName: data[i][7] || "",
          lastUpdate: data[i][5] instanceof Date ? data[i][5].toISOString() : String(data[i][5] || "")
        };
      }
    }
    return { success: false, data: {}, message: "ไม่พบข้อมูลปีนี้" };
  } catch(e) { return { success: false, data: {}, message: e.toString() }; }
}

// ─────────────────────────────────────────
// 10. DASHBOARD ANALYTICS
// ─────────────────────────────────────────
function getDashboardAnalytics(year) {
  try {
    const dashResult     = fetchDashboardData();
    const sectionsResult = getAllSections();
    const allData        = dashResult.data     || [];
    const allSections    = sectionsResult.data  || [];
    const yearData       = allData.filter(d => String(d.year) === String(year));

    const total          = allSections.filter(s => s.isActive).length;
    const submitted      = yearData.length;
    const pending        = total - submitted;
    const submissionRate = total > 0 ? Math.round((submitted / total) * 100) : 0;

    const monthlyTrend = {};
    yearData.forEach(d => {
      if (d.lastUpdate) {
        const month = new Date(d.lastUpdate).getMonth() + 1;
        monthlyTrend[month] = (monthlyTrend[month] || 0) + 1;
      }
    });

    return {
      success:   true,
      analytics: {
        total, submitted, pending, submissionRate,
        monthlyTrend: monthlyTrend,
        sections: allSections.map(s => ({
          ...s,
          hasData:    yearData.some(d => d.sectionId === s.sectionId),
          submission: yearData.find(d => d.sectionId === s.sectionId) || null
        }))
      }
    };
  } catch(e) {
    return { success: false, analytics: {}, message: e.toString() };
  }
}

// ── CacheService helpers ──────────────────────────────────
function _getCache(key) {
  try {
    var c = CacheService.getScriptCache();
    var v = c.get(key);
    return v ? JSON.parse(v) : null;
  } catch(e) { return null; }
}

function _putCache(key, data, ttl) {
  try {
    var c = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    if (json.length < 95000) c.put(key, json, ttl || 300);
  } catch(e) {}
}

function _clearAppCache() {
  try {
    CacheService.getScriptCache().removeAll(['app_static_data']);
  } catch(e) {}
}

// ─────────────────────────────────────────
// 11. LOAD ALL INITIAL DATA (Single Call)
// ─────────────────────────────────────────
function loadInitialData() {
  try {
    initializeDatabase();

    // Always fetch fresh: dashboard and visits
    var dashboard = fetchDashboardData();
    var visits    = getVisitStats();
    try { recordVisit("view", "guest"); } catch(e) {}

    // Try cache for static data (changes only on admin saves)
    var staticData = _getCache('app_static_data');
    if (!staticData) {
      staticData = {
        settings:    getSystemSettings().data    || [],
        sections:    getAllSections().data       || [],
        departments: getAllDepartments().data    || [],
        forms:       getAllForms().data          || [],
        theme:       getThemeSettings().data     || {},
        iqaMapping:  getIqaMappingData().data    || {},
        iqaSchema:   getIqaSchema().data        || { phases: [] },
        allUsers:    getAllUsers().data          || { admins: [], managers: [], users: [] }
      };
      _putCache('app_static_data', staticData, 300); // cache 5 min
    }

    return {
      success:     true,
      settings:    staticData.settings,
      sections:    staticData.sections,
      departments: staticData.departments,
      forms:       staticData.forms,
      theme:       staticData.theme,
      iqaMapping:  staticData.iqaMapping,
      iqaSchema:   staticData.iqaSchema,
      allUsers:    staticData.allUsers,
      dashboard:   dashboard.data || [],
      visits:      visits
    };
  } catch(e) {
    console.error("🔥 ERROR in loadInitialData: " + e.stack);
    return {
      success:     false,
      message:     e.toString(),
      settings:    [],
      sections:    [],
      departments: [],
      forms:       [],
      theme:       {},
      visits:      { total: 0, uniqueUsers: 0, today: 0 },
      iqaMapping:  {},
      allUsers:    { admins: [], managers: [], users: [] },
      dashboard:   []
    };
  }
}

function getVisitStats() {
  try {
    const sh = ensureSheetExists("tb_visits");
    const data = sh.getDataRange().getValues();
    let total = 0, today = 0;
    const uniq = new Set();
    const todayStr = new Date().toDateString();
    for (let i = 1; i < data.length; i++) {
      total++;
      if (data[i][2] && data[i][2] !== "guest") uniq.add(String(data[i][2]));
      if (data[i][0] instanceof Date && data[i][0].toDateString() === todayStr) today++;
    }
    return { total, uniqueUsers: uniq.size, today };
  } catch(e) {
    return { total: 0, uniqueUsers: 0, today: 0 };
  }
}

function recordVisit(type, email) {
  try {
    ensureSheetExists("tb_visits").appendRow([new Date(), type||"view", email||"guest", "", ""]);
  } catch(e) {}
}

function TEST_fetchData() {
  // จำลองการดึงข้อมูลตอนโหลดแอป
  const result = loadInitialData();
  
  if (result.success) {
    Logger.log("✅ ดึงข้อมูลสำเร็จ!");
    Logger.log("📊 จำนวน Section: " + result.sections.length);
    Logger.log("📋 จำนวน Form: " + result.forms.length);
    Logger.log("📈 จำนวน Data: " + result.dashboard.length);
    
    if(result.sections.length === 0) {
      Logger.log("⚠️ หมายเหตุ: ตาราง Section ยังว่างเปล่า (ไม่มีข้อมูลให้ดึง)");
    }
  } else {
    Logger.log("❌ เกิดข้อผิดพลาด: " + result.message);
  }
}

function DIAGNOSTIC_CHECK() {
  const result = loadInitialData();
  
  Logger.log("=========================================");
  Logger.log("🔍 SYSTEM DIAGNOSTIC (การวิเคราะห์ระบบ)");
  Logger.log("=========================================");
  
  if (!result.success) {
    Logger.log("❌ ระบบหลังบ้านล้มเหลว: " + result.message);
    return;
  }

  Logger.log("1. ตาราง tb_settings (ปีการศึกษา): พบ " + result.settings.length + " รายการ");
  if (result.settings.length > 0) {
    Logger.log("   -> ปีที่มีในระบบ: " + result.settings.map(s => s.year).join(", "));
  } else {
    Logger.log("   ❌ ปัญหา: คุณยังไม่ได้ตั้งค่าปีการศึกษาใน tb_settings ระบบจะไม่มี Dropdown ให้เลือกปี");
  }

  Logger.log("2. ตาราง tb_sections (หมวดหมู่): พบ " + result.sections.length + " รายการ");
  if (result.sections.length === 0) {
    Logger.log("   ❌ ปัญหา: ไม่มี Section ระบบจึงไม่สามารถสร้าง Card เพื่อแสดงผลได้");
  } else {
    const activeSecs = result.sections.filter(s => s.isActive);
    Logger.log("   -> Section ที่เปิดใช้งาน (Active): " + activeSecs.length + " หมวดหมู่");
  }

  Logger.log("3. ตาราง tb_data (ข้อมูลที่บันทึก): พบ " + result.dashboard.length + " รายการ");
  if (result.dashboard.length === 0) {
    Logger.log("   ❌ ปัญหา: Backend ดึงข้อมูลจาก tb_data ไม่ได้ (อาจไม่มี Year หรือ SectionID ในแถวนั้น)");
  } else {
    Logger.log("   -> ตัวอย่างข้อมูลที่ดึงได้: SectionID = " + result.dashboard[0].sectionId + " | Year = " + result.dashboard[0].year);
  }
  
  Logger.log("=========================================");
}

// ─────────────────────────────────────────
// IQA MAPPING MANAGEMENT
// ─────────────────────────────────────────
function getIqaMappingData() {
  try {
    const sheet = ensureSheetExists("tb_iqa_mapping");
    const data = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        try { result[data[i][0]] = JSON.parse(data[i][1]); } catch(e) {}
      }
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, data: {} }; }
}

function saveIqaMappingData(mappingData) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet = ensureSheetExists("tb_iqa_mapping");
    const data = sheet.getDataRange().getValues();
    const keys = Object.keys(mappingData);
    
    // อัปเดตหรือเพิ่มข้อมูลใหม่ตาม PageID
    keys.forEach(pageId => {
      let valStr = JSON.stringify(mappingData[pageId]);
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === pageId) {
          sheet.getRange(i + 1, 2, 1, 3).setValues([[valStr, new Date(), adminEmail]]);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([pageId, valStr, new Date(), adminEmail]);
    });
    writeAuditLog(adminEmail, "SAVE_IQA_MAPPING", "tb_iqa_mapping", `Updated ${keys.length} pages`);
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─────────────────────────────────────────
// IQA SCHEMA (Dynamic Builder) — Task 1
// เก็บโครงสร้าง Phase / Page / Item แบบ Dynamic ใน tb_iqa_mapping
// ─────────────────────────────────────────
function getIqaSchema() {
  try {
    const sheet = ensureSheetExists("tb_iqa_mapping");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === "iqa_schema_config") {
        try { return { success: true, data: JSON.parse(data[i][1] || '{"phases":[]}') }; } catch(e) {}
      }
    }
    return { success: true, data: { phases: [] } };
  } catch(e) { return { success: false, data: { phases: [] } }; }
}

function saveIqaSchema(schemaObj) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet = ensureSheetExists("tb_iqa_mapping");
    const data  = sheet.getDataRange().getValues();
    const valStr = JSON.stringify(schemaObj);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === "iqa_schema_config") {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[valStr, new Date(), adminEmail]]);
        writeAuditLog(adminEmail, "SAVE_IQA_SCHEMA", "tb_iqa_mapping", "Schema updated");
        return { success: true, message: "บันทึก IQA Schema สำเร็จ" };
      }
    }
    sheet.appendRow(["iqa_schema_config", valStr, new Date(), adminEmail]);
    writeAuditLog(adminEmail, "SAVE_IQA_SCHEMA", "tb_iqa_mapping", "Schema created");
    return { success: true, message: "บันทึก IQA Schema สำเร็จ" };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ─────────────────────────────────────────
// AUDIT LOG EXPORT — Task 5
// ส่ง log ดิบจาก tb_audit กลับไปให้ Client Download
// ─────────────────────────────────────────
function getAuditLog(limitRows) {
  try {
    const sheet = ensureSheetExists("tb_audit");
    const data  = sheet.getDataRange().getValues();
    const limit = Math.min(limitRows || 3000, 5000);
    const startRow = Math.max(1, data.length - limit);
    const rows = [];
    for (let i = startRow; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({
        timestamp: data[i][0] instanceof Date ? data[i][0].toISOString() : String(data[i][0]),
        email:     String(data[i][1] || ""),
        action:    String(data[i][2] || ""),
        target:    String(data[i][3] || ""),
        detail:    String(data[i][4] || "")
      });
    }
    rows.reverse(); // เรียงใหม่ล่าสุดก่อน
    return { success: true, data: rows, total: rows.length };
  } catch(e) { return { success: false, data: [], message: e.toString() }; }
}

// ─────────────────────────────────────────
// NEW: SAVE LOGO AS BASE64 (ไม่ต้องพึ่งพิง Google Drive)
// ─────────────────────────────────────────
function saveBase64Logo(dataUrl) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    
    // บันทึกรูปภาพ (ที่เป็น Text Data URL) ลงในฐานข้อมูล Theme โดยตรง
    saveThemeSettings({ logoUrl: dataUrl, logoFileId: "base64_embedded" });
    writeAuditLog(adminEmail, "UPLOAD_LOGO", "Base64", "Uploaded Embedded Logo");
    
    return { success: true, url: dataUrl };
  } catch(e) { 
    return { success: false, message: e.toString() }; 
  }
}

// ─────────────────────────────────────────
// DEPARTMENT MANAGEMENT
// ─────────────────────────────────────────
function getAllDepartments() {
  try {
    const sheet = ensureSheetExists("tb_departments");
    const data = sheet.getDataRange().getValues();
    const depts = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      let assigned = [];
      try { assigned = data[i][3] ? JSON.parse(data[i][3]) : []; } catch(e) {}
      
      let cDate = data[i][7];
      if (cDate instanceof Date) { cDate = cDate.toISOString(); }

      depts.push({
        deptId:           data[i][0],
        deptName:         data[i][1],
        description:      data[i][2],
        assignedSections: assigned,
        iconClass:        data[i][4] || "bi-building",
        colorClass:       data[i][5] || "blue",
        isActive:         data[i][6] !== false,
        createdAt:        cDate,
        createdBy:        data[i][8]
      });
    }
    return { success: true, data: depts };
  } catch(e) {
    return { success: false, data: [], message: e.toString() };
  }
}

function saveDepartment(deptData) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet      = ensureSheetExists("tb_departments");
    const data       = sheet.getDataRange().getValues();
    const assignedStr= JSON.stringify(deptData.assignedSections || []);

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === deptData.deptId) {
        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          deptData.deptId, deptData.deptName, deptData.description, assignedStr,
          deptData.iconClass, deptData.colorClass, deptData.isActive, data[i][7], data[i][8]
        ]]);
        writeAuditLog(adminEmail, "UPDATE_DEPT", deptData.deptId, deptData.deptName);
        return { success: true, message: "อัปเดต Department สำเร็จ" };
      }
    }

    const newId = "DEPT_" + Date.now();
    sheet.appendRow([
      newId, deptData.deptName, deptData.description, assignedStr,
      deptData.iconClass, deptData.colorClass, true, new Date(), adminEmail
    ]);
    writeAuditLog(adminEmail, "CREATE_DEPT", newId, deptData.deptName);
    return { success: true, message: "สร้าง Department สำเร็จ", deptId: newId };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteDepartment(deptId) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet = ensureSheetExists("tb_departments");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === deptId) {
        sheet.deleteRow(i + 1);
        writeAuditLog(adminEmail, "DELETE_DEPT", deptId, data[i][1]);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบ Department" };
  } catch(e) { return { success: false, message: e.toString() }; }
}