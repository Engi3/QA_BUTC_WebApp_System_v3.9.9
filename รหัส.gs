

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
    "tb_departments": ["DeptID", "DeptName", "Description", "AssignedSections", "IconClass", "ColorClass", "IsActive", "CreatedAt", "CreatedBy"], // ✅ เพิ่ม tb_departments
    "tb_report_forms": ["FileID", "DocName", "FileName", "DownloadURL", "ViewURL", "UploadedAt", "UploadedBy"] // ✅ เพิ่ม tb_report_forms สำหรับเก็บ Metadata แบบฟอร์มรายงาน
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
      "tb_settings", "tb_audit", "tb_theme", "tb_visits", "tb_iqa_mapping", "tb_departments", "tb_report_forms" // ✅ เพิ่ม tb_report_forms
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

    // ไม่พบผู้ใช้ในระบบ: @butc.ac.th → ลงทะเบียนเองได้, อื่นๆ → ต้องให้ Admin เพิ่ม
    if (activeEmail.endsWith("@butc.ac.th")) {
      return { status: "unregistered", email: activeEmail, switchUrl: switchUrl };
    }
    return { status: "wrong_domain", email: activeEmail, url: switchUrl, switchUrl: switchUrl };

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
// ═══════════════════════════════════════════════════════════════
// Dynamic Chunking Helpers (tb_data)
//
// โครงสร้าง column ของ tb_data:
//   A(1): timestamp_created  B(2): userEmail   C(3): year
//   D(4): sectionId          E(5): sectionName  F(6): lastUpdate
//   G(7): data_chunk_1       H(8): userName
//   I(9): data_chunk_2  J(10): data_chunk_3  ... (ขยายไปเรื่อยๆ)
//
// JSON ข้อมูลทั้งหมด (รวมรูปภาพ base64) จะถูกแบ่งเป็น chunks
// ขนาดไม่เกิน CHUNK_SIZE ตัวอักษรต่อ chunk และจัดเก็บต่อกัน
// _readChunkedData รวม chunks กลับเป็น string แล้ว JSON.parse
// ═══════════════════════════════════════════════════════════════

const DATA_CHUNK_SIZE = 45000;   // ตัวอักษรต่อ chunk (< 50,000 limit ของ Sheets)
const DATA_MAX_CHUNKS = 20;      // รองรับสูงสุด 20 chunks = ~900KB JSON
const DATA_COL_CHUNK1 = 7;       // column index (1-based) ของ chunk แรก = col G
const DATA_COL_UNAME  = 8;       // column index ของ userName = col H
const DATA_COL_OVERFLOW = 9;     // column index เริ่มต้นของ overflow chunks = col I

/**
 * อ่าน chunks จาก row array แล้วคืน Object
 * รองรับทั้งข้อมูลเก่า (ไม่มี chunk) และใหม่ (multi-chunk)
 */
function _readChunkedData(rowArr) {
  // chunk แรกอยู่ที่ index 6 (col G, 0-based)
  let fullJson = String(rowArr[DATA_COL_CHUNK1 - 1] || "{}");
  // overflow chunks อยู่ที่ index 8+ (col I+, 0-based)
  for (let c = DATA_COL_OVERFLOW - 1; c < rowArr.length; c++) {
    if (!rowArr[c]) break;
    fullJson += String(rowArr[c]);
  }
  try { return JSON.parse(fullJson); } catch(e) { return {}; }
}

/**
 * บันทึก mergedData ลง sheet row ด้วย dynamic chunking
 * chunk แรกเขียนที่ col G (DATA_COL_CHUNK1)
 * overflow เขียนที่ col I+ (DATA_COL_OVERFLOW+)
 * คืน error message string ถ้าข้อมูลใหญ่เกิน limit, null ถ้าสำเร็จ
 */
function _writeChunkedData(sheet, rowNum, mergedData) {
  const fullJson = JSON.stringify(mergedData);
  if (fullJson.length > DATA_CHUNK_SIZE * DATA_MAX_CHUNKS) {
    return `ข้อมูลมีขนาดใหญ่เกินไป (${Math.round(fullJson.length / 1024)} KB / สูงสุด ${Math.round(DATA_CHUNK_SIZE * DATA_MAX_CHUNKS / 1024)} KB) กรุณาลดขนาดรูปภาพหรือจำนวนข้อมูลในตาราง`;
  }
  // แบ่ง chunks
  const chunks = [];
  for (let i = 0; i < fullJson.length; i += DATA_CHUNK_SIZE) {
    chunks.push(fullJson.substring(i, i + DATA_CHUNK_SIZE));
  }
  // เขียน chunk แรกที่ col G
  sheet.getRange(rowNum, DATA_COL_CHUNK1).setValue(chunks[0] || "{}");
  // เขียน overflow chunks ที่ col I+ (clear ทุก slot ที่ไม่ใช้ด้วย setValues ครั้งเดียว)
  const overflowSlots = DATA_MAX_CHUNKS - 1;
  const overflowRow = [];
  for (let c = 0; c < overflowSlots; c++) {
    overflowRow.push(c + 1 < chunks.length ? chunks[c + 1] : "");
  }
  sheet.getRange(rowNum, DATA_COL_OVERFLOW, 1, overflowSlots).setValues([overflowRow]);
  return null; // success
}

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
        existingData = _readChunkedData(data[i]);
        break;
      }
    }

    const mergedData = { ...existingData, ...formDataObj };

    if (rowToUpdate > 0) {
      sheet.getRange(rowToUpdate, 2).setValue(userEmail);
      sheet.getRange(rowToUpdate, 6).setValue(timestamp);
      sheet.getRange(rowToUpdate, DATA_COL_UNAME).setValue(userName);
      const err = _writeChunkedData(sheet, rowToUpdate, mergedData);
      if (err) return { success: false, message: err };
    } else {
      // สร้าง row ใหม่: เขียน metadata ก่อน แล้วค่อย write chunks แยก
      // เพราะ appendRow ไม่รองรับ dynamic overflow ได้ดีเท่า
      sheet.appendRow([
        timestamp, userEmail, year,
        sectionId, sectionName,
        timestamp, "{}", userName
      ]);
      const newRowNum = sheet.getLastRow();
      const err = _writeChunkedData(sheet, newRowNum, mergedData);
      if (err) return { success: false, message: err };
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
      // ── รวม chunks กลับเป็น JSON string ──
      let fullJson = String(data[i][DATA_COL_CHUNK1 - 1] || "{}");
      for (let c = DATA_COL_OVERFLOW - 1; c < data[i].length; c++) {
        if (!data[i][c]) break;
        fullJson += String(data[i][c]);
      }
      summary.push({
        year:          data[i][2],
        sectionId:     data[i][3],
        sectionName:   data[i][4],
        lastUpdate:    lUpdate,
        submitterName: data[i][DATA_COL_UNAME - 1] || data[i][1],
        dataJson:      fullJson
      });
    }
    return { success: true, data: summary };
  } catch(e) {
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
        return {
          success:       true,
          data:          _readChunkedData(data[i]),
          submitterName: data[i][DATA_COL_UNAME - 1],
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
        authType:  uData[i][1] || 'SSO_GOOGLE',
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
    if (!email || !email.includes("@")) return { success: false, message: "รูปแบบอีเมลไม่ถูกต้อง" };

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
      const authType = email.endsWith("@butc.ac.th") ? "SSO_GOOGLE" : "ADMIN_ADDED";
      const uSheet = ensureSheetExists("tb_users");
      uSheet.appendRow([email, authType, data.prefix, data.fname, data.lname, data.phone, positionsStr, sectionsStr, new Date()]);
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
        // ล้างข้อมูลทั้งหมด: chunk แรก + overflow chunks ทั้งหมด (dynamic)
        dataSheet.getRange(i + 1, DATA_COL_CHUNK1).setValue("{}");
        const clearRow = new Array(DATA_MAX_CHUNKS - 1).fill("");
        dataSheet.getRange(i + 1, DATA_COL_OVERFLOW, 1, DATA_MAX_CHUNKS - 1).setValues([clearRow]);
        writeAuditLog(adminEmail, "CLEAR_DATA", `${targetSectionId}:${targetYear}`, `ALL fields wiped`);
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
        return {
          success:       true,
          data:          _readChunkedData(data[i]),
          submitterName: data[i][DATA_COL_UNAME - 1] || "",
          lastUpdate:    data[i][5] instanceof Date ? data[i][5].toISOString() : String(data[i][5] || "")
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
        settings:     getSystemSettings().data    || [],
        sections:     getAllSections().data        || [],
        departments:  getAllDepartments().data     || [],
        forms:        getAllForms().data           || [],
        theme:        getThemeSettings().data      || {},
        iqaMapping:   getIqaMappingData().data     || {},
        iqaSchema:    getIqaSchema().data          || { phases: [] },
        allUsers:     getAllUsers().data           || { admins: [], managers: [], users: [] },
        reportForms:      getReportFormsList().data        || [], // ✅ โหลดรายการแบบฟอร์มรายงาน
        reportFormGroups: (function() { try { return getReportFormGroupsConfig().groups || []; } catch(e) { return []; } })(),
        reportFormOrder:  (function() { try { return getReportFormGroupsConfig().order  || {}; } catch(e) { return {}; } })(),
        docTracking:      (function() { try { return getDocTracking().data || { groups: [] }; } catch(e) { return { groups: [] }; } })()
      };
      _putCache('app_static_data', staticData, 300); // cache 5 min
    }

    return {
      success:          true,
      settings:         staticData.settings,
      sections:         staticData.sections,
      departments:      staticData.departments,
      forms:            staticData.forms,
      theme:            staticData.theme,
      iqaMapping:       staticData.iqaMapping,
      iqaSchema:        staticData.iqaSchema,
      allUsers:         staticData.allUsers,
      reportForms:      staticData.reportForms,
      reportFormGroups: staticData.reportFormGroups,
      reportFormOrder:  staticData.reportFormOrder,
      docTracking:      staticData.docTracking,
      dashboard:    dashboard.data || [],
      visits:       visits
    };
  } catch(e) {
    console.error("🔥 ERROR in loadInitialData: " + e.stack);
    return {
      success:      false,
      message:      e.toString(),
      settings:     [],
      sections:     [],
      departments:  [],
      forms:        [],
      theme:        {},
      visits:       { total: 0, uniqueUsers: 0, today: 0 },
      iqaMapping:   {},
      allUsers:         { admins: [], managers: [], users: [] },
      reportForms:      [],
      reportFormGroups: [],
      reportFormOrder:  {},
      docTracking:      { groups: [] },
      dashboard:        []
    };
  }
}

// ─────────────────────────────────────────
// 14. DOC TRACKING
// ─────────────────────────────────────────
function getDocTracking() {
  try {
    const sh = ensureSheetExists("tb_theme");
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === "doc_tracking") {
        try { return { success: true, data: JSON.parse(data[i][1] || '{"groups":[]}') }; }
        catch(e) { return { success: true, data: { groups: [] } }; }
      }
    }
    return { success: true, data: { groups: [] } };
  } catch(e) {
    return { success: false, data: { groups: [] }, message: e.toString() };
  }
}

function saveDocTracking(jsonStr) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sh = ensureSheetExists("tb_theme");
    const data = sh.getDataRange().getValues();
    const val = (typeof jsonStr === "string") ? jsonStr : JSON.stringify(jsonStr);
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === "doc_tracking") {
        sh.getRange(i + 1, 1, 1, 4).setValues([["doc_tracking", val, new Date(), adminEmail]]);
        writeAuditLog(adminEmail, "SAVE_DOC_TRACKING", "doc_tracking", "updated");
        return { success: true };
      }
    }
    sh.appendRow(["doc_tracking", val, new Date(), adminEmail]);
    writeAuditLog(adminEmail, "SAVE_DOC_TRACKING", "doc_tracking", "created");
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getVisitStats() {
  try {
    const sh = ensureSheetExists("tb_visits");
    const data = sh.getDataRange().getValues();
    let total = 0, today = 0;
    const uniq = new Set();
    const tz = Session.getScriptTimeZone();
    const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

    // tb_visits ไม่มี header row → เริ่มจาก i=0
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === "daily_agg") {
        const count = Number(data[i][2]) || 0;
        total += count;
        if (String(data[i][0]) === todayKey) today += count;
        try {
          JSON.parse(data[i][3] || "[]").forEach(u => { if (u) uniq.add(u); });
        } catch(e) {}
      } else {
        // raw row (backward compatible)
        total++;
        if (data[i][2] && data[i][2] !== "guest") uniq.add(String(data[i][2]));
        if (data[i][0] instanceof Date) {
          if (Utilities.formatDate(data[i][0], tz, "yyyy-MM-dd") === todayKey) today++;
        }
      }
    }
    return { total, uniqueUsers: uniq.size, today };
  } catch(e) {
    return { total: 0, uniqueUsers: 0, today: 0 };
  }
}

/**
 * [OPTIMIZED] บันทึกการเข้าชมแบบ Daily Aggregate
 * tb_visits ไม่มี header → loop เริ่มจาก i=0
 */
function recordVisit(type, email) {
  try {
    const sh       = ensureSheetExists("tb_visits");
    const tz       = Session.getScriptTimeZone();
    const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const data     = sh.getDataRange().getValues();

    // tb_visits ไม่มี header row → เริ่มจาก i=0
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] === "daily_agg" && String(data[i][0]) === todayKey) {
        const newCount = (Number(data[i][2]) || 0) + 1;
        let uniqueSet = new Set();
        try { uniqueSet = new Set(JSON.parse(data[i][3] || "[]")); } catch(e) {}
        if (email && email !== "guest") uniqueSet.add(email);
        sh.getRange(i + 1, 3, 1, 2).setValues([[newCount, JSON.stringify(Array.from(uniqueSet))]]);
        return;
      }
    }

    const uniqueArr = (email && email !== "guest") ? JSON.stringify([email]) : "[]";
    sh.appendRow([todayKey, "daily_agg", 1, uniqueArr, ""]);
  } catch(e) {}
}

// ─────────────────────────────────────────
// 12. REPORT FORMS MANAGEMENT
// จัดการแบบฟอร์มรายงาน: อัปโหลด/ดึง/ลบ/เปลี่ยนชื่อ
// ─────────────────────────────────────────

/**
 * ค้นหาหรือสร้าง Sub-folder "QA_Report_Forms"
 * ภายใน Folder เดียวกับ Script เพื่อเก็บไฟล์แบบฟอร์ม
 */
function getOrCreateReportFormsFolder() {
  const parentFolder = getScriptFolder();
  const folderName   = "QA_Report_Forms";
  // ค้นหา Folder ที่มีชื่อตรงกันก่อน เพื่อไม่สร้างซ้ำ
  const iter = parentFolder.getFoldersByName(folderName);
  if (iter.hasNext()) {
    return iter.next();
  }
  // ไม่พบ → สร้างใหม่ภายใน Folder เดิม
  return parentFolder.createFolder(folderName);
}

/**
 * อัปโหลดไฟล์แบบฟอร์มรายงาน (รับข้อมูล Base64 จาก Frontend)
 * บันทึก Metadata ลงใน tb_report_forms
 * @param {string} base64Data - ข้อมูลไฟล์แบบ Base64
 * @param {string} mimeType   - ประเภทไฟล์ เช่น "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
 * @param {string} fileName   - ชื่อไฟล์จริง เช่น "report_template.docx"
 * @param {string} docName    - ชื่อเอกสารที่แสดงบนหน้าเว็บ (กรอกโดย Admin)
 */
function uploadReportForm(base64Data, mimeType, fileName, docName) {
  try {
    const adminEmail = Session.getActiveUser().getEmail();
    const folder     = getOrCreateReportFormsFolder();

    // แปลง Base64 → Blob แล้วสร้างไฟล์ใน Drive
    const decoded   = Utilities.base64Decode(base64Data);
    const blob      = Utilities.newBlob(decoded, mimeType, fileName);
    const driveFile = folder.createFile(blob);

    // ─────────────────────────────────────────────────────────
    // ตั้งค่าสิทธิ์การแชร์แบบ Fault-tolerant
    // บางองค์กร (Google Workspace) ปิดการแชร์ภายนอก ทำให้ setSharing() throw exception
    // ลำดับการลอง: DOMAIN (ภายในองค์กร) → ANYONE_WITH_LINK → ข้ามถ้าทั้งคู่ไม่ได้
    // ไม่ว่าจะ setSharing ได้หรือไม่ กระบวนการบันทึก Metadata ยังดำเนินต่อเสมอ
    // ─────────────────────────────────────────────────────────
    let sharingMode = "owner_only"; // บันทึกว่า Sharing ถูกตั้งค่าอะไรได้บ้าง
    try {
      // ลองแชร์ภายในโดเมนองค์กรก่อน (เหมาะสำหรับ Google Workspace โรงเรียน)
      driveFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      sharingMode = "domain";
    } catch(domainErr) {
      try {
        // ถ้าแชร์ Domain ไม่ได้ ลองแชร์แบบ Public Link
        driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        sharingMode = "anyone_with_link";
      } catch(anyoneErr) {
        // ถ้าทั้งคู่ไม่ได้ (Workspace Admin ล็อกสิทธิ์ไว้) → ข้ามต่อ ไฟล์เป็น Owner Only
        // Admin ยังเข้าถึงและแชร์ด้วยตัวเองได้จาก Google Drive
        console.warn("setSharing ไม่สำเร็จ (Workspace Admin อาจล็อกสิทธิ์): " + anyoneErr);
      }
    }

    const fileId      = driveFile.getId();
    // ใช้ Direct Download URL ที่รองรับทั้ง Workspace และ Personal Account
    const downloadUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
    const viewUrl     = driveFile.getUrl();

    // บันทึก Metadata ลงใน Sheet tb_report_forms
    const sheet = ensureSheetExists("tb_report_forms");
    sheet.appendRow([
      fileId,
      docName  || fileName,  // ชื่อเอกสารที่แสดง (ถ้าไม่ระบุ ใช้ชื่อไฟล์แทน)
      fileName,
      downloadUrl,
      viewUrl,
      new Date(),
      adminEmail
    ]);

    // เพิ่ม fileId เข้ากลุ่ม "ไม่จัดกลุ่ม" (key = "") ใน Config ลำดับกลุ่ม
    // เพื่อให้ไฟล์ใหม่แสดงผลตามลำดับที่กำหนดทันทีในหน้าสาธารณะ
    try {
      const rfConfig = getReportFormGroupsConfig();
      const rfOrder  = rfConfig.order || {};
      if (!Array.isArray(rfOrder[""])) rfOrder[""] = [];
      rfOrder[""].push(fileId);
      saveReportFormGroupsConfig(rfConfig.groups || [], rfOrder);
    } catch(orderErr) {
      console.warn("ไม่สามารถบันทึกลำดับกลุ่มได้ (ไม่กระทบการอัปโหลด): " + orderErr);
    }

    _clearAppCache(); // ล้าง Cache เพื่อให้ Frontend โหลดข้อมูลใหม่ได้ทันที
    writeAuditLog(adminEmail, "UPLOAD_REPORT_FORM", fileId, docName + " (" + fileName + ")");

    return {
      success:     true,
      fileId:      fileId,
      docName:     docName || fileName,
      downloadUrl: downloadUrl
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ดึงรายการแบบฟอร์มรายงานทั้งหมดจาก tb_report_forms
 * คืนค่าเป็น Array ของ Object สำหรับแสดงผลบน Frontend
 */
function getReportFormsList() {
  try {
    const sheet = ensureSheetExists("tb_report_forms");
    const data  = sheet.getDataRange().getValues();
    const forms = [];

    // วนอ่านข้อมูลตั้งแต่แถวที่ 2 (ข้ามหัวตาราง)
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // ข้ามแถวที่ FileID ว่าง
      forms.push({
        fileId:      String(data[i][0]),
        docName:     String(data[i][1] || ""),
        fileName:    String(data[i][2] || ""),
        downloadUrl: String(data[i][3] || ""),
        viewUrl:     String(data[i][4] || ""),
        uploadedAt:  data[i][5] instanceof Date ? data[i][5].toISOString() : String(data[i][5] || ""),
        uploadedBy:  String(data[i][6] || "")
      });
    }
    return { success: true, data: forms };
  } catch(e) {
    return { success: true, data: [] };
  }
}

/**
 * ลบแบบฟอร์มรายงาน:
 * 1. ย้ายไฟล์ใน Drive ไปที่ Trash
 * 2. ลบแถว Metadata ออกจาก tb_report_forms
 * @param {string} fileId - Drive File ID ที่ต้องการลบ
 */
function deleteReportForm(fileId) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();

    // ลบไฟล์จาก Drive (ย้ายไป Trash ไม่ได้ลบถาวร)
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch(driveErr) {
      // ถ้าไม่พบไฟล์ใน Drive (อาจถูกลบแล้ว) ให้ทำงานต่อได้
      console.warn("ไม่พบไฟล์ใน Drive: " + driveErr);
    }

    // ลบแถว Metadata ออกจาก Sheet (วนจากล่างขึ้นบนเพื่อ index ไม่เลื่อน)
    const sheet = ensureSheetExists("tb_report_forms");
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(fileId)) {
        sheet.deleteRow(i + 1); // +1 เพราะ getValues นับจาก 0 แต่ Sheet นับจาก 1
        break;
      }
    }

    // ลบ fileId ออกจากทุก group ใน Config ลำดับกลุ่ม
    // ป้องกัน fileId ค้างอยู่ใน order หลังลบแบบฟอร์มแล้ว
    try {
      const rfConfig = getReportFormGroupsConfig();
      const rfOrder  = rfConfig.order || {};
      Object.keys(rfOrder).forEach(function(gKey) {
        if (Array.isArray(rfOrder[gKey])) {
          rfOrder[gKey] = rfOrder[gKey].filter(function(id) { return String(id) !== String(fileId); });
        }
      });
      saveReportFormGroupsConfig(rfConfig.groups || [], rfOrder);
    } catch(orderErr) {
      console.warn("ไม่สามารถอัปเดตลำดับกลุ่มหลังลบ (ไม่กระทบการลบ): " + orderErr);
    }

    writeAuditLog(adminEmail, "DELETE_REPORT_FORM", fileId, "Deleted");
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * เปลี่ยนชื่อเอกสาร (DocName) ที่แสดงบนหน้าเว็บ
 * ไม่ได้เปลี่ยนชื่อไฟล์จริงใน Drive เพื่อหลีกเลี่ยง Link เสีย
 * @param {string} fileId     - Drive File ID ที่ต้องการแก้ไข
 * @param {string} newDocName - ชื่อเอกสารใหม่
 */
function renameReportForm(fileId, newDocName) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const sheet = ensureSheetExists("tb_report_forms");
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(fileId)) {
        // คอลัมน์ที่ 2 (index 1 → Sheet Column B) คือ DocName
        sheet.getRange(i + 1, 2).setValue(newDocName);
        writeAuditLog(adminEmail, "RENAME_REPORT_FORM", fileId, newDocName);
        return { success: true };
      }
    }
    return { success: false, message: "ไม่พบแบบฟอร์มที่ต้องการแก้ไข" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * แทนที่ไฟล์แบบฟอร์มรายงาน (Replace):
 * ลบไฟล์เก่าออกจาก Drive + อัปโหลดไฟล์ใหม่ + อัปเดต Metadata ใน Sheet
 * @param {string} oldFileId  - Drive File ID ของไฟล์เก่าที่จะถูกแทนที่
 * @param {string} base64Data - ข้อมูลไฟล์ใหม่แบบ Base64
 * @param {string} mimeType   - ประเภทไฟล์ใหม่
 * @param {string} fileName   - ชื่อไฟล์ใหม่
 * @param {string} newDocName - ชื่อเอกสารใหม่ที่แสดงบนหน้าเว็บ (ถ้าว่างใช้ชื่อเดิม)
 */
function replaceReportForm(oldFileId, base64Data, mimeType, fileName, newDocName) {
  try {
    _clearAppCache();
    const adminEmail = Session.getActiveUser().getEmail();
    const folder     = getOrCreateReportFormsFolder();

    // ลบไฟล์เก่าออกจาก Drive ก่อน
    try {
      DriveApp.getFileById(oldFileId).setTrashed(true);
    } catch(driveErr) {
      console.warn("ไม่พบไฟล์เก่าใน Drive: " + driveErr);
    }

    // อัปโหลดไฟล์ใหม่ (ใช้ Fault-tolerant setSharing เหมือน uploadReportForm)
    const decoded    = Utilities.base64Decode(base64Data);
    const blob       = Utilities.newBlob(decoded, mimeType, fileName);
    const driveFile  = folder.createFile(blob);
    try {
      driveFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    } catch(e1) {
      try { driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e2) {}
    }

    const newFileId     = driveFile.getId();
    const newDownloadUrl = "https://drive.google.com/uc?export=download&id=" + newFileId;
    const newViewUrl     = driveFile.getUrl();

    // อัปเดต Metadata ใน Sheet (แทนที่แถวเดิมของ oldFileId)
    const sheet = ensureSheetExists("tb_report_forms");
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(oldFileId)) {
        const docName = newDocName || String(data[i][1]); // ถ้าไม่ระบุชื่อใหม่ ใช้ชื่อเดิม
        sheet.getRange(i + 1, 1, 1, 7).setValues([[
          newFileId, docName, fileName, newDownloadUrl, newViewUrl, new Date(), adminEmail
        ]]);
        writeAuditLog(adminEmail, "REPLACE_REPORT_FORM", newFileId, "แทนที่ " + oldFileId);
        return { success: true, fileId: newFileId, downloadUrl: newDownloadUrl };
      }
    }
    return { success: false, message: "ไม่พบแบบฟอร์มเดิมในฐานข้อมูล" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// 13. REPORT FORM GROUP MANAGEMENT
// จัดการกลุ่มแบบฟอร์มรายงาน: อ่าน/บันทึก Config
// ─────────────────────────────────────────

/**
 * ดึง Config กลุ่มแบบฟอร์มรายงานจาก tb_theme
 * - groups: Array ชื่อกลุ่มตามลำดับ เช่น ["งานบุคลากร","งานวิชาการ"]
 * - order:  Object { "": [fileId,...], "ชื่อกลุ่ม": [fileId,...] }
 *           key = "" หมายถึงไฟล์ "ไม่จัดกลุ่ม"
 */
function getReportFormGroupsConfig() {
  try {
    const themeData = getThemeSettings().data || {};
    let groups = [];
    let order  = {};
    try { groups = JSON.parse(themeData.report_form_groups || "[]"); } catch(e) { groups = []; }
    try { order  = JSON.parse(themeData.report_form_order  || "{}"); } catch(e) { order  = {}; }
    // ตรวจสอบ Type ให้ถูกต้องก่อนส่งออก
    if (!Array.isArray(groups)) groups = [];
    if (typeof order !== "object" || Array.isArray(order)) order = {};
    return { success: true, groups: groups, order: order };
  } catch(e) {
    return { success: false, groups: [], order: {} };
  }
}

/**
 * บันทึก Config กลุ่มแบบฟอร์มรายงานลงใน tb_theme
 * @param {Array}  groups - ลำดับกลุ่ม เช่น ["งานบุคลากร","งานวิชาการ"]
 * @param {Object} order  - { "": [fileIds...], "ชื่อกลุ่ม": [fileIds...] }
 */
function saveReportFormGroupsConfig(groups, order) {
  try {
    _clearAppCache(); // ล้าง Cache ก่อนบันทึก เพื่อให้ Frontend โหลดข้อมูลใหม่ทันที
    const adminEmail = Session.getActiveUser().getEmail();
    const sh   = ensureSheetExists("tb_theme");
    const data = sh.getDataRange().getValues();

    // ตรวจสอบ Type ก่อนแปลงเป็น JSON
    const groupsJson = JSON.stringify(Array.isArray(groups) ? groups : []);
    const orderJson  = JSON.stringify((typeof order === "object" && !Array.isArray(order)) ? order : {});

    const toSave = {
      "report_form_groups": groupsJson,
      "report_form_order":  orderJson
    };

    // วน upsert ทีละ key ใน tb_theme
    Object.keys(toSave).forEach(function(key) {
      const val  = toSave[key];
      let found  = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
          sh.getRange(i + 1, 1, 1, 4).setValues([[key, val, new Date(), adminEmail]]);
          found = true; break;
        }
      }
      if (!found) sh.appendRow([key, val, new Date(), adminEmail]);
    });

    writeAuditLog(adminEmail, "SAVE_RF_GROUPS", "", "Updated RF groups config");
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ─────────────────────────────────────────
// TEST & DIAGNOSTIC FUNCTIONS
// ─────────────────────────────────────────

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

// ============================================================
// 13. MAINTENANCE MODULE
// ระบบดูแลรักษาฐานข้อมูลอัตโนมัติ — ป้องกัน Google Sheet เกิน 10 ล้าน Cell
// ใช้งาน: เรียก setupMaintenanceTriggers() ครั้งเดียวจาก Apps Script Editor
// ============================================================

/**
 * ดูสถิติขนาดของแต่ละ Sheet (สำหรับ Admin)
 * แสดงจำนวน rows, cols, cells และ % การใช้งาน
 */
function getDatabaseStats() {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheets = ss.getSheets();
    const stats = [];
    let totalCells = 0;

    sheets.forEach(sheet => {
      const rows  = sheet.getLastRow();
      const cols  = sheet.getLastColumn();
      const cells = rows * cols;
      totalCells += cells;
      stats.push({ name: sheet.getName(), rows, cols, cells });
    });

    stats.sort((a, b) => b.cells - a.cells);
    const LIMIT = 10000000;
    return {
      success:      true,
      totalCells,
      limitCells:   LIMIT,
      usagePercent: Math.round((totalCells / LIMIT) * 1000) / 10,
      sheets:       stats
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * บีบอัดข้อมูล tb_visits เก่า (raw format) ให้เป็น daily summary
 * [BATCH MODE] เขียนทับทั้ง sheet + ลบแถวส่วนเกิน — เร็วกว่า deleteRow ทีละแถวมาก
 * tb_visits ไม่มี header row
 */
function compressOldVisits() {
  try {
    const sh   = ensureSheetExists("tb_visits");
    const data = sh.getDataRange().getValues();
    const tz   = Session.getScriptTimeZone();

    // แยก existing daily_agg rows และ raw rows
    const aggMap  = {}; // dateKey → { count, users: Set }
    let   rawCount = 0;

    for (let i = 0; i < data.length; i++) { // i=0: no header
      if (data[i][1] === "daily_agg") {
        // Merge existing agg rows ด้วย (กรณีมี duplicate daily_agg)
        const dk = String(data[i][0]);
        if (!aggMap[dk]) aggMap[dk] = { count: 0, users: new Set() };
        aggMap[dk].count += Number(data[i][2]) || 0;
        try { JSON.parse(data[i][3] || "[]").forEach(u => { if (u) aggMap[dk].users.add(u); }); } catch(e) {}
      } else {
        // raw row → aggregate
        rawCount++;
        let dateKey;
        if (data[i][0] instanceof Date) {
          dateKey = Utilities.formatDate(data[i][0], tz, "yyyy-MM-dd");
        } else {
          dateKey = String(data[i][0]).substring(0, 10);
        }
        if (!dateKey || dateKey.length < 8) continue;
        if (!aggMap[dateKey]) aggMap[dateKey] = { count: 0, users: new Set() };
        aggMap[dateKey].count++;
        if (data[i][2] && data[i][2] !== "guest") aggMap[dateKey].users.add(String(data[i][2]));
      }
    }

    if (rawCount === 0) {
      return { success: true, message: "ไม่มีข้อมูลเก่าที่ต้องบีบอัด — ระบบพร้อมแล้ว" };
    }

    // สร้าง final rows เรียงตามวันที่
    const finalRows = Object.entries(aggMap)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([dk, { count, users }]) => [dk, "daily_agg", count, JSON.stringify(Array.from(users)), ""]);

    // [BATCH] เขียนทับ sheet ทั้งหมด
    sh.clearContents();
    if (finalRows.length > 0) {
      sh.getRange(1, 1, finalRows.length, 5).setValues(finalRows);
    }
    // ลบแถวส่วนเกิน (ลด cell count จริง)
    const totalRows = sh.getMaxRows();
    const keepRows  = Math.max(finalRows.length, 1);
    if (totalRows > keepRows) {
      sh.deleteRows(keepRows + 1, totalRows - keepRows);
    }

    writeAuditLog(
      Session.getActiveUser().getEmail(),
      "MAINTENANCE", "tb_visits",
      "Compressed " + rawCount + " raw rows → " + finalRows.length + " daily summaries"
    );
    return {
      success:   true,
      deleted:   rawCount,
      summaries: finalRows.length,
      message:   "บีบอัดสำเร็จ: ลบ " + rawCount + " rows → เหลือ " + finalRows.length + " daily summaries"
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ลบ Audit Log เก่าที่เกิน keepDays วัน
 * [BATCH MODE] เขียนทับ sheet + ลบแถวส่วนเกิน — เร็วกว่า deleteRow ทีละแถวมาก
 * @param {number} keepDays - จำนวนวันที่จะเก็บไว้ (default: 90)
 */
function pruneAuditLog(keepDays) {
  try {
    const days   = keepDays || 90;
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);

    const sh   = ensureSheetExists("tb_audit");
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return { success: true, deleted: 0, message: "ไม่มีข้อมูลที่ต้องลบ" };

    const header   = data[0];
    const keepRows = [header];
    let   deleted  = 0;

    for (let i = 1; i < data.length; i++) {
      const ts = data[i][0];
      if (ts instanceof Date && ts < cutoff) {
        deleted++;
      } else {
        keepRows.push(data[i]);
      }
    }

    if (deleted === 0) {
      return { success: true, deleted: 0, message: "ไม่มี Audit Log เก่ากว่า " + days + " วัน" };
    }

    // [BATCH] เขียนทับ sheet ทั้งหมด
    sh.clearContents();
    sh.getRange(1, 1, keepRows.length, header.length).setValues(keepRows);
    sh.getRange(1, 1, 1, header.length).setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);

    // ลบแถวส่วนเกิน (ลด cell count จริง)
    const totalRows = sh.getMaxRows();
    const keepCount = Math.max(keepRows.length, 1);
    if (totalRows > keepCount) {
      sh.deleteRows(keepCount + 1, totalRows - keepCount);
    }

    return {
      success: true,
      deleted: deleted,
      message: "ลบ Audit Log เก่า " + deleted + " rows (เก่ากว่า " + days + " วัน)"
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ลบ Export Sheets ที่ชื่อ "Export_YEAR_TIMESTAMP" เก่ากว่า keepDays วัน
 * @param {number} keepDays - จำนวนวันที่จะเก็บไว้ (default: 60)
 */
function cleanupExportSheets(keepDays) {
  try {
    const days      = keepDays || 60;
    const now       = Date.now();
    const cutoffMs  = days * 24 * 60 * 60 * 1000;

    const ss      = getOrCreateSpreadsheet();
    const sheets  = ss.getSheets();
    const deleted = [];

    sheets.forEach(sheet => {
      const name  = sheet.getName();
      const match = name.match(/^Export_\d+_(\d+)$/); // Export_YEAR_TIMESTAMP
      if (match) {
        const sheetTs = Number(match[1]);
        if (!isNaN(sheetTs) && (now - sheetTs) > cutoffMs) {
          ss.deleteSheet(sheet);
          deleted.push(name);
        }
      }
    });

    return {
      success: true,
      deleted: deleted.length,
      sheets:  deleted,
      message: `ลบ Export Sheets เก่า ${deleted.length} sheets (เก่ากว่า ${days} วัน)`
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Archive ข้อมูล tb_data ของปีที่ระบุ ไปยัง Spreadsheet แยกต่างหาก
 * ใช้สำหรับปีที่ปิดแล้วและต้องการเก็บข้อมูล แต่ไม่ต้องการให้กิน cell quota
 * @param {string|number} year - ปีที่ต้องการ Archive (เช่น "2566")
 */
function archiveYearData(year) {
  try {
    const adminEmail  = Session.getActiveUser().getEmail();
    const folder      = getScriptFolder();
    const archiveName = `QA_BUTC_Archive_${year}`;

    // ค้นหาหรือสร้าง Archive Spreadsheet
    let archiveSS;
    const existingFiles = folder.getFilesByName(archiveName);
    if (existingFiles.hasNext()) {
      archiveSS = SpreadsheetApp.open(existingFiles.next());
    } else {
      archiveSS = SpreadsheetApp.create(archiveName);
      const archiveFile = DriveApp.getFileById(archiveSS.getId());
      folder.addFile(archiveFile);
      try { DriveApp.getRootFolder().removeFile(archiveFile); } catch(e) {}
    }

    // เตรียม Sheet ใน Archive
    let archiveSheet = archiveSS.getSheetByName("tb_data");
    if (!archiveSheet) {
      archiveSheet = archiveSS.insertSheet("tb_data");
      archiveSheet.appendRow(["Timestamp","UserEmail","Year","SectionID","SectionName","LastUpdate","DataJSON","SubmitterName"]);
      archiveSheet.getRange(1, 1, 1, 8).setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
    }

    // หาข้อมูลปีที่ต้องการ Archive
    const mainSheet = ensureSheetExists("tb_data");
    const mainData  = mainSheet.getDataRange().getValues();
    const toArchive = [];
    const toDelete  = [];

    for (let i = 1; i < mainData.length; i++) {
      if (String(mainData[i][2]) === String(year)) {
        toArchive.push(mainData[i]);
        toDelete.push(i + 1);
      }
    }

    if (toArchive.length === 0) {
      return { success: false, message: `ไม่พบข้อมูลปี ${year} ใน tb_data` };
    }

    // เขียนไปยัง Archive
    toArchive.forEach(row => archiveSheet.appendRow(row));

    // ลบจาก Main Sheet (จากล่างขึ้นบน)
    for (let i = toDelete.length - 1; i >= 0; i--) {
      mainSheet.deleteRow(toDelete[i]);
    }

    writeAuditLog(adminEmail, "ARCHIVE_YEAR", String(year),
      `${toArchive.length} records → ${archiveName}`);
    return {
      success:    true,
      archived:   toArchive.length,
      archiveUrl: archiveSS.getUrl(),
      message:    `Archive ปี ${year} สำเร็จ: ${toArchive.length} records → ${archiveName}`
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * รัน Maintenance ทั้งหมดในครั้งเดียว
 * ถูกเรียกโดย Time-based Trigger รายเดือน (วันที่ 1 เวลา 02:00)
 * สามารถรันด้วยตนเองจาก Apps Script Editor ได้เช่นกัน
 */
function runMonthlyMaintenance() {
  const results = {};
  try { results.compressVisits = compressOldVisits();     } catch(e) { results.compressVisits = { error: e.toString() }; }
  try { results.pruneAudit     = pruneAuditLog(90);       } catch(e) { results.pruneAudit     = { error: e.toString() }; }
  try { results.cleanupExports = cleanupExportSheets(60); } catch(e) { results.cleanupExports = { error: e.toString() }; }

  // เคลียร์ Cache หลัง Maintenance เพื่อให้ข้อมูล fresh
  try { _clearAppCache(); } catch(e) {}

  const email = (() => { try { return Session.getActiveUser().getEmail() || "auto_trigger"; } catch(e) { return "auto_trigger"; } })();
  writeAuditLog(email, "MONTHLY_MAINTENANCE", "system", JSON.stringify({
    visits_deleted:  (results.compressVisits && results.compressVisits.deleted)  || 0,
    audit_deleted:   (results.pruneAudit     && results.pruneAudit.deleted)      || 0,
    exports_deleted: (results.cleanupExports && results.cleanupExports.deleted)  || 0
  }));

  Logger.log("=== Monthly Maintenance Results ===");
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}

/**
 * ตั้งค่า Time-based Trigger สำหรับ Monthly Maintenance
 * เรียกฟังก์ชันนี้ **ครั้งเดียว** จาก Apps Script Editor
 * (เลือก setupMaintenanceTriggers ใน Dropdown แล้วกด Run)
 */
function setupMaintenanceTriggers() {
  // ลบ Trigger runMonthlyMaintenance เก่าก่อน (ป้องกัน duplicate)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runMonthlyMaintenance") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // สร้าง Trigger ใหม่: ทุกวันที่ 1 ของเดือน เวลา 02:00–03:00
  ScriptApp.newTrigger("runMonthlyMaintenance")
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();

  Logger.log("✅ Maintenance Trigger ตั้งค่าแล้ว: runMonthlyMaintenance ทุกวันที่ 1 เวลา 02:00");
  return { success: true, message: "ตั้งค่า Trigger สำเร็จ — Maintenance จะรันอัตโนมัติทุกวันที่ 1 ของเดือน" };
}