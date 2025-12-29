/**
 * ========================================
 * Booking Control System - Main Backend
 * ========================================
 * Google Apps Script Backend for Adventure Tour Booking Management System
 *
 * @author Antigravity AI
 * @version 1.0
 * @date 2025-12-29
 */

// ========================================
// CONFIGURATION - Fixed Sheet ID
// ========================================

const CONFIG = {
  SPREADSHEET_ID: "YOUR_SPREADSHEET_ID_HERE", // ⚠️ ใส่ Spreadsheet ID ของคุณที่นี่
  DRIVE_FOLDER_ID: "YOUR_DRIVE_FOLDER_ID_HERE", // ⚠️ ใส่ Drive Folder ID สำหรับเก็บสลิป

  // Sheet Names
  SHEETS: {
    USERS: "Users",
    BOOKING_RAW: "Booking_Raw",
    LOCATIONS: "Locations",
    PROGRAMS: "Programs",
    BOOKING_STATUS_HISTORY: "Booking_Status_History",
  },

  // User Roles
  ROLES: {
    SALES: "Sales",
    OP: "OP",
    ADMIN: "Admin",
    AR_AP: "AR_AP",
    COST: "Cost",
    OWNER: "Owner",
  },

  // Booking Status
  STATUS: {
    PENDING: "Pending",
    CONFIRM: "Confirm",
    COMPLETE: "Complete",
    CANCEL: "Cancel",
  },

  // Default Password
  DEFAULT_PASSWORD: "password123",
};

// ========================================
// UTILITY FUNCTIONS
// ========================================

/**
 * Get Spreadsheet by ID
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Get Sheet by Name
 */
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}

/**
 * Hash Password using SHA-256
 */
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return rawHash
    .map((byte) => {
      const v = byte < 0 ? 256 + byte : byte;
      return ("0" + v.toString(16)).slice(-2);
    })
    .join("");
}

/**
 * Generate Unique ID
 */
function generateUniqueId(prefix = "") {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000);
  return `${prefix}${timestamp}${random}`;
}

/**
 * Get Current Timestamp
 */
function getCurrentTimestamp() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
}

/**
 * Format Date to String
 */
function formatDate(date, format = "yyyy-MM-dd") {
  if (!date) return "";
  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    format
  );
}

// ========================================
// SESSION MANAGEMENT
// ========================================

/**
 * Set User Session
 */
function setSession(userId, username, role) {
  const userProperties = PropertiesService.getUserProperties();
  const sessionData = {
    userId: userId,
    username: username,
    role: role,
    loginTime: new Date().getTime(),
  };
  userProperties.setProperty("session", JSON.stringify(sessionData));
}

/**
 * Get User Session
 */
function getSession() {
  const userProperties = PropertiesService.getUserProperties();
  const sessionString = userProperties.getProperty("session");
  if (!sessionString) return null;
  return JSON.parse(sessionString);
}

/**
 * Clear User Session
 */
function clearSession() {
  PropertiesService.getUserProperties().deleteProperty("session");
}

/**
 * Check if User is Logged In
 */
function isLoggedIn() {
  return getSession() !== null;
}

/**
 * Check User Role
 */
function hasRole(requiredRole) {
  const session = getSession();
  if (!session) return false;

  // Owner has access to everything
  if (session.role === CONFIG.ROLES.OWNER) return true;

  // Check specific role
  if (Array.isArray(requiredRole)) {
    return requiredRole.includes(session.role);
  }
  return session.role === requiredRole;
}

// ========================================
// AUTHENTICATION FUNCTIONS
// ========================================

/**
 * Login User
 */
function loginUser(username, password) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dbUsername = row[2]; // Column C: ชื่อผู้ใช้
      const dbPassword = row[3]; // Column D: รหัสผ่าน
      const fullName = row[4]; // Column E: ชื่อ-นามสกุล
      const role = row[5]; // Column F: บทบาท
      const status = row[6]; // Column G: สถานะการใช้งาน

      if (dbUsername === username && status === "Active") {
        const hashedPassword = hashPassword(password);
        if (dbPassword === hashedPassword) {
          const userId = row[0]; // Column A: รหัสผู้ใช้
          setSession(userId, username, role);
          return {
            success: true,
            message: "เข้าสู่ระบบสำเร็จ",
            data: {
              userId: userId,
              username: username,
              fullName: fullName,
              role: role,
            },
          };
        }
      }
    }

    return {
      success: false,
      message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง",
    };
  } catch (error) {
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Logout User
 */
function logoutUser() {
  clearSession();
  return {
    success: true,
    message: "ออกจากระบบสำเร็จ",
  };
}

/**
 * Get Current User Info
 */
function getCurrentUser() {
  const session = getSession();
  if (!session) {
    return {
      success: false,
      message: "ไม่พบข้อมูลผู้ใช้",
    };
  }

  return {
    success: true,
    data: session,
  };
}

/**
 * Setup Initial Owner User (Run this once)
 * ฟังก์ชันนี้ใช้สำหรับสร้าง User Owner คนแรกในระบบ
 * รันครั้งเดียวเมื่อเริ่มต้นใช้งานระบบ
 */
function setupInitialOwner() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);

    // Check if sheet is empty (only header or no data)
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      Logger.log("มี User อยู่ในระบบแล้ว");
      return {
        success: false,
        message: "มี User อยู่ในระบบแล้ว ไม่สามารถสร้าง Owner ใหม่ได้",
      };
    }

    // Create header if not exists
    if (data.length === 0) {
      const headers = [
        "รหัสผู้ใช้",
        "อีเมล",
        "ชื่อผู้ใช้",
        "รหัสผ่าน",
        "ชื่อ-นามสกุล",
        "บทบาท",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ];
      sheet.appendRow(headers);
    }

    // Create Owner user
    const userId = "USR001";
    const hashedPassword = hashPassword("password123");
    const timestamp = getCurrentTimestamp();

    const ownerRow = [
      userId,
      "owner@example.com",
      "admin",
      hashedPassword,
      "ผู้ดูแลระบบ",
      "Owner",
      "Active",
      timestamp,
      timestamp,
    ];

    sheet.appendRow(ownerRow);

    Logger.log("สร้าง Owner สำเร็จ!");
    Logger.log("Username: admin");
    Logger.log("Password: password123");

    return {
      success: true,
      message: "สร้าง Owner สำเร็จ!\nUsername: admin\nPassword: password123",
      data: {
        username: "admin",
        password: "password123",
      },
    };
  } catch (error) {
    Logger.log("Error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Setup All Sheets with Headers
 * สร้าง Headers สำหรับทุก Sheet
 */
function setupAllSheets() {
  try {
    const ss = getSpreadsheet();

    // Setup Users Sheet
    let usersSheet = ss.getSheetByName(CONFIG.SHEETS.USERS);
    if (!usersSheet) {
      usersSheet = ss.insertSheet(CONFIG.SHEETS.USERS);
    }
    if (usersSheet.getLastRow() === 0) {
      usersSheet.appendRow([
        "รหัสผู้ใช้",
        "อีเมล",
        "ชื่อผู้ใช้",
        "รหัสผ่าน",
        "ชื่อ-นามสกุล",
        "บทบาท",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Booking_Raw Sheet
    let bookingSheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_RAW);
    if (!bookingSheet) {
      bookingSheet = ss.insertSheet(CONFIG.SHEETS.BOOKING_RAW);
    }
    if (bookingSheet.getLastRow() === 0) {
      bookingSheet.appendRow([
        "รหัสการจอง",
        "วันที่จอง",
        "วันที่เดินทาง",
        "ชื่อสถานที่",
        "โปรแกรม",
        "ผู้ใหญ่ (คน)",
        "เด็ก (คน)",
        "ราคาผู้ใหญ่",
        "ราคาเด็ก",
        "ส่วนลด (บาท)",
        "สถานะ",
        "URL สลิปการชำระเงิน",
        "Agent",
        "หมายเหตุ",
        "ยอดขายต่อรายการ",
        "ผู้สร้าง",
        "วันที่สร้าง",
        "ผู้แก้ไขล่าสุด",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Locations Sheet
    let locationsSheet = ss.getSheetByName(CONFIG.SHEETS.LOCATIONS);
    if (!locationsSheet) {
      locationsSheet = ss.insertSheet(CONFIG.SHEETS.LOCATIONS);
    }
    if (locationsSheet.getLastRow() === 0) {
      locationsSheet.appendRow([
        "รหัสสถานที่",
        "ชื่อสถานที่",
        "ชื่อเซลล์",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Programs Sheet
    let programsSheet = ss.getSheetByName(CONFIG.SHEETS.PROGRAMS);
    if (!programsSheet) {
      programsSheet = ss.insertSheet(CONFIG.SHEETS.PROGRAMS);
    }
    if (programsSheet.getLastRow() === 0) {
      programsSheet.appendRow([
        "รหัสโปรแกรม",
        "ชื่อโปรแกรม",
        "รายละเอียด",
        "ราคาผู้ใหญ่",
        "ราคาเด็ก",
        "สถานะการใช้งาน",
        "วันที่สร้าง",
        "วันที่แก้ไขล่าสุด",
      ]);
    }

    // Setup Booking_Status_History Sheet
    let historySheet = ss.getSheetByName(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    if (!historySheet) {
      historySheet = ss.insertSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    }
    if (historySheet.getLastRow() === 0) {
      historySheet.appendRow([
        "รหัสประวัติ",
        "รหัสการจอง",
        "สถานะเดิม",
        "สถานะใหม่",
        "ผู้เปลี่ยนสถานะ",
        "วันที่เปลี่ยนสถานะ",
        "เหตุผล",
        "URL เอกสาร",
      ]);
    }

    Logger.log("Setup all sheets สำเร็จ!");
    return {
      success: true,
      message: "สร้าง Sheets และ Headers สำเร็จทั้งหมด",
    };
  } catch (error) {
    Logger.log("Error: " + error.message);
    return {
      success: false,
      message: "เกิดข้อผิดพลาด: " + error.message,
    };
  }
}

/**
 * Debug Login - ตรวจสอบปัญหาการ Login
 * รันฟังก์ชันนี้เพื่อดูข้อมูล User ในระบบ
 */
function debugLogin() {
  try {
    Logger.log("=== เริ่มตรวจสอบระบบ ===");

    // 1. ตรวจสอบ Spreadsheet ID
    Logger.log("1. Spreadsheet ID: " + CONFIG.SPREADSHEET_ID);
    if (CONFIG.SPREADSHEET_ID === "YOUR_SPREADSHEET_ID_HERE") {
      Logger.log("❌ ERROR: คุณยังไม่ได้ใส่ SPREADSHEET_ID ใน Code.gs");
      return {
        success: false,
        message: "กรุณาใส่ SPREADSHEET_ID ใน Code.gs",
      };
    }

    // 2. ตรวจสอบ Sheet Users
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    Logger.log("2. Sheet Users: พบแล้ว");

    // 3. ตรวจสอบข้อมูล User
    const data = sheet.getDataRange().getValues();
    Logger.log("3. จำนวนแถวทั้งหมด: " + data.length);

    if (data.length === 0) {
      Logger.log("❌ ERROR: Sheet Users ว่างเปล่า");
      Logger.log("แก้ไข: รันฟังก์ชัน setupAllSheets() และ setupInitialOwner()");
      return {
        success: false,
        message: "Sheet Users ว่างเปล่า กรุณารัน setupInitialOwner()",
      };
    }

    if (data.length === 1) {
      Logger.log("❌ ERROR: มีแค่ Header ไม่มี User");
      Logger.log("แก้ไข: รันฟังก์ชัน setupInitialOwner()");
      return {
        success: false,
        message: "ไม่มี User ในระบบ กรุณารัน setupInitialOwner()",
      };
    }

    // 4. แสดงข้อมูล User ทั้งหมด
    Logger.log("4. User ในระบบ:");
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      Logger.log(`   - User ${i}:`);
      Logger.log(`     รหัส: ${row[0]}`);
      Logger.log(`     Username: ${row[2]}`);
      Logger.log(`     Role: ${row[5]}`);
      Logger.log(`     Status: ${row[6]}`);
      Logger.log(`     Password Hash: ${row[3].substring(0, 20)}...`);
    }

    // 5. ทดสอบ Login
    Logger.log("5. ทดสอบ Login...");
    const testResult = loginUser("admin", "password123");
    Logger.log("   ผลลัพธ์: " + JSON.stringify(testResult));

    if (testResult.success) {
      Logger.log("✅ SUCCESS: Login สำเร็จ!");
    } else {
      Logger.log("❌ ERROR: Login ไม่สำเร็จ - " + testResult.message);
    }

    Logger.log("=== สิ้นสุดการตรวจสอบ ===");

    return {
      success: true,
      message: "ตรวจสอบเสร็จสิ้น ดู Execution log สำหรับรายละเอียด",
      data: {
        totalUsers: data.length - 1,
        loginTest: testResult,
      },
    };
  } catch (error) {
    Logger.log("❌ CRITICAL ERROR: " + error.message);
    Logger.log("Stack: " + error.stack);
    return {
      success: false,
      message: "เกิดข้อผิดพลาดร้ายแรง: " + error.message,
    };
  }
}

// ========================================
// USER MANAGEMENT (Owner Only)
// ========================================

/**
 * Get All Users
 */
function getAllUsers() {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();
    const users = [];

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      users.push({
        userId: row[0],
        email: row[1],
        username: row[2],
        fullName: row[4],
        role: row[5],
        status: row[6],
        createdAt: row[7],
        updatedAt: row[8],
      });
    }

    return { success: true, data: users };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Create New User
 */
function createUser(userData) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);

    // Check if username or email already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userData.email || data[i][2] === userData.username) {
        return { success: false, message: "อีเมลหรือชื่อผู้ใช้นี้มีอยู่แล้ว" };
      }
    }

    const userId = generateUniqueId("USR");
    const hashedPassword = hashPassword(CONFIG.DEFAULT_PASSWORD);
    const timestamp = getCurrentTimestamp();

    const newRow = [
      userId, // A: รหัสผู้ใช้
      userData.email, // B: อีเมล
      userData.username, // C: ชื่อผู้ใช้
      hashedPassword, // D: รหัสผ่าน
      userData.fullName, // E: ชื่อ-นามสกุล
      userData.role, // F: บทบาท
      "Active", // G: สถานะการใช้งาน
      timestamp, // H: วันที่สร้าง
      timestamp, // I: วันที่แก้ไขล่าสุด
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มพนักงานสำเร็จ",
      data: { userId: userId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Update User
 */
function updateUser(userId, userData) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const timestamp = getCurrentTimestamp();

        // Update only provided fields
        if (userData.email) sheet.getRange(i + 1, 2).setValue(userData.email);
        if (userData.username)
          sheet.getRange(i + 1, 3).setValue(userData.username);
        if (userData.fullName)
          sheet.getRange(i + 1, 5).setValue(userData.fullName);
        if (userData.role) sheet.getRange(i + 1, 6).setValue(userData.role);
        if (userData.status) sheet.getRange(i + 1, 7).setValue(userData.status);

        // Update timestamp
        sheet.getRange(i + 1, 9).setValue(timestamp);

        return { success: true, message: "อัพเดทข้อมูลสำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Reset User Password
 */
function resetUserPassword(userId) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        const hashedPassword = hashPassword(CONFIG.DEFAULT_PASSWORD);
        const timestamp = getCurrentTimestamp();

        sheet.getRange(i + 1, 4).setValue(hashedPassword);
        sheet.getRange(i + 1, 9).setValue(timestamp);

        return {
          success: true,
          message:
            "รีเซ็ตรหัสผ่านสำเร็จ รหัสผ่านใหม่: " + CONFIG.DEFAULT_PASSWORD,
        };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Delete User (Soft Delete)
 */
function deleteUser(userId) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        // Check if user is Owner
        if (data[i][5] === CONFIG.ROLES.OWNER) {
          return { success: false, message: "ไม่สามารถลบ Owner ได้" };
        }

        const timestamp = getCurrentTimestamp();
        sheet.getRange(i + 1, 7).setValue("Inactive");
        sheet.getRange(i + 1, 9).setValue(timestamp);

        return { success: true, message: "ลบพนักงานสำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบผู้ใช้" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// BOOKING MANAGEMENT
// ========================================

/**
 * Get All Bookings
 */
function getAllBookings(filters = {}) {
  if (!hasRole([CONFIG.ROLES.OP, CONFIG.ROLES.OWNER, CONFIG.ROLES.AR_AP])) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();
    const bookings = [];
    const session = getSession();

    // Skip header row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      bookings.push({
        rowIndex: i + 1,
        bookingId: row[0],
        bookingDate: row[1],
        travelDate: row[2],
        location: row[3],
        program: row[4],
        adults: row[5],
        children: row[6],
        adultPrice: row[7],
        childPrice: row[8],
        discount: row[9],
        status: row[10],
        slipUrl: row[11],
        agent: row[12],
        note: row[13],
        totalAmount: row[14],
        createdBy: row[15],
        createdAt: row[16],
        updatedBy: row[17],
        updatedAt: row[18],
        canEdit:
          session.role === CONFIG.ROLES.OWNER || row[15] === session.username,
      });
    }

    return { success: true, data: bookings };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Create New Booking
 */
function createBooking(bookingData) {
  if (!hasRole([CONFIG.ROLES.OP, CONFIG.ROLES.OWNER])) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const session = getSession();
    const bookingId = generateUniqueId("BK");
    const timestamp = getCurrentTimestamp();

    // Calculate total amount
    const totalAmount =
      bookingData.adults * bookingData.adultPrice +
      bookingData.children * bookingData.childPrice -
      bookingData.discount;

    const newRow = [
      bookingId, // A: รหัสการจอง
      bookingData.bookingDate, // B: วันที่จอง
      bookingData.travelDate, // C: วันที่เดินทาง
      bookingData.location, // D: ชื่อสถานที่
      bookingData.program, // E: โปรแกรม
      bookingData.adults, // F: ผู้ใหญ่ (คน)
      bookingData.children, // G: เด็ก (คน)
      bookingData.adultPrice, // H: ราคาผู้ใหญ่
      bookingData.childPrice, // I: ราคาเด็ก
      bookingData.discount || 0, // J: ส่วนลด (บาท)
      CONFIG.STATUS.PENDING, // K: สถานะ
      bookingData.slipUrl || "", // L: URL สลิปการชำระเงิน
      bookingData.agent || "", // M: Agent
      bookingData.note || "", // N: หมายเหตุ
      totalAmount, // O: ยอดขายต่อรายการ
      session.username, // P: ผู้สร้าง
      timestamp, // Q: วันที่สร้าง
      session.username, // R: ผู้แก้ไขล่าสุด
      timestamp, // S: วันที่แก้ไขล่าสุด
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "สร้างการจองสำเร็จ",
      data: { bookingId: bookingId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Update Booking
 */
function updateBooking(bookingId, bookingData) {
  if (!hasRole([CONFIG.ROLES.OP, CONFIG.ROLES.OWNER])) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();
    const session = getSession();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        // Check permission: OP can only edit their own bookings
        if (
          session.role === CONFIG.ROLES.OP &&
          data[i][15] !== session.username
        ) {
          return { success: false, message: "คุณไม่มีสิทธิ์แก้ไขรายการนี้" };
        }

        const timestamp = getCurrentTimestamp();
        const rowNum = i + 1;

        // Update fields
        if (bookingData.bookingDate)
          sheet.getRange(rowNum, 2).setValue(bookingData.bookingDate);
        if (bookingData.travelDate)
          sheet.getRange(rowNum, 3).setValue(bookingData.travelDate);
        if (bookingData.location)
          sheet.getRange(rowNum, 4).setValue(bookingData.location);
        if (bookingData.program)
          sheet.getRange(rowNum, 5).setValue(bookingData.program);
        if (bookingData.adults !== undefined)
          sheet.getRange(rowNum, 6).setValue(bookingData.adults);
        if (bookingData.children !== undefined)
          sheet.getRange(rowNum, 7).setValue(bookingData.children);
        if (bookingData.adultPrice !== undefined)
          sheet.getRange(rowNum, 8).setValue(bookingData.adultPrice);
        if (bookingData.childPrice !== undefined)
          sheet.getRange(rowNum, 9).setValue(bookingData.childPrice);
        if (bookingData.discount !== undefined)
          sheet.getRange(rowNum, 10).setValue(bookingData.discount);
        if (bookingData.slipUrl)
          sheet.getRange(rowNum, 12).setValue(bookingData.slipUrl);
        if (bookingData.agent)
          sheet.getRange(rowNum, 13).setValue(bookingData.agent);
        if (bookingData.note)
          sheet.getRange(rowNum, 14).setValue(bookingData.note);

        // Recalculate total amount
        const adults =
          bookingData.adults !== undefined ? bookingData.adults : data[i][5];
        const children =
          bookingData.children !== undefined
            ? bookingData.children
            : data[i][6];
        const adultPrice =
          bookingData.adultPrice !== undefined
            ? bookingData.adultPrice
            : data[i][7];
        const childPrice =
          bookingData.childPrice !== undefined
            ? bookingData.childPrice
            : data[i][8];
        const discount =
          bookingData.discount !== undefined
            ? bookingData.discount
            : data[i][9];

        const totalAmount =
          adults * adultPrice + children * childPrice - discount;
        sheet.getRange(rowNum, 15).setValue(totalAmount);

        // Update metadata
        sheet.getRange(rowNum, 18).setValue(session.username);
        sheet.getRange(rowNum, 19).setValue(timestamp);

        return { success: true, message: "อัพเดทการจองสำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบการจอง" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Update Booking Status (OP and AR/AP)
 */
function updateBookingStatus(bookingId, newStatus, reason = "") {
  if (!hasRole([CONFIG.ROLES.OP, CONFIG.ROLES.AR_AP, CONFIG.ROLES.OWNER])) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();
    const session = getSession();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === bookingId) {
        const oldStatus = data[i][10];
        const timestamp = getCurrentTimestamp();

        // Update status
        sheet.getRange(i + 1, 11).setValue(newStatus);
        sheet.getRange(i + 1, 18).setValue(session.username);
        sheet.getRange(i + 1, 19).setValue(timestamp);

        // Log status change to history
        logStatusChange(bookingId, oldStatus, newStatus, reason);

        return { success: true, message: "อัพเดทสถานะสำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบการจอง" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Log Status Change to History
 */
function logStatusChange(
  bookingId,
  oldStatus,
  newStatus,
  reason = "",
  documentUrl = ""
) {
  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_STATUS_HISTORY);
    const session = getSession();
    const historyId = generateUniqueId("HST");
    const timestamp = getCurrentTimestamp();

    const newRow = [
      historyId, // A: รหัสประวัติ
      bookingId, // B: รหัสการจอง
      oldStatus, // C: สถานะเดิม
      newStatus, // D: สถานะใหม่
      session.username, // E: ผู้เปลี่ยนสถานะ
      timestamp, // F: วันที่เปลี่ยนสถานะ
      reason, // G: เหตุผล
      documentUrl, // H: URL เอกสาร
    ];

    sheet.appendRow(newRow);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Upload Slip to Google Drive
 */
function uploadSlip(bookingId, fileBlob, fileName) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const file = folder.createFile(fileBlob);
    file.setName(`${bookingId}_${fileName}`);

    // Make file accessible
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = file.getUrl();

    // Update booking with slip URL
    updateBooking(bookingId, { slipUrl: fileUrl });

    return {
      success: true,
      message: "อัพโหลดสลิปสำเร็จ",
      data: { url: fileUrl },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// LOCATIONS MANAGEMENT
// ========================================

/**
 * Get All Locations
 */
function getAllLocations() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const data = sheet.getDataRange().getValues();
    const locations = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      locations.push({
        locationId: row[0],
        locationName: row[1],
        cellName: row[2],
        createdAt: row[3],
        updatedAt: row[4],
      });
    }

    return { success: true, data: locations };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Create Location (Admin/Owner)
 */
function createLocation(locationData) {
  if (!hasRole([CONFIG.ROLES.ADMIN, CONFIG.ROLES.OWNER])) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.LOCATIONS);
    const locationId = generateUniqueId("LOC");
    const timestamp = getCurrentTimestamp();

    const newRow = [
      locationId,
      locationData.locationName,
      locationData.cellName || "",
      timestamp,
      timestamp,
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มสถานที่สำเร็จ",
      data: { locationId: locationId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// PROGRAMS MANAGEMENT
// ========================================

/**
 * Get All Programs
 */
function getAllPrograms() {
  try {
    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    const data = sheet.getDataRange().getValues();
    const programs = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      programs.push({
        programId: row[0],
        programName: row[1],
        description: row[2],
        adultPrice: row[3],
        childPrice: row[4],
        isActive: row[5],
        createdAt: row[6],
        updatedAt: row[7],
      });
    }

    return { success: true, data: programs };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Create Program (Owner Only)
 */
function createProgram(programData) {
  if (!hasRole(CONFIG.ROLES.OWNER)) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.PROGRAMS);
    const programId = generateUniqueId("PRG");
    const timestamp = getCurrentTimestamp();

    const newRow = [
      programId,
      programData.programName,
      programData.description || "",
      programData.adultPrice,
      programData.childPrice,
      true, // isActive
      timestamp,
      timestamp,
    ];

    sheet.appendRow(newRow);

    return {
      success: true,
      message: "เพิ่มโปรแกรมสำเร็จ",
      data: { programId: programId },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// DASHBOARD & REPORTS
// ========================================

/**
 * Get Dashboard Data
 */
function getDashboardData() {
  if (!hasRole([CONFIG.ROLES.COST, CONFIG.ROLES.OWNER])) {
    return { success: false, message: "ไม่มีสิทธิ์เข้าถึง" };
  }

  try {
    const sheet = getSheet(CONFIG.SHEETS.BOOKING_RAW);
    const data = sheet.getDataRange().getValues();

    const today = new Date();
    const thisMonth = today.getMonth();
    const thisYear = today.getFullYear();

    let salesToday = 0;
    let salesThisMonth = 0;
    let pendingAmount = 0;
    let totalBookings = 0;
    let cancelledBookings = 0;

    const programStats = {};
    const agentStats = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const bookingDate = new Date(row[1]);
      const status = row[10];
      const totalAmount = row[14];
      const program = row[4];
      const agent = row[12];

      totalBookings++;

      // Count cancelled bookings
      if (status === CONFIG.STATUS.CANCEL) {
        cancelledBookings++;
      }

      // Sales today (Complete only)
      if (
        status === CONFIG.STATUS.COMPLETE &&
        bookingDate.toDateString() === today.toDateString()
      ) {
        salesToday += totalAmount;
      }

      // Sales this month (Complete only)
      if (
        status === CONFIG.STATUS.COMPLETE &&
        bookingDate.getMonth() === thisMonth &&
        bookingDate.getFullYear() === thisYear
      ) {
        salesThisMonth += totalAmount;
      }

      // Pending amount (Confirm status)
      if (status === CONFIG.STATUS.CONFIRM) {
        pendingAmount += totalAmount;
      }

      // Program stats
      if (status === CONFIG.STATUS.COMPLETE) {
        if (!programStats[program]) {
          programStats[program] = { count: 0, amount: 0 };
        }
        programStats[program].count++;
        programStats[program].amount += totalAmount;
      }

      // Agent stats
      if (status === CONFIG.STATUS.COMPLETE && agent) {
        if (!agentStats[agent]) {
          agentStats[agent] = 0;
        }
        agentStats[agent] += totalAmount;
      }
    }

    // Calculate cancel rate
    const cancelRate =
      totalBookings > 0
        ? ((cancelledBookings / totalBookings) * 100).toFixed(2)
        : 0;

    // Get top 5 programs
    const topPrograms = Object.entries(programStats)
      .sort((a, b) => b[1].amount - a[1].amount)
      .slice(0, 5)
      .map(([name, stats]) => ({ name, ...stats }));

    return {
      success: true,
      data: {
        salesToday: salesToday,
        salesThisMonth: salesThisMonth,
        pendingAmount: pendingAmount,
        cancelRate: cancelRate,
        topPrograms: topPrograms,
        salesByAgent: agentStats,
      },
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ========================================
// HTML TEMPLATE FUNCTIONS
// ========================================

/**
 * Include HTML file (for templating)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========================================
// WEB APP ENTRY POINT
// ========================================

/**
 * Serve HTML Pages
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Booking Control System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
