const SHEET_ID = '1VUjJRh-x6e89EYDMdWgWt49HhavsINnIo9Ld948cVIk'; // ใส่ ID Google Sheet ของคุณ
const TARGET_F1_SS_ID = "1VUjJRh-x6e89EYDMdWgWt49HhavsINnIo9Ld948cVIk";
const FOLDER_ID_SLIPS = '1JYx6uaANy8_0NCzC3qQbfhWRGmO0TRHK'; // *** ใส่ ID โฟลเดอร์เก็บรูปสลิปที่นี่ ***
const FOLDER_ID = '1JYx6uaANy8_0NCzC3qQbfhWRGmO0TRHK'; 
const QR_FOLDER_ID = '1JYx6uaANy8_0NCzC3qQbfhWRGmO0TRHK'; // โฟลเดอร์สำหรับเก็บรูป QR Code โดยเฉพาะ
// *** เพิ่มบรรทัดนี้เข้าไปครับ เพื่อแก้ปัญหา SPREADSHEET_ID is not defined ***
const SPREADSHEET_ID = SHEET_ID; 
const USERS_SHEET_NAME = 'Users';
const LOG_SHEET_NAME = 'UserLog';
const EmployeeSheetName = 'Employee';
const SENIOR_CHECKIN_SHEET_NAME = 'SeniorCheckIn';
const SENIOR_CHECKIN_FOLDER_ID = '1JYx6uaANy8_0NCzC3qQbfhWRGmO0TRHK'; 
const SHEET_USERS = USERS_SHEET_NAME; // ใช้ค่าเดียวกับ USERS_SHEET_NAME ที่มีอยู่แล้ว
const LOGIN_EXPIRATION_SECONDS = 8 * 60 * 60; // 8 ชั่วโมง

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  const page = (e && e.parameter.page) ? e.parameter.page : null; // Get page parameter

  // [แก้ไข v4] Check login status from CacheService
  const cache = CacheService.getUserCache();
  const user = cache.get('loggedInUser');

  if (user) {
    // --- User is Logged In ---

    // Handle Logout request
    if (page === 'logout') {
      return doLogout(); // This function will return the login page HTML
    }

    // --- Serve the main application layout or specific content pages ---
    let fileName;
    let title;
    let targetPage = page || 'Index'; // Default to 'Dashboard' if no specific page is requested

    // Determine which HTML file to load into the iframe or as the main page
    switch (targetPage) {
      case 'Dashboard_02':
        fileName = 'Dashboard_02';
        title = 'Dashboard_02';
        break;  
      case 'ImportF1':
        fileName = 'ImportF1';
        title = 'ImportF1';
        break;
      case 'High':
        fileName = 'High';
        title = 'High';
        break;         
      case 'Low':
        fileName = 'Low';
        title = 'Low';
        break; 
       case 'Work':
        fileName = 'Work';
        title = 'Work';
        break;  
      case 'Checkinsenior':
        fileName = 'Checkinsenior';
        title = 'ลงเวลา Senior';
        break;                               
      case 'UserProfile': // <-- เพิ่ม case นี้
        fileName = 'UserProfile';
        title = 'แก้ไขข้อมูลส่วนตัว';
        break;
      case 'AssignedTasks': // <-- เพิ่ม case นี้
        fileName = 'AssignedTasks';
        title = 'ชีตทำงาน';
        break;        
      case 'UserManagement': 
        fileName = 'UserManagement';
        title = 'จัดการผู้ใช้งาน';
        break;
      // --- [สิ้นสุดการเพิ่มใหม่] ---      
      default:
        fileName = 'MainLayout'; // Load the main layout file
        title = 'STORE SERVICE SYSTEM'; // Title for the main layout
        // No need to pass targetPage here, MainLayout handles default content
        break;
        // *** END MODIFICATION ***
    }

    // Serve MainLayout or specific content pages requested directly via URL
    const template = HtmlService.createTemplateFromFile(fileName);
    template.scriptUrl = getScriptUrl();
    template.userEmail = user; // Pass the logged-in username

    return template.evaluate()
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } else {
    // --- User is Not Logged In ---
    // Always show the Login page (handles registration/forgot password too)
    const loginTemplate = HtmlService.createTemplateFromFile('Login');
    loginTemplate.scriptUrl = getScriptUrl();
    return loginTemplate.evaluate()
      .setTitle('Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * [แก้ไข v4] เปลี่ยนไปใช้ CacheService
 */
 function doLogout() {
  try {
    // [แก้ไข v4]
    const cache = CacheService.getUserCache();
    cache.remove('loggedInUser');
    Logger.log("User logged out successfully (Cache removed).");

    // Prepare login page HTML to return
    const loginTemplate = HtmlService.createTemplateFromFile('Login');
    loginTemplate.scriptUrl = getScriptUrl();
    const loginHtml = loginTemplate.evaluate()
      .setTitle('Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    // If called directly from doGet (e.g., accessing ?page=logout)
    // Note: The global 'e' variable might not be accessible here directly if doLogout
    // is called from another context. Check if 'e' is defined before using it.
    if (typeof e !== 'undefined' && e && e.parameter && e.parameter.page === 'logout') {
        return loginHtml;
    }


    // If called from client-side JS (e.g., logout button)
    return {
      success: true,
      message: 'Logged out successfully.',
      redirectUrl: getScriptUrl() // URL to redirect to (the Login page)
    };

  } catch (error) {
    Logger.log(`Error during logout: ${error.message}`);
    // Handle error case for client-side call
     return {
       success: false,
       message: `Logout failed: ${error.message}`,
       redirectUrl: getScriptUrl() // Still provide login URL
     };
     // If doGet call failed, we might want to return a generic error page or login page
     // Returning login page is safer. Consider adding more robust error handling if needed.
    // const loginTemplate = HtmlService.createTemplateFromFile('Login');
    // loginTemplate.scriptUrl = getScriptUrl();
    // return loginTemplate.evaluate()... (as above)
  }
}

/**
 * [ฟังก์ชันใหม่]
 * เข้ารหัสผ่านแบบทางเดียว (One-way Hash) เพื่อความปลอดภัย
 * @param {string} password รหัสผ่านที่ยังไม่เข้ารหัส
 * @returns {string} รหัสผ่านที่เข้ารหัสแล้ว (Hashed)
 */
function hashPassword_(password) {
  const SPREADSHEET_ID = '1yGD48iiUlqk7xo3Xl6zjRhO9eHGJSKP5oL1Qu_uY32I'
  // ใช้ ID ของ Spreadsheet เป็น "Salt" เพื่อให้การ hash นี้ไม่เหมือนกับที่อื่น
  const salt = SPREADSHEET_ID || "default_salt"; 
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return Utilities.base64Encode(bytes);
}

/**
 * [ฟังก์ชันใหม่] [แก้ไข v4]
 * ตรวจสอบข้อมูลการ Login และใช้ CacheService
 * @param {object} credentials - {username, password}
 * @returns {object} {success: boolean, message: string, redirectUrl?: string}
 */
function doLogin(credentials) {
  try {
    const { username, password } = credentials;
    if (!username || !password) {
      return { success: false, message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน' };
    }

    const sheet = getSheetByName_(USERS_SHEET_NAME);
    if (sheet.getLastRow() < 2) {
      return { success: false, message: 'ไม่พบผู้ใช้ในระบบ' };
    }
    
    const headers = getSheetHeaders_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    
    const h = {
      username: headers.indexOf('Username'),
      passwordHash: headers.indexOf('PasswordHash')
    };
    
    const passwordHash = hashPassword_(password);

    for (const row of data) {
      if (row[h.username].toLowerCase() === username.toLowerCase()) {
        if (row[h.passwordHash] === passwordHash) {
          // --- Login สำเร็จ ---
          // [แก้ไข v4] เก็บชื่อผู้ใช้ไว้ใน Cache (มีหมดอายุ)
          const cache = CacheService.getUserCache();
          cache.put('loggedInUser', row[h.username], LOGIN_EXPIRATION_SECONDS); // 8 ชั่วโมง
          
          return { 
            success: true, 
            message: 'Login successful',
            redirectUrl: getScriptUrl() // ส่ง URL สำหรับ redirect กลับไป
          };
        } else {
          return { success: false, message: 'The password is incorrect.' };
        }
      }
    }

    return { success: false, message: 'This username was not found in the system.' };
  } catch (e) {
    Logger.log(`Error in doLogin: ${e.stack}`);
    return { success: false, message: `An error occurred.: ${e.message}` };
  }
}

/**
 * [ฟังก์ชันใหม่]
 * ลงทะเบียนผู้ใช้ใหม่
 * @param {object} userData - ข้อมูลจากฟอร์มลงทะเบียน
 * @returns {object} {success: boolean, message: string}
 */
function doRegister(userData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const { employeeName, employeeId, phone, gmail, username, password } = userData;
    
    // ตรวจสอบข้อมูลซ้ำ
    const sheet = getSheetByName_(USERS_SHEET_NAME);
    const headers = getSheetHeaders_(sheet);
    const h = {
      employeeId: headers.indexOf('EmployeeID'),
      phone: headers.indexOf('Phone'),
      username: headers.indexOf('Username')
    };

    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
      for (const row of data) {
        if (row[h.username].toLowerCase() === username.toLowerCase()) {
          return { success: false, message: 'ชื่อผู้ใช้นี้ (Username) ถูกใช้ไปแล้ว' };
        }
        if (row[h.employeeId] === employeeId) {
          return { success: false, message: 'รหัสพนักงานนี้ถูกใช้ไปแล้ว' };
        }
        if (row[h.phone] === phone) {
          return { success: false, message: 'เบอร์โทรศัพท์นี้ถูกใช้ไปแล้ว' };
        }
      }
    }
    
    // เข้ารหัสผ่าน
    const passwordHash = hashPassword_(password);
    
    // บันทึกข้อมูลใหม่
    const newRow = [
      employeeName,
      employeeId,
      phone,
      gmail,
      username,
      passwordHash
    ];
    
    // ตรวจสอบว่า newRow มีจำนวนคอลัมน์ตรงกับ headers หรือไม่
    const finalRow = headers.map(header => {
        const headerMap = {
            'EmployeeName': employeeName,
            'EmployeeID': employeeId,
            'Phone': phone,
            'CompanyGmail': gmail,
            'Username': username,
            'PasswordHash': passwordHash
        };
        return headerMap[header] || ''; // ถ้า header ไม่ตรง ให้เป็นค่าว่าง
    });
    
    sheet.appendRow(finalRow);
    
    return { success: true, message: 'สมัครสมาชิกสำเร็จ' };
    
  } catch (e) {
    Logger.log(`Error in doRegister: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * [ฟังก์ชันใหม่]
 * ตรวจสอบข้อมูลสำหรับลืมรหัสผ่าน (ขั้นตอนที่ 1)
 * @param {object} verificationData - {username, phone}
 * @returns {object} {success: boolean, message: string}
 */
function doForgotPassword(verificationData) {
  try {
    const { username, phone } = verificationData; // phone จาก client จะมี '0' นำหน้าแล้ว
    const sheet = getSheetByName_(USERS_SHEET_NAME);

    if (sheet.getLastRow() < 2) {
      return { success: false, message: 'ไม่พบผู้ใช้ในระบบ' };
    }

    const headers = getSheetHeaders_(sheet);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

    const h = {
      phone: headers.indexOf('Phone'),
      username: headers.indexOf('Username')
    };

    // --- เพิ่ม: ฟังก์ชันสำหรับปรับรูปแบบเบอร์โทร ---
    const normalizePhone = (numStr) => {
      if (!numStr || typeof numStr !== 'string') return '';
      const trimmedNum = numStr.trim();
      if (trimmedNum.startsWith('66')) {
        return trimmedNum.substring(2); // เอา 66 ออก เหลือ 9 หลัก
      }
      if (trimmedNum.startsWith('0')) {
        return trimmedNum.substring(1); // เอา 0 ออก เหลือ 9 หลัก
      }
      return trimmedNum; // คืนค่าเดิมถ้าไม่มี 0 หรือ 66 นำหน้า
    };
    // --- สิ้นสุดการเพิ่ม ---

    for (const row of data) {
      // --- แก้ไข: เปรียบเทียบ username แบบ case-insensitive ---
      if (String(row[h.username] || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()) {
        // --- แก้ไข: ปรับรูปแบบเบอร์โทรทั้งคู่ก่อนเปรียบเทียบ ---
        const sheetPhone = String(row[h.phone] || '').trim();
        const clientPhone = String(phone || '').trim(); // phone ที่รับมามี '0' นำหน้า

        const normalizedSheetPhone = normalizePhone(sheetPhone);
        const normalizedClientPhone = normalizePhone(clientPhone);

        Logger.log(`Comparing Phones - Sheet (Raw): "${sheetPhone}", Client (Raw): "${clientPhone}"`);
        Logger.log(`Comparing Phones - Sheet (Normalized): "${normalizedSheetPhone}", Client (Normalized): "${normalizedClientPhone}"`);

        if (normalizedSheetPhone === normalizedClientPhone && normalizedSheetPhone !== '') { // เพิ่มเช็คว่าไม่เป็นค่าว่างหลัง normalize
          // --- สิ้นสุดการแก้ไข ---
          return { success: true, message: 'ยืนยันตัวตนสำเร็จ' };
        } else {
          return { success: false, message: 'เบอร์โทรศัพท์ไม่ตรงกับชื่อผู้ใช้' };
        }
      }
    }

    return { success: false, message: 'ไม่พบชื่อผู้ใช้นี้' };
  } catch (e) {
    Logger.log(`Error in doForgotPassword: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  }
} // <-- [แก้ไขครั้งที่ 2] เพิ่ม } ที่ขาดไป

/**
 * [ฟังก์ชันใหม่]
 * ตั้งรหัสผ่านใหม่ (ขั้นตอนที่ 2)
 * @param {object} resetData - {username, phone, newPassword}
 * @returns {object} {success: boolean, message: string}
 */
function doResetPassword(resetData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const { username, phone, newPassword } = resetData; // phone จาก client จะมี '0' นำหน้า
    const sheet = getSheetByName_(USERS_SHEET_NAME);
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0];

    const h = {
      phone: headers.indexOf('Phone'),
      username: headers.indexOf('Username'),
      passwordHash: headers.indexOf('PasswordHash')
    };

    // --- เพิ่ม: ฟังก์ชันสำหรับปรับรูปแบบเบอร์โทร (เหมือนเดิม) ---
    const normalizePhone = (numStr) => {
      if (!numStr || typeof numStr !== 'string') return '';
      const trimmedNum = numStr.trim();
      if (trimmedNum.startsWith('66')) { return trimmedNum.substring(2); }
      if (trimmedNum.startsWith('0')) { return trimmedNum.substring(1); }
      return trimmedNum;
    };
    // --- สิ้นสุดการเพิ่ม ---

    let updated = false;
    const clientUsernameLower = String(username || '').trim().toLowerCase();
    const normalizedClientPhone = normalizePhone(String(phone || '').trim());

    for (let i = 1; i < data.length; i++) { // เริ่มที่ 1 (ข้าม header)
      const sheetUsernameLower = String(data[i][h.username] || '').trim().toLowerCase();
      const normalizedSheetPhone = normalizePhone(String(data[i][h.phone] || '').trim());

      // --- แก้ไข: ใช้ username ตัวเล็ก และ phone ที่ปรับรูปแบบแล้วในการเปรียบเทียบ ---
      if (sheetUsernameLower === clientUsernameLower && normalizedSheetPhone === normalizedClientPhone && normalizedClientPhone !== '') {
      // --- สิ้นสุดการแก้ไข ---
        const newPasswordHash = hashPassword_(newPassword);
        data[i][h.passwordHash] = newPasswordHash;
        updated = true;
        break; // เจอและอัปเดตแล้ว
      }
    }

    if (updated) {
      dataRange.setValues(data); // บันทึกข้อมูลที่แก้ไขกลับลงชีต
      return { success: true, message: 'ตั้งรหัสผ่านใหม่สำเร็จ' };
    } else {
      Logger.log(`Reset Password Failed: No matching user found for username "${username}" and normalized phone "${normalizedClientPhone}"`);
      return { success: false, message: 'ไม่สามารถอัปเดตรหัสผ่านได้ (ไม่พบข้อมูลผู้ใช้ที่ตรงกัน)' };
    }

  } catch (e) {
    Logger.log(`Error in doResetPassword: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

// --- CLIENT-CALLABLE FUNCTIONS ---

/**
 * [ฟังก์ชันใหม่] [แก้ไข v4]
 * ใช้สำหรับบันทึกกิจกรรมของผู้ใช้ลงใน Sheet 'UserLog'
 * @param {string} actionType ประเภทของกิจกรรม (เช่น 'PageLoad', 'Login', 'SaveInventory')
 * @param {string} details รายละเอียดของกิจกรรม
 */
function logUserActivity(actionType, details) {
  try {
    // [แก้ไข v4] --- ดึง Username จาก Cache ---
    const cache = CacheService.getUserCache();
    const username = cache.get('loggedInUser');
    if (!username) {
      Logger.log("logUserActivity: ไม่สามารถบันทึก Log ได้ ไม่พบ Username ของผู้ใช้ที่ล็อกอิน (Cache miss or expired)");
      return; // ออกจากการทำงานถ้าไม่รู้ว่าผู้ใช้คือใคร
    }

    let employeeId = 'N/A'; // ค่าเริ่มต้นถ้าหาไม่เจอ

    // --- ค้นหา EmployeeID จากชีต Users ---
    try {
        const usersSheet = getSheetByName_(USERS_SHEET_NAME);
        if (usersSheet.getLastRow() > 1) {
            const headers = getSheetHeaders_(usersSheet);
            const usernameCol = headers.indexOf('Username');
            const employeeIdCol = headers.indexOf('EmployeeID');

            if (usernameCol !== -1 && employeeIdCol !== -1) {
                const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, headers.length).getValues();
                const usernameLower = username.trim().toLowerCase();
                for (const row of usersData) {
                    if (String(row[usernameCol] || '').trim().toLowerCase() === usernameLower) {
                        employeeId = row[employeeIdCol] || 'N/A';
                        break;
                    }
                }
            } else {
                 Logger.log("logUserActivity: ไม่พบคอลัมน์ Username หรือ EmployeeID ในชีต Users");
            }
        }
    } catch (lookupError) {
         Logger.log(`logUserActivity: เกิดข้อผิดพลาดขณะค้นหา EmployeeID: ${lookupError.message}`);
         // ดำเนินการต่อไปโดยใช้ EmployeeID เป็น 'N/A'
    }

    // --- เตรียมข้อมูลสำหรับบันทึก ---
    const logSheet = getSheetByName_(LOG_SHEET_NAME);
    const timestamp = new Date();

    // --- ปรับปรุง Array ที่จะบันทึก ---
    // ลำดับคอลัมน์ใหม่: Timestamp, UserEmail (เดิม), Username, EmployeeID, ActionType, Details
    logSheet.appendRow([
      timestamp,
      Session.getActiveUser().getEmail(), // ยังคงเก็บ Email เดิมไว้ด้วย
      username,          // เพิ่ม Username ที่ดึงมา
      employeeId,        // เพิ่ม EmployeeID ที่หาเจอ
      actionType,
      details
    ]);

  } catch (e) {
    Logger.log(`Error in logUserActivity: ${e.message}\nStack: ${e.stack}`);
  }
}

/**
 * [ฟังก์ชันใหม่]
 * ใช้สำหรับให้หน้าเว็บ (Client) ดึงอีเมลของผู้ใช้ที่ล็อกอินอยู่
 * @returns {string} อีเมลของผู้ใช้
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log(`Error in getUserEmail: ${e.message}`);
    return '';
  }
}

// Function for Inventory Receiving Page
function getInventoryPageData(filters) {
  Logger.log(`--- Starting getInventoryPageData with filters: ${JSON.stringify(filters)} ---`);
  try {
    const sheet = getSheetByName_(INVENTORY_SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      Logger.log("Sheet is empty. Returning default structure.");
      return {
        dashboard: { stats: { totalReceivings: 0, totalF1: 0, totalItems: 0 }, history: [] },
        filters: { allMonths: [], allWarehouses: [] }
      };
    }
    
    const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const headers = getSheetHeaders_(sheet);
    
    const uniqueMonths = [...new Set(allData.map(row => {
      const dateValue = row[headers.indexOf('Timestamped')];
      const date = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
      return !isNaN(date.getTime()) ? Utilities.formatDate(date, TIMEZONE, "yyyy-MM") : null;
    }).filter(Boolean))].sort().reverse();
    
    const uniqueWarehouses = [...new Set(allData.map(row => row[headers.indexOf('Warehouse')]).filter(Boolean))].sort();

    const filteredData = allData.filter(row => {
      const dateValue = row[headers.indexOf('Timestamped')];
      const rowDate = (dateValue instanceof Date) ? dateValue : new Date(dateValue);
      if (isNaN(rowDate.getTime())) return false;

      const rowMonth = Utilities.formatDate(new Date(rowDate), TIMEZONE, "yyyy-MM");
      const rowWarehouse = row[headers.indexOf('Warehouse')];
      
      const monthMatch = !filters || !filters.months || filters.months.length === 0 || filters.months.includes(rowMonth);
      const warehouseMatch = !filters || !filters.warehouses || filters.warehouses.length === 0 || filters.warehouses.includes(rowWarehouse);
      
      if (!(monthMatch && warehouseMatch)) return false;

      if (filters && filters.searchText) {
        const searchText = String(filters.searchText).toLowerCase().trim();
        if(searchText) {
          const f1Value = String(row[headers.indexOf('F1JobNumber')]).toLowerCase();
          const articleIdValue = String(row[headers.indexOf('ArticleID')]).toLowerCase();
          const itemDescValue = String(row[headers.indexOf('ItemDescription')]).toLowerCase();
          return f1Value.includes(searchText) || articleIdValue.includes(searchText) || itemDescValue.includes(searchText);
        }
      }

      return true;
    });
    Logger.log(`Filtered data count: ${filteredData.length} rows.`);

    const f1Index = headers.indexOf('F1JobNumber');
    const qtyIndex = headers.indexOf('Quantity');
    const timestampIndex = headers.indexOf('Timestamped');
    const warehouseIndex = headers.indexOf('Warehouse');
    const articleIdIndex = headers.indexOf('ArticleID');
    const itemDescIndex = headers.indexOf('ItemDescription');
    const photoUrlIndex = headers.indexOf('PhotoURL');
    const receiverIndex = headers.indexOf('Receiver'); 

    const uniqueF1s = new Set();
    const uniqueTimestamps = new Set();
    let totalItems = 0;
    filteredData.forEach(row => {
      if (row[f1Index]) uniqueF1s.add(row[f1Index]);
      if (row[timestampIndex]) uniqueTimestamps.add(row[timestampIndex].toString());
      const quantity = Number(row[qtyIndex]);
      if (!isNaN(quantity)) totalItems += quantity;
    });
    const stats = {
      totalReceivings: uniqueTimestamps.size,
      totalF1: uniqueF1s.size,
      totalItems: totalItems
    };
    Logger.log(`Calculated Stats: ${JSON.stringify(stats)}`);

    const groupedByTimestamp = filteredData.reduce((acc, row) => {
      const timestampKey = row[timestampIndex].toString();
      if (!timestampKey) return acc;

      if (!acc[timestampKey]) {
        acc[timestampKey] = {
          timestamp: row[timestampIndex].toString(),
          warehouse: row[warehouseIndex],
          f1: row[f1Index],
          receiver: row[receiverIndex],
          photos: row[photoUrlIndex] ? String(row[photoUrlIndex]).split(',').map(url => convertToDirectUrl_(url.trim())).filter(Boolean) : [],
          items: []
        };
      }
      acc[timestampKey].items.push({
        name: `${row[articleIdIndex]} - ${row[itemDescIndex]}`,
        quantity: row[qtyIndex]
      });
      return acc;
    }, {});
    
    const sortedHistory = Object.values(groupedByTimestamp)
                                .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
                                .slice(0, 50);
    Logger.log(`Grouped and sorted history. Returning ${sortedHistory.length} items.`);

    Logger.log("--- Finished getInventoryPageData successfully. ---");
    return {
      dashboard: { stats: stats, history: sortedHistory },
      filters: { allMonths: uniqueMonths, allWarehouses: uniqueWarehouses }
    };

  } catch (error) {
    Logger.log(`!!! CRITICAL ERROR in getInventoryPageData: ${error.stack}`);
    return { error: true, message: `Server Error: ${error.message}` };
  }
}

function saveInventoryData(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); 

  try {
    const { rows, files } = payload;
    if (!rows || rows.length === 0) {
      throw new Error("No data rows provided.");
    }

    const sheet = getSheetByName_(INVENTORY_SHEET_NAME);
    const headers = getSheetHeaders_(sheet);
    const timestamp = new Date();
    
    let attachmentUrls = '';
    if (files && files.length > 0) {
      attachmentUrls = processFiles_(files);
    }

    const sheetData = rows.map(row => {
      return headers.map(header => {
        const trimmedHeader = header.trim();
        if (trimmedHeader === 'Timestamped') {
          return Utilities.formatDate(timestamp, TIMEZONE, "dd/MM/yyyy, HH:mm:ss");
        }
        if (trimmedHeader === 'ReceiveDate') {
          return Utilities.formatDate(timestamp, TIMEZONE, "dd/MM/yyyy");
        }
        if (trimmedHeader === 'PhotoURL') {
          return attachmentUrls;
        }
        return row[trimmedHeader] !== undefined ? row[trimmedHeader] : '';
      });
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, sheetData.length, headers.length).setValues(sheetData);

    Logger.log("Successfully saved inventory data.");
    return { success: true, message: "บันทึกข้อมูลสำเร็จ!", savedData: payload };

  } catch (e) {
    Logger.log(`Error in saveInventoryData: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function getPickupPageData(filters) {
    Logger.log(`--- Starting getPickupPageData with filters: ${JSON.stringify(filters)} ---`);
    try {
        const sheet = getSheetByName_(PICKUP_SHEET_NAME);
        if (sheet.getLastRow() <= 1) {
            return { dashboard: { stats: { totalPickups: 0, totalF1: 0, totalItems: 0 }, history: [] }, filters: { allMonths: [], allWarehouses: [], allTeams: [] }, currentPage: 1, totalPages: 1 };
        }

        const headers = getSheetHeaders_(sheet);
        let allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

        // --- Filtering Logic ---
        const filteredData = allData.filter(row => {
            const dateValue = row[headers.indexOf('Timestamped')];
            if (!dateValue) return false;
            const rowDate = new Date(dateValue);
            if (isNaN(rowDate.getTime())) return false;

            if (filters.startDate && rowDate < new Date(filters.startDate).setHours(0,0,0,0)) return false;
            if (filters.endDate && rowDate > new Date(filters.endDate).setHours(23,59,59,999)) return false;
            
            if (!filters.startDate && !filters.endDate) {
                 const rowMonth = Utilities.formatDate(rowDate, TIMEZONE, "yyyy-MM");
                 const monthMatch = !filters || !filters.months || filters.months.length === 0 || filters.months.includes(rowMonth);
                 if (!monthMatch) return false;
            }

            const warehouseMatch = !filters || !filters.warehouses || filters.warehouses.length === 0 || filters.warehouses.includes(row[headers.indexOf('Warehouse')]);
            const teamMatch = !filters || !filters.teams || filters.teams.length === 0 || filters.teams.includes(row[headers.indexOf('TeamName')]);
            if (!(warehouseMatch && teamMatch)) return false;

            if (filters && filters.searchText) {
                const searchText = String(filters.searchText).toLowerCase().trim();
                return ['PickupID', 'F1JobNumber', 'TeamName', 'Picker1', 'Picker2'].some(header => {
                    const headerIndex = headers.indexOf(header);
                    return headerIndex > -1 && String(row[headerIndex]).toLowerCase().includes(searchText);
                });
            }
            return true;
        });

        // --- Grouping Logic ---
        const groupedByPickupId = filteredData.reduce((acc, row) => {
            const pickupId = row[headers.indexOf('PickupID')];
            if (!pickupId) return acc;

            if (!acc[pickupId]) {
                acc[pickupId] = {
                    pickupId: pickupId,
                    timestamp: row[headers.indexOf('Timestamped')].toString(),
                    warehouse: row[headers.indexOf('Warehouse')],
                    f1: row[headers.indexOf('F1JobNumber')],
                    teamName: row[headers.indexOf('TeamName')],
                    picker1: row[headers.indexOf('Picker1')],
                    picker2: row[headers.indexOf('Picker2')],
                    payer: row[headers.indexOf('Payer')],
                    branchCode: row[headers.indexOf('BranchCode')],
                    branchName: row[headers.indexOf('BranchName')],
                    sapNumber: row[headers.indexOf('SAPTrackingNumber')], // Initial value
                    photos: [],
                    itemsDetail: []
                };
            }
            
            // --- THIS IS THE FIX ---
            // If the current row has an SAP number and the group object doesn't, update it.
            const currentSap = row[headers.indexOf('SAPTrackingNumber')];
            if (currentSap && !acc[pickupId].sapNumber) {
                acc[pickupId].sapNumber = currentSap;
            }
            
            acc[pickupId].itemsDetail.push({
                name: `${row[headers.indexOf('ArticleID')]} - ${row[headers.indexOf('ItemDescription')]}`,
                quantity: Number(row[headers.indexOf('Quantity')]) || 0
            });
            
            const photoUrls = row[headers.indexOf('PhotoURL')] ? String(row[headers.indexOf('PhotoURL')]).split(',') : [];
            photoUrls.forEach(url => {
                const directUrl = convertToDirectUrl_(url.trim());
                if (directUrl && !acc[pickupId].photos.includes(directUrl)) {
                    acc[pickupId].photos.push(directUrl);
                }
            });

            return acc;
        }, {});
        
        // --- Post-Grouping Processing ---
        Object.values(groupedByPickupId).forEach(group => {
            const itemsSummaryMap = group.itemsDetail.reduce((summary, item) => {
                summary.set(item.name, (summary.get(item.name) || 0) + item.quantity);
                return summary;
            }, new Map());
            group.items = Array.from(itemsSummaryMap, ([name, quantity]) => ({ name, quantity }));
        });

        const sortedHistory = Object.values(groupedByPickupId).sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        
        // --- Pagination Logic ---
        const page = filters.page || 1;
        const pageSize = filters.pageSize || 10;
        const totalRecords = sortedHistory.length;
        const totalPages = Math.ceil(totalRecords / pageSize) || 1;
        const startIndex = (page - 1) * pageSize;
        const paginatedHistory = sortedHistory.slice(startIndex, startIndex + pageSize);

        // --- Stats Calculation (on filtered data) ---
        const uniqueF1s = new Set();
        let totalItems = 0;
        sortedHistory.forEach(group => {
            if (group.f1) uniqueF1s.add(group.f1);
            group.itemsDetail.forEach(item => { totalItems += item.quantity; });
        });
        const stats = { totalPickups: sortedHistory.length, totalF1: uniqueF1s.size, totalItems: totalItems };
        
        // --- Filter Options (from all data, not just filtered) ---
        const uniqueMonths = [...new Set(allData.map(row => Utilities.formatDate(new Date(row[headers.indexOf('Timestamped')]), TIMEZONE, "yyyy-MM")).filter(Boolean))].sort().reverse();
        const uniqueWarehouses = [...new Set(allData.map(row => row[headers.indexOf('Warehouse')]).filter(Boolean))].sort();
        const uniqueTeams = [...new Set(allData.map(row => row[headers.indexOf('TeamName')]).filter(Boolean))].sort();

        return {
            dashboard: { stats: stats, history: paginatedHistory },
            filters: { allMonths: uniqueMonths, allWarehouses: uniqueWarehouses, allTeams: uniqueTeams },
            currentPage: page,
            totalPages: totalPages
        };
    } catch (error) {
        Logger.log(`!!! CRITICAL ERROR in getPickupPageData: ${error.stack}`);
        return { error: true, message: `Server Error: ${error.message}` };
    }
}

function savePickupData(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { rows, files } = payload;
    if (!rows || rows.length === 0) throw new Error("No data rows provided.");

    const sheet = getSheetByName_(PICKUP_SHEET_NAME);
    const headers = getSheetHeaders_(sheet);
    const timestamp = new Date();
    const attachmentUrls = (files && files.length > 0) ? processFiles_(files) : '';
    
    // --- START: NEW ID GENERATION LOGIC ---
    const today = Utilities.formatDate(timestamp, TIMEZONE, "yyyyMMdd");
    const idPrefix = `SPPIT${today}`;
    
    const idColumnIndex = headers.indexOf('PickupID') + 1;
    let lastSequence = 0;
    
    if (sheet.getLastRow() > 1) {
        const allIDs = sheet.getRange(2, idColumnIndex, sheet.getLastRow() - 1, 1).getValues();
        const todaySequences = allIDs
            .map(idCell => idCell[0])
            .filter(id => id && String(id).startsWith(idPrefix))
            .map(id => parseInt(String(id).slice(-3), 10));

        if (todaySequences.length > 0) {
            lastSequence = Math.max(...todaySequences);
        }
    }
    const newSequence = (lastSequence + 1).toString().padStart(3, '0');
    const pickupId = `${idPrefix}${newSequence}`;
    // --- END: NEW ID GENERATION LOGIC ---


    const sheetData = rows.map(row => {
      // เพิ่ม 'PickupID' เข้าไปในข้อมูลที่จะบันทึก
      const rowObject = { ...row, 'PickupID': pickupId };
      return headers.map(header => {
        if (header === 'Timestamped') return Utilities.formatDate(timestamp, TIMEZONE, "dd/MM/yyyy, HH:mm:ss");
        if (header === 'PickupDate') return Utilities.formatDate(timestamp, TIMEZONE, "dd/MM/yyyy");
        if (header === 'PhotoURL') return attachmentUrls;
        // ใช้ข้อมูลจาก rowObject ที่มี PickupID แล้ว
        return rowObject[header] !== undefined ? rowObject[header] : '';
      });
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, sheetData.length, headers.length).setValues(sheetData);
    return { success: true, message: "บันทึกข้อมูลการเบิกสำเร็จ!" };
  } catch (e) {
    Logger.log(`Error in savePickupData: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

function getImageAsBase64(fileId) {
  try {
    if (!fileId) throw new Error("File ID is required.");
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    return { success: true, base64Data: base64Data, mimeType: blob.getContentType() };
  } catch (e) {
    Logger.log(`Error in getImageAsBase64: ${e.message}`);
    return { success: false, message: e.message };
  }
}

/**
 * Processes files uploaded from the client, saves them to Google Drive,
 * and returns a comma-separated string of their URLs.
 * @param {Array<Object>} files An array of file objects, each with {fileName, mimeType, data}.
 * @returns {string} A comma-separated string of direct-access URLs for the uploaded files.
 */
function processFiles_(files) {
  try {
    if (!files || files.length === 0) {
      Logger.log('processFiles_ was called with no files.');
      return '';
    }
    
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    const urlList = files.map(fileInfo => {
      if (!fileInfo || !fileInfo.data || !fileInfo.mimeType || !fileInfo.fileName) {
        Logger.log(`Skipping invalid fileInfo object: ${JSON.stringify(fileInfo)}`);
        return null; 
      }

      // --- NEW FIX ---
      // This logic makes the function robust. It handles both full Data URLs 
      // (e.g., "data:image/png;base64,iVBOR...") and pure Base64 strings.
      let base64Data = String(fileInfo.data);
      const commaIndex = base64Data.indexOf(',');
      if (commaIndex !== -1) {
        // If a comma is found, it's a Data URL. Get the part after the comma.
        base64Data = base64Data.substring(commaIndex + 1);
      }
      // --- END NEW FIX ---
      
      const decodedData = Utilities.base64Decode(base64Data);
      const blob = Utilities.newBlob(decodedData, fileInfo.mimeType, fileInfo.fileName);
      
      const newFileName = `${Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd_HH-mm-ss")}_${fileInfo.fileName}`;
      const newFile = folder.createFile(blob);
      
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      return `https://drive.google.com/thumbnail?id=${newFile.getId()}&sz=w1024`;

    }).filter(Boolean);

    return urlList.join(',');

  } catch (e) {
    Logger.log(`CRITICAL ERROR in processFiles_: ${e.stack}`);
    // Re-throw the error with a more descriptive message for the client-side.
    throw new Error(`Server failed to process files. Original error: ${e.message}`);
  }
}
function getSheetByName_(name) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(name);
    const expectedHeaders = {
    [LOG_SHEET_NAME]: [ // <-- เพิ่ม case นี้เข้าไป
      'Timestamp', 'UserEmail','Username', 'EmployeeID', 'ActionType', 'Details'
    ],
        [USERS_SHEET_NAME]: [
      'EmployeeName', 'EmployeeID', 'Phone', 'CompanyGmail', 'Username', 'PasswordHash', 'ProfilePicURL'
    ],    
            [EmployeeSheetName]: [
      'EmployeeName', 'EmployeeID', 'access'
    ]
    
  };

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    if (expectedHeaders[name]) {
      sheet.appendRow(expectedHeaders[name]);
      Logger.log(`Created new sheet '${name}' with headers.`);
    }
  } else {
    if (expectedHeaders[name]) {
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const headersToAdd = expectedHeaders[name].filter(h => !currentHeaders.includes(h));
      if (headersToAdd.length > 0) {
        // เพิ่มคอลัมน์ที่ขาดหายไปอัตโนมัติ
        sheet.getRange(1, currentHeaders.length + 1, 1, headersToAdd.length).setValues([headersToAdd]);
        Logger.log(`Added missing headers to sheet '${name}': ${headersToAdd.join(', ')}`);
      }
    }
  }
  return sheet;
}

function getSheetHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
}

function convertToDirectUrl_(url) {
  if (!url || typeof url !== 'string') return '';
  const regex = /(?:drive\.google\.com\/(?:file\/d\/|uc\?id=|thumbnail\?id=))([a-zA-Z0-9_-]{28,})/;
  const match = url.match(regex);
  if (match && match[1]) {
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1024`;
  }
  return url;
}
function checkImageAccessibility(url) {
  Logger.log(`Checking URL: "${url}"`);

  if (!url || typeof url !== 'string') {
    return { success: false, status: 'Invalid URL Input', message: 'ได้รับค่า URL ที่ไม่ถูกต้อง (ไม่ใช่ข้อความ)' };
  }
  
  if (!url.includes('drive.google.com')) {
    return { success: false, status: 'Not a Google Drive URL', message: 'URL ที่ส่งมาไม่ใช่ Google Drive URL', details: { receivedUrl: url } };
  }

  const fileIdMatch = url.match(/id=([a-zA-Z0-9_-]{28,})/);
  
  if (!fileIdMatch || !fileIdMatch[1]) {
    Logger.log(`Failed to extract File ID from URL: "${url}"`);
    return { 
      success: false, 
      status: 'No File ID Found', 
      message: 'ไม่สามารถดึง File ID จาก URL ได้ อาจเป็นเพราะรูปแบบ URL ไม่ถูกต้อง',
      details: { receivedUrl: url }
    };
  }

  const fileId = fileIdMatch[1];
  
  try {
    const file = DriveApp.getFileById(fileId);
    const sharingAccess = file.getSharingAccess().toString();
    const owner = file.getOwner().getEmail();
    const size = file.getSize();

    if (sharingAccess !== 'ANYONE_WITH_LINK' && sharingAccess !== 'ANYONE') {
        return { 
          success: false, 
          status: 'Permission Denied', 
          message: `ไฟล์ไม่ได้แชร์สาธารณะ (ANYONE_WITH_LINK) สถานะปัจจุบันคือ: ${sharingAccess}`,
          details: { fileId: fileId, owner: owner, size: size }
        };
    }
    
    const directUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w1024`;
    const response = UrlFetchApp.fetch(directUrl, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      return { 
        success: false, 
        status: `HTTP Error ${responseCode}`, 
        message: `เซิร์ฟเวอร์ตอบกลับด้วยสถานะผิดปกติ (${responseCode}) เมื่อพยายามเข้าถึงไฟล์โดยตรง`,
        details: { fileId: fileId, owner: owner, size: size, sharing: sharingAccess }
      };
    }

    const contentType = response.getHeaders()['Content-Type'];
    if (!contentType || !contentType.startsWith('image/')) {
        return { 
          success: false, 
          status: 'Not an Image', 
          message: `ไฟล์ที่พบไม่ใช่ไฟล์รูปภาพ แต่เป็นประเภท: ${contentType}`,
          details: { fileId: fileId, owner: owner, size: size, sharing: sharingAccess, contentType: contentType }
        };
    }

    return { 
      success: true, 
      status: 'OK', 
      message: 'สามารถเข้าถึงรูปภาพได้ปกติ',
      details: { fileId: fileId, owner: owner, size: size, sharing: sharingAccess, contentType: contentType }
    };

  } catch (e) {
    Logger.log(`EXCEPTION for fileId "${fileId}" from URL "${url}". Error: ${e.message}`);
    let message = e.message;
    if (e.message.includes("Invalid argument")) {
      message = `File ID ที่ได้มา (${fileId}) ไม่ถูกต้อง อาจจะสั้นหรือยาวเกินไป หรือมีอักขระที่ไม่ได้รับอนุญาต`;
    } else if (e.message.includes("not found")) {
      message = `ไม่พบไฟล์ที่มี ID: ${fileId} ใน Google Drive`;
    }
    return { 
      success: false, 
      status: 'Exception', 
      message: `เกิดข้อผิดพลาดขณะตรวจสอบไฟล์: ${message}`,
      details: { fileId: fileId, originalUrl: url }
    };
  }
}
  
// --- NEW FUNCTIONS FOR USER PROFILE ---

/**
 * [ฟังก์ชันใหม่]
 * ดึงข้อมูลโปรไฟล์ของผู้ใช้ที่ล็อกอินอยู่
 * @returns {object} {success: boolean, profile?: object, message?: string}
 */
function getUserProfile() {
  try {
    // [แก้ไข v4] --- ใช้ Username จาก Cache ---
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
     if (!loggedInUsername) {
      return { success: false, message: 'ไม่สามารถระบุตัวตนผู้ใช้ได้ (กรุณาล็อกอินใหม่)' };
    }
    // --- สิ้นสุดการแก้ไข ---


    // *** ใช้ SPREADSHEET_ID และ SHEET_USERS ที่กำหนดไว้ด้านบน ***
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS); // ใช้ตัวแปร SHEET_USERS
    if (!sheet) {
      return { success: false, message: `ไม่พบชีต '${SHEET_USERS}'` };
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    // --- หาตำแหน่งคอลัมน์ (เผื่อมีการสลับ) ---
    const colIndices = {
      username: headers.indexOf("Username"),       // E
      employeeName: headers.indexOf("EmployeeName"), // A
      employeeID: headers.indexOf("EmployeeID"),   // B
      phone: headers.indexOf("Phone"),           // C
      companyGmail: headers.indexOf("CompanyGmail"), // F (เดิม) -> D
      profilePicURL: headers.indexOf("ProfilePicURL") // G
    };

    // --- ตรวจสอบว่าพบคอลัมน์ที่ต้องการทั้งหมดหรือไม่ ---
    for (let key in colIndices) {
      if (colIndices[key] === -1) {
        return { success: false, message: `ไม่พบคอลัมน์ '${key}' ในชีต '${SHEET_USERS}'` };
      }
    }

    let userRowData = null;
    // --- แก้ไข: ค้นหาแถวของผู้ใช้โดยเทียบจาก Username (คอลัมน์ E) ---
    const usernameLower = loggedInUsername.toLowerCase();
    for (let i = 1; i < values.length; i++) {
      if (values[i][colIndices.username].toString().trim().toLowerCase() === usernameLower) {
        userRowData = values[i];
        break;
      }
    }
    // --- สิ้นสุดการแก้ไข ---

    if (userRowData) {
      const profileData = {
        username: userRowData[colIndices.username],
        employeeName: userRowData[colIndices.employeeName],
        employeeID: userRowData[colIndices.employeeID],
        phone: userRowData[colIndices.phone],
        companyGmail: userRowData[colIndices.companyGmail],
        profilePicURL: userRowData[colIndices.profilePicURL]
      };
      return { success: true, data: profileData };
    } else {
      // logUserActivity('Error', `getUserProfile: User '${loggedInUsername}' not found in sheet.`); // Optional
      return { success: false, message: 'ไม่พบข้อมูลผู้ใช้ในระบบ' };
    }

  } catch (e) {
    console.error(`getUserProfile Error: ${e.message} \n ${e.stack}`); // Log stack trace
    // แสดงข้อความ error ที่ชัดเจนขึ้นถ้าตัวแปรไม่ได้ตั้งค่า
    if (e.message.includes("is not defined")) {
       return { success: false, message: `เกิดข้อผิดพลาด: ไม่สามารถเชื่อมต่อชีตเซิร์ฟเวอร์: ${e.message}` };
    }
    return { success: false, message: `เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ${e.message}` };
  }
}


function updateProfileAndPicture(profileData, pictureData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait up to 30 seconds for the lock

  let profileUpdateResult = null;
  let pictureUpdateResult = null;
  let finalImageUrl = null;

  try {
    // [แก้ไข v4]
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
    if (!loggedInUsername) {
      throw new Error('ไม่สามารถระบุตัวตนผู้ใช้ได้ (กรุณาล็อกอินใหม่)');
    }

    // --- Step 1: Update Profile Data (เรียกใช้ Logic เดิมจาก updateUserProfile) ---
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_USERS);
        if (!sheet) throw new Error(`ไม่พบชีต '${SHEET_USERS}'`);

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        const headers = values[0];
        const colIndices = {
            username: headers.indexOf("Username"), // Use Username for lookup now
            employeeName: headers.indexOf("EmployeeName"),
            employeeID: headers.indexOf("EmployeeID"),
            phone: headers.indexOf("Phone"),
            companyGmail: headers.indexOf("CompanyGmail")
        };
        if (colIndices.username === -1) throw new Error(`ไม่พบคอลัมน์ 'Username' สำหรับค้นหา`);

        let userRowIndex = -1;
        const usernameLower = loggedInUsername.toLowerCase();
        for (let i = 1; i < values.length; i++) {
            if (values[i][colIndices.username].toString().trim().toLowerCase() === usernameLower) {
                userRowIndex = i;
                break;
            }
        }

        if (userRowIndex > -1) {
            const sheetRow = userRowIndex + 1;
            if (colIndices.employeeName !== -1) sheet.getRange(sheetRow, colIndices.employeeName + 1).setValue(profileData.employeeName);
            if (colIndices.employeeID !== -1) sheet.getRange(sheetRow, colIndices.employeeID + 1).setValue(profileData.employeeID);
            if (colIndices.phone !== -1) sheet.getRange(sheetRow, colIndices.phone + 1).setValue(profileData.phone);
            if (colIndices.companyGmail !== -1) sheet.getRange(sheetRow, colIndices.companyGmail + 1).setValue(profileData.companyGmail);
            profileUpdateResult = { success: true };
            Logger.log(`Profile data updated successfully for ${loggedInUsername}`);
        } else {
             throw new Error(`ไม่พบ Username '${loggedInUsername}' ในชีต Users`);
        }
    } catch (profileError) {
        throw new Error(`เกิดข้อผิดพลาดในการอัปเดตข้อมูลโปรไฟล์: ${profileError.message}`);
    }

    // --- Step 2: Update Profile Picture (เรียกใช้ Logic เดิมจาก updateProfilePicture) ---
    try {
        const folderName = "ProfilePictures";
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        if (folders.hasNext()) {
            folder = folders.next();
        } else {
            folder = DriveApp.createFolder(folderName);
            folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            Logger.log(`Created new folder: '${folderName}'. Ensure sharing permissions.`);
        }

        const uniqueFilename = `${loggedInUsername}_${new Date().getTime()}_${pictureData.filename}`;
        const blob = Utilities.newBlob(Utilities.base64Decode(pictureData.base64Data), pictureData.mimeType, uniqueFilename);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileId = file.getId();
        
        // --- *** แก้ไข URL Format ตรงนี้ *** ---
        finalImageUrl = `https://drive.google.com/thumbnail?id=${fileId}`; // ใช้ /thumbnail?id=
        // --- *** สิ้นสุดการแก้ไข *** ---
        
        Logger.log(`Picture uploaded for ${loggedInUsername}, URL: ${finalImageUrl}`);

        // Update URL in Sheet
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // Re-open or reuse sheet object
        const sheet = ss.getSheetByName(SHEET_USERS);
        const values = sheet.getDataRange().getValues(); // Re-fetch values in case sheet changed
        const headers = values[0];
        const usernameCol = headers.indexOf("Username");
        const picUrlCol = headers.indexOf("ProfilePicURL");

        if (usernameCol === -1 || picUrlCol === -1) throw new Error("ไม่พบคอลัมน์ Username หรือ ProfilePicURL");

        let userRowIndex = -1;
        const usernameLower = loggedInUsername.toLowerCase();
        for (let i = 1; i < values.length; i++) {
            if (values[i][usernameCol].toString().trim().toLowerCase() === usernameLower) {
                userRowIndex = i;
                break;
            }
        }
        if (userRowIndex > -1) {
            const sheetRow = userRowIndex + 1;
            sheet.getRange(sheetRow, picUrlCol + 1).setValue(finalImageUrl);
            pictureUpdateResult = { success: true };
            Logger.log(`Profile picture URL updated successfully for ${loggedInUsername}`);
        } else {
            // This case should ideally not happen if profile update succeeded, but handle defensively
             Logger.log(`Warning: Could not find user ${loggedInUsername} to update picture URL after successful profile data update.`);
             // Consider if this should be an error or just a warning
        }

    } catch (pictureError) {
         throw new Error(`เกิดข้อผิดพลาดในการอัปเดตรูปโปรไฟล์: ${pictureError.message}`);
    }

    // --- Step 3: Return Combined Result ---
    SpreadsheetApp.flush(); // Ensure all changes are saved
    return { success: true, newImageUrl: finalImageUrl, message: 'บันทึกข้อมูลและอัปเดตรูปโปรไฟล์สำเร็จ' };

  } catch (e) {
    // Catch errors from either step
    console.error(`updateProfileAndPicture Error: ${e.message} \n ${e.stack}`);
    Logger.log(`ERROR in updateProfileAndPicture: ${e.stack}`);
    // Return a specific error message
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}


// --- ฟังก์ชัน updateUserProfile และ updateProfilePicture เดิมควรคงอยู่ ---
// --- แต่ updateUserProfile จะถูกเรียกใช้เมื่อไม่มีการเปลี่ยนรูปเท่านั้น ---
// --- และ updateProfilePicture จะไม่ถูกเรียกจาก client โดยตรงอีกต่อไป ---

/**
 * อัปเดตข้อมูลโปรไฟล์ของผู้ใช้ (กรณีไม่มีการเปลี่ยนรูป)
 * @param {Object} profileData - ข้อมูลที่ส่งมาจากหน้าเว็บ
 * @returns {ServerResponse} Object ที่มีสถานะ (success) และข้อความ (message)
 */
function updateUserProfile(profileData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
     // [แก้ไข v4]
     const cache = CacheService.getUserCache();
     const loggedInUsername = cache.get('loggedInUser');
     if (!loggedInUsername) {
       return { success: false, message: 'ไม่สามารถระบุตัวตนผู้ใช้ได้ (กรุณาล็อกอินใหม่)' };
     }

     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
     const sheet = ss.getSheetByName(SHEET_USERS);
     if (!sheet) {
       return { success: false, message: `ไม่พบชีต '${SHEET_USERS}'` };
     }

     const dataRange = sheet.getDataRange();
     const values = dataRange.getValues();
     const headers = values[0];
     const colIndices = {
         username: headers.indexOf("Username"), // ใช้ Username สำหรับค้นหา
         employeeName: headers.indexOf("EmployeeName"),
         employeeID: headers.indexOf("EmployeeID"),
         phone: headers.indexOf("Phone"),
         companyGmail: headers.indexOf("CompanyGmail")
     };
     if (colIndices.username === -1) {
        return { success: false, message: `ไม่พบคอลัมน์ 'Username' สำหรับค้นหา` };
     }

     let userRowIndex = -1;
     const usernameLower = loggedInUsername.toLowerCase();
     for (let i = 1; i < values.length; i++) {
       if (values[i][colIndices.username].toString().trim().toLowerCase() === usernameLower) {
         userRowIndex = i;
         break;
       }
     }

     if (userRowIndex > -1) {
       const sheetRow = userRowIndex + 1;
       if (colIndices.employeeName !== -1) sheet.getRange(sheetRow, colIndices.employeeName + 1).setValue(profileData.employeeName);
       if (colIndices.employeeID !== -1) sheet.getRange(sheetRow, colIndices.employeeID + 1).setValue(profileData.employeeID);
       if (colIndices.phone !== -1) sheet.getRange(sheetRow, colIndices.phone + 1).setValue(profileData.phone);
       if (colIndices.companyGmail !== -1) sheet.getRange(sheetRow, colIndices.companyGmail + 1).setValue(profileData.companyGmail);

       SpreadsheetApp.flush();
       return { success: true, message: 'อัปเดตข้อมูลโปรไฟล์สำเร็จ' };
     } else {
       return { success: false, message: `ไม่พบ Username '${loggedInUsername}' ในระบบ` };
     }
  } catch (e) {
    console.error(`updateUserProfile Error: ${e.message}`);
     Logger.log(`ERROR in updateUserProfile: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ${e.message}` };
  } finally {
     lock.releaseLock();
  }
}

/**
 * [แก้ไข v4] เปลี่ยนไปใช้ CacheService
 */
function getProfilePicForHeader() {
  try {
    // [แก้ไข v4]
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
    if (!loggedInUsername) {
      return { success: false, message: 'ไม่สามารถระบุตัวตนผู้ใช้ได้' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS); // ใช้ SHEET_USERS
    if (!sheet) {
      return { success: false, message: `ไม่พบชีต '${SHEET_USERS}'` };
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    const usernameCol = headers.indexOf("Username");
    const profilePicUrlCol = headers.indexOf("ProfilePicURL");

    if (usernameCol === -1 || profilePicUrlCol === -1) {
      return { success: false, message: "ไม่พบคอลัมน์ 'Username' หรือ 'ProfilePicURL' ในชีต 'Users'" };
    }

    let userProfilePicURL = null;
    const usernameLower = loggedInUsername.toLowerCase();
    for (let i = 1; i < values.length; i++) {
      if (values[i][usernameCol].toString().trim().toLowerCase() === usernameLower) {
        userProfilePicURL = values[i][profilePicUrlCol];
        break;
      }
    }

    if (userProfilePicURL) {
      return { success: true, profilePicURL: userProfilePicURL };
    } else {
      // คืนค่า default ถ้าไม่มีรูป
      return { success: true, profilePicURL: 'https://placehold.co/160x160/EFEFEF/AAAAAA?text=Profile' };
    }

  } catch (e) {
    console.error(`getProfilePicForHeader Error: ${e.message} \n ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาด: ${e.message}` };
  }
}


// --- แก้ไขฟังก์ชันนี้ใน Code.gs ของคุณ ---
/**
 * อัปโหลดรูปโปรไฟล์ใหม่ไปที่ Google Drive และอัปเดต URL ในชีต [แก้ไข v4]
 * @param {string} base64Data - ข้อมูลรูปภาพแบบ Base64
 * @param {string} mimeType - Mime type ของไฟล์ (เช่น 'image/png')
 * @param {string} filename - ชื่อไฟล์
 * @returns {ServerResponse} Object ที่มีสถานะ (success), ข้อความ (message), และ URL รูปใหม่ (newImageUrl)
 */
function updateProfilePicture(base64Data, mimeType, filename) {
   const lock = LockService.getScriptLock();
   lock.waitLock(30000);
  try {
     // [แก้ไข v4]
     const cache = CacheService.getUserCache();
     const loggedInUsername = cache.get('loggedInUser');
     if (!loggedInUsername) {
      return { success: false, message: 'ไม่สามารถระบุตัวตนผู้ใช้ได้ (กรุณาล็อกอินใหม่)' };
    }

    const folderName = "ProfilePictures";
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      console.warn(`สร้างโฟลเดอร์ใหม่: '${folderName}' กรุณาตรวจสอบสิทธิ์การเข้าถึง`);
       Logger.log(`Created new folder: '${folderName}'. Ensure sharing permissions are correct.`);
    }

    const uniqueFilename = `${loggedInUsername}_${new Date().getTime()}_${filename}`;

    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, uniqueFilename);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const imageUrl = `https://drive.google.com/uc?id=${fileId}`;
     Logger.log(`Uploaded file ${uniqueFilename} (ID: ${fileId}), URL: ${imageUrl}`);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USERS); // ใช้ SHEET_USERS
    if (!sheet) {
      return { success: false, message: `ไม่พบชีต '${SHEET_USERS}'` };
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];

    const usernameCol = headers.indexOf("Username");
    const picUrlCol = headers.indexOf("ProfilePicURL");

    if (usernameCol === -1 || picUrlCol === -1) {
      return { success: false, message: "ไม่พบคอลัมน์ 'Username' หรือ 'ProfilePicURL'" };
    }

    let userRowIndex = -1;
     const usernameLower = loggedInUsername.toLowerCase();
    for (let i = 1; i < values.length; i++) {
      if (values[i][usernameCol].toString().trim().toLowerCase() === usernameLower) {
        userRowIndex = i;
        break;
      }
    }

    if (userRowIndex > -1) {
      const sheetRow = userRowIndex + 1;
      sheet.getRange(sheetRow, picUrlCol + 1).setValue(imageUrl); // อัปเดต URL ใหม่
      
      // --- การเปลี่ยนแปลงสำคัญอยู่ตรงนี้: ส่ง newImageUrl กลับไป ---
      return { success: true, newImageUrl: imageUrl, message: 'อัปเดตรูปโปรไฟล์สำเร็จ' };
    } else {
      return { success: false, message: 'ไม่พบผู้ใช้ในชีตเพื่ออัปเดต URL รูปภาพ' };
    }

  } catch (e) {
    console.error(`updateProfilePicture Error: ${e.message} \n ${e.stack}`);
     Logger.log(`ERROR in updateProfilePicture: ${e.stack}`);
    if (e.message.includes("is not defined")) {
       return { success: false, message: `เกิดข้อผิดพลาด: ไม่สามารถเชื่อมต่อชีตเซิร์ฟเวอร์: ${e.message}` };
    }
     if (e.message.includes("You do not have permission")) {
       return { success: false, message: `เกิดข้อผิดพลาด: ไม่มีสิทธิ์เข้าถึงโฟลเดอร์ Google Drive '${folderName}'. กรุณาตรวจสอบสิทธิ์.` };
     }
    return { success: false, message: `เกิดข้อผิดพลาดขณะอัปโหลดรูป: ${e.message}` };
  } finally {
     lock.releaseLock();
  }
}

/**
 * [ฟังก์ชันใหม่ - Private Helper] [แก้ไข v4]
 * ดึงระดับสิทธิ์ (GM, Admin, Tech) ของผู้ใช้ที่ล็อกอินปัจจุบันจากชีต Users (จาก Cache)
 * @returns {string|null} ระดับสิทธิ์ของผู้ใช้ หรือ null ถ้าไม่พบ
 */
function getUserAccessLevel_() {
  try {
    // [แก้ไข v4]
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
    if (!loggedInUsername) {
      Logger.log("getUserAccessLevel_: Could not find loggedInUser property (Cache miss or expired)");
      return null;
    }

    // --- Step 1: Find EmployeeID from Users sheet ---
    const usersSheet = getSheetByName_(USERS_SHEET_NAME); // Use your existing helper function
    if (usersSheet.getLastRow() < 2) {
      Logger.log("getUserAccessLevel_: Users sheet has no data");
      return null;
    }

    const usersHeaders = getSheetHeaders_(usersSheet); // Use your existing helper function
    const usernameColUsers = usersHeaders.indexOf('Username');
    const employeeIdColUsers = usersHeaders.indexOf('EmployeeID'); // EmployeeID column in Users sheet

    if (usernameColUsers === -1 || employeeIdColUsers === -1) {
      Logger.log("getUserAccessLevel_: Could not find Username or EmployeeID column in Users sheet");
      return null;
    }

    const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, usersHeaders.length).getValues();
    const loggedInUsernameLower = loggedInUsername.trim().toLowerCase();
    let userEmployeeId = null;

    for (const row of usersData) {
      const sheetUsername = String(row[usernameColUsers] || '').trim().toLowerCase();
      if (sheetUsername === loggedInUsernameLower) {
        userEmployeeId = String(row[employeeIdColUsers] || '').trim();
        break;
      }
    }

    if (!userEmployeeId) {
      Logger.log(`getUserAccessLevel_: Could not find EmployeeID for Username "${loggedInUsername}" in Users sheet`);
      return null; // EmployeeID not found for this user
    }

    // --- Step 2: Find Access Level from Employee sheet using EmployeeID ---
    const employeeSheet = getSheetByName_(EmployeeSheetName); // Use the new sheet name
    if (!employeeSheet || employeeSheet.getLastRow() < 2) {
      Logger.log(`getUserAccessLevel_: Sheet '${EmployeeSheetName}' has no data or was not found`);
      return null; // Employee sheet is empty or not found
    }

    const employeeHeaders = getSheetHeaders_(employeeSheet);
    const employeeIdColEmp = employeeHeaders.indexOf('EmployeeID'); // EmployeeID column in Employee sheet
    const accessColEmp = employeeHeaders.indexOf('access');        // access column in Employee sheet

    if (employeeIdColEmp === -1 || accessColEmp === -1) {
      Logger.log(`getUserAccessLevel_: Could not find EmployeeID or access column in sheet '${EmployeeSheetName}'`);
      return null;
    }

    const employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, employeeHeaders.length).getValues();

    for (const row of employeeData) {
      const sheetEmployeeId = String(row[employeeIdColEmp] || '').trim();
      if (sheetEmployeeId === userEmployeeId) {
        const accessLevel = String(row[accessColEmp] || '').trim().toUpperCase();
        Logger.log(`getUserAccessLevel_: Found EmployeeID ${userEmployeeId} in Employee sheet, Access Level: ${accessLevel || ' (blank)'}`);
        return accessLevel || null; // Return the access level or null if the access column is blank
      }
    }

    Logger.log(`getUserAccessLevel_: EmployeeID "${userEmployeeId}" not found in sheet '${EmployeeSheetName}'`);
    return null; // EmployeeID not found in Employee sheet

  } catch (e) {
    Logger.log(`getUserAccessLevel_ Error: ${e.stack}`);
    return null; // Return null in case of other errors
  }
}

/**
 * [ฟังก์ชันเดิม - ไม่เปลี่ยนแปลง]
 * ตรวจสอบว่าผู้ใช้ปัจจุบันมีสิทธิ์เข้าถึงหน้าที่ร้องขอหรือไม่ (ใช้สำหรับการเข้าถึง URL โดยตรง)
 * @param {string} pageName ชื่อของไฟล์ HTML ที่ต้องการเข้าถึง (เช่น 'FieldService', 'Dashboard')
 * @returns {object} { allow: boolean, message?: string }
 */
function checkPageAccess(pageName) {
  try {
    const accessLevel = getUserAccessLevel_();

    // Allow access to Dashboard and UserProfile for everyone logged in
    // Note: Make sure 'DASHBOARD' and 'USERPROFILE' are uppercased for comparison
    const alwaysAllowedPages = ['DASHBOARD', 'USERPROFILE'];
    if (alwaysAllowedPages.includes(pageName.toUpperCase())) {
        return { allow: true };
    }

    // --- [แก้ไข v5] ปรับตรรกะ GM / Admin ---
    if (!accessLevel) {
      // If no specific level, deny access to restricted pages
      return { allow: false, message: 'ไม่สามารถระบุระดับสิทธิ์ผู้ใช้ได้ กรุณาติดต่อผู้ดูแลระบบ' };
    }

    // GM เข้าถึงได้ทุกหน้า
    if (accessLevel === 'GM') {
      return { allow: true };
    }

    // Admin เข้าถึงได้ทุกหน้า *ยกเว้น* UserManagement
    if (accessLevel === 'ADMIN') {
      if (pageName.toUpperCase() === 'PRINTER') {
        return { allow: false, message: 'Admin ไม่มีสิทธิ์เข้าถึงหน้านี้' };
      }
      return { allow: true };
    }
    // --- [สิ้นสุดการแก้ไข v5] ---
    // --- [แก้ไข] ปรับสิทธิ์สำหรับ SALE ให้เข้าได้เฉพาะหน้า Sale และ Dashboard (Dashboard รวมใน alwaysAllowedPages แล้ว) ---
    if (accessLevel === 'SALE') {
      const allowedSalePages = ['Sale','Dashboard_02', 'Reports', 'USERPROFILE']; // รวมหน้า Dashboard และ UserProfile ไว้เพื่อความชัดเจน
      if (allowedSalePages.includes(pageName.toUpperCase())) {
        return { allow: true };
      } else {
        return { allow: false, message: 'คุณไม่มีสิทธิ์เข้าถึงหน้านี้' };
      }
    }
    // Tech เข้าถึงได้เฉพาะหน้าที่กำหนด
    if (accessLevel === 'RESERVATION') {
      const allowedTechPages = [
        'Reports',
        'Accounting',
        'VoucherUsage',
        'Dashboard_02'

      ];
      // Check against allowed list for Tech
      if (allowedTechPages.includes(pageName.toUpperCase())) {
        return { allow: true };
      } else {
        return { allow: false, message: 'คุณไม่มีสิทธิ์เข้าถึงหน้านี้' };
      }
    }
    
    // --- [เพิ่มใหม่] ตรรกะสำหรับ ITSETUP ---
    else if (accessLevel === 'Sale') {
      const allowedItSetupPages = [
          'Sale',
          'Reports',
          'Dashboard_02',
      ];
      if (allowedItSetupPages.includes(pageName.toUpperCase())) {
          return { allow: true };
      } else {
          return { allow: false, message: 'คุณไม่มีสิทธิ์เข้าถึงหน้านี้' };
      }
    }
    // --- [สิ้นสุดการเพิ่มใหม่] ---


    // Default deny for unrecognized roles
    return { allow: false, message: 'ระดับสิทธิ์ของคุณไม่ได้รับอนุญาตให้เข้าใช้งานหน้านี้' };

  } catch (e) {
    Logger.log(`checkPageAccess Error for page "${pageName}": ${e.stack}`);
    return { allow: false, message: `เกิดข้อผิดพลาดในการตรวจสอบสิทธิ์: ${e.message}` };
  }
}


/**
 * [แก้ไขครั้งที่ 1 & 4 & 5] อัปเดต getAccessiblePages()
 * - เพิ่ม 'UserManagement'
 * - เปลี่ยนไปใช้ CacheService
 * - แยกสิทธิ์ GM / Admin
 */
function getAccessiblePages() {
  try {
    const accessLevel = getUserAccessLevel_(); // [แก้ไข v4] ฟังก์ชันนี้ถูกอัปเดตให้ใช้ Cache แล้ว
    let accessiblePages = [];

    // --- [แก้ไข อัปเดตครั้งที่ 2] เพิ่ม 'UpdatePlan' ---
    const allPages = [
        'Dashboard','ImportF1','UserManagement','Low','High','Work','Checkinsenior','AssignedTasks','Dashboard_02'
    ];
    // --- [สิ้นสุดการแก้ไข] ---

    // Pages always allowed for any logged-in user
    const baseAllowed = ['Dashboard', 'UserProfile'];

    // --- [แก้ไข v5] ปรับตรรกะ GM / Admin ---
    if (!accessLevel) {
       // If no access level defined, only allow base pages
       accessiblePages = [...baseAllowed];
       Logger.log("getAccessiblePages: No access level found, returning base pages.");
    
    } else if (accessLevel === 'GM') {
      // GM gets all pages
      accessiblePages = [...allPages];
      Logger.log(`getAccessiblePages: User is ${accessLevel}, returning all pages.`);
    
    } else if (accessLevel === 'ADMIN') {
      // Admin gets all pages EXCEPT UserManagement
      accessiblePages = allPages.filter(page => page.toUpperCase() !== 'PRINTER');
      Logger.log(`getAccessiblePages: User is ${accessLevel}, returning all pages except Printer.`);

    } else if (accessLevel === 'RESERVATION') {
      // Tech gets base pages + specific Tech pages
      const techSpecificPages = [
        'Reports',
        'Accounting',
        'VoucherUsage',
        'Dashboard_02',
        'Deletes'
      ];
      // Dashboard is already included in baseAllowed, which is always added for TECH.
      accessiblePages = [...baseAllowed, ...techSpecificPages];
      Logger.log("getAccessiblePages: User is TECH, returning specific pages including Dashboard via baseAllowed.");
    // --- [แก้ไข] ปรับสิทธิ์สำหรับ SALE ให้เห็นแค่ Dashboard และ Sale ---
       } else if (accessLevel === 'SALE') {
        const saleSpecificPages = ['Sale','Reports','Dashboard_02']; // Dashboard และ UserProfile อยู่ใน baseAllowed แล้ว
        accessiblePages = [...baseAllowed, ...saleSpecificPages];
        Logger.log("getAccessiblePages: User is SALE, returning Dashboard and Sale.");
    // --- [สิ้นสุดการแก้ไข] --- 

    // --- [เพิ่มใหม่] ตรรกะสำหรับ ITSETUP ---
    } else if (accessLevel === 'ITSETUP') {
        // IT Setup จะได้หน้าพื้นฐาน + หน้าที่กำหนด
        const itSetupSpecificPages = [
           'DC_Out',
           'Sale',
            'Financial',
            'Products', // เบิกอะไหล่
            'StockSetUp',    // รับ Stock (IT Setup)
            'ITSetupAsset',  // เตรียมทรัพย์สิน Setup
            'NewStore',      // ภาพรวมสาขาใหม่
            'UpdatePlan',    // รายชื่อสาขาใหม่
            'POS',           // งานเครื่อง POS
            'TasksCalendar',
            'PANTUM'
        ];
        accessiblePages = [...baseAllowed, ...itSetupSpecificPages];
        Logger.log("getAccessiblePages: User is ITSETUP, returning specific pages.");
    // --- [สิ้นสุดการเพิ่มใหม่] ---

    } else {
        // Unrecognized role, only allow base pages
        accessiblePages = [...baseAllowed];
        Logger.log(`getAccessiblePages: Unrecognized access level "${accessLevel}", returning base pages.`);
    }
    // --- [สิ้นสุดการแก้ไข v5] ---

    // Ensure uniqueness (though defined lists should be unique) and convert to uppercase for reliable comparison later
    const finalPages = [...new Set(accessiblePages)].map(page => page.toUpperCase());

    return { success: true, pages: finalPages };

  } catch (e) {
    Logger.log(`getAccessiblePages Error: ${e.stack}`);
    return { success: false, message: `เกิดข้อผิดพลาดในการดึงรายการหน้า: ${e.message}` };
  }
}

// --- Public: Send User Info to Frontend ---
function getCurrentUserInfo() {
  return getUpdaterInfo_();
}

/**
 * 1. Get User Data
 * Reads 'Recorded By' (Col H) and 'Last Updated' (Col I)
 */
function getUserManagementData() {
  try {
    const usersSheet = getSheetByName_(USERS_SHEET_NAME);
    const employeeSheet = getSheetByName_(EmployeeSheetName);

    // 1. Read Employee Access
    const employeeAccessMap = new Map();
    if (employeeSheet.getLastRow() > 1) {
      const empHeaders = getSheetHeaders_(employeeSheet);
      const empIdCol = empHeaders.indexOf('EmployeeID');
      const accessCol = empHeaders.indexOf('access');
      if (empIdCol !== -1 && accessCol !== -1) {
        const empData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, empHeaders.length).getValues();
        empData.forEach(row => {
          const empId = String(row[empIdCol]).trim();
          if (empId) employeeAccessMap.set(empId, row[accessCol] || 'N/A');
        });
      }
    }

    // 2. Read Users
    const users = [];
    if (usersSheet.getLastRow() > 1) {
      const userData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, usersSheet.getLastColumn()).getDisplayValues();
      const userHeaders = getSheetHeaders_(usersSheet);
      
      const colIndices = {
        EmployeeName: userHeaders.indexOf('EmployeeName'),
        EmployeeID: userHeaders.indexOf('EmployeeID'),
        Department: userHeaders.indexOf('Department'), // <--- เพิ่มคอลัมน์ Department
        Phone: userHeaders.indexOf('Phone'),
        CompanyGmail: userHeaders.indexOf('CompanyGmail'),
        Username: userHeaders.indexOf('Username')
      };

      let recordByIdx = userHeaders.indexOf('Recorded By');
      let recordDateIdx = userHeaders.indexOf('Recorded Date');
      if (recordByIdx === -1) recordByIdx = 7;
      if (recordDateIdx === -1) recordDateIdx = 8;

      userData.forEach(row => {
        const employeeID = String(row[colIndices.EmployeeID]).trim();
        users.push({
          employeeName: row[colIndices.EmployeeName],
          employeeID: employeeID,
          department: colIndices.Department !== -1 ? row[colIndices.Department] : '-', // <--- ส่งค่า Department กลับไป
          phone: row[colIndices.Phone],
          companyGmail: row[colIndices.CompanyGmail],
          username: row[colIndices.Username],
          access: employeeAccessMap.get(employeeID) || 'N/A',
          recordedBy: row[recordByIdx] || '-', 
          recordedDate: row[recordDateIdx] || '-' 
        });
      });
    }
    
    users.sort((a, b) => String(a.employeeName || '').localeCompare(String(b.employeeName || '')));

    return { success: true, users: users };

  } catch (e) {
    Logger.log(`Error in getUserManagementData: ${e.stack}`);
    return { success: false, message: `Server Error: ${e.message}` };
  }
}

/**
 * [ฟังก์ชันใหม่ที่เพิ่มเข้าไป]
 * Helper function สำหรับค้นหาแถว (1-based index) จากค่าในคอลัมน์ที่ระบุ
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {*} valueToFind The value to look for.
 * @param {number} colIndex The 0-based column index to search in.
 * @returns {number} The 1-based row index (e.g., 1, 2, 3...) or -1 if not found.
 */
function findRowIndex_(sheet, valueToFind, colIndex) {
  if (colIndex === -1) {
    Logger.log(`findRowIndex_: Column index is -1. Search cannot be performed.`);
    return -1; // Column not found
  }
  if (sheet.getLastRow() < 2) {
    Logger.log(`findRowIndex_: Sheet '${sheet.getName()}' has no data.`);
    return -1; // No data
  }
  
  // ดึงข้อมูลเฉพาะคอลัมน์ที่ต้องการค้นหา (ตั้งแต่แถวแรก)
  const columnValues = sheet.getRange(1, colIndex + 1, sheet.getLastRow(), 1).getValues();
  const searchValue = String(valueToFind).trim().toLowerCase();

  // วนลูปตั้งแต่แถวที่ 2 (index 1 in array)
  for (let i = 1; i < columnValues.length; i++) {
    if (String(columnValues[i][0]).trim().toLowerCase() === searchValue) {
      return i + 1; // คืนค่า 1-based sheet row number (e.g., แถวที่ 5)
    }
  }
  return -1; // Not found
}

/**
 * 2. Add New User
 * Saves Log to Col H & I
 */
function addNewUser(userData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const { employeeName, employeeID, department, phone, companyGmail, username, newPassword, access } = userData;
    
    const usersSheet = getSheetByName_(USERS_SHEET_NAME);
    const employeeSheet = getSheetByName_(EmployeeSheetName);
    
    // หากยังไม่มีคอลัมน์ Department ในชีต ให้สร้างขึ้นมาอัตโนมัติต่อท้าย
    let headers = getSheetHeaders_(usersSheet);
    if (headers.indexOf('Department') === -1) {
        usersSheet.getRange(1, usersSheet.getLastColumn() + 1).setValue('Department');
        headers = getSheetHeaders_(usersSheet); // ดึง Header ใหม่อีกรอบ
    }

    // Check Duplicates
    if (findRowIndex_(usersSheet, username, headers.indexOf('Username')) !== -1) throw new Error(`Username '${username}' already exists.`);
    if (findRowIndex_(usersSheet, employeeID, headers.indexOf('EmployeeID')) !== -1) throw new Error(`ID '${employeeID}' already exists.`);

    // Prepare Log
    const updaterInfo = getUpdaterInfo_();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    // Ensure headers for logs exist
    let recordByIdx = headers.indexOf('Recorded By');
    let recordDateIdx = headers.indexOf('Recorded Date');
    
    if (recordByIdx === -1 || recordDateIdx === -1) {
       const nextCol = usersSheet.getLastColumn() + 1;
       usersSheet.getRange(1, nextCol).setValue('Recorded By'); 
       usersSheet.getRange(1, nextCol + 1).setValue('Recorded Date');
       headers = getSheetHeaders_(usersSheet); // อัปเดต headers ใหม่
    }

    const passwordHash = hashPassword_(newPassword);
    const newUserRow = new Array(headers.length).fill(''); // สร้างอาร์เรย์ว่างๆ ความยาวเท่าจำนวนคอลัมน์
    
    // Map data to columns based on header names
    headers.forEach((header, index) => {
        if(header === 'EmployeeName') newUserRow[index] = employeeName;
        else if(header === 'EmployeeID') newUserRow[index] = employeeID;
        else if(header === 'Department') newUserRow[index] = department; // <--- ใส่ข้อมูลแผนก
        else if(header === 'Phone') newUserRow[index] = phone;
        else if(header === 'CompanyGmail') newUserRow[index] = companyGmail;
        else if(header === 'Username') newUserRow[index] = username;
        else if(header === 'PasswordHash') newUserRow[index] = passwordHash;
        else if(header === 'Recorded By') newUserRow[index] = updaterInfo;
        else if(header === 'Recorded Date') newUserRow[index] = timestamp;
    });

    usersSheet.appendRow(newUserRow);

    // Add to Employee Sheet
    employeeSheet.appendRow([employeeName, employeeID, access]);

    return { success: true, message: `Added user '${username}' successfully.` };
    
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
 * 3. Update User
 * Updates Log in Col H & I
 */
function updateUser(userData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const { originalUsername, employeeName, employeeID, department, phone, companyGmail, username, access, newPassword } = userData;
    
    const usersSheet = getSheetByName_(USERS_SHEET_NAME);
    const employeeSheet = getSheetByName_(EmployeeSheetName);
    
    // หากยังไม่มีคอลัมน์ Department ในชีต ให้สร้างขึ้นมาอัตโนมัติต่อท้าย
    let usersHeaders = getSheetHeaders_(usersSheet);
    if (usersHeaders.indexOf('Department') === -1) {
        usersSheet.getRange(1, usersSheet.getLastColumn() + 1).setValue('Department');
        usersHeaders = getSheetHeaders_(usersSheet); // ดึง Header ใหม่อีกรอบ
    }
    
    const hUsers = usersHeaders.reduce((acc, h, i) => (acc[h] = i, acc), {});
    const userRowIndex = findRowIndex_(usersSheet, originalUsername, hUsers.Username); 
    if (userRowIndex === -1) throw new Error(`User '${originalUsername}' not found.`);
    
    // Prepare Log
    const updaterInfo = getUpdaterInfo_();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    // Update Users Sheet
    const row = userRowIndex; 
    if(hUsers.EmployeeName !== undefined) usersSheet.getRange(row, hUsers.EmployeeName + 1).setValue(employeeName);
    if(hUsers.EmployeeID !== undefined) usersSheet.getRange(row, hUsers.EmployeeID + 1).setValue(employeeID);
    if(hUsers.Department !== undefined) usersSheet.getRange(row, hUsers.Department + 1).setValue(department || '-'); // <--- บันทึกแก้ไขแผนก
    if(hUsers.Phone !== undefined) usersSheet.getRange(row, hUsers.Phone + 1).setValue(phone);
    if(hUsers.CompanyGmail !== undefined) usersSheet.getRange(row, hUsers.CompanyGmail + 1).setValue(companyGmail);
    if(hUsers.Username !== undefined) usersSheet.getRange(row, hUsers.Username + 1).setValue(username);
    
    if (newPassword && newPassword.length >= 6 && hUsers.PasswordHash !== undefined) {
      usersSheet.getRange(row, hUsers.PasswordHash + 1).setValue(hashPassword_(newPassword));
    }

    // Update History
    let recordByIdx = usersHeaders.indexOf('Recorded By');
    let recordDateIdx = usersHeaders.indexOf('Recorded Date');
    if(recordByIdx !== -1) usersSheet.getRange(row, recordByIdx + 1).setValue(updaterInfo);
    if(recordDateIdx !== -1) usersSheet.getRange(row, recordDateIdx + 1).setValue(timestamp);
    
    // Update Employee Sheet
    const empHeaders = getSheetHeaders_(employeeSheet);
    const hEmp = empHeaders.reduce((acc, h, i) => (acc[h] = i, acc), {});
    const empRowIndex = findRowIndex_(employeeSheet, employeeID, hEmp.EmployeeID);
    
    if (empRowIndex !== -1) {
      employeeSheet.getRange(empRowIndex, hEmp.EmployeeName + 1).setValue(employeeName);
      employeeSheet.getRange(empRowIndex, hEmp.access + 1).setValue(access);
    } else {
      // Create if missing
      employeeSheet.appendRow([employeeName, employeeID, access]);
    }

    return { success: true, message: `Updated user '${username}' successfully.` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 4. Delete User
 * Backs up to 'DeletedUsers' sheet
 */
function deleteUser(username) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const usersSheet = getSheetByName_(USERS_SHEET_NAME);
    const employeeSheet = getSheetByName_(EmployeeSheetName);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // --- Backup Sheet ---
    let deletedSheet = ss.getSheetByName('DeletedUsers');
    if (!deletedSheet) {
      deletedSheet = ss.insertSheet('DeletedUsers');
      deletedSheet.appendRow(['Deleted Date', 'Deleted By', 'EmployeeName', 'EmployeeID', 'Username', 'Phone', 'Gmail']);
    }

    const usersHeaders = getSheetHeaders_(usersSheet);
    const hUsers = usersHeaders.reduce((acc, h, i) => (acc[h] = i, acc), {});
    const userRowIndex = findRowIndex_(usersSheet, username, hUsers.Username); 
    
    if (userRowIndex === -1) throw new Error(`User not found.`);

    // Get Data
    const userData = usersSheet.getRange(userRowIndex, 1, 1, usersSheet.getLastColumn()).getValues()[0];
    const deleterInfo = getUpdaterInfo_();
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    // Save Backup
    deletedSheet.appendRow([
      timestamp,
      deleterInfo,
      userData[hUsers.EmployeeName],
      userData[hUsers.EmployeeID],
      userData[hUsers.Username],
      userData[hUsers.Phone],
      userData[hUsers.CompanyGmail]
    ]);

    // Delete
    const employeeID = userData[hUsers.EmployeeID];
    usersSheet.deleteRow(userRowIndex);

    // Delete from Employee Sheet too
    const empHeaders = getSheetHeaders_(employeeSheet);
    const hEmp = empHeaders.reduce((acc, h, i) => (acc[h] = i, acc), {});
    const empRowIndex = findRowIndex_(employeeSheet, employeeID, hEmp.EmployeeID);
    if (empRowIndex !== -1) employeeSheet.deleteRow(empRowIndex);

    return { success: true, message: `Deleted successfully.` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// --- IMAGE HELPER (Modified to accept Folder ID) ---
function saveImageToDrive(base64Data, fileName, targetFolderId) {
  try {
    if (!base64Data) return null;
    
    // ใช้ Folder ID ที่ส่งมา หรือถ้าไม่ส่งให้ใช้ค่า Default (FOLDER_ID)
    const destFolderId = targetFolderId || FOLDER_ID;
    
    const contentType = base64Data.split(',')[0].split(':')[1].split(';')[0];
    const decoded = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(decoded, contentType, fileName);
    
    const folder = DriveApp.getFolderById(destFolderId);
    const file = folder.createFile(blob);
    
    // -------------------------------------------------------------
    // แก้ไข: ป้องกัน Error จาก Policy การแชร์ของ Google Workspace (องค์กร)
    // -------------------------------------------------------------
    try {
      // พยายามแชร์แบบสาธารณะก่อน
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareError) {
      Logger.log("ติด Policy องค์กร ไม่สามารถแชร์สาธารณะได้: " + shareError.toString());
      try {
        // หากติด Policy ให้พยายามแชร์ให้ดูได้เฉพาะคนในองค์กรแทน
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (domainShareError) {
        Logger.log("ไม่สามารถแชร์ระดับองค์กรได้ ข้ามการเซ็ตสิทธิ์: " + domainShareError.toString());
        // ปล่อยผ่านไป (ไฟล์ถูกสร้างลง Drive แล้ว แค่สิทธิ์เป็นส่วนตัว)
      }
    }
    // -------------------------------------------------------------
    
    // *** ใช้ฟังก์ชันแปลง URL เพื่อให้ได้ลิงก์ที่แสดงผลรูปภาพได้ทันที ***
    return convertToDirectUrl_(file.getUrl());
  } catch (e) {
    Logger.log("Error saving image: " + e.toString());
    return null;
  }
}

/**
 * [REVISED] แปลง URL ของ Google Drive ทุกรูปแบบให้เป็น URL สำหรับแสดงผลภาพ Thumbnail
 * ช่วยให้รูปภาพแสดงบนหน้าเว็บ HTML ได้โดยไม่ติดปัญหา Permission หรือ iFrame
 */
function convertToDirectUrl_(url) {
  if (!url || typeof url !== 'string') return '';
  
  // Regex จับ File ID จาก URL รูปแบบต่างๆ ของ Drive
  const regex = /(?:drive\.google\.com\/(?:file\/d\/|uc\?id=|thumbnail\?id=))([a-zA-Z0-9_-]{28,})/;
  const match = url.match(regex);
  
  if (match && match[1]) {
    // แปลงเป็น Thumbnail URL (sz=w1024 คือกำหนดขนาดความกว้างสูงสุด 1024px)
    return `https://drive.google.com/thumbnail?id=${match[1]}&sz=w1024`;
  }
  
  return url; // ถ้าไม่ตรงรูปแบบ ให้คืนค่าเดิม
}

/**
 * ตรวจสอบการเข้าถึงรูปภาพ (สำหรับ Debugging)
 */
function checkImageAccessibility(url) {
  Logger.log(`Checking accessibility for: ${url}`);
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return { 
      status: response.getResponseCode(), 
      contentType: response.getHeaders()['Content-Type'] 
    };
  } catch (e) {
    return { error: e.message };
  }
}

// --- USER HELPER ---
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  return {
    email: email,
    username: email ? email.split('@')[0] : 'Guest'
  };
}

/**
 * ดึง EmployeeID ของผู้ใช้ที่ล็อกอินอยู่
 * @returns {object} { success: boolean, employeeId: string }
 */
function getUserEmployeeId() {
  try {
    const cache = CacheService.getUserCache();
    const username = cache.get('loggedInUser');
    
    if (!username) return { success: false, message: 'No user logged in' };
    
    // ใช้ฟังก์ชัน getUpdaterInfo_() ซึ่งได้ EmployeeID - EmployeeName ตามที่คุณต้องการพอดี
    const updaterInfo = getUpdaterInfo_();
    
    if (updaterInfo && updaterInfo !== 'Unknown User') {
        return { success: true, employeeId: updaterInfo, fullInfo: updaterInfo };
    }
    
    return { success: false, message: 'User not found' };

  } catch (e) {
    Logger.log('Error getting employee ID: ' + e.message);
    return { success: false, message: e.message };
  }
}

// =======================================================
// --- ส่วนที่เพิ่มใหม่สำหรับ Reports.html (Edit & Delete) ---
// =======================================================
// --- เพิ่ม Helper Function สำหรับดึงข้อมูล User ปัจจุบันให้ Frontend ---
function getCurrentUserInfo() {
  return getUpdaterInfo_();
}
/**
 * Helper: ดึงข้อมูลผู้อัปเดต (EmployeeID - Name) จาก Cache
 */
function getUpdaterInfo_() {
  try {
    const cache = CacheService.getUserCache();
    const username = cache.get('loggedInUser');
    if (!username) return 'Unknown User';

    const sheet = getSheetByName_(USERS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const userCol = headers.indexOf('Username');
    const empIdCol = headers.indexOf('EmployeeID');
    const nameCol = headers.indexOf('EmployeeName');

    if (userCol === -1) return username;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][userCol]).toLowerCase() === username.toLowerCase()) {
        const empID = data[i][empIdCol] || '-';
        const empName = data[i][nameCol] || '-';
        return `${empID} - ${empName}`;
      }
    }
    return username;
  } catch (e) {
    return 'System Error';
  }
}

/**
 * =======================================================
 * --- ส่วนการนำเข้าข้อมูล Import F1 ---
 * =======================================================
 */

/**
 * 2. ฟังก์ชันอัปโหลดและจ่ายงาน (เวอร์ชัน Optimized รวดเร็ว 1-2 วินาที)
 * - ใช้ Batch Processing อัปโหลดรวดเดียว
 * - ใช้ getValues() แทน getDisplayValues() เพื่อความเร็วสูงสุด
 */
function processImportF1Data(dataArray, fileName, selectedWorkers) {
  try {
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
    
    if (!loggedInUsername) {
      throw new Error("หมดเวลาเชื่อมต่อ หรือผู้ใช้ไม่ได้ล็อกอิน กรุณารีเฟรชเพื่อล็อกอินใหม่");
    }

    let uploaderName = getUpdaterInfo_() || loggedInUsername;

    if (!dataArray || dataArray.length <= 1) {
      throw new Error("ไม่มีข้อมูลสำหรับนำเข้า");
    }

    // --- 1. คัดกรองข้อมูลใน Memory (เร็วมาก 0.01 วินาที) ---
    const headers = dataArray[0];
    const catalogMainIndex = headers.findIndex(h => String(h || "").trim() === "Services Catalog Main");
    let filteredData = [headers];

    if (catalogMainIndex !== -1) {
      for (let i = 1; i < dataArray.length; i++) {
        if (String(dataArray[i][catalogMainIndex] || "").trim() === "งานไอที") {
          filteredData.push(dataArray[i]);
        }
      }
      if (filteredData.length <= 1) throw new Error("ไม่พบข้อมูลที่เป็น 'งานไอที'");
    } else {
      throw new Error("ไม่พบคอลัมน์ 'Services Catalog Main'");
    }

    // --- 2. เปิด Spreadsheet แค่ครั้งเดียว (ลดเวลาโหลด) ---
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);

    // --- 3. บันทึก Sheet "Data" (เขียนรวดเดียว) ---
    let dataSheet = ss.getSheetByName("Data") || ss.insertSheet("Data");
    dataSheet.clearContents(); 
    
    const numRows = filteredData.length;
    const numCols = filteredData[0].length;
    
    const normalizedData = filteredData.map(row => {
      let newRow = [...row];
      while (newRow.length < numCols) newRow.push("");
      return newRow.slice(0, numCols); 
    });
    
    // สาดข้อมูลลง Data ทีเดียว
    dataSheet.getRange(1, 1, numRows, numCols).setValues(normalizedData);

    // ====================================================================
    // --- 4. ระบบจ่ายงานลง Sheet "AddName" (เช็คซ้ำและเขียนรวดเดียว) ---
    // ====================================================================
    let assignMsg = "";
    let assignedCount = 0;
    let duplicateCount = 0;
    let duplicateAssignedCount = 0;
    let summaryText = "ไม่มีการจ่ายงาน"; 

    selectedWorkers = selectedWorkers || [];

    if (selectedWorkers.length === 0) {
      assignMsg = "ข้ามการจ่ายงาน (คุณไม่ได้เลือกพนักงาน)";
    } else {
      let taskNoColIndex = headers.findIndex(h => String(h || "").trim().toLowerCase().includes("task no"));
      if (taskNoColIndex === -1) taskNoColIndex = 3; 

      // ตรวจข้อมูลซ้ำใน AddName (อ่านรวดเดียวด้วย getValues)
      const existingTasksAddName = new Set();
      let addNameSheet = ss.getSheetByName("AddName");
      
      if (!addNameSheet) {
        addNameSheet = ss.insertSheet("AddName");
        const addNameHeaders = [...headers, "Assignee", "Timestamp"]; 
        addNameSheet.appendRow(addNameHeaders);
        addNameSheet.getRange(1, 1, 1, addNameHeaders.length).setBackground("#374151").setFontColor("white").setFontWeight("bold");
        addNameSheet.setFrozenRows(1);
      } else {
        const lastRowAddName = addNameSheet.getLastRow();
        if (lastRowAddName > 1) {
          const addNameData = addNameSheet.getRange(2, taskNoColIndex + 1, lastRowAddName - 1, 1).getValues();
          for (let r = 0; r < addNameData.length; r++) {
            if (addNameData[r][0]) existingTasksAddName.add(String(addNameData[r][0]).trim());
          }
        }
      }

      // ตรวจข้อมูลซ้ำใน Assigned_Tasks (อ่านรวดเดียวด้วย getValues)
      const existingTasksAssigned = new Set();
      let assignedSheet = ss.getSheetByName("Assigned_Tasks");
      if (assignedSheet) {
        const lastRowAssigned = assignedSheet.getLastRow();
        if (lastRowAssigned > 1) {
          const assignedData = assignedSheet.getRange(2, 16, lastRowAssigned - 1, 1).getValues();
          for (let r = 0; r < assignedData.length; r++) {
            if (assignedData[r][0]) existingTasksAssigned.add(String(assignedData[r][0]).trim());
          }
        }
      }

      // คัดแยกงานใน Memory
      const tasksToAssign = [];
      for (let i = 1; i < filteredData.length; i++) {
        const row = filteredData[i];
        const taskNo = String(row[taskNoColIndex] || "").trim();

        if (!taskNo) continue;

        if (existingTasksAssigned.has(taskNo)) {
          duplicateAssignedCount++;
        } else if (existingTasksAddName.has(taskNo)) {
          duplicateCount++;
        } else {
          tasksToAssign.push(row);
          if (selectedWorkers.length === 1 && tasksToAssign.length >= 25) break; 
        }
      }

      // เขียนข้อมูลลง AddName รวดเดียว
      if (tasksToAssign.length > 0) {
        const rowsToSave = [];
        const now = new Date();
        const timeZone = Session.getScriptTimeZone() || "GMT+7";
        const formattedNow = Utilities.formatDate(now, timeZone, "dd/MM/yyyy HH:mm:ss");

        const workerTaskCount = {};
        selectedWorkers.forEach(w => workerTaskCount[w] = 0);

        for (let i = 0; i < tasksToAssign.length; i++) {
          const assignedWorker = selectedWorkers[i % selectedWorkers.length];
          const newRow = [...tasksToAssign[i]];
          
          while(newRow.length < headers.length) newRow.push(""); 
          newRow.push(assignedWorker); 
          newRow.push(formattedNow);   
          
          rowsToSave.push(newRow);
          workerTaskCount[assignedWorker]++; 
        }

        // เขียนรวดเดียว (Batch Write)
        addNameSheet.getRange(addNameSheet.getLastRow() + 1, 1, rowsToSave.length, rowsToSave[0].length).setValues(rowsToSave);
        assignedCount = rowsToSave.length;
        assignMsg = `จ่ายงานใหม่ ${assignedCount} รายการ ให้ Senior ${selectedWorkers.length} คน เรียบร้อยแล้ว`;
        
        summaryText = Object.entries(workerTaskCount)
          .filter(([name, count]) => count > 0)
          .map(([name, count]) => `${name.split(" - ")[1] || name}: ${count} งาน`)
          .join(', ');

      } else {
        assignMsg = "ไม่มีงานใหม่ให้จ่าย (งานในไฟล์ถูกดำเนินการ หรือมีค้างในระบบทั้งหมดแล้ว)";
      }
    }

    // --- 5. บันทึกชีต Datalog ---
    let logSheet = ss.getSheetByName("Datalog");
    if (!logSheet) {
      logSheet = ss.insertSheet("Datalog");
      logSheet.appendRow(["Timestamp", "วันที่", "ผู้บันทึก", "ชื่อไฟล์", "จำนวนครั้ง/วัน", "รายละเอียดการจ่ายงาน"]);
    } else if (logSheet.getLastColumn() < 6) {
      logSheet.getRange(1, 6).setValue("รายละเอียดการจ่ายงาน");
    }

    const now = new Date();
    const timeZone = Session.getScriptTimeZone() || "GMT+7";
    const todayStrLog = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    let dailyImportCount = 1;
    
    const lastRowLog = logSheet.getLastRow();
    if (lastRowLog > 1) {
      const timestamps = logSheet.getRange(2, 1, lastRowLog - 1, 1).getValues();
      for (let i = 0; i < timestamps.length; i++) {
        const tsVal = timestamps[i][0];
        if (tsVal && (tsVal instanceof Date || !isNaN(new Date(tsVal).getTime()))) {
          if (Utilities.formatDate(new Date(tsVal), timeZone, "yyyy-MM-dd") === todayStrLog) dailyImportCount++;
        }
      }
    }
    
    logSheet.appendRow([now, todayStrLog, uploaderName, fileName, dailyImportCount, summaryText]);

    return {
      success: true,
      rowCount: filteredData.length > 0 ? filteredData.length - 1 : 0, 
      dailyCount: dailyImportCount,
      assignMsg: assignMsg,           
      assignedCount: assignedCount,
      duplicateCount: duplicateCount,                 
      duplicateAssignedCount: duplicateAssignedCount, 
      activeWorkers: selectedWorkers.length
    };

  } catch (e) {
    Logger.log("Error in processImportF1Data: " + e.stack);
    return { success: false, message: e.message };
  }
}

/**
 * 2. ฟังก์ชันดึงข้อมูลประวัติการ Import พร้อมระบบ Pagination
 */
function getImportF1History(page, limit, filterDate) {
  page = page || 1;
  limit = limit || 6;
  
  try {
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
    
    if (!loggedInUsername) {
      throw new Error("หมดเวลาเชื่อมต่อ หรือผู้ใช้ไม่ได้ล็อกอิน");
    }

    const targetSS = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const logSheet = targetSS.getSheetByName("Datalog");
    
    if (!logSheet) {
      return { success: true, data: [], totalPages: 0, currentPage: 1, totalItems: 0 };
    }

    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) { 
      return { success: true, data: [], totalPages: 0, currentPage: 1, totalItems: 0 };
    }

    const data = logSheet.getRange(1, 1, lastRow, logSheet.getLastColumn()).getValues();
    let rows = data.slice(1); 
    rows.reverse(); 

    const timeZone = Session.getScriptTimeZone() || "GMT+7";

    // --- กรองวันที่ด้วยตัวแปร filterDate ---
    if (filterDate) {
      rows = rows.filter(row => {
        let dateStr = row[1];
        if (dateStr && (dateStr instanceof Date || !isNaN(new Date(dateStr).getTime()))) {
            dateStr = Utilities.formatDate(new Date(dateStr), timeZone, "yyyy-MM-dd");
        }
        return dateStr === filterDate;
      });
    }

    const totalItems = rows.length;
    const totalPages = Math.ceil(totalItems / limit) || 1;
    const startIndex = (page - 1) * limit;
    const endIndex = startIndex + limit;

    const paginatedRows = rows.slice(startIndex, endIndex).map(row => {
      let ts = row[0];
      if (ts && (ts instanceof Date || !isNaN(new Date(ts).getTime()))) {
         ts = Utilities.formatDate(new Date(ts), timeZone, "dd/MM/yyyy HH:mm:ss");
      } else { ts = "-"; }
      
      let dateStr = row[1];
      if (dateStr && (dateStr instanceof Date || !isNaN(new Date(dateStr).getTime()))) {
          dateStr = Utilities.formatDate(new Date(dateStr), timeZone, "dd/MM/yyyy");
      } else { dateStr = String(row[1] || ""); }

      return {
        timestamp: ts,
        date: dateStr,
        uploader: String(row[2] || "ไม่ระบุ"),
        filename: String(row[3] || "ไม่ระบุชื่อไฟล์"),
        count: row[4] || 0,
        summary: String(row[5] || "") // ดึงสรุปการจ่ายงานจากคอลัมน์ที่ 6
      };
    });

    return {
      success: true,
      data: paginatedRows,
      totalPages: totalPages,
      currentPage: page,
      totalItems: totalItems
    };

  } catch (e) {
    Logger.log("Error in getImportF1History: " + e.message);
    return { success: false, message: e.message };
  }
}

// =======================================================
// --- ส่วนหน้า Low.html (Task Management) ---
// =======================================================

/**
 * ดึงข้อมูลจากชีต Data เพื่อนำไปแสดงผลที่หน้า Low.html
 */
function getLowPageData() {
  try {
    // 1. ตรวจสอบผู้ใช้งาน
    const userInfo = getUpdaterInfo_(); // จะได้ค่าเช่น "12345 - สมชาย ใจดี"
    
    // 2. ดึงข้อมูลจากชีต
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('Low');
    
    if (!sheet) {
      return { success: false, message: "ไม่พบชีต 'Low' ในระบบ" };
    }

    const data = sheet.getDataRange().getDisplayValues(); // ใช้ getDisplayValues เพื่อรักษารูปแบบวันที่
    
    if (data.length <= 1) {
      return { success: true, data: [], currentUserInfo: userInfo };
    }

    const headers = data[0];
    const rows = data.slice(1).map(row => {
      // เติมช่องว่างให้ครบ 18 คอลัมน์ป้องกัน Error กรณีข้อมูลแหว่ง
      while(row.length < 18) row.push(""); 
      
      return {
        empId: String(row[0] || "").trim(),        // คอลัมน์ A (Index 0) : รหัสพนักงาน
        empName: String(row[1] || "").trim(),      // คอลัมน์ B (Index 1) : ชื่อพนักงาน
        highLow: String(row[2] || "").trim(),      // คอลัมน์ C (Index 2) : High/Low
        taskNo: String(row[3] || "").trim(),       // คอลัมน์ D (Index 3) : Task No.
        createBy: String(row[4] || "").trim(),     // คอลัมน์ E (Index 4) : Create By
        branchName: String(row[5] || "").trim(),   // คอลัมน์ F (Index 5) : ชื่อสาขา
        detail: String(row[6] || "").trim(),       // คอลัมน์ G (Index 6) : Detail
        workGroup: String(row[7] || "").trim(),    // คอลัมน์ H (Index 7) : กลุ่มงาน
        catalogSub: String(row[8] || "").trim(),   // คอลัมน์ I (Index 8) : Services Catalog Sub
        issue: String(row[9] || "").trim(),        // คอลัมน์ J (Index 9) : Issue
        reportDate: String(row[10] || "").trim(),  // คอลัมน์ K (Index 10) : วันที่แจ้งงาน
        createTime: String(row[11] || "").trim(),  // คอลัมน์ L (Index 11) : Create Time
        dueDate: String(row[12] || "").trim(),     // คอลัมน์ M (Index 12) : กำหนดวันแล้วเสร็จ
        dueTime: String(row[13] || "").trim(),     // คอลัมน์ N (Index 13) : Due Time
        province: String(row[14] || "").trim(),    // คอลัมน์ O (Index 14) : Province
        phone: String(row[15] || "").trim(),       // คอลัมน์ P (Index 15) : เบอร์โทรสาขา
        openDate: String(row[16] || "").trim(),    // คอลัมน์ Q (Index 16) : วันเปิดสาขา
        assignee: String(row[17] || "").trim(),    // คอลัมน์ R (Index 17) : Assignee
        rawData: row // เก็บ Array ดิบไว้สำหรับก๊อปปี้
      };
    });

// --- ส่วนที่เพิ่มใหม่: ดึงเวลาอัปเดตล่าสุดจากชีต Datalog ---
    let lastUpdateTime = "-";
    const logSheet = ss.getSheetByName('Datalog');
    if (logSheet && logSheet.getLastRow() > 1) {
      const lastTimestamp = logSheet.getRange(logSheet.getLastRow(), 1).getValue(); // ดึงจากคอลัมน์ A (Timestamp)
      if (lastTimestamp && (lastTimestamp instanceof Date || !isNaN(new Date(lastTimestamp).getTime()))) {
        lastUpdateTime = Utilities.formatDate(new Date(lastTimestamp), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");
      }
    }

    return { 
      success: true, 
      data: rows, 
      currentUserInfo: userInfo,
      lastUpdateTime: lastUpdateTime // ส่งเวลากลับไปให้หน้าเว็บ
    };

  } catch (e) {
    Logger.log("Error in getLowPageData: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ฟังก์ชันสำหรับปุ่ม "Add F1" 
 * คัดลอกข้อมูลแถวนั้นไปยังชีตรวม (Shared Sheet) พร้อมดึงพิกัดจากชีต Branch
 */
function copyTaskToUserSheet(rawData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    const userInfo = getUpdaterInfo_();
    if (userInfo === 'Unknown User' || !userInfo) {
      throw new Error("หมดเวลาเชื่อมต่อ กรุณารีเฟรชเพื่อล็อกอินใหม่");
    }

    const targetSheetName = "Assigned_Tasks";
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    let userSheet = ss.getSheetByName(targetSheetName);
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");

    // ถ้ายังไม่มีชีตรวม ให้สร้างใหม่และตั้งค่า Header
    if (!userSheet) {
      userSheet = ss.insertSheet(targetSheetName);
      const dataSheet = ss.getSheetByName('High') || ss.getSheetByName('Low'); 
      let headers = dataSheet ? dataSheet.getRange(1, 1, 1, 19).getValues()[0] : new Array(19).fill("");
      
      headers[17] = "AssigZone"; 
      if (!headers[18]) headers[18] = "Pull Time";
      headers[19] = "Latitude";  // เพิ่ม Header ใหม่
      headers[20] = "Longitude"; // เพิ่ม Header ใหม่
      headers[21] = "Zone";      // เพิ่ม Header ใหม่

      userSheet.getRange(1, 1, 1, 22).setValues([headers]);
      userSheet.getRange(1, 1, 1, 22).setBackground("#e0e7ff").setFontWeight("bold");
      userSheet.setFrozenRows(1);
    } else {
      // ตรวจสอบว่ามีคอลัมน์พิกัดและ Zone หรือยัง ถ้ายังให้สร้างเพิ่ม
      const lastCol = userSheet.getLastColumn();
      if (lastCol < 22) {
        userSheet.getRange(1, 20, 1, 3).setValues([["Latitude", "Longitude", "Zone"]]).setBackground("#e0e7ff").setFontWeight("bold");
      }
    }

    const taskNo = rawData[3];
    if (taskNo) {
      const existingData = userSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][3] === taskNo) {
          throw new Error(`งาน Task No: ${taskNo} มีผู้ดึงงานไปไว้ในชีตเรียบร้อยแล้ว`);
        }
      }
    }

    // --- ระบบดึงพิกัด (Latitude / Longitude) จากชีต Branch ---
    let latitude = "";
    let longitude = "";
    const branchSheet = ss.getSheetByName('Branch');
    
    if (branchSheet) {
      const branchText = String(rawData[5] || "").trim(); // คอลัมน์ F (Index 5)
      const match = branchText.match(/^(\d+)/); // ดึงเฉพาะตัวเลขด้านหน้า เช่น "70-ด่าน..." จะได้ "70"
      
      if (match) {
        const branchCodeToFind = match[1].padStart(4, '0'); // แปลงเป็น 4 หลัก (เช่น "0070")
        const branchData = branchSheet.getDataRange().getValues();
        
        for (let i = 1; i < branchData.length; i++) {
          let rowCode = String(branchData[i][0] || "").trim(); // คอลัมน์ A (Index 0)
          if (/^\d+$/.test(rowCode)) rowCode = rowCode.padStart(4, '0'); // จัดฟอร์แมตเผื่อข้อมูลในชีตเป็น "70"
          
          if (rowCode === branchCodeToFind) {
            latitude = branchData[i][11] || "";  // คอลัมน์ L (Index 11)
            longitude = branchData[i][12] || ""; // คอลัมน์ M (Index 12)
            break;
          }
        }
      }
    }

    // --- ระบบดึงข้อมูล Zone จากชีต F1 โดยอ้างอิงจาก Task No. ---
    let zone = "";
    const f1Sheet = ss.getSheetByName('F1');
    if (f1Sheet && taskNo) {
      const taskNoToFind = String(taskNo).trim();
      const f1Data = f1Sheet.getDataRange().getValues();
      for (let i = 1; i < f1Data.length; i++) {
        if (String(f1Data[i][0] || "").trim() === taskNoToFind) { // คอลัมน์ A (Index 0)
          zone = f1Data[i][52] || ""; // คอลัมน์ BA คือคอลัมน์ที่ 53 (Index 52)
          break;
        }
      }
    }

    // --- ประทับตราผู้ดึงงาน เวลา และเพิ่มพิกัด, Zone ---
    let newRow = [...rawData];
    while (newRow.length < 19) newRow.push(""); 
    newRow[17] = userInfo;  // R
    newRow[18] = timestamp; // S
    newRow[19] = latitude;  // T (Latitude)
    newRow[20] = longitude; // U (Longitude)
    newRow[21] = zone;      // V (Zone)

    userSheet.appendRow(newRow);

    return { 
      success: true, 
      message: `นำงานเข้าชีตรวม <b>"${targetSheetName}"</b><br><span class="text-sm">บันทึกงาน ดึงพิกัด และ Zone เรียบร้อยแล้ว</span>` 
    };

  } catch (e) {
    Logger.log("Error in copyTaskToUserSheet: " + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ===============================================
// ฟังก์ชันบันทึกข้อมูลหลายรายการพร้อมกัน (Bulk Action) ลงชีตรวม พร้อมดึงพิกัด
// ===============================================
function copyMultipleTasksToUserSheet(rawDataArray) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); 
  
  try {
    const userInfo = getUpdaterInfo_();
    if (userInfo === 'Unknown User' || !userInfo) {
      throw new Error("หมดเวลาเชื่อมต่อ กรุณารีเฟรชเพื่อล็อกอินใหม่");
    }

    const targetSheetName = "Assigned_Tasks";
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    let userSheet = ss.getSheetByName(targetSheetName);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");

    if (!userSheet) {
      userSheet = ss.insertSheet(targetSheetName);
      const dataSheet = ss.getSheetByName('High') || ss.getSheetByName('Low');
      let headers = dataSheet ? dataSheet.getRange(1, 1, 1, 19).getValues()[0] : new Array(19).fill("");
      
      headers[17] = "AssigZone";
      if (!headers[18]) headers[18] = "Pull Time";
      headers[19] = "Latitude";
      headers[20] = "Longitude";
      headers[21] = "Zone";
      
      userSheet.getRange(1, 1, 1, 22).setValues([headers]);
      userSheet.getRange(1, 1, 1, 22).setBackground("#e0e7ff").setFontWeight("bold");
      userSheet.setFrozenRows(1);
    } else {
      const lastCol = userSheet.getLastColumn();
      if (lastCol < 22) {
        userSheet.getRange(1, 20, 1, 3).setValues([["Latitude", "Longitude", "Zone"]]).setBackground("#e0e7ff").setFontWeight("bold");
      }
    }

    // --- โหลดข้อมูลพิกัดสาขาทั้งหมดเตรียมไว้ใน Map (ช่วยให้ทำงานไวขึ้นตอนเลือกหลายรายการ) ---
    const branchSheet = ss.getSheetByName('Branch');
    const branchMap = new Map();
    if (branchSheet) {
      const branchData = branchSheet.getDataRange().getValues();
      for (let i = 1; i < branchData.length; i++) {
        let code = String(branchData[i][0] || "").trim();
        if (/^\d+$/.test(code)) code = code.padStart(4, '0');
        branchMap.set(code, {
          lat: branchData[i][11] || "",
          lon: branchData[i][12] || ""
        });
      }
    }

    // --- โหลดข้อมูล Zone จากชีต F1 เตรียมไว้ใน Map ---
    const f1Sheet = ss.getSheetByName('F1');
    const f1Map = new Map();
    if (f1Sheet) {
      const f1Data = f1Sheet.getDataRange().getValues();
      for (let i = 1; i < f1Data.length; i++) {
        let tNo = String(f1Data[i][0] || "").trim(); // คอลัมน์ A (Index 0)
        if (tNo) {
          f1Map.set(tNo, f1Data[i][52] || ""); // คอลัมน์ BA (Index 52)
        }
      }
    }

    const existingData = userSheet.getDataRange().getValues();
    const existingTaskNos = new Set(existingData.map(row => row[3]).filter(String));
    
    let rowsToAppend = [];
    let duplicateCount = 0;

    for (let i = 0; i < rawDataArray.length; i++) {
      let rawData = rawDataArray[i];
      let taskNo = rawData[3];
      
      if (taskNo && existingTaskNos.has(taskNo)) {
        duplicateCount++;
        continue; 
      }

      // ตรวจสอบรหัสสาขาเพื่อ Map พิกัด
      let latitude = "";
      let longitude = "";
      const branchText = String(rawData[5] || "").trim();
      const match = branchText.match(/^(\d+)/);
      if (match) {
        const branchCodeToFind = match[1].padStart(4, '0');
        if (branchMap.has(branchCodeToFind)) {
          latitude = branchMap.get(branchCodeToFind).lat;
          longitude = branchMap.get(branchCodeToFind).lon;
        }
      }

      // ตรวจสอบ Task No. เพื่อ Map Zone จากชีต F1
      let zone = "";
      let taskNoStr = String(taskNo || "").trim();
      if (taskNoStr && f1Map.has(taskNoStr)) {
        zone = f1Map.get(taskNoStr);
      }

      let newRow = [...rawData];
      while (newRow.length < 19) newRow.push(""); 
      newRow[17] = userInfo;  
      newRow[18] = timestamp; 
      newRow[19] = latitude;  
      newRow[20] = longitude; 
      newRow[21] = zone;
      
      rowsToAppend.push(newRow);
      existingTaskNos.add(taskNo); 
    }

    if (rowsToAppend.length > 0) {
      // กำหนด 22 เพราะมี 22 คอลัมน์ (นับถึง V)
      userSheet.getRange(userSheet.getLastRow() + 1, 1, rowsToAppend.length, 22).setValues(rowsToAppend);
    }

    let resultMsg = `นำงานเข้าชีต <b>"${targetSheetName}"</b> สำเร็จ <span class="text-emerald-600 font-bold">${rowsToAppend.length}</span> รายการ`;
    if (duplicateCount > 0) {
      resultMsg += `<br><span class="text-[12px] text-amber-500 mt-1 block"><i class="fa-solid fa-triangle-exclamation"></i> ข้ามงานที่ผู้อื่นดึงไปแล้ว ${duplicateCount} รายการ</span>`;
    }

    return { 
      success: true, 
      message: resultMsg,
      addedCount: rowsToAppend.length
    };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getHighPageData() {
  try {
    const userInfo = getUpdaterInfo_(); 
    
    // ดึงข้อมูลจากชีต 'High'
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('High');
    
    if (!sheet) {
      return { success: false, message: "ไม่พบชีต 'High' ในระบบ" };
    }

    const data = sheet.getDataRange().getDisplayValues(); 
    
    if (data.length <= 1) {
      return { success: true, data: [], currentUserInfo: userInfo };
    }

    const headers = data[0];
    const rows = data.slice(1).map(row => {
      while(row.length < 18) row.push(""); 
      
      return {
        empId: String(row[0] || "").trim(),        
        empName: String(row[1] || "").trim(),      
        highLow: String(row[2] || "").trim(),      
        taskNo: String(row[3] || "").trim(),       
        createBy: String(row[4] || "").trim(),     
        branchName: String(row[5] || "").trim(),   
        detail: String(row[6] || "").trim(),       
        workGroup: String(row[7] || "").trim(),    
        catalogSub: String(row[8] || "").trim(),   
        issue: String(row[9] || "").trim(),        
        reportDate: String(row[10] || "").trim(),  
        createTime: String(row[11] || "").trim(),  
        dueDate: String(row[12] || "").trim(),     
        dueTime: String(row[13] || "").trim(),     
        province: String(row[14] || "").trim(),    
        phone: String(row[15] || "").trim(),       
        openDate: String(row[16] || "").trim(),    
        assignee: String(row[17] || "").trim(),    
        rawData: row 
      };
    });

// --- ส่วนที่เพิ่มใหม่: ดึงเวลาอัปเดตล่าสุดจากชีต Datalog ---
    let lastUpdateTime = "-";
    const logSheet = ss.getSheetByName('Datalog');
    if (logSheet && logSheet.getLastRow() > 1) {
      const lastTimestamp = logSheet.getRange(logSheet.getLastRow(), 1).getValue(); // ดึงจากคอลัมน์ A (Timestamp)
      if (lastTimestamp && (lastTimestamp instanceof Date || !isNaN(new Date(lastTimestamp).getTime()))) {
        lastUpdateTime = Utilities.formatDate(new Date(lastTimestamp), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");
      }
    }

    return { 
      success: true, 
      data: rows, 
      currentUserInfo: userInfo,
      lastUpdateTime: lastUpdateTime // ส่งเวลากลับไปให้หน้าเว็บ
    };

  } catch (e) {
    Logger.log("Error in getHighPageData: " + e.message);
    return { success: false, message: e.message };
  }
}

// =======================================================
// --- ส่วนหน้า Work.html (Overview Low & High) ---
// =======================================================

function getWorkPageData() {
  try {
    const userInfo = getUpdaterInfo_(); 
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    
    // 1. หาเวลาอัปเดตล่าสุดของ Sheet "DATA" (อ้างอิงจากชีต Datalog)
    let dataLastUpdated = "ไม่พบประวัติการอัปเดต";
    const logSheet = ss.getSheetByName('Datalog');
    if (logSheet && logSheet.getLastRow() > 1) {
      const lastRow = logSheet.getLastRow();
      const lastTimestamp = logSheet.getRange(lastRow, 1).getValue(); // คอลัมน์ A (Timestamp)
      if (lastTimestamp && (lastTimestamp instanceof Date || !isNaN(new Date(lastTimestamp).getTime()))) {
        dataLastUpdated = Utilities.formatDate(new Date(lastTimestamp), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");
      }
    }

    let combinedData = [];

    // 2. ดึงข้อมูลจากชีต "Low&High" โดยตรง
    const combinedSheet = ss.getSheetByName('Low&High');
    if (combinedSheet && combinedSheet.getLastRow() > 1) {
      const combinedSheetData = combinedSheet.getDataRange().getDisplayValues().slice(1);
      combinedSheetData.forEach(row => {
        while(row.length < 19) row.push(""); 
        combinedData.push({
          empId: String(row[0] || "").trim(),
          empName: String(row[1] || "").trim(),
          highLow: String(row[2] || "").trim(), // คอลัมน์ C (Index 2) : High/Low
          taskNo: String(row[3] || "").trim(),
          createBy: String(row[4] || "").trim(),
          branchName: String(row[5] || "").trim(),
          detail: String(row[6] || "").trim(),
          workGroup: String(row[7] || "").trim(),
          catalogSub: String(row[8] || "").trim(),
          issue: String(row[9] || "").trim(),
          reportDate: String(row[10] || "").trim(),
          createTime: String(row[11] || "").trim(),
          dueDate: String(row[12] || "").trim(),
          dueTime: String(row[13] || "").trim(),
          phone: String(row[15] || "").trim(),
          assignee: String(row[17] || "").trim(), // คอลัมน์ R (Index 17) : Assignee
          statusColS: String(row[18] || "").trim(), // คอลัมน์ S (Index 18) : F1&Sheet (Status)
          sourceSheet: 'Low&High',
          rawData: row
        });
      });
    }

    return { 
      success: true, 
      data: combinedData, 
      dataLastUpdated: dataLastUpdated,
      currentUserInfo: userInfo 
    };

  } catch (e) {
    Logger.log("Error in getWorkPageData: " + e.message);
    return { success: false, message: e.message };
  }
}

// =======================================================
// --- ส่วนฟังก์ชันสำหรับระบบลงเวลา Checkinsenior.html ---
// =======================================================

/**
 * ดึงข้อมูลประวัติการลงเวลาและชื่อผู้ใช้ปัจจุบันส่งให้ Frontend
 */
function getSeniorHistoryData(userToken) {
  try {
    const cache = CacheService.getUserCache();
    const loggedInUsername = cache.get('loggedInUser');
    
    if (!loggedInUsername) {
      return { success: false, message: 'กรุณาล็อกอินใหม่ เซสชันหมดอายุ' };
    }

    const userFullName = getUpdaterInfo_();
    
    if (userFullName === 'Unknown User' || userFullName === 'System Error') {
        return { success: false, message: 'ไม่สามารถระบุตัวตนได้ กรุณาล็อกอินใหม่' };
    }

    const sheet = getSheetByName_(SENIOR_CHECKIN_SHEET_NAME);
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Date', 'TimeIn', 'TimeOut', 'Recorder', 'Shift', 
        'RemarkIn', 'RemarkOut', 'LatIn', 'LonIn', 'LatOut', 'LonOut', 
        'PicIn', 'PicOut', 'IsLateIn', 'IsEarlyOut'
      ]);
    }

    let historyData = [];
    if (sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getDisplayValues();
      const headers = data[0];
      
      const col = {
        date: headers.indexOf('Date') !== -1 ? headers.indexOf('Date') : 1,
        recorder: headers.indexOf('Recorder') !== -1 ? headers.indexOf('Recorder') : 4,
        shift: headers.indexOf('Shift') !== -1 ? headers.indexOf('Shift') : 5,
        timeIn: headers.indexOf('TimeIn') !== -1 ? headers.indexOf('TimeIn') : 2,
        timeOut: headers.indexOf('TimeOut') !== -1 ? headers.indexOf('TimeOut') : 3,
        remarkIn: headers.indexOf('RemarkIn') !== -1 ? headers.indexOf('RemarkIn') : 6,
        remarkOut: headers.indexOf('RemarkOut') !== -1 ? headers.indexOf('RemarkOut') : 7,
        latIn: headers.indexOf('LatIn') !== -1 ? headers.indexOf('LatIn') : 8,
        lonIn: headers.indexOf('LonIn') !== -1 ? headers.indexOf('LonIn') : 9,
        latOut: headers.indexOf('LatOut') !== -1 ? headers.indexOf('LatOut') : 10,
        lonOut: headers.indexOf('LonOut') !== -1 ? headers.indexOf('LonOut') : 11,
        isLateIn: headers.indexOf('IsLateIn') !== -1 ? headers.indexOf('IsLateIn') : 14,
        isEarlyOut: headers.indexOf('IsEarlyOut') !== -1 ? headers.indexOf('IsEarlyOut') : 15,
        picIn: 12, // <--- บังคับดึงรูป Check In จากคอลัมน์ M (Index 12)
        picOut: 13 // <--- บังคับดึงรูป Check Out จากคอลัมน์ N (Index 13)
      };

      const startRow = Math.max(1, data.length - 200);
      for (let i = data.length - 1; i >= startRow; i--) {
        historyData.push({
          date: data[i][col.date] || '-',
          recorder: data[i][col.recorder],
          shift: data[i][col.shift],
          timeIn: data[i][col.timeIn] || '-',
          timeOut: data[i][col.timeOut] || 'ยังไม่ Check Out',
          remarkIn: data[i][col.remarkIn] || '', 
          remarkOut: data[i][col.remarkOut] || '', 
          latIn: data[i][col.latIn] || '',       
          lonIn: data[i][col.lonIn] || '',       
          latOut: data[i][col.latOut] || '',     
          lonOut: data[i][col.lonOut] || '',     
          isLateIn: String(data[i][col.isLateIn]).toUpperCase() === 'TRUE',
          isEarlyOut: String(data[i][col.isEarlyOut]).toUpperCase() === 'TRUE',
          picIn: data[i][col.picIn],
          picOut: data[i][col.picOut]
        });
      }
    }

    return { 
      success: true, 
      userFullName: userFullName,
      data: historyData 
    };

  } catch (e) {
    Logger.log(`Error in getSeniorHistoryData: ${e.message}`);
    return { success: false, message: e.message };
  }
}


/**
 * บันทึกข้อมูล Check In
 */
function processCheckInSenior(formObject) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    const sheet = getSheetByName_(SENIOR_CHECKIN_SHEET_NAME);
    const now = new Date();
    const tz = Session.getScriptTimeZone() || "GMT+7";
    const dateStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");
    const timeStr = Utilities.formatDate(now, tz, "HH:mm");

    let isLate = false;
    if (formObject.shift.includes("06:00") && timeStr > "06:00") isLate = true;
    if (formObject.shift.includes("13:00") && timeStr > "13:00") isLate = true;
    if (formObject.shift.includes("08:00") && timeStr > "08:00") isLate = true;

    // อัปโหลดรูปภาพ โดยใช้ ID โฟลเดอร์ใหม่ที่คุณให้มา
    let picUrl = "";
    if (formObject.file && formObject.file.fileData) {
      picUrl = saveImageToDrive(formObject.file.fileData, formObject.file.fileName, SENIOR_CHECKIN_FOLDER_ID);
      // ถ้าไม่ได้ URL กลับมา แปลว่าอัปโหลดลง Drive ไม่สำเร็จ
      if (!picUrl) throw new Error("ไม่สามารถบันทึกรูปลง Google Drive ได้ กรุณาลองอีกครั้ง");
    }

    const newRow = [
      now,                  // A (1): Timestamp
      dateStr,              // B (2): Date
      timeStr,              // C (3): TimeIn
      "",                   // D (4): TimeOut
      formObject.recorder,  // E (5): Recorder
      formObject.shift,     // F (6): Shift
      formObject.remarkIn || "", // G (7): RemarkIn
      "",                   // H (8): RemarkOut
      formObject.lat || "", // I (9): LatIn
      formObject.lon || "", // J (10): LonIn
      "",                   // K (11): LatOut
      "",                   // L (12): LonOut
      picUrl,               // M (13): PicIn (รูปภาพ Check In)
      "",                   // N (14): PicOut
      isLate,               // O (15): IsLateIn
      false                 // P (16): IsEarlyOut
    ];

    sheet.appendRow(newRow);
    logUserActivity('CheckIn Senior', `ผู้ใช้ ${formObject.recorder} Check In เวลา ${timeStr}`);

    return { success: true };

  } catch (e) {
    Logger.log(`Error in processCheckInSenior: ${e.stack}`);
    throw new Error(`เกิดข้อผิดพลาดในการบันทึกข้อมูล: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * โหลดข้อมูล Check In ล่าสุดที่ยังไม่ได้ Check Out (ข้ามวันได้)
 */
function loadSeniorCheckInToday(recorder) {
  try {
    const sheet = getSheetByName_(SENIOR_CHECKIN_SHEET_NAME);
    if (sheet.getLastRow() <= 1) return { found: false };

    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];

    const dateCol = headers.indexOf('Date');
    const recorderCol = headers.indexOf('Recorder');
    const timeOutCol = headers.indexOf('TimeOut');
    const shiftCol = headers.indexOf('Shift');
    const timeInCol = headers.indexOf('TimeIn');

    // ค้นหาย้อนหลังจากล่างขึ้นบน จะได้รายการล่าสุดที่ยังไม่ได้ Check Out (ถอดการจำกัดเฉพาะ "วันนี้" ออกแล้ว)
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][recorderCol]).trim() === String(recorder).trim() && data[i][timeOutCol] === "") {
        return {
          found: true,
          rowIndex: i + 1, // +1 เพราะ Index อาร์เรย์เริ่มที่ 0 แต่ชีตเริ่มที่ 1
          shift: data[i][shiftCol],
          timeIn: data[i][timeInCol],
          dateIn: data[i][dateCol] // เพิ่ม Date ให้ UI รู้ว่ากำลัง Check out ของวันไหน
        };
      }
    }

    return { found: false };

  } catch (e) {
    Logger.log(`Error in loadSeniorCheckInToday: ${e.message}`);
    throw new Error(`ไม่สามารถดึงข้อมูลได้: ${e.message}`);
  }
}

/**
 * บันทึกข้อมูล Check Out (รองรับการข้ามวันอย่างปลอดภัย)
 */
function processCheckOutSenior(formObject) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getSheetByName_(SENIOR_CHECKIN_SHEET_NAME);
    const rowId = parseInt(formObject.rowId);
    
    if (!rowId || rowId <= 1) throw new Error("ไม่พบ ID ของแถวที่ต้องการอัปเดต");

    const now = new Date();
    const tz = Session.getScriptTimeZone() || "GMT+7";
    const timeStr = Utilities.formatDate(now, tz, "HH:mm");
    const todayStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");

    const dateIn = sheet.getRange(rowId, 2).getDisplayValue(); // ดึงวันที่เช็คอินมาตรวจสอบข้ามวัน
    const shift = sheet.getRange(rowId, 6).getValue(); 
    
    let isEarly = false;
    // ป้องกันบัค isEarly เมื่อทำ OT ข้ามวัน (ถ้าข้ามวัน ถือว่าทำเลยเวลา ไม่ใช่ออกก่อน)
    if (dateIn === todayStr) {
        if (shift.includes("กะเช้า") && timeStr < "14:30") isEarly = true;
        if (shift.includes("กะบ่าย") && timeStr < "23:59") isEarly = true;
        if (shift.includes("กะปกติ") && timeStr < "18:00") isEarly = true;
    }

    // อัปโหลดรูปภาพ
    let picUrl = "";
    if (formObject.file && formObject.file.fileData) {
      picUrl = saveImageToDrive(formObject.file.fileData, formObject.file.fileName, SENIOR_CHECKIN_FOLDER_ID);
      if (!picUrl) throw new Error("ไม่สามารถบันทึกรูปลง Google Drive ได้ กรุณาลองอีกครั้ง");
    }

    sheet.getRange(rowId, 4).setValue(timeStr);                    // D (4): TimeOut
    sheet.getRange(rowId, 8).setValue(formObject.remarkOut || ""); // H (8): RemarkOut
    sheet.getRange(rowId, 11).setValue(formObject.lat || "");      // K (11): LatOut
    sheet.getRange(rowId, 12).setValue(formObject.lon || "");      // L (12): LonOut
    sheet.getRange(rowId, 14).setValue(picUrl);                    // N (14): PicOut (รูปภาพ Check Out)
    sheet.getRange(rowId, 16).setValue(isEarly);                   // P (16): IsEarlyOut

    logUserActivity('CheckOut Senior', `ผู้ใช้ ${formObject.recorder} Check Out เวลา ${timeStr} ของรายการวันที่ ${dateIn}`);

    return { success: true };

  } catch (e) {
    Logger.log(`Error in processCheckOutSenior: ${e.stack}`);
    throw new Error(`เกิดข้อผิดพลาดในการบันทึกข้อมูล Check Out: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

function getAssignedTasksData() {
  try {
    const userInfo = getUpdaterInfo_();
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('Assigned_Tasks');
    
    if (!sheet) return { success: false, message: "ไม่พบชีต 'Assigned_Tasks'" };

    // ดึง Header เพื่อหา Index คอลัมน์แบบไดนามิก
    const headers = getSheetHeaders_(sheet);
    
    // ตรวจสอบและสร้างคอลัมน์สำหรับ Workflow ถ้าย้อนยังไม่มี (เพิ่มคอลัมน์แยกตามที่ต้องการ)
    const requiredCols = [
      'WF_Status', 'WF_Log', 'WF_Vendor', 'WF_VendorType', 'WF_Cost', 'WF_PO', 'WF_Dates',
      'WF_Mat_Cost', 'WF_Labor_Cost', 'WF_Travel_Cost', 'WF_Total_Cost',
      'WF_Cancel_Reason', 'WF_Entry_Date', 'WF_Complete_Date', 'WF_Cancel_Date',
      'WF_Last_Update_Date', 'WF_Last_Update_Time', 'WF_Last_Remark'
    ];
    
    let needsUpdate = false;
    requiredCols.forEach(col => {
      if (headers.indexOf(col) === -1) {
        sheet.getRange(1, headers.length + 1).setValue(col);
        headers.push(col);
        needsUpdate = true;
      }
    });
    if(needsUpdate) sheet.getRange(1, 1, 1, headers.length).setBackground("#374151").setFontColor("white");

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, data: [], currentUserInfo: userInfo };

    // Map ข้อมูลส่งให้ Frontend
    const rows = data.slice(1).map((row, index) => {
      return {
        rowId: index + 2,
        createBy: row[4] || "-",
        taskNo: row[3] || "-",
        branchName: row[5] || "-",
        detail: row[6] || "-",          // เพิ่ม: ดึง Detail
        workGroup: row[7] || "-",
        issue: row[9] || "-",
        highLow: row[2] || "-",
        reportDate: row[10] || "-",
        dueDate: row[12] || "-",        // เพิ่ม: ดึง กำหนดวันแล้วเสร็จ (Due Date)
        phone: row[15] || "-",          // เพิ่ม: ดึง เบอร์โทรสาขา
        assignZone: row[17] || "-",
        pullTime: row[18] || "-",
        branchCode: String(row[5] || "").match(/^(\d+)/) ? String(row[5]).match(/^(\d+)/)[1].padStart(4, '0') : "",
        
        // ข้อมูล Workflow
        wfStatus: row[headers.indexOf('WF_Status')] || "_00_รอจ่ายงาน",
        wfLog: row[headers.indexOf('WF_Log')] || "[]",
        wfVendor: row[headers.indexOf('WF_Vendor')] || "",
        wfVendorType: row[headers.indexOf('WF_VendorType')] || "",
        wfCost: row[headers.indexOf('WF_Cost')] || "{}",
        wfPO: row[headers.indexOf('WF_PO')] || "",
        wfDates: row[headers.indexOf('WF_Dates')] || "{}"
      };
    });

    return { success: true, data: rows.reverse(), currentUserInfo: userInfo };
  } catch (e) {
    Logger.log("Error in getAssignedTasksData: " + e.message);
    return { success: false, message: e.message };
  }
}


/**
 * ดึงข้อมูลผู้รับเหมา (Vendor) โครงสร้าง
 */
function getStructuralVendors() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('vendor'); // ต้องมีชีตชื่อ vendor
    if (!sheet) return { success: false, data: [] };
    
    const data = sheet.getDataRange().getValues().slice(1);
    const vendors = data.map(row => ({
      province: String(row[0]).trim(),
      workType: String(row[1]).trim(),
      vendorName: String(row[2]).trim(),
      vendorCode: String(row[3]).trim()
    })).filter(v => v.vendorName);
    
    return { success: true, data: vendors };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * ดึงข้อมูลช่างเครื่องเย็นประจำสาขา
 */
function getCoolingVendors(branchCode) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('โครงสร้างสาขา'); 
    if (!sheet) return { success: false, data: [] };
    
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      let code = String(data[i][4]).trim(); // คอลัมน์ E (Index 4)
      if(/^\d+$/.test(code)) code = code.padStart(4, '0');
      
      if(code === branchCode) {
        return { 
          success: true, 
          pm: data[i][29] || "-", // คอลัมน์ AD (Index 29)
          turnkey: data[i][30] || "-" // คอลัมน์ AE (Index 30)
        };
      }
    }
    return { success: true, pm: "ไม่พบข้อมูล", turnkey: "ไม่พบข้อมูล" };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * บันทึกการเปลี่ยนสถานะ Workflow กลับลงชีต Assigned_Tasks
 */
function updateTaskWorkflow(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const { rowId, status, newLogEntry, updates } = payload;
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('Assigned_Tasks');
    const headers = getSheetHeaders_(sheet);
    
    // ฟังก์ชันค้นหาหมายเลขคอลัมน์อย่างปลอดภัย
    const col = (name) => headers.indexOf(name) !== -1 ? headers.indexOf(name) + 1 : null;
    
    // อัปเดตสถานะ
    if (status && col('WF_Status')) sheet.getRange(rowId, col('WF_Status')).setValue(status);
    
    // อัปเดตประวัติการเปลี่ยนแปลง
    if (newLogEntry && col('WF_Log')) {
      let currentLog = sheet.getRange(rowId, col('WF_Log')).getValue();
      let logArr = [];
      try { logArr = JSON.parse(currentLog); } catch(e) {}
      if (!Array.isArray(logArr)) logArr = [];
      logArr.push(newLogEntry);
      sheet.getRange(rowId, col('WF_Log')).setValue(JSON.stringify(logArr));

      // บันทึก วันที่ เวลา และรายละเอียดล่าสุด ลงคอลัมน์แยก
      const dateTimeParts = newLogEntry.date.split(' ');
      const updateDate = dateTimeParts[0] || '';
      const updateTime = dateTimeParts[1] || '';
      
      if(col('WF_Last_Update_Date')) {
        sheet.getRange(rowId, col('WF_Last_Update_Date')).setValue(updateDate);
        if(col('WF_Last_Update_Time')) sheet.getRange(rowId, col('WF_Last_Update_Time')).setValue(updateTime);
        if(col('WF_Last_Remark')) sheet.getRange(rowId, col('WF_Last_Remark')).setValue(newLogEntry.remark || '');
      }
    }
    
    // อัปเดตฟิลด์อื่นๆ (Vendor, PO, ราคา, วันที่) แบบแยกคอลัมน์ให้ชัดเจน ป้องกัน Error คอลัมน์ไม่มี
    if (updates) {
      if(updates.vendor !== undefined && col('WF_Vendor')) sheet.getRange(rowId, col('WF_Vendor')).setValue(updates.vendor);
      if(updates.vendorType !== undefined && col('WF_VendorType')) sheet.getRange(rowId, col('WF_VendorType')).setValue(updates.vendorType);
      if(updates.po !== undefined && col('WF_PO')) sheet.getRange(rowId, col('WF_PO')).setValue(updates.po);
      
      if(updates.cancelReason !== undefined && col('WF_Cancel_Reason')) {
        sheet.getRange(rowId, col('WF_Cancel_Reason')).setValue(updates.cancelReason);
      }
      
      if(updates.cost !== undefined && col('WF_Cost')) {
        sheet.getRange(rowId, col('WF_Cost')).setValue(JSON.stringify(updates.cost));
        if(col('WF_Mat_Cost')) sheet.getRange(rowId, col('WF_Mat_Cost')).setValue(updates.cost.mat || 0);
        if(col('WF_Labor_Cost')) sheet.getRange(rowId, col('WF_Labor_Cost')).setValue(updates.cost.labor || 0);
        if(col('WF_Travel_Cost')) sheet.getRange(rowId, col('WF_Travel_Cost')).setValue(updates.cost.travel || 0);
        if(col('WF_Total_Cost')) sheet.getRange(rowId, col('WF_Total_Cost')).setValue(updates.cost.total || 0);
      }
      
      if(updates.dates && col('WF_Dates')) {
        let currentDates = sheet.getRange(rowId, col('WF_Dates')).getValue();
        let datesObj = {};
        try { datesObj = JSON.parse(currentDates); } catch(e) {}
        datesObj = { ...datesObj, ...updates.dates };
        sheet.getRange(rowId, col('WF_Dates')).setValue(JSON.stringify(datesObj));

        // แยกคอลัมน์วันที่
        if(col('WF_Entry_Date') && updates.dates.entryDate) sheet.getRange(rowId, col('WF_Entry_Date')).setValue(updates.dates.entryDate);
        if(col('WF_Complete_Date') && updates.dates.completeDate) sheet.getRange(rowId, col('WF_Complete_Date')).setValue(updates.dates.completeDate);
        if(col('WF_Cancel_Date') && updates.dates.cancelDate) sheet.getRange(rowId, col('WF_Cancel_Date')).setValue(updates.dates.cancelDate);
      }
    }

    return { success: true, message: "บันทึกข้อมูลเรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// =======================================================
// ส่วนที่เพิ่มใหม่สำหรับหน้ารายละเอียดทีมทำงาน (Low.html / High.html)
// =======================================================

/**
 * ดึงข้อมูลจากชีต "dropdown" สำหรับฟอร์ม
 */
function getDropdownDataForTasks() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('dropdown');
    
    if (!sheet) return { success: false, message: "ไม่พบชีต 'dropdown'" };
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { success: true, data: [] };

    let dropdownList = [];
    for (let i = 1; i < data.length; i++) {
      dropdownList.push({
        repairType: String(data[i][0] || "").trim(),     // Col A
        subRepairType: String(data[i][1] || "").trim(),  // Col B
        forwardStatus: String(data[i][3] || "").trim(),  // *** เปลี่ยนเป็นดึงจาก Col D (Index 3)
        typeAsset: String(data[i][6] || "").trim()       // Col G
      });
    }
    return { success: true, data: dropdownList };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * ฟังก์ชันสำหรับบันทึกข้อมูลจากฟอร์มทีมทำงาน
 * (คุณสามารถนำ payload ไปบันทึกลงชีตที่คุณต้องการต่อได้เลย)
 */
function saveTeamAction(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    let assignedSheet = ss.getSheetByName('Assigned_Tasks');
    
    if (!assignedSheet) {
      throw new Error("ไม่พบ Sheet ชื่อ Assigned_Tasks โปรดตรวจสอบ");
    }

    // จัดเรียง Array ข้อมูล โดยนำข้อมูลใหม่ขึ้นก่อน แล้วตามด้วย เวลาบันทึก และข้อมูลจากฟอร์ม
    const newRow = [
      payload.highLow,            // 1. High/Low
      payload.createBy,           // 2. Create By
      payload.branchName,         // 3. ชื่อสาขา
      payload.catalogSub,         // 4. Services Catalog Sub
      payload.detail,             // 5. Detail
      payload.issue,              // 6. Issue
      payload.reportDate,         // 7. วันที่แจ้งงาน
      payload.createTime,         // 8. Create Time
      payload.dueDate,            // 9. กำหนดวันแล้วเสร็จ
      payload.dueTime,            // 10. Due Time
      payload.province,           // 11. Province
      payload.phone,              // 12. เบอร์โทรสาขา
      payload.openDate,           // 13. วันเปิดสาขา
      
      payload.actionTime,         // 14. เวลาบันทึก (เวลาดำเนินการ)
      payload.user,               // 15. ผู้บันทึก (ผู้รับงาน)
      payload.taskNo,             // 16. รหัสงาน
      payload.branchContact,      // 17. ผู้รับเรื่องสาขา
      payload.repairType,         // 18. ประเภทงานซ่อม
      payload.subRepairType,      // 19. ประเภทงานย่อย
      payload.assetType,          // 20. Type Asset
      payload.symptom,            // 21. อาการ
      payload.teamStatus,         // 22. สถานะทีม
      payload.actionResult,       // 23. ผลดำเนินการ
      payload.resolution,         // 24. การแก้ไข
      payload.forwardTeam,        // 25. ทีมส่งต่อ
      payload.forwardStatus,      // 26. สถานะส่งต่อ
      payload.remark,              // 27. หมายเหตุ
      payload.assetId             // 28. ข้อมูล Asset (คอลัมน์ AB) <--- เพิ่มบรรทัดนี้ต่อท้าย
    ];

    assignedSheet.appendRow(newRow);

    return { 
      success: true, 
      message: `บันทึกข้อมูลงาน <b>${payload.taskNo}</b> ลง Sheet Assigned_Tasks เรียบร้อยแล้ว` 
    };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}
/**
 * [อัปเดตใหม่ เพื่อโหลดข้อมูลมารอตั้งแต่ต้น]
 * อ่านข้อมูลประวัติจาก Assigned_Tasks แบบกวาดมาทั้งหมดแล้วจัดกลุ่มตามสาขา (Pre-load)
 */
function getAllBranchHistoryMap() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sheet = ss.getSheetByName('Assigned_Tasks');
    if (!sheet) return { success: false, message: "ไม่พบชีต Assigned_Tasks" };

    const data = sheet.getDataRange().getDisplayValues();
    let historyMap = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const branchName = String(row[2] || "").trim(); // คอลัมน์ C คือสาขา
      
      if (!branchName || branchName === "-") continue;

      if (!historyMap[branchName]) {
        historyMap[branchName] = { count: 0, data: [] };
      }

      // ดึงข้อมูลประวัติ (จัดเรียงตรงตาม Index ใน Assigned_Tasks)
      historyMap[branchName].data.push({
        date: String(row[6] || "") + " " + String(row[7] || ""), // วันที่แจ้งงาน G, H
        catalog: String(row[3] || ""),                           // Catalog D
        detail: String(row[4] || ""),                            // Detail E
        issue: String(row[5] || ""),                             // Issue F
        taskNo: String(row[15] || "")                            // Task No (รหัสงาน) คอลัมน์ P (Index 15)
      });
      historyMap[branchName].count++;
    }

    // เรียงประวัติของแต่ละสาขาให้โชว์รายการใหม่ล่าสุดขึ้นก่อน
    for (let branch in historyMap) {
       historyMap[branch].data.reverse();
    }

    return { success: true, data: historyMap };

  } catch(e) {
    Logger.log("Error in getAllBranchHistoryMap: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ให้นำฟังก์ชันนี้ไปวางต่อท้ายสุดในไฟล์ Code.gs ของคุณครับ
 * ฟังก์ชันนี้ใช้สำหรับดึงรายชื่อ Senior ที่เข้ากะ (Check-In) แต่ยังไม่ได้ (Check-Out)
 */
function getActiveSeniors() {
  try {
    const currentUserInfo = getUpdaterInfo_(); 
    const mainSS = SpreadsheetApp.openById(SPREADSHEET_ID); 
    const checkinSheet = mainSS.getSheetByName(SENIOR_CHECKIN_SHEET_NAME);
    
    let isUserCheckedIn = false;

    if (!checkinSheet || checkinSheet.getLastRow() <= 1) {
      return { success: true, data: [], isUserCheckedIn: false };
    }

    // แก้ปัญหาโหลดช้า: กำหนดขนาด Range ให้พอดี ไม่เอาช่องว่าง
    const lastRow = checkinSheet.getLastRow();
    const lastCol = checkinSheet.getLastColumn();
    const ciData = checkinSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = ciData[0];
    
    const dateCol = headers.indexOf('Date') !== -1 ? headers.indexOf('Date') : 1;
    const timeInCol = headers.indexOf('TimeIn') !== -1 ? headers.indexOf('TimeIn') : 2;
    const timeOutCol = headers.indexOf('TimeOut') !== -1 ? headers.indexOf('TimeOut') : 3;
    const recorderCol = headers.indexOf('Recorder') !== -1 ? headers.indexOf('Recorder') : 4;

    const activeWorkers = [];
    const seenNames = new Set();

    for (let i = 1; i < ciData.length; i++) {
      const timeIn = String(ciData[i][timeInCol] || "").trim();
      const timeOut = String(ciData[i][timeOutCol] || "").trim();
      const workerName = String(ciData[i][recorderCol] || "").trim();
      const dateIn = String(ciData[i][dateCol] || "").trim(); // ดึงวันที่เช็คอิน
      
      if (timeIn !== "" && timeOut === "" && workerName !== "") {
        if (!seenNames.has(workerName)) {
           seenNames.add(workerName);
           activeWorkers.push({
             name: workerName,
             timeIn: timeIn,
             dateIn: dateIn // ส่งกลับไปหน้าบ้าน
           });
        }
      }
    }

    if (seenNames.has(currentUserInfo)) {
      isUserCheckedIn = true; // ระบุว่าคนล็อกอิน ลงเวลาเข้างานหรือยัง
    }

    return { success: true, data: activeWorkers, isUserCheckedIn: isUserCheckedIn };

  } catch (err) {
    Logger.log("Error in getActiveSeniors: " + err.stack);
    return { success: false, message: err.message };
  }
}

/**
 * ดึงข้อมูลจากชีต Assigned_Tasks เพื่อไปแสดงในหน้า Work.html (ประวัติทั้งหมด)
 */
function getHistoryPageData() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID); // หรือใส่ SPREADSHEET_ID ของคุณ
    const sheet = ss.getSheetByName('Assigned_Tasks');
    
    if (!sheet) throw new Error("ไม่พบชีต 'Assigned_Tasks' โปรดตรวจสอบชื่อชีต");

    const data = sheet.getDataRange().getDisplayValues();
    const rows = [];
    
    // เริ่มวนลูปอ่านข้อมูลตั้งแต่บรรทัดที่ 2 (ข้าม Header)
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      if (!row[15]) continue; // ถ้าไม่มีรหัสงาน (Task No) ให้ข้ามแถวนั้นไป

      let assigneeFull = String(row[14] || "").trim(); // ผู้รับงาน (User)
      let empId = "";
      let empName = assigneeFull;
      
      // แยก รหัสพนักงาน และ ชื่อ ออกจากกัน
      if (assigneeFull.includes(' - ')) {
        let parts = assigneeFull.split(' - ');
        empId = parts[0].trim();
        empName = parts[1].trim();
      }

      // สร้าง rawData เพื่อให้หน้าเว็บ Work.html นำไปดึงข้อมูลใส่ฟอร์มฝั่งขวา
      let syntheticRawData = new Array(30).fill("");
      syntheticRawData[13] = String(row[13] || ""); // actionTime
      syntheticRawData[14] = String(row[14] || ""); // user (Assignee)
      syntheticRawData[16] = String(row[16] || ""); // branchContact
      syntheticRawData[17] = String(row[17] || ""); // repairType
      syntheticRawData[18] = String(row[18] || ""); // subRepairType
      syntheticRawData[19] = String(row[19] || ""); // assetType
      syntheticRawData[20] = String(row[20] || ""); // symptom
      syntheticRawData[21] = String(row[21] || ""); // teamStatus
      syntheticRawData[22] = String(row[22] || ""); // actionResult
      syntheticRawData[23] = String(row[23] || ""); // resolution
      syntheticRawData[24] = String(row[24] || ""); // forwardTeam
      syntheticRawData[25] = String(row[25] || ""); // forwardStatus
      syntheticRawData[26] = String(row[26] || ""); // remark
      syntheticRawData[27] = String(row[27] || ""); // assetId

      rows.push({
        highLow: String(row[0] || "").trim(),
        createBy: String(row[1] || "").trim(),
        branchName: String(row[2] || "").trim(),
        catalogSub: String(row[3] || "").trim(),
        detail: String(row[4] || "").trim(),
        issue: String(row[5] || "").trim(),
        reportDate: String(row[6] || "").trim(),
        createTime: String(row[7] || "").trim(),
        dueDate: String(row[8] || "").trim(),
        dueTime: String(row[9] || "").trim(),
        province: String(row[10] || "").trim(),
        phone: String(row[11] || "").trim(),
        openDate: String(row[12] || "").trim(),
        
        assignee: assigneeFull,
        empId: empId,
        empName: empName,
        taskNo: String(row[15] || "").trim(),
        workGroup: "IT",
        
        rawData: syntheticRawData
      });
    }
    
    // เรียงประวัติจากใหม่ล่าสุดไปเก่าสุด (เอาแถวล่างสุดขึ้นก่อน)
    rows.reverse();

    let userInfo = "ผู้ใช้งานทั่วไป";
    try {
      if (typeof getUserInfo === "function") {
        userInfo = getUserInfo(); 
      }
    } catch(e) {}

    // ดึงเวลาล่าสุดจาก Datalog มาแสดง
    let lastUpdateTime = "-";
    const logSheet = ss.getSheetByName('Datalog');
    if (logSheet && logSheet.getLastRow() > 1) {
      const lastTimestamp = logSheet.getRange(logSheet.getLastRow(), 1).getValue();
      if (lastTimestamp && (lastTimestamp instanceof Date || !isNaN(new Date(lastTimestamp).getTime()))) {
        lastUpdateTime = Utilities.formatDate(new Date(lastTimestamp), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");
      }
    }

    return { 
      success: true, 
      data: rows, 
      currentUserInfo: userInfo,
      lastUpdateTime: lastUpdateTime
    };

  } catch (e) {
    Logger.log("Error in getHistoryPageData: " + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * ดึงรายละเอียดสาขาจาก Sheet 'Branch' ด้วยรหัสสาขา
 * สำหรับหน้า Low / High
 */
function getBranchDetailsByCode(code) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID); 
    const sheet = ss.getSheetByName('Branch');
    
    if (!sheet) throw new Error("ไม่พบชีต 'Branch' ในระบบ");
    
    const data = sheet.getDataRange().getDisplayValues();
    const searchCode = String(code).trim();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === searchCode) { 
        return {
          success: true,
          data: {
            name: String(data[i][13] || "").trim(),     // คอลัมน์ N คือ Index 13 (ชื่อสาขา)
            province: String(data[i][5] || "").trim(),  // คอลัมน์ F คือ Index 5 (จังหวัด)
            phone: String(data[i][9] || "").trim()      // คอลัมน์ J คือ Index 9 (เบอร์โทร)
          }
        };
      }
    }
    
    return { success: false, message: "ไม่พบรหัสสาขานี้ในระบบ" };

  } catch (e) {
    Logger.log("Error in getBranchDetailsByCode: " + e.message);
    return { success: false, message: e.message };
  }
}

// ==========================================
// ฟังก์ชันสำหรับ อัปเดตข้อมูลเดิม (แทนการเพิ่มใหม่)
// ==========================================
function updateTeamAction(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID); // หรือ SPREADSHEET_ID ของคุณ
    const sheet = ss.getSheetByName('Assigned_Tasks');
    if (!sheet) throw new Error("ไม่พบชีต Assigned_Tasks");

    const data = sheet.getDataRange().getValues();
    let rowIndexToUpdate = -1;

    // ค้นหาบรรทัดที่มี Task No ตรงกัน (ค้นหาจากล่างขึ้นบน เพื่อเอาอันล่าสุด)
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][15]).trim() === String(payload.taskNo).trim()) { // คอลัมน์ P (Index 15) คือ Task No
        rowIndexToUpdate = i + 1; // +1 เพราะ Index เริ่มจาก 0 แต่แถว Sheet เริ่มจาก 1
        break;
      }
    }

    if (rowIndexToUpdate === -1) {
      throw new Error(`ไม่พบงานหมายเลข ${payload.taskNo} ในระบบเพื่อทำการแก้ไข`);
    }

    // ข้อมูลที่จะนำไปทับบรรทัดเดิม (โครงสร้างเดิมของคุณ)
    const updateRow = [
      payload.highLow,            // 1. A
      payload.createBy,           // 2. B
      payload.branchName,         // 3. C
      payload.catalogSub,         // 4. D
      payload.detail,             // 5. E
      payload.issue,              // 6. F
      payload.reportDate,         // 7. G
      payload.createTime,         // 8. H
      payload.dueDate,            // 9. I
      payload.dueTime,            // 10. J
      payload.province,           // 11. K
      payload.phone,              // 12. L
      payload.openDate,           // 13. M
      payload.actionTime,         // 14. N
      payload.user,               // 15. O
      payload.taskNo,             // 16. P
      payload.branchContact,      // 17. Q
      payload.repairType,         // 18. R
      payload.subRepairType,      // 19. S
      payload.assetType,          // 20. T
      payload.symptom,            // 21. U
      payload.teamStatus,         // 22. V
      payload.actionResult,       // 23. W
      payload.resolution,         // 24. X
      payload.forwardTeam,        // 25. Y
      payload.forwardStatus,      // 26. Z
      payload.remark,             // 27. AA
      payload.assetId             // 28. AB
    ];

    sheet.getRange(rowIndexToUpdate, 1, 1, updateRow.length).setValues([updateRow]);

    return { success: true, message: `อัปเดตข้อมูลงาน <b>${payload.taskNo}</b> เรียบร้อยแล้ว` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ฟังก์ชันสำหรับ ลบงานและย้ายไปชีต Delete
// ==========================================
function deleteTeamAction(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(TARGET_F1_SS_ID);
    const sourceSheet = ss.getSheetByName('Assigned_Tasks');
    if (!sourceSheet) throw new Error("ไม่พบชีต Assigned_Tasks");

    // ตรวจสอบและสร้างชีต "Delete" หากยังไม่มี
    let deleteSheet = ss.getSheetByName('Delete');
    if (!deleteSheet) {
      deleteSheet = ss.insertSheet('Delete');
      // คัดลอก Header จากชีตเดิมมา
      const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
      headers.push("Deleted By", "Deleted Time"); // เพิ่มหัวข้อคนลบและเวลา
      deleteSheet.appendRow(headers);
    }

    const data = sourceSheet.getDataRange().getValues();
    let rowIndexToDelete = -1;
    let rowDataToMove = null;

    // ค้นหาข้อมูล (จากล่างขึ้นบน)
    for (let i = data.length - 1; i > 0; i--) {
      if (String(data[i][15]).trim() === String(payload.taskNo).trim()) {
        rowIndexToDelete = i + 1;
        rowDataToMove = data[i];
        break;
      }
    }

    if (rowIndexToDelete === -1) {
      throw new Error(`ไม่พบงานหมายเลข ${payload.taskNo} ในระบบเพื่อทำการลบ`);
    }

    // เพิ่มข้อมูลคนลบและเวลาลบเข้าไปต่อท้าย Array
    const deletedTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "dd/MM/yyyy HH:mm:ss");
    rowDataToMove.push(payload.deletedBy, deletedTime);

    // 1. นำข้อมูลไปวางในชีต Delete
    deleteSheet.appendRow(rowDataToMove);
    
    // 2. ลบบรรทัดในชีต Assigned_Tasks
    sourceSheet.deleteRow(rowIndexToDelete);

    return { success: true, message: `ลบงาน <b>${payload.taskNo}</b> และย้ายไปที่ชีต Delete สำเร็จ!` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}