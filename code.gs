// ===================================================
// === 1. การตั้งค่ารวม (Configuration) ===
// ===================================================
const CONFIG = {
  // ++ ส่วนของข้อมูล (Data Sheets) ++
  receiptData: {
    dataSheetName: "ข้อมูลรับเข้า"
  },
  withdrawalData: {
    dataSheetName: "ข้อมูลการเบิก"
  },
  sortingData: {
    dataSheetName: "ข้อมูลการคัดแยก"
  },
  randomCheckData: {
    dataSheetName: "ข้อมูลสุ่มเช็ค"
  },
  maintenanceData: {
    dataSheetName: "ข้อมูลซ่อมบำรุง"
  },
  // ++ ส่วนของข้อมูลเสริมสำหรับ Web App ++
  webAppInfo: {
    stockSheet: "คลัง",
    contactsSheet: "Contacts",      // ++ เพิ่มเข้ามา ++
    trayStockSheet: "TrayStock",    // ++ เพิ่มเข้ามา ++
    configSheet: "config",
    documentInfoSheet: "ข้อมูลเอกสาร",
    settingsSheet: "Settings",
    logSheet: "Log"
  }
};

// ===================================================
// === 2. ฟังก์ชันหลักสำหรับ Web App (Entry & Shared) ===
// ===================================================
function doGet(e) {
  // 1. ตรวจสอบสิทธิ์ผู้ใช้งานก่อน
  if (checkUserAccess_()) {

    // 2. ตรวจสอบว่าผู้ใช้ต้องการไปหน้าทำงานหลักหรือไม่
    if (e.parameter.page) {
      const template = HtmlService.createTemplateFromFile('WebApp');
      template.initialPage = e.parameter.page;
      template.dashboardUrl = ScriptApp.getService().getUrl(); 
      
      return template.evaluate()
        .setTitle("ระบบจัดการสต็อก") // ตั้งชื่อบนแท็บ
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    // 3. ถ้าไม่ใช่ ให้ไปที่หน้า Dashboard
    } else {
      return HtmlService.createTemplateFromFile('Dashboard')  
        .evaluate()
        .setTitle("Dashboard | ระบบจัดการสต็อก") // ตั้งชื่อบนแท็บ
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    }
    
  // 4. ถ้าไม่มีสิทธิ์ ให้แสดงหน้า Access Denied
  } else {
    return HtmlService.createHtmlOutputFromFile('AccessDenied')
      .setTitle("Access Denied")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }
}

// ===================================================
// === [NEW] ฟังก์ชันสำหรับหน้า Dashboard ===========
// ===================================================

function getDashboardStats() {
  try {
    const settings = getAppSettings();
    const LOW_STOCK_THRESHOLD = settings.lowStockThreshold || 10;

    const stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    if (!stockSheet || stockSheet.getLastRow() < 2) {
      return { totalProducts: 0, lowStockCount: 0, outOfStockCount: 0 };
    }

    const data = stockSheet.getDataRange().getValues();
    const headers = data.shift();
    
    const idColIndex = headers.indexOf("รหัสสินค้า");
    const qtyColIndex = headers.indexOf("คลังกลาง/ฟอง");

    if (idColIndex === -1 || qtyColIndex === -1) {
      throw new Error("ไม่พบคอลัมน์ 'รหัสสินค้า' หรือ 'คลังกลาง/ฟอง'");
    }

    let totalProducts = 0;
    let lowStockCount = 0;
    let outOfStockCount = 0;

    data.forEach(row => {
      if (row[idColIndex]) {
        totalProducts++;
        const quantity = Number(row[qtyColIndex]) || 0;

        if (quantity === 0) {
          outOfStockCount++;
        } else if (quantity <= LOW_STOCK_THRESHOLD) {
          lowStockCount++;
        }
      }
    });

    return { 
      totalProducts: totalProducts, 
      lowStockCount: lowStockCount, 
      outOfStockCount: outOfStockCount 
    };

  } catch (e) {
    console.error("getDashboardStats Error: " + e.toString());
    return { error: e.message };
  }
}

/**
 * [NEW] ดึงการตั้งค่าทั้งหมดของแอปจากชีต 'ข้อมูลเอกสาร' ในครั้งเดียว
 */
function getAppSettings() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'appSettings';
  const cachedSettings = cache.get(cacheKey);
  if (cachedSettings) {
    return JSON.parse(cachedSettings);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.documentInfoSheet);
  const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  
  const settings = {};
  const keyMapping = {
    'ชื่อบริษัท': 'companyName',
    'ที่อยู่ 1': 'address1',
    'ที่อยู่ 2': 'address2',
    'ข้อมูลติดต่อ': 'contactInfo',
    'Low Stock Threshold': 'lowStockThreshold'
  };

  data.forEach(row => {
    const key = row[0].toString().trim();
    if (keyMapping[key]) {
      settings[keyMapping[key]] = row[1];
    }
  });

  cache.put(cacheKey, JSON.stringify(settings), 3600); // Cache for 1 hour
  return settings;
}

/**
 * [NEW] บันทึกการตั้งค่าทั้งหมดลงในชีต 'ข้อมูลเอกสาร'
 */
function saveAppSettings(settingsData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.documentInfoSheet);
    const data = sheet.getRange("A1:A" + sheet.getLastRow()).getValues().flat();
    
    const keyMapping = {
      'companyName': 'ชื่อบริษัท',
      'address1': 'ที่อยู่ 1',
      'address2': 'ที่อยู่ 2',
      'contactInfo': 'ข้อมูลติดต่อ',
      'lowStockThreshold': 'Low Stock Threshold'
    };

    for (const key in settingsData) {
      const settingName = keyMapping[key];
      if (settingName) {
        const rowIndex = data.findIndex(item => item.toString().trim() === settingName);
        if (rowIndex !== -1) {
          sheet.getRange(rowIndex + 1, 2).setValue(settingsData[key]);
        } else {
          // ถ้าไม่เจอ setting นั้น ให้เพิ่มใหม่ท้ายชีต
          sheet.appendRow([settingName, settingsData[key]]);
        }
      }
    }

    CacheService.getScriptCache().remove('appSettings');
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * [NEW] ฟังก์ชันสำหรับดึงเนื้อหาไฟล์ HTML อื่นเข้ามาในเทมเพลตหลัก
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * [OPTIMIZED] ตรวจสอบว่าอีเมลของผู้ใช้ปัจจุบันมีสิทธิ์เข้าถึงแอปหรือไม่ โดยใช้ Cache
 */
function checkUserAccess_() {
  try {
    const cache = CacheService.getScriptCache();
    const currentUser = Session.getActiveUser().getEmail().toLowerCase();
    
    // 1. ลองดึงรายชื่ออีเมลจาก Cache ก่อน
    const cachedEmails = cache.get('allowedEmails');
    if (cachedEmails) {
        console.log("User access checked from CACHE.");
        const allowedEmails = JSON.parse(cachedEmails);
        return allowedEmails.includes(currentUser);
    }

    // 2. ถ้าไม่มีใน Cache ให้ไปดึงจาก Sheet
    console.log("User access checked from SHEET.");
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.settingsSheet);
    if (!settingsSheet) return true; // ถ้าไม่มีชีต Settings ให้เข้าได้ทุกคน
    
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) return false;
    
    const allowedEmails = settingsSheet.getRange(`A2:A${lastRow}`).getValues()
      .flat()
      .map(email => email.toString().trim().toLowerCase())
      .filter(email => email);
    
    // 3. นำรายชื่ออีเมลไปเก็บใน Cache (เก็บไว้ 1 ชั่วโมง) ก่อนส่งค่ากลับ
    cache.put('allowedEmails', JSON.stringify(allowedEmails), 3600);
    
    return allowedEmails.includes(currentUser);

  } catch (e) {
    console.error("checkUserAccess_ Error: " + e.toString());
    return false;
  }
}

/**
 * READ: ดึงข้อมูลสต็อกทั้งหมดจากชีต 'คลัง'
 */
function getStockData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    if (!sheet) throw new Error(`ไม่พบชีต '${CONFIG.webAppInfo.stockSheet}'`);
    if (sheet.getLastRow() < 2) return { headers: [], data: [] };
    
    const wholeData = sheet.getDataRange().getValues();
    const headers = wholeData[0];
    const data = wholeData.slice(1);
    
    return { headers: headers, data: data };
  } catch (e) {
    console.error("getStockData Error: " + e.toString());
    return { headers: [], data: [], error: e.message };
  }
}

/**
 * [NEW] อ่านข้อมูลสต็อกทั้งหมดจากชีตแล้วจัดเก็บใน Cache
 * @returns {Map<string, {quantity: number}>} - Map ที่มี key เป็น Product ID
 */
function getStockDataFromCache_() {
  const cache = CacheService.getScriptCache();
  const cachedStock = cache.get('fullStockData');

  if (cachedStock) {
    console.log("Stock data retrieved from CACHE.");
    return new Map(Object.entries(JSON.parse(cachedStock)));
  }

  console.log("Stock data retrieved from SHEET.");
  const stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
  if (!stockSheet || stockSheet.getLastRow() < 2) return new Map();

  const headers = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
  const idColIndex = headers.indexOf("รหัสสินค้า");
  const qtyColIndex = headers.indexOf("คลังกลาง/ฟอง");

  if (idColIndex === -1 || qtyColIndex === -1) return new Map();

  const stockData = stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, Math.max(idColIndex, qtyColIndex) + 1).getValues();
  const stockMap = new Map();

  stockData.forEach(row => {
    const productId = row[idColIndex].toString().trim();
    if (productId) {
      stockMap.set(productId, { quantity: row[qtyColIndex] });
    }
  });
  
  // เก็บข้อมูลใน Cache เป็นเวลา 10 นาที
  // แปลง Map เป็น Object ก่อนเก็บ
  cache.put('fullStockData', JSON.stringify(Object.fromEntries(stockMap)), 600);

  return stockMap;
}

/**
 * [OPTIMIZED] ดึงจำนวนสต็อกคงเหลือจากรหัสสินค้า โดยใช้ข้อมูลจาก Cache
 */
function getStockByProductId(productId) {
  try {
    if (!productId) return null;
    const stockMap = getStockDataFromCache_();
    const productInfo = stockMap.get(productId.toString().trim());
    
    // อย่าลืมอัปเดต cache เมื่อมีการเบิกจ่ายด้วยนะครับ
    // อาจจะต้องมีฟังก์ชัน clearStockCache() เพื่อเรียกใช้หลังการเบิกสำเร็จ
    
    return productInfo ? productInfo.quantity : null;
  } catch (e) {
    console.error("getStockByProductId Error: " + e.toString());
    return null;
  }
} 

// ===================================================
// === 3. ฟังก์ชัน Web App: DROPDOWN HELPERS ========
// ===================================================
function getWebAppInitialData() {
    const allContacts = getContacts_();
    return {
        products: getProductList(),
        suppliers: allContacts.filter(c => c.type === 'Supplier'),
        branches: allContacts.filter(c => c.type === 'Branch'),
        allReturnContacts: allContacts,
        employees: getEmployeeList(),
        contactBalances: getContactDashboardData() // <-- [ เพิ่มบรรทัดนี้ ]
    };
}


/**
 * [NEW] ฟังก์ชันกลางสำหรับดึงข้อมูลจากชีต Contacts ทั้งหมด (ใช้ Cache)
 * @param {string} type - (Optional) กรองประเภท 'Supplier' หรือ 'Branch'
 * @returns {Array<Object>}
 */
function getContacts_(type = null) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'allContactsData';
    const cached = cache.get(cacheKey);

    if (cached) {
        console.log("Contacts data from CACHE.");
        const allContacts = JSON.parse(cached);
        if (type) {
            return allContacts.filter(c => c.type === type);
        }
        return allContacts;
    }

    console.log("Contacts data from SHEET.");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contacts");
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getRange(`A2:D${sheet.getLastRow()}`).getValues();
    const contacts = data.map(row => ({
        id: row[0],
        name: row[1],
        type: row[2],
        tel: row[3]
    })).filter(c => c.id && c.name);

    cache.put(cacheKey, JSON.stringify(contacts), 3600); // Cache for 1 hour

    if (type) {
        return contacts.filter(c => c.type === type);
    }
    return contacts;
}

/**
 * [REVISED] อัปเดตยอดแผงคงค้าง พร้อมบันทึกประวัติด้วยระบบ Log ใหม่
 */
function updateContactTrayStock_(contactId, contactName, quantity) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TrayStock");
        if (!sheet) return;

        if (sheet.getLastRow() === 0) {
            sheet.getRange("A1:C1").setValues([["ContactID", "ContactName", "TrayBalance"]]);
        }

        const data = sheet.getRange("A2:C" + (sheet.getLastRow() || 1)).getValues();
        let contactFound = false;
        let newBalance = 0;

        for (let i = 0; i < data.length; i++) {
            if (data[i][0].toString().trim() === contactId.toString().trim()) {
                const currentBalance = parseInt(data[i][2], 10) || 0;
                newBalance = currentBalance + quantity;
                sheet.getRange(i + 2, 3).setValue(newBalance);
                contactFound = true;
                
                const finalContactName = contactName || data[i][1];
                
                // [แก้] เปลี่ยนมาเรียกใช้ logTrayUpdate_ ตัวใหม่
                logTrayUpdate_(contactId, finalContactName, quantity, newBalance);
                break;
            }
        }

        if (!contactFound) {
            newBalance = quantity;
            sheet.appendRow([contactId, contactName, newBalance]);
            
            // [แก้] เปลี่ยนมาเรียกใช้ logTrayUpdate_ ตัวใหม่
            logTrayUpdate_(contactId, contactName, quantity, newBalance);
        }
    } finally {
        lock.releaseLock();
    }
}
/**
 * [UPGRADED] READ: ดึงข้อมูลสินค้า "ทุกคอลัมน์" จากชีตคลัง
 */
function getProductsWithDetails() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    if (!sheet || sheet.getLastRow() < 2) return { headers: [], data: [] };

    const wholeData = sheet.getDataRange().getValues();
    const headers = wholeData.shift(); // ดึงหัวข้อออกมา
    const data = wholeData;

    // แปลงข้อมูล Array ธรรมดาให้เป็น Array of Objects
    const products = data.map(row => {
      const productObject = {};
      headers.forEach((header, index) => {
        productObject[header] = row[index];
      });
      return productObject;
    });

    // ส่งกลับไปทั้ง Headers และ Data ที่แปลงแล้ว
    return { headers: headers, data: products };

  } catch (e) {
    console.error("getProductsWithDetails Error: " + e.toString());
    return { headers: [], data: [], error: e.message };
  }
}

function addNewProduct(productData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const productCodeCol = headers.indexOf("รหัสสินค้า");
    if (productCodeCol !== -1) {
        const allProductCodes = sheet.getRange(2, productCodeCol + 1, sheet.getLastRow() - 1 || 1, 1).getValues().flat();
        if (allProductCodes.includes(productData.productCode.trim())) {
          throw new Error(`รหัสสินค้า '${productData.productCode}' นี้มีอยู่แล้วในระบบ`);
        }
    }

    const newRow = new Array(headers.length).fill('');
    newRow[headers.indexOf("รหัสสินค้า")] = productData.productCode.trim();
    newRow[headers.indexOf("ชื่อสินค้า")] = productData.productName.trim();
    newRow[headers.indexOf("หมวดหมู่")] = productData.category.trim();
    
    headers.forEach((header, index) => {
        if(header.includes('คลัง') || header.includes('จำนวน')){
            newRow[index] = 0;
        }
    });

    sheet.appendRow(newRow);

    // [แก้] บันทึก Log
    logActivity_({
        action: 'เพิ่ม',
        target: 'สินค้า',
        docId: productData.productCode.trim(),
        details: { productName: productData.productName.trim() }
    });

    clearServerCache();
    return { success: true, message: `เพิ่มสินค้า '${productData.productName}' สำเร็จ` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateProductDetails(productData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    
    const productCodeCol = headers.indexOf("รหัสสินค้า");
    const productNameCol = headers.indexOf("ชื่อสินค้า");
    const categoryCol = headers.indexOf("หมวดหมู่");

    const rowIndexToUpdate = data.findIndex(row => row[productCodeCol].toString().trim() === productData.productCode.trim());

    if (rowIndexToUpdate === -1) {
      throw new Error(`ไม่พบรหัสสินค้า '${productData.productCode}'`);
    }

    const rowToUpdate = rowIndexToUpdate + 2;
    sheet.getRange(rowToUpdate, productNameCol + 1).setValue(productData.productName.trim());
    sheet.getRange(rowToUpdate, categoryCol + 1).setValue(productData.category.trim());

    // [แก้] บันทึก Log
    logActivity_({
        action: 'แก้ไข',
        target: 'สินค้า',
        docId: productData.productCode.trim(),
        details: { productName: productData.productName.trim() }
    });

    clearServerCache();
    return { success: true, message: `อัปเดตข้อมูลสินค้า '${productData.productCode}' สำเร็จ` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteProduct(productCode) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const productCodeCol = headers.indexOf("รหัสสินค้า");
    const stockQtyCol = headers.indexOf("คลังกลาง/ฟอง");
    
    const rowIndexToDelete = data.findIndex(row => row[productCodeCol].toString().trim() === productCode.trim());

    if (rowIndexToDelete === -1) {
      throw new Error(`ไม่พบรหัสสินค้า '${productCode}'`);
    }

    const stockQty = data[rowIndexToDelete][stockQtyCol];
    
    if (Number(stockQty) > 0) {
      return { success: false, message: `ไม่สามารถลบได้! สินค้า '${productCode}' ยังมีสต็อกคงเหลือ (${stockQty})` };
    }

    sheet.deleteRow(rowIndexToDelete + 2);

    // [แก้] บันทึก Log
    logActivity_({
      action: 'ลบ',
      target: 'สินค้า',
      docId: productCode.trim(),
      details: {}
    });

    clearServerCache();
    return { success: true, message: `ลบสินค้า '${productCode}' สำเร็จ` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


function getProductList() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get('productList');
    if (cachedData) {
      console.log("Product list from CACHE");
      return JSON.parse(cachedData);
    }

    console.log("Product list from SHEET");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    // ++ ส่วนที่แก้ไขให้ฉลาดขึ้น ++
    const wholeData = sheet.getDataRange().getValues();
    const headers = wholeData.shift(); // เอาหัวข้อออกมา
    const idColIndex = headers.indexOf("รหัสสินค้า");
    const nameColIndex = headers.indexOf("ชื่อสินค้า");

    if (idColIndex === -1 || nameColIndex === -1) {
        throw new Error("ไม่พบคอลัมน์ 'รหัสสินค้า' หรือ 'ชื่อสินค้า' ในชีต 'คลัง'");
    }

    const productList = wholeData
      .filter(row => row[idColIndex] && row[nameColIndex])
      .map(row => ({ 
        id: row[idColIndex].toString().trim(), 
        name: row[nameColIndex].toString().trim() 
      }));
    // ++ สิ้นสุดส่วนที่แก้ไข ++
    
    cache.put('productList', JSON.stringify(productList), 3600);
    return productList;
  } catch (e) {
    console.error("getProductList Error: " + e.toString());
    return [];
  }
}

function getEmployeeList() {
  try {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.configSheet);
    if (!configSheet) throw new Error(`ไม่พบชีต '${CONFIG.webAppInfo.configSheet}'`);
    const lastRow = configSheet.getLastRow();
    if (lastRow < 3) return [];
    const employeeNames = configSheet.getRange(`N3:N${lastRow}`).getValues().flat().filter(name => name.toString().trim() !== '');
    return [...new Set(employeeNames)];
  } catch (e) { console.error("getEmployeeList Error: " + e.toString()); return []; }
}


// ===================================================
// === 4. ฟังก์ชัน Web App: RECEIPT (รับสินค้า) ========
// ===================================================

function saveReceiptDataFromWebApp(formData) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.receiptData.dataSheetName);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
    const timestamp = new Date();
    const docId = generateDocId_('GRN', CONFIG.receiptData.dataSheetName);

    if (dataSheet.getLastRow() === 0) {
      const headers = [
        ['เลขเอกสาร', 'วันที่รับเข้า', 'ชื่อซัพพลายเออร์', 'เบอร์ติดต่อ', 'รายการสินค้า', 'จำนวน', 'หน่วย', 'น้ำหนัก', 'หมายเหตุ', 'ชื่อผู้รับสินค้า', 'แผงที่รับ', 'แผงที่คืน', 'ผู้บันทึก', 'เวลาที่สร้าง']
      ];
      dataSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    }

    const recordsToSave = formData.items.map(item => [
        docId, timestamp, formData.supplier, formData.tel,
        item.name, item.quantity, item.unit, item.weight, item.note,
        formData.receiver, formData.traysReceived || 0, formData.traysReturned || 0,
        currentUser, timestamp
    ]);

    const netTrayChange = (parseInt(formData.traysReceived, 10) || 0) - (parseInt(formData.traysReturned, 10) || 0);
    if (netTrayChange !== 0 && formData.contactId) {
        updateContactTrayStock_(formData.contactId, formData.supplier, netTrayChange);
    }

    if (recordsToSave.length > 0) {
        dataSheet.getRange(dataSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
        
        // [แก้] บันทึก Log
        logActivity_({
            action: 'สร้าง',
            target: 'เอกสารรับเข้า',
            docId: docId,
            details: { 
                contactName: formData.supplier, 
                itemCount: recordsToSave.length 
            }
        });
        
        const productList = getProductList();
        const itemsToUpdate = formData.items.map(item => {
            const product = productList.find(p => p.name === item.name);
            return { id: product ? product.id : null, quantityChange: Math.abs(Number(item.quantity)) };
        }).filter(item => item.id);

        if (itemsToUpdate.length > 0) {
            updateStockLevels_(itemsToUpdate);
        }
        
        clearServerCache(); 
        return { success: true, docId: docId };
    } else {
        throw new Error("ไม่พบรายการสินค้าที่จะบันทึก");
    }
  } catch (e) {
      console.error("saveReceiptDataFromWebApp Error: " + e.message);
      return { success: false, message: e.message };
  }
}

function updateReceiptDataFromWebApp(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.receiptData.dataSheetName);
    const allData = dataSheet.getDataRange().getValues();
    const headers = allData.shift();

    const docIdCol = headers.indexOf('เลขเอกสาร');
    const itemNameCol = headers.indexOf('รายการสินค้า');
    const qtyCol = headers.indexOf('จำนวน');
    const trayReceivedCol = headers.indexOf('แผงที่รับ');
    const trayReturnedCol = headers.indexOf('แผงที่คืน');
    const supplierNameCol = headers.indexOf('ชื่อซัพพลายเออร์');

    const oldRecords = allData.filter(row => row[docIdCol].toString().trim() === formData.docId.toString().trim());

    if (oldRecords.length > 0) {
      const productList = getProductList();
      const itemsToReturn = oldRecords.map(row => {
        const product = productList.find(p => p.name === row[itemNameCol]);
        return { id: product ? product.id : null, quantityChange: -Math.abs(Number(row[qtyCol])) };
      }).filter(item => item.id);
      
      if (itemsToReturn.length > 0) {
        updateStockLevels_(itemsToReturn);
      }

      const oldSupplierName = oldRecords[0][supplierNameCol];
      const allContacts = getContacts_();
      const oldContact = allContacts.find(c => c.name === oldSupplierName);
      if(oldContact) {
        const oldTraysReceived = parseInt(oldRecords[0][trayReceivedCol], 10) || 0;
        const oldTraysReturned = parseInt(oldRecords[0][trayReturnedCol], 10) || 0;
        const oldNetChange = oldTraysReceived - oldTraysReturned;
        if (oldNetChange !== 0) {
            updateContactTrayStock_(oldContact.id, oldContact.name, -oldNetChange);
        }
      }
    }

    const rowsToKeep = allData.filter(row => row[docIdCol].toString().trim() !== formData.docId.toString().trim());
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
    const timestamp = new Date();
    const originalDate = oldRecords.length > 0 ? oldRecords[0][headers.indexOf('วันที่รับเข้า')] : timestamp;

    const recordsToUpdate = formData.items.map(item => [
      formData.docId, originalDate, formData.supplier, formData.tel,
      item.name, item.quantity, item.unit, item.weight, item.note,
      formData.receiver, formData.traysReceived || 0, formData.traysReturned || 0,
      currentUser, timestamp
    ]);
    
    const finalData = [headers, ...rowsToKeep, ...recordsToUpdate];
    dataSheet.clearContents();
    if (finalData.length > 0) {
      dataSheet.getRange(1, 1, finalData.length, headers.length).setValues(finalData);
    }
    
    const productList = getProductList();
    const itemsToAdd = formData.items.map(item => {
        const product = productList.find(p => p.name === item.name);
        return { id: product ? product.id : null, quantityChange: Math.abs(Number(item.quantity)) };
    }).filter(item => item.id);
    
    if (itemsToAdd.length > 0) {
        updateStockLevels_(itemsToAdd);
    }
    
    const newNetChange = (parseInt(formData.traysReceived, 10) || 0) - (parseInt(formData.traysReturned, 10) || 0);
    if (newNetChange !== 0 && formData.contactId) {
        updateContactTrayStock_(formData.contactId, formData.supplier, newNetChange);
    }

    // [แก้] บันทึก Log
    logActivity_({
        action: 'แก้ไข',
        target: 'เอกสารรับเข้า',
        docId: formData.docId,
        details: { 
            contactName: formData.supplier, 
            itemCount: recordsToUpdate.length 
        }
    });
    clearServerCache();
    return { success: true, docId: formData.docId };

  } catch (e) {
    console.error("updateReceiptDataFromWebApp Error: " + e.toString());
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteReceiptByDocId(docId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.receiptData.dataSheetName);
    const sortingSheet = ss.getSheetByName(CONFIG.sortingData.dataSheetName);

    if (sortingSheet && sortingSheet.getLastRow() > 1) {
      const refDocIds = sortingSheet.getRange(2, 3, sortingSheet.getLastRow() - 1, 1).getValues().flat();
      if (refDocIds.includes(docId.toString().trim())) {
        throw new Error(`เอกสาร GRN '${docId}' ถูกใช้ในการคัดแยกแล้ว! กรุณาลบเอกสารคัดแยกที่อ้างอิงถึงเอกสารนี้ก่อน`);
      }
    }

    const allData = dataSheet.getDataRange().getValues();
    const headers = allData.shift();
    const nameCol = headers.indexOf('รายการสินค้า');
    const qtyCol = headers.indexOf('จำนวน');
    const trayReceivedCol = headers.indexOf('แผงที่รับ');
    const trayReturnedCol = headers.indexOf('แผงที่คืน');
    const supplierNameCol = headers.indexOf('ชื่อซัพพลายเออร์');

    const itemsToDelete = allData.filter(row => row[0].toString().trim() === docId.toString().trim());
    let oldSupplierName = "";

    if (itemsToDelete.length > 0) {
        oldSupplierName = itemsToDelete[0][supplierNameCol];
        const productList = getProductList();
        const itemsToUpdate = itemsToDelete.map(row => {
            const product = productList.find(p => p.name === row[nameCol]);
            return { id: product ? product.id : null, quantityChange: -Math.abs(Number(row[qtyCol])) };
        }).filter(item => item.id);

        if (itemsToUpdate.length > 0) {
            updateStockLevels_(itemsToUpdate);
        }

        const allContacts = getContacts_();
        const contact = allContacts.find(c => c.name === oldSupplierName);

        if (contact) {
            const traysReceived = parseInt(itemsToDelete[0][trayReceivedCol], 10) || 0;
            const traysReturned = parseInt(itemsToDelete[0][trayReturnedCol], 10) || 0;
            const netChange = traysReceived - traysReturned;
            if (netChange !== 0) {
                updateContactTrayStock_(contact.id, contact.name, -netChange);
            }
        }
    }
    
    deleteRowsByDocId_(dataSheet, docId);
    // [แก้] บันทึก Log
    logActivity_({
        action: 'ลบ',
        target: 'เอกสารรับเข้า',
        docId: docId,
        details: { 
            contactName: oldSupplierName, 
            itemCount: itemsToDelete.length 
        }
    });
    
    clearServerCache();
    return { success: true, message: `ลบเอกสารรับเข้า ${docId} และปรับสต็อก/ยอดแผงสำเร็จ` };

  } catch (e) { 
    console.error("deleteReceiptByDocId Error: " + e.message);
    return { success: false, message: e.message }; 
  }
}



/**
 * [REVISED] ดึงข้อมูลการรับสินค้าตามลำดับคอลัมน์ใหม่
 */
function getReceiptData(page = 1, rowsPerPage = 10, searchTerm = "") {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.receiptData.dataSheetName);
    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) return { data: [], totalItems: 0 };

    // ดึงข้อมูลทั้งหมดมาทีเดียว (เหมือนเดิม)
    const allData = dataSheet.getRange(2, 1, lastRow - 1, dataSheet.getLastColumn()).getValues();

    // กรองข้อมูล (เหมือนเดิม)
    const lowerCaseSearchTerm = searchTerm.trim().toLowerCase();
    const filteredData = lowerCaseSearchTerm
      ? allData.filter(row => 
          row[0].toString().toLowerCase().includes(lowerCaseSearchTerm) ||
          row[2].toString().toLowerCase().includes(lowerCaseSearchTerm) ||
          row[9].toString().toLowerCase().includes(lowerCaseSearchTerm)
        )
      : allData;

    // จัดกลุ่มข้อมูล
    const groupedData = {};
    filteredData.forEach(row => {
      const docId = row[0];
      if (!groupedData[docId]) {
        groupedData[docId] = {
          docId: docId,
          date: new Date(row[1]).toLocaleDateString('th-TH'),
          supplier: row[2],
          tel: row[3] || '',
          receiver: row[9] || '',
          // ++ แก้ไขตำแหน่ง Index ของคอลัมน์ ++
          traysReceived: row[10] || 0, // คอลัมน์ที่ 11
          traysReturned: row[11] || 0, // คอลัมน์ที่ 12
          items: []
        };
      }
      groupedData[docId].items.push({
        name: row[4], quantity: row[5], unit: row[6], weight: row[7], note: row[8]
      });
    });
    
    const allGroupedRecords = Object.values(groupedData).reverse();
    const totalItems = allGroupedRecords.length;

    const startIndex = (page - 1) * rowsPerPage;
    const paginatedData = allGroupedRecords.slice(startIndex, startIndex + rowsPerPage);

    return { data: paginatedData, totalItems: totalItems };
  } catch(e) {
    console.error("getReceiptData Error: " + e.toString());
    return { data: [], totalItems: 0, error: e.message };
  }
}

// ===================================================
// === 5. ฟังก์ชัน Web App: WITHDRAWAL (เบิกสินค้า) ====
// ===================================================
function saveDataFromWebApp(formData) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const withdrawalSheet = ss.getSheetByName(CONFIG.withdrawalData.dataSheetName);
        const stockSheet = ss.getSheetByName(CONFIG.webAppInfo.stockSheet);
        if (!stockSheet) throw new Error("ไม่พบชีต 'คลัง' สำหรับตรวจสอบสต็อก");

        const stockMap = getStockDataFromCache_();
        for (const item of formData.items) {
            if (!item.id) throw new Error(`ไม่พบรหัสสินค้าสำหรับรายการ '${item.name}'`);
            const stockInfo = stockMap.get(item.id.trim());
            if (!stockInfo) throw new Error(`ไม่พบรหัสสินค้า '${item.id}' ในคลัง`);
            if (Number(stockInfo.quantity) < Number(item.quantity)) {
                throw new Error(`สินค้าไม่พอ! '${item.name}' มีในคลัง ${stockInfo.quantity} แต่ต้องการเบิก ${item.quantity}`);
            }
        }
        
        const traysSent = parseInt(formData.traysSent, 10) || 0;
        if (traysSent > 0 && formData.departmentId) {
            updateContactTrayStock_(formData.departmentId, formData.department, traysSent);
        }
        
        const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
        const timestamp = new Date();
        const docId = generateDocId_('WD', CONFIG.withdrawalData.dataSheetName);

        if (withdrawalSheet.getLastRow() === 0) {
            const headers = [['เลขเอกสาร', 'วันที่', 'ขอผู้เบิก', 'รหัสสาขา', 'สาขา', 'รหัสสินค้า', 'สินค้า', 'จำนวน', 'หน่วย', 'หมายเหตุ', 'ผู้อนุมัติ', 'แผงที่ส่ง', 'เมล์ที่สร้าง', 'วันที่สร้าง']];
            withdrawalSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
        }

        const recordsToSave = formData.items.map(item => [
          docId, timestamp, formData.requester, formData.departmentId, formData.department,
          item.id, item.name, item.quantity, item.unit, item.note,
          formData.approver, formData.traysSent || 0, currentUser, timestamp
        ]);
        
        if (recordsToSave.length > 0) {
            withdrawalSheet.getRange(withdrawalSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
            // [แก้] บันทึก Log
            logActivity_({
                action: 'สร้าง',
                target: 'เอกสารเบิก',
                docId: docId,
                details: {
                    contactName: formData.department,
                    itemCount: recordsToSave.length
                }
            });
        } else {
            throw new Error("ไม่พบรายการสินค้าที่จะบันทึก");
        }

        const itemsToUpdate = formData.items.map(item => ({
            id: item.id,
            quantityChange: -Math.abs(Number(item.quantity))
        }));
        updateStockLevels_(itemsToUpdate);

        clearServerCache();
        return { success: true, docId: docId };

    } catch (e) {
        console.error("saveDataFromWebApp Error: " + e.toString());
        return { success: false, message: e.message };
    }
}

function updateWithdrawalDataFromWebApp(formData) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const withdrawalSheet = ss.getSheetByName(CONFIG.withdrawalData.dataSheetName);
        const allData = withdrawalSheet.getDataRange().getValues();
        const headers = allData.shift();

        const docIdCol = headers.indexOf('เลขเอกสาร');
        const itemIdCol = headers.indexOf('รหัสสินค้า');
        const qtyCol = headers.indexOf('จำนวน');
        const branchIdCol = headers.indexOf('รหัสสาขา');
        const traysSentCol = headers.indexOf('แผงที่ส่ง');

        const oldRecords = allData.filter(row => row[docIdCol].toString().trim() === formData.docId.toString().trim());

        if (oldRecords.length > 0) {
            const itemsToReturn = oldRecords.map(row => ({
                id: row[itemIdCol],
                quantityChange: Math.abs(Number(row[qtyCol]))
            })).filter(item => item.id);
            if (itemsToReturn.length > 0) {
                updateStockLevels_(itemsToReturn);
            }

            const oldBranchId = oldRecords[0][branchIdCol];
            const oldTraysSent = parseInt(oldRecords[0][traysSentCol], 10) || 0;
            if (oldTraysSent > 0 && oldBranchId) {
                updateContactTrayStock_(oldBranchId.toString().trim(), "", -Math.abs(oldTraysSent));
            }
        }

        const stockMap = getStockDataFromCache_();
        for (const item of formData.items) {
            const stockInfo = stockMap.get(item.id.trim());
            const currentStock = (stockInfo ? Number(stockInfo.quantity) : 0);
            if (currentStock < Number(item.quantity)) {
                throw new Error(`[แก้ไข] สินค้าไม่พอ! '${item.name}' มีในคลัง ${currentStock} แต่ต้องการเบิก ${item.quantity}`);
            }
        }

        const rowsToKeep = allData.filter(row => row[docIdCol].toString().trim() !== formData.docId.toString().trim());
        const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
        const timestamp = new Date();
        const originalDate = oldRecords.length > 0 ? oldRecords[0][headers.indexOf('วันที่')] : timestamp;

        const recordsToUpdate = formData.items.map(item => [
            formData.docId, originalDate, formData.requester, formData.departmentId, formData.department,
            item.id, item.name, item.quantity, item.unit, item.note, formData.approver,
            formData.traysSent || 0, currentUser, timestamp
        ]);
        
        const finalData = [headers, ...rowsToKeep, ...recordsToUpdate];
        withdrawalSheet.clearContents();
        if (finalData.length > 0) {
            withdrawalSheet.getRange(1, 1, finalData.length, headers.length).setValues(finalData);
        }

        const itemsToDeduct = formData.items.map(item => ({
            id: item.id,
            quantityChange: -Math.abs(Number(item.quantity))
        }));
        if (itemsToDeduct.length > 0) {
            updateStockLevels_(itemsToDeduct);
        }

        const newTraysSent = parseInt(formData.traysSent, 10) || 0;
        if (newTraysSent > 0 && formData.departmentId) {
            updateContactTrayStock_(formData.departmentId, formData.department, newTraysSent);
        }
        
        // [แก้] บันทึก Log
        logActivity_({
            action: 'แก้ไข',
            target: 'เอกสารเบิก',
            docId: formData.docId,
            details: {
                contactName: formData.department,
                itemCount: recordsToUpdate.length
            }
        });
        clearServerCache();
        return { success: true, docId: formData.docId };

    } catch (e) {
        console.error("updateWithdrawalDataFromWebApp Error: " + e.toString());
        return { success: false, message: e.message };
    } finally {
        lock.releaseLock();
    }
}

function deleteRecordByDocId(docId) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.withdrawalData.dataSheetName);
    const allData = dataSheet.getDataRange().getValues();
    const headers = allData.shift();
    
    const idCol = headers.indexOf('รหัสสินค้า'); 
    const qtyCol = headers.indexOf('จำนวน');
    const branchIdCol = headers.indexOf('รหัสสาขา');
    const branchNameCol = headers.indexOf('สาขา');
    const traysSentCol = headers.indexOf('แผงที่ส่ง');

    if (idCol === -1) throw new Error("ไม่พบคอลัมน์ 'รหัสสินค้า'");
    if (branchIdCol === -1) throw new Error("ไม่พบคอลัมน์ 'รหัสสาขา'");
    if (traysSentCol === -1) throw new Error("ไม่พบคอลัมน์ 'แผงที่ส่ง'");
    
    const itemsToDelete = allData.filter(row => row[0].toString().trim() === docId.toString().trim());
    let branchName = "";

    if (itemsToDelete.length > 0) {
        branchName = itemsToDelete[0][branchNameCol];
        const itemsToUpdate = itemsToDelete.map(row => ({
            id: row[idCol],
            quantityChange: Math.abs(Number(row[qtyCol]))
        })).filter(item => item.id);

        if (itemsToUpdate.length > 0) {
            updateStockLevels_(itemsToUpdate);
        }

        const branchId = itemsToDelete[0][branchIdCol];
        const traysSent = parseInt(itemsToDelete[0][traysSentCol], 10) || 0;

        if (traysSent > 0 && branchId) {
            updateContactTrayStock_(branchId.toString().trim(), "", -Math.abs(traysSent));
        }
    }

    deleteRowsByDocId_(dataSheet, docId);
    // [แก้] บันทึก Log
    logActivity_({
        action: 'ลบ',
        target: 'เอกสารเบิก',
        docId: docId,
        details: {
            contactName: branchName,
            itemCount: itemsToDelete.length
        }
    });
    
    clearServerCache();
    return { success: true, message: `ลบเอกสาร ${docId} และปรับปรุงข้อมูลคลัง/แผงสำเร็จ` };

  } catch (e) { 
      console.error("deleteRecordByDocId Error: " + e.message);
      return { success: false, message: `เกิดข้อผิดพลาดในการลบ: ${e.message}` }; 
  }
}




/**
 * [CORRECTED LOGIC] ดึงข้อมูลการเบิกสินค้า พร้อมแก้ไขตรรกะการจัดกลุ่มข้อมูล
 */
function getWithdrawalData(page = 1, rowsPerPage = 10, searchTerm = "") {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.withdrawalData.dataSheetName);
    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) return { data: [], totalItems: 0 };

    const allDataWithHeaders = dataSheet.getDataRange().getValues();
    const headers = allDataWithHeaders.shift();

    // ค้นหาตำแหน่งคอลัมน์จากชื่อ Header
    const docIdCol = headers.indexOf('เลขเอกสาร');
    const dateCol = headers.indexOf('วันที่');
    const requesterCol = headers.indexOf('ขอผู้เบิก');
    const departmentCol = headers.indexOf('สาขา');
    const approverCol = headers.indexOf('ผู้อนุมัติ');
    const traysSentCol = headers.indexOf('แผงที่ส่ง');
    const itemIdCol = headers.indexOf('รหัสสินค้า');
    const itemNameCol = headers.indexOf('สินค้า');
    const qtyCol = headers.indexOf('จำนวน');
    const unitCol = headers.indexOf('หน่วย');
    const noteCol = headers.indexOf('หมายเหตุ');

    const allData = allDataWithHeaders;

    const lowerCaseSearchTerm = searchTerm.trim().toLowerCase();
    const filteredData = lowerCaseSearchTerm
      ? allData.filter(row => 
          (row[docIdCol] && row[docIdCol].toString().toLowerCase().includes(lowerCaseSearchTerm)) ||
          (row[requesterCol] && row[requesterCol].toString().toLowerCase().includes(lowerCaseSearchTerm)) ||
          (row[approverCol] && row[approverCol].toString().toLowerCase().includes(lowerCaseSearchTerm))
        )
      : allData;

    // --- ตรรกะการจัดกลุ่มข้อมูลที่ถูกต้อง ---
    const groupedData = {};
    filteredData.forEach(row => {
      const docId = row[docIdCol];
      
      // 1. ถ้ายังไม่เคยเจอเอกสารนี้ ให้สร้างข้อมูลหลักขึ้นมาก่อน
      if (!groupedData[docId]) {
        groupedData[docId] = { 
          docId: docId, 
          date: new Date(row[dateCol]).toLocaleDateString('th-TH'), 
          requester: row[requesterCol], 
          department: row[departmentCol], 
          approver: row[approverCol], 
          traysSent: row[traysSentCol] || 0,
          items: [] // เริ่มต้นด้วยรายการสินค้าว่างๆ
        };
      }
      
      // 2. เพิ่มรายการสินค้าเข้าไปในเอกสารนั้นๆ (บรรทัดนี้ต้องอยู่นอก if เสมอ)
      groupedData[docId].items.push({ 
        id: row[itemIdCol],
        name: row[itemNameCol], 
        quantity: row[qtyCol], 
        unit: row[unitCol], 
        note: row[noteCol] 
      });
    });
    
    const allGroupedRecords = Object.values(groupedData).reverse();
    const totalItems = allGroupedRecords.length;

    const startIndex = (page - 1) * rowsPerPage;
    const paginatedData = allGroupedRecords.slice(startIndex, startIndex + rowsPerPage);

    return { data: paginatedData, totalItems: totalItems };

  } catch(e) {
    console.error("getWithdrawalData Error: " + e.toString());
    return { data: [], totalItems: 0, error: e.message };
  }
}

// ===================================================
// === 6. ฟังก์ชัน Web App: SORTING (คัดแยก) =========
// ===================================================
/**
 * [REVISED] ดึงข้อมูลทุกรายการจาก GRN พร้อมคำนวณยอดรวม
 * @param {string} grn_id - เลขที่เอกสาร GRN
 * @returns {object} Object ที่มีทั้งข้อมูลสรุป (summary) และข้อมูลรายตัว (items)
 */
function fetchReceiptForSortingWebApp(grn_id) {
    try {
        if (!grn_id) {
            throw new Error("กรุณาป้อนเลขเอกสารรับเข้า (GRN)");
        }
        
        const receiptDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.receiptData.dataSheetName);
        const allReceiptData = receiptDataSheet.getRange(2, 1, receiptDataSheet.getLastRow() - 1, 6).getValues(); // A:F
        
        // 1. กรองหาทุกแถวที่ตรงกับ GRN
        const foundRecords = allReceiptData.filter(row => row[0].toString().trim() === grn_id.toString().trim());
        
        if (foundRecords.length === 0) {
            throw new Error(`ไม่พบข้อมูลสำหรับเอกสารเลขที่ ${grn_id}`);
        }
        
        // 2. แปลงข้อมูลเป็น Array ของ Object (เหมือนเดิม)
        const items = foundRecords.map(record => ({
            itemName: record[4],
            quantity: Number(record[5]) || 0
        }));

        // 3. ***ส่วนที่เพิ่มเข้ามา: คำนวณยอดรวม***
        // (สมมติว่าสินค้าใน GRN ที่จะคัดแยกเป็นชนิดเดียวกันทั้งหมด)
        const totalQuantity = items.reduce((sum, currentItem) => sum + currentItem.quantity, 0);
        const summary = {
            itemName: items[0].itemName, // ใช้ชื่อสินค้าจากรายการแรก
            totalQuantity: totalQuantity
        };

        // 4. ส่งข้อมูลกลับในรูปแบบใหม่ที่มีทั้ง summary และ items
        return { success: true, data: { summary: summary, items: items } };

    } catch (e) {
        return { success: false, message: e.message };
    }
}

function getSortingHistory(page = 1, rowsPerPage = 10, searchTerm = "") {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sortingData.dataSheetName);
    if (dataSheet.getLastRow() < 2) return { data: [], totalItems: 0 };

    const allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();

    // 1. กรองข้อมูลตามคำค้นหา
    const lowerCaseSearchTerm = searchTerm.trim().toLowerCase();
    const filteredData = lowerCaseSearchTerm
      ? allData.filter(row =>
          row[0].toString().toLowerCase().includes(lowerCaseSearchTerm) || // เลขเอกสารคัดแยก
          row[2].toString().toLowerCase().includes(lowerCaseSearchTerm) || // เอกสารรับเข้าอ้างอิง
          row[3].toString().toLowerCase().includes(lowerCaseSearchTerm)    // สินค้าต้นทาง
        )
      : allData;

    // 2. จัดกลุ่มข้อมูลที่ผ่านการกรองแล้ว
    const groupedData = {};
    filteredData.forEach(row => {
      const docId = row[0];
      if (!groupedData[docId]) {
        groupedData[docId] = {
          docId: docId,
          date: new Date(row[1]).toLocaleDateString('th-TH'),
          refDocId: row[2],
          sourceItem: `${row[3]} (${row[4]})`,
          items: []
        };
      }
      groupedData[docId].items.push(`${row[5]} (${row[6]})`);
    });

    const allGroupedRecords = Object.values(groupedData).reverse();
    const totalItems = allGroupedRecords.length;

    // 3. แบ่งหน้าข้อมูล
    const startIndex = (page - 1) * rowsPerPage;
    const paginatedData = allGroupedRecords.slice(startIndex, startIndex + rowsPerPage);

    // 4. ส่งข้อมูลกลับในรูปแบบที่ถูกต้อง
    return { data: paginatedData, totalItems: totalItems };

  } catch(e) {
    console.error("getSortingHistory Error: " + e.toString());
    return { data: [], totalItems: 0, error: e.message };
  }
}

function saveSortingDataFromWebApp(formData) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sortingData.dataSheetName);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
    const timestamp = new Date();
    const docId = generateDocId_('SORT', CONFIG.sortingData.dataSheetName);

    if (dataSheet.getLastRow() === 0) {
      const headers = [['เลขเอกสารคัดแยก', 'วันที่คัดแยก', 'เอกสารรับเข้าอ้างอิง', 'สินค้าต้นทาง', 'จำนวนต้นทาง', 'สินค้าคัดแยก', 'จำนวนที่ได้', 'ผู้บันทึก']];
      dataSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    }

    const recordsToSave = formData.sortedItems.map(item => [
      docId, timestamp, formData.refDocId, 
      formData.sourceItem, formData.sourceQty, 
      item.name, item.quantity, currentUser
    ]);

    if (recordsToSave.length > 0) {
      dataSheet.getRange(dataSheet.getLastRow() + 1, 1, recordsToSave.length, recordsToSave[0].length).setValues(recordsToSave);
      
      // [แก้] บันทึก Log
      logActivity_({
          action: 'สร้าง',
          target: 'เอกสารคัดแยก',
          docId: docId,
          details: {
              refDocId: formData.refDocId,
              sourceItem: formData.sourceItem
          }
      });
      
      const productList = getProductList();
      const itemsToUpdate = [];

      const sourceProduct = productList.find(p => p.name === formData.sourceItem);
      if (sourceProduct) {
        itemsToUpdate.push({ id: sourceProduct.id, quantityChange: -Math.abs(Number(formData.sourceQty)) });
      } else { console.warn(`[Sorting] ไม่พบรหัสสินค้าสำหรับ '${formData.sourceItem}' จึงไม่ได้ลดสต็อก`); }

      formData.sortedItems.forEach(item => {
        const sortedProduct = productList.find(p => p.name === item.name);
        if (sortedProduct) {
          itemsToUpdate.push({ id: sortedProduct.id, quantityChange: Math.abs(Number(item.quantity)) });
        } else { console.warn(`[Sorting] ไม่พบรหัสสินค้าสำหรับ '${item.name}' จึงไม่ได้เพิ่มสต็อก`); }
      });
      
      if(itemsToUpdate.length > 0){
        updateStockLevels_(itemsToUpdate);
      }
      
      clearServerCache();
      return { success: true, docId: docId };

    } else { 
      throw new Error("ไม่พบรายการสินค้าคัดแยกที่จะบันทึก"); 
    }
  } catch (e) { 
    console.error("saveSortingDataFromWebApp Error: " + e.toString());
    return { success: false, message: e.message }; 
  }
}

/**
 * [CRITICAL BUG FIX] UPDATE: อัปเดตข้อมูลการคัดแยก พร้อมคืนสต็อกเก่าและตัดสต็อกใหม่
 */
function updateSortingDataFromWebApp(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.sortingData.dataSheetName);
    const allData = dataSheet.getDataRange().getValues();
    const headers = allData.shift();
    const docIdCol = 0;

    // --- 1. คืนสต็อกของเก่า ---
    const oldRecords = allData.filter(row => row[docIdCol].toString().trim() === formData.docId.toString().trim());
    
    if (oldRecords.length > 0) {
      const productList = getProductList();
      const itemsToReverse = [];

      // คืนสต็อกสินค้าต้นทาง (ทำให้เป็นบวก)
      const oldSourceItemName = oldRecords[0][3];
      const oldSourceQty = Number(oldRecords[0][4]);
      const oldSourceProduct = productList.find(p => p.name === oldSourceItemName);
      if (oldSourceProduct) {
        itemsToReverse.push({ id: oldSourceProduct.id, quantityChange: Math.abs(oldSourceQty) });
      }

      // คืนสต็อกสินค้าที่คัดแยกได้ (ทำให้เป็นลบ)
      oldRecords.forEach(row => {
        const oldSortedItemName = row[5];
        const oldSortedQty = Number(row[6]);
        const oldSortedProduct = productList.find(p => p.name === oldSortedItemName);
        if (oldSortedProduct) {
          itemsToReverse.push({ id: oldSortedProduct.id, quantityChange: -Math.abs(oldSortedQty) });
        }
      });
      
      if (itemsToReverse.length > 0) {
        updateStockLevels_(itemsToReverse);
      }
    }

    // --- 2. ตรวจสอบและตัดสต็อกใหม่ ---
    const productList = getProductList();
    const stockMap = getStockDataFromCache_();
    
    // ตรวจสอบว่าสินค้าต้นทางมีพอให้ตัดหรือไม่
    const sourceProduct = productList.find(p => p.name === formData.sourceItem);
    if (!sourceProduct) throw new Error(`ไม่พบสินค้าต้นทาง '${formData.sourceItem}' ในคลัง`);

    const stockInfo = stockMap.get(sourceProduct.id.trim());
    const currentStock = stockInfo ? Number(stockInfo.quantity) : 0;
    if (currentStock < Number(formData.sourceQty)) {
      throw new Error(`[แก้ไข] สต็อกไม่พอ! '${formData.sourceItem}' มีในคลัง ${currentStock} แต่ต้องการใช้ ${formData.sourceQty}`);
    }

    const itemsToApply = [];
    // ตัดสต็อกสินค้าต้นทาง (ค่าลบ)
    itemsToApply.push({ id: sourceProduct.id, quantityChange: -Math.abs(Number(formData.sourceQty)) });
    // เพิ่มสต็อกสินค้าที่คัดแยกได้ (ค่าบวก)
    formData.sortedItems.forEach(item => {
      const sortedProduct = productList.find(p => p.name === item.name);
      if (sortedProduct) {
        itemsToApply.push({ id: sortedProduct.id, quantityChange: Math.abs(Number(item.quantity)) });
      }
    });

    if (itemsToApply.length > 0) {
      updateStockLevels_(itemsToApply);
    }
    
    // --- 3. อัปเดตข้อมูลในชีตคัดแยก (Delete and Re-create) ---
    const rowsToKeep = allData.filter(row => row[docIdCol].toString().trim() !== formData.docId.toString().trim());
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
    const timestamp = new Date();
    
    const recordsToUpdate = formData.sortedItems.map(item => [
      formData.docId, timestamp,
      formData.refDocId, formData.sourceItem, formData.sourceQty,
      item.name, item.quantity, currentUser
    ]);
    
    const finalData = [headers, ...rowsToKeep, ...recordsToUpdate];
    dataSheet.clearContents();
    if (finalData.length > 0) {
      dataSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
    }

    // --- 4. บันทึก Log ---
    logActivity_({
      action: 'แก้ไข',
      target: 'เอกสารคัดแยก',
      docId: formData.docId,
      details: {
        refDocId: formData.refDocId,
        sourceItem: formData.sourceItem
      }
    });
    
    clearServerCache();
    return { success: true, docId: formData.docId };

  } catch (e) {
    console.error("updateSortingDataFromWebApp Error: " + e.toString());
    // หมายเหตุ: หากเกิด Error ระหว่างทาง ควรมีการทำ Rollback แต่ใน Apps Script จะซับซ้อนมาก
    // การคืนสต็อกเก่าไปก่อนเป็นวิธีลดความเสี่ยงเบื้องต้น
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteSortingByDocId(docId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sortingData.dataSheetName);
    const allData = dataSheet.getDataRange().getValues();
    const headers = allData.shift();

    const itemsToDelete = allData.filter(row => row[0].toString().trim() === docId.toString().trim());
    let refDocId = "";
    let sourceItem = "";

    if (itemsToDelete.length > 0) {
        refDocId = itemsToDelete[0][2];
        sourceItem = itemsToDelete[0][3];
      const productList = getProductList();
      const itemsToReverse = [];

      const sourceItemName = itemsToDelete[0][3];
      const sourceQty = Number(itemsToDelete[0][4]);
      const sourceProduct = productList.find(p => p.name === sourceItemName);
      if (sourceProduct) {
        itemsToReverse.push({ id: sourceProduct.id, quantityChange: Math.abs(sourceQty) });
      }

      itemsToDelete.forEach(row => {
        const sortedItemName = row[5];
        const sortedQty = Number(row[6]);
        const sortedProduct = productList.find(p => p.name === sortedItemName);
        if (sortedProduct) {
          itemsToReverse.push({ id: sortedProduct.id, quantityChange: -Math.abs(sortedQty) });
        }
      });

      if (itemsToReverse.length > 0) {
        const stockMap = getStockDataFromCache_(); 

        for (const item of itemsToReverse) {
            const productInfo = stockMap.get(item.id.toString().trim());
            const currentQty = productInfo ? Number(productInfo.quantity) : 0;
            const projectedQty = currentQty + item.quantityChange;

            if (projectedQty < 0) {
                const productName = (productList.find(p => p.id === item.id) || { name: item.id }).name;
                throw new Error(`ไม่สามารถลบได้! หากลบแล้ว สินค้า '${productName}' จะติดลบ (${projectedQty.toLocaleString()})`);
            }
        }
        
        updateStockLevels_(itemsToReverse);
      }
    }

    deleteRowsByDocId_(dataSheet, docId);
    // [แก้] บันทึก Log
    logActivity_({
        action: 'ลบ',
        target: 'เอกสารคัดแยก',
        docId: docId,
        details: {
            refDocId: refDocId,
            sourceItem: sourceItem
        }
    });
    
    clearServerCache();
    return { success: true, message: `ลบเอกสารคัดแยก ${docId} และปรับสต็อกคืนสำเร็จ` };

  } catch (e) { 
    console.error("deleteSortingByDocId Error: " + e.message);
    return { success: false, message: e.message }; 
  } finally {
    lock.releaseLock();
  }
}




// ===================================================
// === 7. ฟังก์ชัน Web App: RANDOM CHECK (สุ่มเช็ค) ======
// ===================================================
function getAllCheckData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.randomCheckData.dataSheetName);
    if (!dataSheet || dataSheet.getLastRow() < 2) {
      return {};
    }
    const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
    const checksByReceiptId = {};
    data.forEach(row => {
      const refDocId = row[2];
      if (!checksByReceiptId[refDocId]) {
        checksByReceiptId[refDocId] = [];
      }
      checksByReceiptId[refDocId].push({
        checkId: row[0], timestamp: row[1], refDocId: row[2],
        itemName: row[3], docWeight: row[4], actualWeight: row[5],
        docQuantity: row[7], actualQuantity: row[8], checkResult: row[10],
        notes: row[11], checkerName: row[12]
      });
    });
    return checksByReceiptId;
  } catch (e) {
    console.error("getAllCheckData Error: " + e.toString());
    return { error: e.message };
  }
}

function saveOrUpdateCheckData(formData) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.randomCheckData.dataSheetName);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      if (dataSheet.getLastRow() < 1) {
        const headers = [['ID สุ่มเช็ค', 'เวลาที่บันทึก', 'เอกสารรับเข้าอ้างอิง', 'รายการสินค้า', 'น้ำหนักตามเอกสาร', 'น้ำหนักจริง', 'ผลต่างน้ำหนัก', 'จำนวนตามเอกสาร', 'จำนวนจริง', 'ผลต่างจำนวน', 'ผลการตรวจสอบ', 'หมายเหตุ', 'ชื่อผู้สุ่มนับ', 'ผู้บันทึก']];
        dataSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
      }
      const allData = dataSheet.getDataRange().getValues();
      const headers = allData.shift();
      const rowsToKeep = allData.filter(row => row[2] && row[2].toString().trim() !== formData.refDocId.toString().trim());
      const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
      const timestamp = new Date();
      const newRowsToAdd = [];
      formData.items.forEach(item => {
        if (item.actualQuantity || item.actualWeight) {
          const checkId = generateDocId_('CHK', CONFIG.randomCheckData.dataSheetName);
          const docWeight = parseFloat(item.docWeight) || 0;
          const actualWeight = parseFloat(item.actualWeight) || 0;
          const docQuantity = parseFloat(item.docQuantity) || 0;
          const actualQuantity = parseFloat(item.actualQuantity) || 0;
          const weightDiff = (docWeight > 0 || actualWeight > 0) ? (actualWeight - docWeight) : '';
          const quantityDiff = (docQuantity > 0 || actualQuantity > 0) ? (actualQuantity - docQuantity) : '';
          newRowsToAdd.push([
            checkId, timestamp, formData.refDocId, item.itemName,
            item.docWeight, item.actualWeight, weightDiff,
            item.docQuantity, item.actualQuantity, quantityDiff,
            item.checkResult, item.notes, formData.checkerName, currentUser
          ]);
        }
      });
      const finalData = [headers, ...rowsToKeep, ...newRowsToAdd];
      dataSheet.clearContents();
      if (finalData.length > 0) {
        dataSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
      }
      
      // [แก้] บันทึก Log
      logActivity_({
          action: 'บันทึก/อัปเดต',
          target: 'เอกสารสุ่มเช็ค',
          docId: formData.refDocId,
          details: { checkerName: formData.checkerName }
      });

      clearServerCache();
      return { success: true, docId: formData.refDocId };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    console.error("saveOrUpdateCheckData Error: " + e.toString());
    return { success: false, message: e.message };
  }
}

function deleteRandomCheckData(checkId) {
  try {
    if (!checkId) throw new Error("ไม่พบ Check ID");
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.randomCheckData.dataSheetName);
    deleteRowsByDocId_(dataSheet, checkId, 0);
    
    // [แก้] บันทึก Log
    logActivity_({
        action: 'ลบ',
        target: 'เอกสารสุ่มเช็ค',
        docId: checkId,
        details: {}
    });

    clearServerCache();
    return { success: true, message: `ลบข้อมูลการเช็ค ${checkId} สำเร็จ` };
  } catch (e) {
    console.error("deleteRandomCheckData Error: " + e.toString());
    return { success: false, message: e.message };
  }
}



function getCheckHistoryList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.randomCheckData.dataSheetName);
    if (!dataSheet || dataSheet.getLastRow() < 2) {
      return [];
    }
    const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
    return data.map(row => ({
      checkId: row[0], timestamp: new Date(row[1]).toLocaleDateString('th-TH'), refDocId: row[2],
      itemName: row[3], docWeight: row[4], actualWeight: row[5], weightDiff: row[6],
      docQuantity: row[7], actualQuantity: row[8], quantityDiff: row[9], checkResult: row[10],
      notes: row[11], checkerName: row[12]
    })).reverse();
  } catch (e) {
    console.error("getCheckHistoryList Error: " + e.toString());
    return [];
  }
}

// ===================================================
// === 8. ฟังก์ชัน Web App: MAINTENANCE (ซ่อมบำรุง) ====
// ===================================================
function saveMaintenanceData(formData) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.maintenanceData.dataSheetName);
    const currentUser = Session.getActiveUser().getEmail() || 'Unknown User';
    const timestamp = new Date();
    const docId = generateDocId_('MA', CONFIG.maintenanceData.dataSheetName);
    dataSheet.appendRow([
      docId, new Date(formData.date), formData.vehicleId, formData.mileage,
      formData.type, formData.details, formData.cost, currentUser, timestamp
    ]);

    // ++ เคลียร์ Cache ++
    clearServerCache();
    return { success: true, docId: docId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getMaintenanceHistory() {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.maintenanceData.dataSheetName);
    if (dataSheet.getLastRow() < 2) return [];
    const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();
    return data.map(row => ({
      docId: row[0], date: new Date(row[1]).toLocaleDateString('th-TH'), vehicleId: row[2],
      mileage: row[3], type: row[4], details: row[5], cost: row[6]
    })).reverse();
  } catch (e) {
    return [];
  }
}

function updateMaintenanceData(formData) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.maintenanceData.dataSheetName);
    const data = dataSheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0] === formData.docId);
    if (rowIndex === -1) throw new Error("ไม่พบข้อมูลที่จะแก้ไข");
    const updatedRow = [
      formData.docId, new Date(formData.date), formData.vehicleId, formData.mileage,
      formData.type, formData.details, formData.cost,
      data[rowIndex][7], new Date()
    ];
    dataSheet.getRange(rowIndex + 1, 1, 1, updatedRow.length).setValues([updatedRow]);

    // ++ เคลียร์ Cache ++
    clearServerCache();
    return { success: true, docId: formData.docId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteMaintenanceData(docId) {
  try {
    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.maintenanceData.dataSheetName);
    deleteRowsByDocId_(dataSheet, docId);

    // ++ เคลียร์ Cache ++
    clearServerCache();
    return { success: true, message: `ลบข้อมูล ${docId} สำเร็จ` };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===================================================
// === 9. ฟังก์ชันเสริมที่ใช้ร่วมกัน (Shared Helpers) ===
// ===================================================
function generateDocId_(prefix, sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const today = new Date();
    const datePart = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    const datePrefix = `${prefix}-${datePart}-`;
    if (!sheet || sheet.getLastRow() < 2) {
      return datePrefix + "1";
    }
    const allDocIds = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    let maxNum = 0;
    allDocIds.forEach(id => {
      if (String(id).startsWith(datePrefix)) {
        const num = parseInt(id.split('-')[2]);
        if (num > maxNum) maxNum = num;
      }
    });
    return datePrefix + (maxNum + 1);
  } catch (e) {
    console.error(`generateDocId_ Error for prefix ${prefix}: ${e.toString()}`);
    return `${prefix}-${Date.now()}`;
  }
}

/**
 * [OPTIMIZED] ลบแถวทั้งหมดที่มี ID ที่ตรงกันในคอลัมน์ที่กำหนด
 */
function deleteRowsByDocId_(sheet, docId, idColumnIndex = 0) {
  if (!sheet || !docId) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const allData = sheet.getDataRange().getValues();
  const headers = allData.shift();
  const trimmedDocId = docId.toString().trim();

  const rowsToKeep = allData.filter(row => row[idColumnIndex].toString().trim() !== trimmedDocId);
  
  sheet.clearContents();
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rowsToKeep.length > 0) {
    sheet.getRange(2, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  }
}

/**
 * [NEW & IMPROVED] บันทึกกิจกรรมของผู้ใช้ลงในชีต 'Log' ในรูปแบบที่มนุษย์อ่านเข้าใจง่าย
 * @param {object} logData - อ็อบเจกต์ข้อมูลสำหรับบันทึก Log
 * { action: string, target: string, docId: string, details: object }
 * ตัวอย่าง: { action: 'สร้าง', target: 'เอกสารรับเข้า', docId: 'GRN-123', details: { contactName: 'ฟาร์ม A', itemCount: 5 } }
 */
function logActivity_(logData) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.logSheet);
    if (!logSheet) return;

    // หากเป็นชีต Log ใหม่ ให้สร้าง Header ก่อน
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange("A1:C1").setValues([['Timestamp', 'User', 'Activity Description']]).setFontWeight('bold');
      logSheet.setColumnWidth(3, 450); // ขยายคอลัมน์ Description ให้กว้างขึ้น
    }

    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail() || 'Unknown User';
    let message = '';
    const { action, target, docId, details } = logData;

    // สร้างข้อความ Log ตามประเภทของข้อมูล
    switch (target) {
      case 'เอกสารรับเข้า':
        message = `✅ ${action}${target} #${docId} จาก "${details.contactName}" (${details.itemCount} รายการ)`;
        break;
      case 'เอกสารเบิก':
        message = `📤 ${action}${target} #${docId} ไปยัง "${details.contactName}" (${details.itemCount} รายการ)`;
        break;
      case 'เอกสารคัดแยก':
        message = `✨ ${action}${target} #${docId} (อ้างอิง: ${details.refDocId}) จาก "${details.sourceItem}"`;
        break;
      case 'เอกสารสุ่มเช็ค':
        message = `📋 ${action}ผล${target} สำหรับ #${docId} โดยคุณ "${details.checkerName}"`;
        break;
      case 'สินค้า':
        if (action === 'เพิ่ม') {
          message = `📦 ${action}${target}ใหม่: ${details.productName} (รหัส: ${docId})`;
        } else if (action === 'ลบ') {
          message = `🗑️ ${action}${target}: ${docId}`;
        } else { // แก้ไข
          message = `✏️ ${action}ข้อมูล${target}: ${details.productName} (รหัส: ${docId})`;
        }
        break;
      case 'ผู้ติดต่อ':
         message = `👤 ${action}${target}ใหม่: ${details.contactName} (ประเภท: ${details.type})`;
         break;
      case 'ยอดแผงคงค้าง':
         const quantity = details.quantity > 0 ? `+${details.quantity}` : details.quantity;
         message = `🔄 อัปเดต${target} ของ "${details.contactName}" จำนวน ${quantity} แผง (ยอดใหม่: ${details.newBalance})`;
         break;
      default:
        message = `${action} ${target} #${docId}`;
    }

    logSheet.appendRow([timestamp, userEmail, message]);
  } catch (e) {
    console.error("Failed to write activity log: " + e.toString());
  }
}

/**
 * [NEW] ฟังก์ชันสำหรับบันทึกประวัติการเคลื่อนไหวของแผงไข่ (แยกจาก updateContactTrayStock_)
 */
function logTrayUpdate_(contactId, contactName, quantityChange, newBalance) {
  try {
    logActivity_({ 
        action: 'อัปเดต', 
        target: 'ยอดแผงคงค้าง', 
        docId: contactId, 
        details: {
            contactName: contactName,
            quantity: quantityChange,
            newBalance: newBalance
        }
    });
  } catch(e) {
    console.error("logTrayUpdate_ Error: " + e.toString());
  }
}

function clearServerCache() {
  try {
    // เพิ่ม 'allContactsData' เข้าไปในรายการที่จะลบเพื่อให้ครอบคลุม
    CacheService.getScriptCache().removeAll(['productList', 'allowedEmails', 'fullStockData', 'allContactsData']);
    console.log("Server cache cleared for: productList, allowedEmails, fullStockData, allContactsData");
    return { success: true, message: 'ล้างแคชฝั่งเซิร์ฟเวอร์สำเร็จ' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


function returnEggTrays(supplierId, quantity) {
    try {
        if (!supplierId || !quantity) {
            throw new Error("ข้อมูลไม่ครบถ้วน");
        }
        
        const qtyToSubtract = -Math.abs(parseInt(quantity, 10));
        updateContactTrayStock_(supplierId, '', qtyToSubtract);
        
        // ++ เคลียร์ Cache ++
        clearServerCache();
        return { success: true, message: `บันทึกการคืนแผงจำนวน ${Math.abs(qtyToSubtract)} แผงสำเร็จ` };
    } catch (e) {
        return { success: false, message: e.message };
    }
}


/**
 * [NEW] ดึงข้อมูล 'ผู้ติดต่อ' ทั้งหมด (Supplier และ Branch) พร้อมยอดแผงคงค้าง
 */
function getContactDashboardData() {
  try {
    // 1. ดึงรายชื่อ 'ผู้ติดต่อ' ทั้งหมดจากฟังก์ชันกลาง (ไม่กรองประเภท)
    const contacts = getContacts_(); 

    if (!contacts || contacts.length === 0) {
      return [];
    }
    
    // 2. กำหนดค่ายอดแผงเริ่มต้นให้ทุกคนเป็น 0
    contacts.forEach(c => c.trayBalance = 0);

    // 3. ดึงข้อมูลยอดแผงคงค้างทั้งหมด (ถ้ามี)
    const traySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TrayStock");
    if (traySheet && traySheet.getLastRow() >= 2) {
      const trayData = traySheet.getRange("A2:C" + traySheet.getLastRow()).getValues();
      const trayBalanceMap = new Map();
      trayData.forEach(row => {
        const id = row[0];
        const balance = parseInt(row[2], 10) || 0;
        if (id) {
          trayBalanceMap.set(id.toString().trim(), balance);
        }
      });

      // 4. นำยอดแผงไปรวมกับข้อมูลผู้ติดต่อ
      contacts.forEach(c => {
        if (trayBalanceMap.has(c.id.toString().trim())) {
          c.trayBalance = trayBalanceMap.get(c.id.toString().trim());
        }
      });
    }
    
    // 5. เรียงลำดับข้อมูลจากยอดคงค้างมากไปน้อย
    contacts.sort((a, b) => b.trayBalance - a.trayBalance);
    
    return contacts;

  } catch (e) {
    console.error("getContactDashboardData Error: " + e.toString());
    return { error: e.message };
  }
}

// ===================================================
// === 9. ฟังก์ชันเสริมที่ใช้ร่วมกัน (Shared Helpers) ===
// ===================================================

/**
 * [NEW] ฟังก์ชันกลางสำหรับอัปเดตสต็อกสินค้าในชีต 'คลัง'
 * @param {Array<Object>} itemsToUpdate - อาร์เรย์ของอ็อบเจกต์สินค้า [{id: 'รหัส', quantityChange: จำนวน}]
 * (จำนวนเป็นลบสำหรับการเบิก, เป็นบวกสำหรับการรับ/คัดแยกได้)
 */
function updateStockLevels_(itemsToUpdate) {
  if (!itemsToUpdate || itemsToUpdate.length === 0) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // ล็อกสคริปต์ ป้องกันการชนกันของข้อมูล

  try {
    const stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.stockSheet);
    if (!stockSheet) throw new Error("ไม่พบชีต 'คลัง'");

    const headers = stockSheet.getRange(1, 1, 1, stockSheet.getLastColumn()).getValues()[0];
    const idColIndex = headers.indexOf("รหัสสินค้า");
    const qtyColIndex = headers.indexOf("คลังกลาง/ฟอง");

    if (idColIndex === -1 || qtyColIndex === -1) {
      throw new Error("ไม่พบคอลัมน์ 'รหัสสินค้า' หรือ 'คลังกลาง/ฟอง' ในชีต 'คลัง'");
    }

    const stockData = stockSheet.getRange(2, idColIndex + 1, stockSheet.getLastRow() - 1, 1).getValues().flat();
    const stockMap = new Map(stockData.map((id, index) => [id.toString().trim(), index + 2])); // Map[ProductID -> RowIndex]

    for (const item of itemsToUpdate) {
      const row = stockMap.get(item.id.toString().trim());
      if (row) {
        const qtyCell = stockSheet.getRange(row, qtyColIndex + 1);
        const currentQty = Number(qtyCell.getValue()) || 0;
        qtyCell.setValue(currentQty + item.quantityChange); // อัปเดตยอดใหม่
      } else {
        console.warn(`ไม่พบรหัสสินค้า '${item.id}' ในชีต 'คลัง' เพื่ออัปเดตสต็อก`);
      }
    }

    // --- !! สำคัญมาก !! เคลียร์ Cache ของสต็อกทิ้ง ---
    clearServerCache();
    console.log("Stock levels updated and cache cleared.");

  } catch (e) {
    console.error("updateStockLevels_ Error: " + e.toString());
    // อาจจะต้องมีการแจ้งเตือนผู้ใช้หรือบันทึก Log ที่ละเอียดขึ้น
    throw e; // ส่ง error ต่อเพื่อให้ฟังก์ชันที่เรียกใช้รู้ว่ามีปัญหา
  } finally {
    lock.releaseLock(); // ปลดล็อกทุกครั้ง
  }
}


function updateContactTrayStock_(contactId, contactName, quantity) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TrayStock");
        if (!sheet) return;

        if (sheet.getLastRow() === 0) {
            sheet.getRange("A1:C1").setValues([["ContactID", "ContactName", "TrayBalance"]]);
        }

        const data = sheet.getRange("A2:C" + (sheet.getLastRow() || 1)).getValues();
        let contactFound = false;
        let newBalance = 0;

        for (let i = 0; i < data.length; i++) {
            if (data[i][0].toString().trim() === contactId.toString().trim()) {
                const currentBalance = parseInt(data[i][2], 10) || 0;
                newBalance = currentBalance + quantity;
                sheet.getRange(i + 2, 3).setValue(newBalance);
                contactFound = true;
                
                const finalContactName = contactName || data[i][1];
                // [แก้] บันทึก Log การเคลื่อนไหวของแผง
                logTrayUpdate_(contactId, finalContactName, quantity, newBalance);
                break;
            }
        }

        if (!contactFound) {
            newBalance = quantity;
            sheet.appendRow([contactId, contactName, newBalance]);
            // [แก้] บันทึก Log การเคลื่อนไหวของแผง
            logTrayUpdate_(contactId, contactName, quantity, newBalance);
        }
    } finally {
        lock.releaseLock();
    }
}

function addNewContact(formData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.webAppInfo.contactsSheet);
    if (!sheet) throw new Error("ไม่พบชีต 'Contacts'");
    
    if (sheet.getLastRow() === 0) {
      sheet.getRange("A1:D1").setValues([["ContactID", "ContactName", "Type", "Tel"]]);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let newId;
    const contactIdCol = headers.indexOf("ContactID") + 1;
    const allContactIds = contactIdCol > 0 ? sheet.getRange(2, contactIdCol, sheet.getLastRow() - 1 || 1, 1).getValues().flat() : [];

    if (formData.type === 'Supplier') {
      const lastRow = sheet.getLastRow();
      const prefix = 'SUP';
      newId = `${prefix}${String(lastRow + 1).padStart(4, '0')}`;
    } else if (formData.type === 'Branch') {
      newId = formData.contactId.trim();
      if (!newId) throw new Error("กรุณากรอกรหัสสาขา");
      if (allContactIds.includes(newId)) throw new Error(`รหัสสาขา '${newId}' นี้มีอยู่แล้วในระบบ`);
    } else {
      throw new Error("ประเภทผู้ติดต่อไม่ถูกต้อง");
    }

    const phoneNumberAsText = formData.tel.trim() ? "'" + formData.tel.trim() : "";
    
    sheet.appendRow([ newId, formData.name.trim(), formData.type, phoneNumberAsText ]);
    
    // [แก้] บันทึก Log
    logActivity_({
        action: 'เพิ่ม',
        target: 'ผู้ติดต่อ',
        docId: newId,
        details: { contactName: formData.name.trim(), type: formData.type }
    });

    clearServerCache();
    return { success: true, message: `เพิ่ม '${formData.name}' สำเร็จ` };

  } catch (e) {
    console.error("addNewContact Error: " + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

