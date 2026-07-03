/**
 * Purchase Order Management System
 */
const CONFIG = { SHEET_NAMES: { USERS: 'Users', VENDORS: 'Vendors', PURCHASE_ORDERS: 'PurchaseOrders', PO_LINE_ITEMS: 'POLineItems', AUDIT_LOG: 'AuditLog' }, STATUS: { DRAFT: 'Draft', PENDING_MANAGER: 'Pending Manager', PENDING_FINANCE: 'Pending Finance', APPROVED: 'Approved', ORDERED: 'Ordered', PARTIAL: 'Partial Received', RECEIVED: 'Received', REJECTED: 'Rejected' }, THRESHOLD: 10000 };
const SHEET_CONFIG = const SHEET_CONFIG = {
  Users: { columns: ['Email', 'Name', 'Role', 'Department', 'Active'] },
  Vendors: { columns: ['VendorID', 'Name', 'ContactPerson', 'Email', 'Phone', 'Address', 'PaymentTerms', 'Rating', 'Active'] },
  PurchaseOrders: { columns: ['POID', 'VendorID', 'Requester', 'Department', 'Status', 'TotalAmount', 'ExpectedDelivery', 'CreatedAt', 'ApprovedBy', 'ApprovedAt', 'Notes'] },
  POLineItems: { columns: ['LineID', 'POID', 'ItemCode', 'Description', 'Quantity', 'UnitPrice', 'Total', 'ReceivedQty'] },
  AuditLog: { columns: ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue'] }
};;

// ============== ENHANCED UTILITIES (v2.0) ==============
const VERSION = '2.0.0';

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  for (const [sheetName, config] of Object.entries(SHEET_CONFIG)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      created.push(sheetName);
      sheet.getRange(1, 1, 1, config.columns.length).setValues([config.columns]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      if (config.sampleData) config.sampleData.forEach(row => sheet.appendRow(row));
    }
  }
  Logger.log('InitializeSheets: Created ' + created.join(', '));
  return { created };
}

function handleError(error, context) {
  const errorMsg = error instanceof Error ? error.message : String(error);
  Logger.log('[ERROR] ' + context + ': ' + errorMsg);
  if (typeof logAction === 'function') logAction('ERROR', context, '', errorMsg);
  return { success: false, error: errorMsg };
}

function backupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = 'Backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  ss.copy(backupName);
  if (typeof logAction === 'function') logAction('BACKUP_CREATED', 'System', '', backupName);
  return { success: true, backupName };
}

function exportToPDF(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  const pdf = DriveApp.getFileById(ss.getId()).getAs('application/pdf');
  const pdfName = sheet.getName() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
  DriveApp.getRootFolder().createFile(pdf).setName(pdfName);
  return { success: true, fileName: pdfName };
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 System Menu')
    .addItem('📊 Initialize Sheets', 'initializeSheets')
    .addItem('💾 Create Backup', 'backupData')
    .addItem('📄 Export to PDF', 'exportToPDF')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('System v' + VERSION, 'Enhanced with:\n- initializeSheets()\n- backupData()\n- exportToPDF()', ui.ButtonSet.OK);
}

function getSheet(name) { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName(name); if (!sheet) { sheet = ss.insertSheet(name); setupHeaders(sheet, name); } return sheet; }
function setupHeaders(sheet, name) { const h = { Users: ['Email', 'Name', 'Role', 'Department', 'Active'], Vendors: ['VendorID', 'Name', 'ContactPerson', 'Email', 'Phone', 'Address', 'PaymentTerms', 'Rating', 'Active'], PurchaseOrders: ['POID', 'VendorID', 'Requester', 'Department', 'Status', 'TotalAmount', 'ExpectedDelivery', 'CreatedAt', 'ApprovedBy', 'ApprovedAt', 'Notes'], POLineItems: ['LineID', 'POID', 'ItemCode', 'Description', 'Quantity', 'UnitPrice', 'Total', 'ReceivedQty'], AuditLog: ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue'] }; if (h[name]) { sheet.getRange(1, 1, 1, h[name].length).setValues([h[name]]); sheet.getRange(1, 1, 1, h[name].length).setFontWeight('bold'); } }
function generateId(prefix) { return prefix + '-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0, 6).toUpperCase(); }
function getCurrentUser() { const email = Session.getActiveUser().getEmail(); const sheet = getSheet(CONFIG.SHEET_NAMES.USERS); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0].toLowerCase() === email.toLowerCase()) return { email: data[i][0], name: data[i][1], role: data[i][2], department: data[i][3] }; } return { email, name: 'Unknown', role: 'Employee' }; }
function requireRole(roles) { const user = getCurrentUser(); if (!roles.includes(user.role)) throw new Error('Unauthorized'); return user; }
function logAction(action, recordId, oldValue, newValue) { const sheet = getSheet(CONFIG.SHEET_NAMES.AUDIT_LOG); sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, recordId, oldValue, newValue]); }
function addVendor(name, contactPerson, email, phone, address, paymentTerms) { requireRole(['Admin', 'Manager']); const vendorID = generateId('VND'); const sheet = getSheet(CONFIG.SHEET_NAMES.VENDORS); sheet.appendRow([vendorID, name, contactPerson, email, phone, address, paymentTerms, 5, true]); logAction('ADD_VENDOR', vendorID, '', name); return { success: true, vendorID }; }
function createPO(vendorID, department, expectedDelivery, notes, lineItems) { requireRole(['Employee', 'Manager']); const user = getCurrentUser(); const poID = generateId('PO'); let total = 0; lineItems.forEach(item => { total += item.quantity * item.unitPrice; }); const sheet = getSheet(CONFIG.SHEET_NAMES.PURCHASE_ORDERS); let status = CONFIG.STATUS.PENDING_MANAGER; sheet.appendRow([poID, vendorID, user.email, department, status, total, expectedDelivery ? new Date(expectedDelivery) : '', new Date(), '', '', notes || '']); const lineSheet = getSheet(CONFIG.SHEET_NAMES.PO_LINE_ITEMS); lineItems.forEach(item => { lineSheet.appendRow([generateId('LI'), poID, item.code, item.description, item.quantity, item.unitPrice, item.quantity * item.unitPrice, 0]); }); logAction('CREATE_PO', poID, '', status); return { success: true, poID, total }; }
function approveByManager(poID) { requireRole(['Manager', 'Admin']); const sheet = getSheet(CONFIG.SHEET_NAMES.PURCHASE_ORDERS); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === poID) { const row = i + 1; const total = data[i][5]; let newStatus = CONFIG.STATUS.APPROVED; if (total > CONFIG.THRESHOLD) newStatus = CONFIG.STATUS.PENDING_FINANCE; sheet.getRange(row, 5, 1, 1).setValue(newStatus); sheet.getRange(row, 9, 1, 1).setValue(Session.getActiveUser().getEmail()); sheet.getRange(row, 10, 1, 1).setValue(new Date()); logAction('MANAGER_APPROVE_PO', poID, CONFIG.STATUS.PENDING_MANAGER, newStatus); return { success: true }; } } throw new Error('PO not found'); }
function approveByFinance(poID) { requireRole(['Finance Admin', 'Admin']); const sheet = getSheet(CONFIG.SHEET_NAMES.PURCHASE_ORDERS); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === poID) { const row = i + 1; sheet.getRange(row, 5, 1, 1).setValue(CONFIG.STATUS.APPROVED); sheet.getRange(row, 9, 1, 1).setValue(Session.getActiveUser().getEmail()); sheet.getRange(row, 10, 1, 1).setValue(new Date()); logAction('FINANCE_APPROVE_PO', poID, CONFIG.STATUS.PENDING_FINANCE, CONFIG.STATUS.APPROVED); return { success: true }; } } throw new Error('PO not found'); }
function markAsReceived(poID, lineItems) { requireRole(['Manager', 'Admin', 'Employee']); const sheet = getSheet(CONFIG.SHEET_NAMES.PURCHASE_ORDERS); const lineSheet = getSheet(CONFIG.SHEET_NAMES.PO_LINE_ITEMS); const lineData = lineSheet.getDataRange().getValues(); let allReceived = true; lineItems.forEach(received => { for (let i = 1; i < lineData.length; i++) { if (lineData[i][0] === received.lineID) { lineSheet.getRange(i + 1, 8, 1, 1).setValue(received.qty); if (received.qty < lineData[i][4]) allReceived = false; break; } } }); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === poID) { sheet.getRange(i + 1, 5, 1, 1).setValue(allReceived ? CONFIG.STATUS.RECEIVED : CONFIG.STATUS.PARTIAL); logAction('RECEIVE_PO', poID, CONFIG.STATUS.APPROVED, allReceived ? CONFIG.STATUS.RECEIVED : CONFIG.STATUS.PARTIAL); return { success: true }; } } throw new Error('PO not found'); }
function getAllPOs() { const sheet = getSheet(CONFIG.SHEET_NAMES.PURCHASE_ORDERS); const data = sheet.getDataRange().getValues(); const pos = []; for (let i = 1; i < data.length; i++) { pos.push({ poID: data[i][0], vendorID: data[i][1], requester: data[i][2], department: data[i][3], status: data[i][4], totalAmount: data[i][5], expectedDelivery: data[i][6], createdAt: data[i][7], approvedBy: data[i][8], approvedAt: data[i][9], notes: data[i][10] }); } return pos; }
function getPOLineItems(poID) { const sheet = getSheet(CONFIG.SHEET_NAMES.PO_LINE_ITEMS); const data = sheet.getDataRange().getValues(); const items = []; for (let i = 1; i < data.length; i++) { if (data[i][1] === poID) items.push({ lineID: data[i][0], itemCode: data[i][2], description: data[i][3], quantity: data[i][4], unitPrice: data[i][5], total: data[i][6], receivedQty: data[i][7] }); } return items; }
function getVendor(vendorID) { const sheet = getSheet(CONFIG.SHEET_NAMES.VENDORS); const data = sheet.getDataRange().getValues(); for (let i = 1; i < data.length; i++) { if (data[i][0] === vendorID) return { vendorID: data[i][0], name: data[i][1], contactPerson: data[i][2], email: data[i][3], phone: data[i][4], address: data[i][5], paymentTerms: data[i][6] }; } return null; }
function getAllVendors() { const sheet = getSheet(CONFIG.SHEET_NAMES.VENDORS); const data = sheet.getDataRange().getValues(); const vendors = []; for (let i = 1; i < data.length; i++) { vendors.push({ vendorID: data[i][0], name: data[i][1], contactPerson: data[i][2], email: data[i][3], phone: data[i][4] }); } return vendors; }
function doGet(e) { return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Purchase Order Management').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
