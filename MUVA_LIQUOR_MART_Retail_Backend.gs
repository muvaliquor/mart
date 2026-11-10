/**
 * MUVA LIQUOR MART — RETAIL GOOGLE SHEETS BACKEND
 * Bound to the Google Sheet used by the shop.
 *
 * Sheets created automatically:
 *   Products
 *   Sales
 *   SaleLines
 *   StockLog
 *
 * Staff (exactly 4):
 *   MUVA   / password set in Script Properties (MUVA_ADMIN_PASSWORD)
 *   Rebeka / existing password retained (default rebeka123)
 *   Isaya  / existing password retained (default isaya123)
 *   Lameck / existing password retained (default lameck123)
 *
 * Deploy as Web app:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * IMPORTANT: keep this Apps Script project private to shop administrators.
 *
 * UPDATE: Reports now support a custom date range (period=range&from=&to=),
 * a single day, a whole month, plus salesCount and best-sellers-first ordering.
 * Reporting is read-only: no business data is changed.
 */

const DEFAULT_STAFF = {
  MUVA: null, // Set MUVA_ADMIN_PASSWORD in Apps Script Script Properties.
  Rebeka: 'rebeka123',
  Isaya: 'isaya123',
  Lameck: 'lameck123'
};

const STAFF_DEFAULTS_VERSION = '2026-11-07-MUVA-RETAIL-01';
const ADMIN_NAME = 'MUVA';


const CACHE_SECONDS = 21600; // 6 hours
const SHEET_NAMES = {
  PRODUCTS: 'Products',
  SALES: 'Sales',
  LINES: 'SaleLines',
  STOCKLOG: 'StockLog',
  STAFF: 'Staff',
  DEVELOPER_ACCESS: 'DeveloperAccess'
};

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '');
    const token = String((e && e.parameter && e.parameter.token) || '');
    requireAuth_(token);

    switch (action) {
      case 'products':
        return json_({ok:true, products:listProducts_()});
      case 'dashboard':
        return json_({ok:true, data:dashboard_(String(e.parameter.customerType || 'retail'))});
      case 'reports':
        return json_({ok:true, data:reports_(
          String(e.parameter.period || 'day'),
          String(e.parameter.date || ''),
          String(e.parameter.from || ''),
          String(e.parameter.to || '')
        )});
      case 'getSale':
        return json_(getSale_(String(e.parameter.saleId || '')));
      case 'developerAccess':
        return json_({ok:true, data:getDeveloperAccess_()});
      default:
        return json_({ok:false, error:'Unknown action.'});
    }
  } catch (err) {
    return jsonError_(err);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '');

    // Login is the only endpoint that does not require an existing token.
    if (action === 'login') {
      return json_(login_(String(body.staff || ''), String(body.password || '')));
    }

    const session = requireAuth_(String(body.token || ''));
    // Central access gate: only Lameck (developer) can continue after the
    // developer-set deadline. Business data and password logic remain intact.
    if (action !== 'developerAccess' && action !== 'clearDeveloperAccess' && action !== 'saveDeveloperAccess') {
      requireDeveloperAccess_(session);
    }

    switch (action) {
      case 'registerStock':
        return json_(registerStock_(body));
      case 'updateStock':
        return json_(updateStock_(body));
      case 'editProduct':
        return json_(editProduct_(body));
      case 'recordSale':
        return json_(recordSale_(body));
      case 'getSale':
        return json_(getSale_(String(body.saleId || '')));
      case 'editSale':
        return json_(editSale_(body));
      case 'changePassword':
        return json_(changePassword_(session, String(body.currentPassword || ''), String(body.newPassword || '')));
      case 'developerAccess':
        return json_({ok:true, data:getDeveloperAccess_()});
      case 'saveDeveloperAccess':
        return json_(saveDeveloperAccess_(session, body));
      case 'clearDeveloperAccess':
        return json_(clearDeveloperAccess_(session));
      default:
        return json_({ok:false, error:'Unknown action.'});
    }
  } catch (err) {
    return jsonError_(err);
  }
}

/* ======================= AUTH ======================= */

function getStaffSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAMES.STAFF);
  if(!sh) sh = ss.insertSheet(SHEET_NAMES.STAFF);

  const headers = [['Name','PasswordHash','Salt','UpdatedAt']];
  const lastRow = sh.getLastRow();
  const existing = lastRow > 0
    ? sh.getRange(1,1,lastRow,4).getValues()
    : [];

  const headerOk = existing.length > 0 &&
    String(existing[0][0]).trim() === 'Name' &&
    String(existing[0][1]).trim() === 'PasswordHash' &&
    String(existing[0][2]).trim() === 'Salt';

  // Preserve existing hashes for every retained staff account.
  const saved = {};
  if(headerOk){
    existing.slice(1).forEach(r => {
      const name = String(r[0] || '').trim().toUpperCase();
      if(name && r[1] && r[2]){
        saved[name] = [name, r[1], r[2], r[3] || new Date()];
      }
    });
  }

  const allowed = Object.keys(DEFAULT_STAFF).map(n => n.toUpperCase());
  const existingNames = Object.keys(saved);
  const needsMigration = !headerOk ||
    existingNames.length !== allowed.length ||
    allowed.some(name => !saved[name]);

  if(needsMigration){
    const rows = [];
    const scriptProps = PropertiesService.getScriptProperties();

    // Do not erase or rewrite the Staff sheet unless a MUVA password is
    // available when the MUVA account still needs to be created.
    if(!saved[ADMIN_NAME]){
      const muvaPassword = scriptProps.getProperty('MUVA_ADMIN_PASSWORD');
      if(!muvaPassword || String(muvaPassword).length < 6){
        throw new Error(
          'MUVA admin account is not initialized. Set MUVA_ADMIN_PASSWORD ' +
          '(at least 6 characters) in Apps Script Project Settings > Script Properties, then retry.'
        );
      }
    }

    allowed.forEach(name => {
      if(saved[name]){
        rows.push(saved[name]);
      } else {
        const password = name === ADMIN_NAME
          ? scriptProps.getProperty('MUVA_ADMIN_PASSWORD')
          : DEFAULT_STAFF[name];
        if(!password){
          throw new Error('Missing initial password for staff account: ' + name);
        }
        const salt = Utilities.getUuid();
        rows.push([name, hashPassword_(password, salt), salt, new Date()]);
      }
    });

    // Only the authentication sheet is changed. Business sheets are untouched.
    sh.clearContents();
    sh.getRange(1,1,1,4).setValues(headers);
    if(rows.length){
      sh.getRange(2,1,rows.length,4).setValues(rows);
    }
    sh.setFrozenRows(1);
    scriptProps.setProperty('staffDefaultsVersion', STAFF_DEFAULTS_VERSION);
  }

  if(!sh.isSheetHidden()) sh.hideSheet();
  return sh;
}

function hashPassword_(password, salt){
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + ':' + String(password), Utilities.Charset.UTF_8);
  return bytes.map(b => { const v = b < 0 ? b + 256 : b; return ('0' + v.toString(16)).slice(-2); }).join('');
}

function findStaff_(name){
  const sh=getStaffSheet_();
  const values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(String(values[i][0]).trim().toUpperCase()===name.trim().toUpperCase()){
      return {row:i+1, name:String(values[i][0]).trim().toUpperCase(), hash:String(values[i][1]), salt:String(values[i][2])};
    }
  }
  return null;
}

function login_(staff, password) {
  const name = staff.trim().toUpperCase();
  const account = findStaff_(name);
  if (!account || hashPassword_(password, account.salt) !== account.hash) {
    return {ok:false, error:'Invalid staff name or password.'};
  }

  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put(
    'session:' + token,
    JSON.stringify({staffName:name}),
    CACHE_SECONDS
  );

  return {ok:true, token:token, staffName:name};
}

function changePassword_(session, currentPassword, newPassword){
  const name = String(session.staffName || '').trim().toUpperCase();
  if(!name) return {ok:false, error:'Invalid staff session.'};
  if(String(newPassword).length < 6) return {ok:false, error:'New password must be at least 6 characters.'};
  if(currentPassword === newPassword) return {ok:false, error:'New password must be different from the current password.'};

  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const account=findStaff_(name);
    if(!account || hashPassword_(currentPassword, account.salt) !== account.hash){
      return {ok:false, error:'Current password is incorrect.'};
    }
    const newSalt=Utilities.getUuid();
    const sh=getStaffSheet_();
    sh.getRange(account.row,2,1,3).setValues([[hashPassword_(newPassword,newSalt), newSalt, new Date()]]);
    return {ok:true, message:'Password changed successfully.'};
  } finally {
    lock.releaseLock();
  }
}

/* ======================= DEVELOPER ACCESS ======================= */
const DEVELOPER_NAME = 'LAMECK';
const DEVELOPER_PHONE = '+255744065662';
const DEVELOPER_SHEET_HEADERS = ['Message','ExpiresAt','UpdatedAt','UpdatedBy','DeveloperPhone'];

function developerAccessSheet_(){
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_NAMES.DEVELOPER_ACCESS);
  if(!sh){
    sh = ss.insertSheet(SHEET_NAMES.DEVELOPER_ACCESS);
    sh.getRange(1,1,1,DEVELOPER_SHEET_HEADERS.length).setValues([DEVELOPER_SHEET_HEADERS]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  } else if(sh.getLastRow() === 0){
    sh.getRange(1,1,1,DEVELOPER_SHEET_HEADERS.length).setValues([DEVELOPER_SHEET_HEADERS]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function getDeveloperAccess_(){
  const sh = developerAccessSheet_();
  if(sh.getLastRow() < 2) return {message:'', expiresAt:0, updatedAt:0, updatedBy:'', developerPhone:DEVELOPER_PHONE};
  const r = sh.getRange(2,1,1,5).getValues()[0];
  const expiresAt = r[1] ? new Date(r[1]).getTime() : 0;
  const updatedAt = r[2] ? new Date(r[2]).getTime() : 0;
  return {
    message: String(r[0] || ''),
    expiresAt: Number.isFinite(expiresAt) ? expiresAt : 0,
    updatedAt: Number.isFinite(updatedAt) ? updatedAt : 0,
    updatedBy: String(r[3] || ''),
    developerPhone: String(r[4] || DEVELOPER_PHONE)
  };
}

function isDeveloperSession_(session){
  return !!(session && String(session.staffName || '').trim().toUpperCase() === DEVELOPER_NAME);
}

function requireDeveloper_(session){
  if(!isDeveloperSession_(session)){
    const err = new Error('Developer access required. Please call the developer ' + DEVELOPER_PHONE + ' to restore access.');
    err.code = 'DEVELOPER_REQUIRED';
    throw err;
  }
}

function saveDeveloperAccess_(session, body){
  requireDeveloper_(session);
  const message = String(body.message || '').trim();
  const duration = Number(body.duration);
  const unit = String(body.unit || 'minutes').toLowerCase();
  if(!message) return {ok:false, error:'Enter a message for users.'};
  if(!Number.isFinite(duration) || duration <= 0) return {ok:false, error:'Enter a time limit greater than zero.'};
  const multiplier = unit === 'days' ? 24*60*60*1000 : (unit === 'hours' ? 60*60*1000 : 60*1000);
  const now = new Date();
  const expires = new Date(now.getTime() + duration * multiplier);
  const sh = developerAccessSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(sh.getLastRow() < 2) sh.insertRowAfter(1);
    sh.getRange(2,1,1,5).setValues([[message, expires, now, String(session.staffName || DEVELOPER_NAME), DEVELOPER_PHONE]]);
  } finally { lock.releaseLock(); }
  return {ok:true, data:getDeveloperAccess_()};
}

function clearDeveloperAccess_(session){
  requireDeveloper_(session);
  const sh = developerAccessSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(sh.getLastRow() >= 2) sh.getRange(2,1,1,5).clearContent();
  } finally { lock.releaseLock(); }
  return {ok:true, data:getDeveloperAccess_()};
}

function requireDeveloperAccess_(session){
  const n = getDeveloperAccess_();
  const expired = !!(n.expiresAt && Date.now() >= n.expiresAt);
  if(expired && !isDeveloperSession_(session)){
    const err = new Error((n.message || 'System access has expired.') + ' Please call the developer ' + DEVELOPER_PHONE + ' to restore access.');
    err.code = 'ACCESS_EXPIRED';
    err.accessExpired = true;
    throw err;
  }
  return n;
}

function requireAuth_(token) {
  if (!token) throw new Error('Authentication required.');
  const raw = CacheService.getScriptCache().get('session:' + token);
  if (!raw) {
    const err = new Error('Your session has expired. Please sign in again.');
    err.auth = false;
    throw err;
  }
  return JSON.parse(raw);
}

function jsonError_(err) {
  return json_({
    ok:false,
    auth: err && err.auth === false ? false : undefined,
    error: String(err && err.message ? err.message : err)
  });
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================= SHEETS ======================= */

function ss_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script from the shop Google Sheet.');
  return ss;
}

function sheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    // Keep existing shop sheets compatible when new reporting columns are added.
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function productsSheet_() {
  return sheet_(SHEET_NAMES.PRODUCTS,
    ['ID','Name','Unit','Stock','BuyPrice','RetailPrice','UpdatedAt']);
}

function salesSheet_() {
  return sheet_(SHEET_NAMES.SALES,
    ['SaleID','Timestamp','Staff','Total','QtyUnits']);
}

function linesSheet_() {
  return sheet_(SHEET_NAMES.LINES,
    ['SaleID','Timestamp','Staff','ProductID','ProductName','Unit','Qty','UnitPrice','LineTotal','CostTotal','Profit']);
}

function stockLogSheet_() {
  return sheet_(SHEET_NAMES.STOCKLOG,
    ['Timestamp','Staff','ProductID','ProductName','QtyChange','BuyPrice','RetailPrice','ExpectedProfit','Action']);
}

function normalize_(v) {
  return String(v == null ? '' : v).trim();
}

function num_(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function iso_(d) {
  return Utilities.formatDate(new Date(d), ss_().getSpreadsheetTimeZone(), "yyyy-MM-dd");
}

function now_() {
  return new Date();
}

/* ======================= PRODUCTS ======================= */

function listProducts_() {
  const sh = productsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(r => normalize_(r[0]) && normalize_(r[1]))
    .map(r => ({
      id: normalize_(r[0]),
      name: normalize_(r[1]),
      unit: normalize_(r[2]) || 'pcs',
      stock: num_(r[3]),
      buyPrice: num_(r[4]),
      retailPrice: num_(r[5]),
      totalProfit: num_(r[3]) * (num_(r[5]) - num_(r[4]))
    }));
}

function productMap_() {
  const sh = productsSheet_();
  const values = sh.getDataRange().getValues();
  const map = {};
  for (let i=1; i<values.length; i++) {
    if (normalize_(values[i][0])) {
      map[normalize_(values[i][0])] = {
        row:i+1,
        id:normalize_(values[i][0]),
        name:normalize_(values[i][1]),
        unit:normalize_(values[i][2]) || 'pcs',
        stock:num_(values[i][3]),
        buyPrice:num_(values[i][4]),
        retailPrice:num_(values[i][5]),
        totalProfit:num_(values[i][3]) * (num_(values[i][5]) - num_(values[i][4]))
      };
    }
  }
  return {sheet:sh, map:map};
}

function registerStock_(b) {
  const staff = String(b.staff || '').trim().toUpperCase();
  requireStaffMatchesSession_(staff, String(b.token || ''));

  const qty = num_(b.qty);
  const buyPrice = num_(b.buyPrice);
  const retailPrice = num_(b.retailPrice);
  const unit = normalize_(b.unit) || 'pcs';

  if (qty <= 0) throw new Error('Quantity received must be greater than zero.');
  if (buyPrice < 0 || retailPrice < 0) throw new Error('Prices cannot be negative.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const pm = productMap_();
    let p = b.productId ? pm.map[normalize_(b.productId)] : null;
    const timestamp = now_();

    if (!p) {
      const name = normalize_(b.name);
      if (!name) throw new Error('Enter a product name.');
      const id = 'P-' + Utilities.getUuid().slice(0,8).toUpperCase();
      const sh = pm.sheet;
      sh.appendRow([id,name,unit,qty,buyPrice,retailPrice,timestamp]);
      stockLogSheet_().appendRow([timestamp,staff,id,name,qty,buyPrice,retailPrice,qty*(retailPrice-buyPrice),'STOCK_IN']);
      p = {id:id,name:name,unit:unit,stock:qty,buyPrice:buyPrice,retailPrice:retailPrice,expectedProfit:qty*(retailPrice-buyPrice),totalProfit:qty*(retailPrice-buyPrice)};
    } else {
      const row = p.row;
      const newStock = p.stock + qty;
      pm.sheet.getRange(row,3,1,5).setValues([[
        unit || p.unit,
        newStock,
        buyPrice,
        retailPrice,
        timestamp
      ]]);
      stockLogSheet_().appendRow([
        timestamp,staff,p.id,p.name,qty,buyPrice,retailPrice,qty*(retailPrice-buyPrice),'STOCK_IN'
      ]);
      p = {
        id:p.id,name:p.name,unit:unit || p.unit,stock:newStock,
        buyPrice:buyPrice,retailPrice:retailPrice,expectedProfit:qty*(retailPrice-buyPrice),totalProfit:newStock*(retailPrice-buyPrice)
      };
    }
    return {ok:true, product:p};
  } finally {
    lock.releaseLock();
  }
}

function editProduct_(b) {
  const staff = String(b.staff || '').trim().toUpperCase();
  const token = String(b.token || '');
  requireStaffMatchesSession_(staff, token);

  const productId = normalize_(b.productId);
  const name = normalize_(b.name);
  const unit = normalize_(b.unit) || 'pcs';
  const stock = num_(b.stock);
  const buyPrice = num_(b.buyPrice);
  const retailPrice = num_(b.retailPrice);
  const password = String(b.password || '');

  if (!productId) throw new Error('Product is required.');
  if (!name) throw new Error('Product name is required.');
  if (!Number.isFinite(stock) || stock < 0) throw new Error('Stock quantity cannot be negative.');
  if (buyPrice < 0 || retailPrice < 0) throw new Error('Prices cannot be negative.');
  if (!password) throw new Error('Password is required to edit product details.');

  // The edit is deliberately password-protected on the server. The password
  // must belong to the currently signed-in staff account.
  const session = requireAuth_(token);
  const sessionStaff = String(session.staffName || '').trim().toUpperCase();
  if (!sessionStaff || sessionStaff !== staff) {
    throw new Error('The signed-in staff account does not match this request.');
  }
  // Product/stock editing is an ADMIN action. MUVA admin password is required.
  const adminAccount = findStaff_(ADMIN_NAME);
  if (!adminAccount || hashPassword_(password, adminAccount.salt) !== adminAccount.hash) {
    throw new Error('Incorrect MUVA admin password. Product details were not changed.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const pm = productMap_();
    const p = pm.map[productId];
    if (!p) throw new Error('Product not found.');

    const timestamp = now_();
    pm.sheet.getRange(p.row,2,1,6).setValues([[
      name,
      unit,
      stock,
      buyPrice,
      retailPrice,
      timestamp
    ]]);

    // Keep an audit trail without altering historical SaleLines.
    stockLogSheet_().appendRow([
      timestamp,
      staff,
      p.id,
      name,
      stock - p.stock,
      buyPrice,
      retailPrice,
      0,
      'PRODUCT_EDIT'
    ]);

    return {
      ok:true,
      product:{
        id:p.id,
        name:name,
        unit:unit,
        stock:stock,
        buyPrice:buyPrice,
        retailPrice:retailPrice,
        totalProfit:stock * (retailPrice - buyPrice)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function updateStock_(b) {
  const staff = String(b.staff || '').trim().toUpperCase();
  const token = String(b.token || '');
  requireStaffMatchesSession_(staff, token);
  const password = String(b.password || '');
  if (!password) throw new Error('MUVA admin password is required to adjust stock.');
  const adminAccount = findStaff_(ADMIN_NAME);
  if (!adminAccount || hashPassword_(password, adminAccount.salt) !== adminAccount.hash) {
    throw new Error('Incorrect MUVA admin password. Stock was not changed.');
  }
  const productId = normalize_(b.productId);
  const newStock = num_(b.stock);
  if (!productId) throw new Error('Product is required.');
  if (newStock < 0) throw new Error('Stock cannot be negative.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const pm = productMap_();
    const p = pm.map[productId];
    if (!p) throw new Error('Product not found.');

    const delta = newStock - p.stock;
    pm.sheet.getRange(p.row,4).setValue(newStock);
    pm.sheet.getRange(p.row,7).setValue(now_());

    if (delta !== 0) {
      stockLogSheet_().appendRow([
        now_(),staff,p.id,p.name,delta,p.buyPrice,p.retailPrice,delta*(p.retailPrice-p.buyPrice),'STOCK_ADJUSTMENT'
      ]);
    }
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}

/* ======================= SALES ======================= */

function recordSale_(b) {
  const staff = String(b.staff || '').trim().toUpperCase();
  requireStaffMatchesSession_(staff, String(b.token || ''));

  const items = Array.isArray(b.items) ? b.items : [];
  if (!items.length) throw new Error('Add at least one item.');
  // Backend deliberately accepts no customerType and always calculates retail.
  // Any old retail/customer-type payload is ignored.

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const pm = productMap_();
    const saleId = 'S-' + Utilities.getUuid().slice(0,8).toUpperCase();
    const ts = now_();
    const lineRows = [];
    let total = 0;
    let totalQty = 0;
    let totalProfit = 0;

    items.forEach(it => {
      const p = pm.map[normalize_(it.productId)];
      const qty = num_(it.qty);
      if (!p) throw new Error('A selected product no longer exists.');
      if (qty <= 0) throw new Error('Sale quantity must be greater than zero.');
      if (qty > p.stock) throw new Error(p.name + ': not enough stock.');

      const priceMode = String(it.priceMode || 'registered').trim().toLowerCase();
      let unitPrice;
      if (priceMode === 'manual') {
        unitPrice = num_(it.unitPrice);
        if (unitPrice < 0) throw new Error(p.name + ': selling price cannot be negative.');
      } else if (priceMode === 'registered') {
        unitPrice = num_(p.retailPrice);
      } else {
        throw new Error(p.name + ': invalid selling price option.');
      }
      const lineTotal = qty * unitPrice;
      const costTotal = qty * p.buyPrice;
      const profit = lineTotal - costTotal;

      lineRows.push([
        saleId,ts,staff,p.id,p.name,p.unit,qty,unitPrice,lineTotal,costTotal,profit
      ]);

      total += lineTotal;
      totalQty += qty;
      totalProfit += profit;
      p.stock -= qty;
    });

    // Write stock changes only after all items have passed validation.
    Object.keys(pm.map).forEach(id => {
      // only changed products are written below using the original stock comparison
    });
    items.forEach(it => {
      const p = pm.map[normalize_(it.productId)];
      pm.sheet.getRange(p.row,4).setValue(p.stock);
      pm.sheet.getRange(p.row,7).setValue(ts);
    });

    salesSheet_().appendRow([saleId,ts,staff,total,totalQty]);
    if (lineRows.length) {
      const sh = linesSheet_();
      sh.getRange(sh.getLastRow()+1,1,lineRows.length,lineRows[0].length).setValues(lineRows);
    }

    const receiptLines = lineRows.map(r => ({
      name:r[4],
      qty:r[6],
      lineTotal:r[8],
      profit:r[10]
    }));

    return {ok:true,saleId:saleId,total:total,totalProfit:totalProfit,lines:receiptLines};
  } finally {
    lock.releaseLock();
  }
}


function findSaleRow_(saleId){
  const sh = salesSheet_();
  const values = sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(normalize_(values[i][0]) === normalize_(saleId)){
      return {
        row:i+1,
        saleId:normalize_(values[i][0]),
        timestamp:values[i][1],
        staff:normalize_(values[i][2]),
        total:num_(values[i][3]),
        qtyUnits:num_(values[i][4])
      };
    }
  }
  return null;
}

function getSale_(saleId){
  const id = normalize_(saleId);
  if(!id) throw new Error('Sale ID is required.');
  const sale = findSaleRow_(id);
  if(!sale) throw new Error('Sale not found.');

  const lines = readLines_().filter(x => x.saleId === id).map(x=>({
    productId:x.productId,
    productName:x.productName,
    unit:x.unit,
    qty:x.qty,
    unitPrice:x.unitPrice,
    lineTotal:x.lineTotal,
    costTotal:x.costTotal,
    profit:x.profit
  }));

  return {ok:true,sale:{
    saleId:sale.saleId,
    timestamp:sale.timestamp,
    staff:sale.staff,
    total:sale.total,
    qtyUnits:sale.qtyUnits,
    items:lines
  }};
}

function editSale_(b){
  // A valid staff session is required, but the person performing the correction
  // does not have to be logged in as MUVA. The separate MUVA admin password is the
  // approval credential required to save a completed-sale correction.
  const token = String(b.token || '');
  const session = requireAuth_(token);
  const password = String(b.password || '');
  if(!password) throw new Error('MUVA admin password is required to edit a completed sale.');

  const account = findStaff_(ADMIN_NAME);
  if(!account){
    throw new Error('MUVA admin account was not found. Please initialize the Staff sheet from the latest backend.');
  }
  if(hashPassword_(password, account.salt) !== account.hash){
    throw new Error('Incorrect MUVA admin password. The sale was not changed.');
  }
  const editingStaff = String(session.staffName || 'UNKNOWN').trim().toUpperCase();

  const saleId = normalize_(b.saleId);
  if(!saleId) throw new Error('Sale ID is required.');

  const incoming = Array.isArray(b.items) ? b.items : [];
  if(!incoming.length) throw new Error('A sale must contain at least one item.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try{
    const sale = findSaleRow_(saleId);
    if(!sale) throw new Error('Sale not found.');

    const linesSh = linesSheet_();
    const lineValues = linesSh.getDataRange().getValues();
    const oldRows = [];
    const oldByProduct = {};
    for(let i=1;i<lineValues.length;i++){
      if(normalize_(lineValues[i][0]) === saleId){
        const productId = normalize_(lineValues[i][3]);
        const qty = num_(lineValues[i][6]);
        const costTotal = num_(lineValues[i][9]);
        oldRows.push(i+1);
        if(!oldByProduct[productId]){
          oldByProduct[productId] = {qty:0,costPerUnit: qty > 0 ? costTotal/qty : 0};
        }
        oldByProduct[productId].qty += qty;
      }
    }
    if(!oldRows.length) throw new Error('The sale has no saved line items.');

    const pm = productMap_();

    // First restore the stock consumed by the original sale.
    Object.keys(oldByProduct).forEach(id=>{
      const p = pm.map[id];
      if(!p) throw new Error('A product from the original sale no longer exists.');
      p.stock += oldByProduct[id].qty;
    });

    // Aggregate the replacement sale so duplicate product lines cannot
    // accidentally bypass the stock check.
    const normalized = [];
    const requestedByProduct = {};
    incoming.forEach(it=>{
      const productId = normalize_(it.productId);
      const qty = num_(it.qty);
      const unitPrice = num_(it.unitPrice);
      if(!productId) throw new Error('Every sale line must have a product.');
      if(qty <= 0) throw new Error('Sale quantity must be greater than zero.');
      if(unitPrice < 0) throw new Error('Selling price cannot be negative.');
      const p = pm.map[productId];
      if(!p) throw new Error('A selected product no longer exists.');
      normalized.push({productId,qty,unitPrice});
      requestedByProduct[productId] = (requestedByProduct[productId] || 0) + qty;
    });

    Object.keys(requestedByProduct).forEach(id=>{
      const p = pm.map[id];
      if(requestedByProduct[id] > p.stock){
        throw new Error(p.name + ': not enough stock for the corrected sale.');
      }
    });

    const ts = sale.timestamp || now_();
    const newRows = [];
    let total = 0;
    let totalQty = 0;

    normalized.forEach(it=>{
      const p = pm.map[it.productId];
      const old = oldByProduct[it.productId];
      const costPerUnit = old && old.costPerUnit > 0 ? old.costPerUnit : p.buyPrice;
      const lineTotal = it.qty * it.unitPrice;
      const costTotal = it.qty * costPerUnit;
      const profit = lineTotal - costTotal;

      p.stock -= it.qty;
      total += lineTotal;
      totalQty += it.qty;
      newRows.push([
        saleId,ts,sale.staff,p.id,p.name,p.unit,it.qty,it.unitPrice,lineTotal,costTotal,profit
      ]);
    });

    // Update all affected stock values.
    const affected = {};
    Object.keys(oldByProduct).forEach(id=>affected[id]=true);
    Object.keys(requestedByProduct).forEach(id=>affected[id]=true);
    Object.keys(affected).forEach(id=>{
      const p = pm.map[id];
      pm.sheet.getRange(p.row,4).setValue(p.stock);
      pm.sheet.getRange(p.row,7).setValue(now_());
    });

    // Replace the old line items with the corrected line items.
    oldRows.sort((a,b)=>b-a).forEach(row=>linesSh.deleteRow(row));
    if(newRows.length){
      const start = linesSh.getLastRow()+1;
      linesSh.getRange(start,1,newRows.length,newRows[0].length).setValues(newRows);
    }

    // Keep the original sale ID/date/staff, but update its totals.
    const salesSh = salesSheet_();
    salesSh.getRange(sale.row,4,1,2).setValues([[total,totalQty]]);

    // Audit the stock movement caused by the correction.
    const logSh = stockLogSheet_();
    const allAffected = {};
    Object.keys(oldByProduct).forEach(id=>allAffected[id]=true);
    Object.keys(requestedByProduct).forEach(id=>allAffected[id]=true);
    Object.keys(allAffected).forEach(id=>{
      const p = pm.map[id];
      const oldQty = oldByProduct[id] ? oldByProduct[id].qty : 0;
      const newQty = requestedByProduct[id] || 0;
      const delta = oldQty - newQty; // positive means stock was returned
      if(delta !== 0){
        logSh.appendRow([
          now_(),editingStaff,p.id,p.name,delta,p.buyPrice,p.retailPrice,0,'SALE_EDIT'
        ]);
      }
    });

    const receiptLines = newRows.map(r=>({
      name:r[4],qty:r[6],lineTotal:r[8],profit:r[10]
    }));

    return {ok:true,saleId:saleId,total:total,totalQty:totalQty,
      totalProfit:newRows.reduce((a,r)=>a+r[10],0),lines:receiptLines};
  } finally {
    lock.releaseLock();
  }
}

/* ======================= REPORTING ======================= */

function requireStaffMatchesSession_(staff, token) {
  const session = requireAuth_(token);
  if (!staff || staff !== session.staffName) {
    throw new Error('Staff identity does not match the signed-in account.');
  }
  if (!findStaff_(staff)) throw new Error('Unknown staff account.');
}

function dateRange_(period, selectedDate) {
  const tz = ss_().getSpreadsheetTimeZone();
  const now = new Date();
  const base = selectedDate ? new Date(selectedDate + 'T12:00:00') : now;
  const day = new Date(base);
  day.setHours(0,0,0,0);

  let start = new Date(day);
  let end = new Date(day);
  if (period === 'week') {
    const dow = day.getDay();
    start.setDate(day.getDate() - dow);
    end = new Date(start); end.setDate(start.getDate()+7);
  } else if (period === 'month') {
    start = new Date(day.getFullYear(), day.getMonth(), 1);
    end = new Date(day.getFullYear(), day.getMonth()+1, 1);
  } else {
    end.setDate(day.getDate()+1);
  }

  return {start:start,end:end,startLabel:Utilities.formatDate(start,tz,'yyyy-MM-dd'),endLabel:Utilities.formatDate(new Date(end.getTime()-1),tz,'yyyy-MM-dd')};
}

/* Parse 'yyyy-MM-dd' into a local date (script timezone). Returns null if invalid. */
function parseYmd_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '').trim());
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
  return isNaN(d.getTime()) ? null : d;
}

/* Custom range: both days included. Start/end swapped if given in reverse. */
function rangeBetween_(fromStr, toStr) {
  let a = parseYmd_(fromStr), b = parseYmd_(toStr);
  if (!a || !b) throw new Error('Choose both a start date and an end date.');
  if (a.getTime() > b.getTime()) { const t = a; a = b; b = t; }
  const tz = ss_().getSpreadsheetTimeZone();
  const end = new Date(b.getFullYear(), b.getMonth(), b.getDate() + 1, 0, 0, 0, 0);
  return {
    start: a,
    end: end,
    startLabel: Utilities.formatDate(a, tz, 'yyyy-MM-dd'),
    endLabel: Utilities.formatDate(b, tz, 'yyyy-MM-dd')
  };
}

function readLines_() {
  const sh = linesSheet_();
  const values = sh.getDataRange().getValues();
  return values.length < 2 ? [] : values.slice(1).filter(r => normalize_(r[0])).map(r => ({
    saleId:normalize_(r[0]),
    timestamp:new Date(r[1]),
    staff:normalize_(r[2]),
    productId:normalize_(r[3]),
    productName:normalize_(r[4]),
    unit:normalize_(r[5]),
    qty:num_(r[6]),
    unitPrice:num_(r[7]),
    lineTotal:num_(r[8]),
    costTotal:num_(r[9]),
    profit:num_(r[10])
  }));
}

function reports_(period, selectedDate, fromDate, toDate) {
  const range = (period === 'range')
    ? rangeBetween_(fromDate, toDate)
    : dateRange_(period, selectedDate);

  const lines = readLines_().filter(x => x.timestamp >= range.start && x.timestamp < range.end);

  const byProduct = {};
  const dailyMap = {};
  const saleIds = {};
  let revenue = 0, profit = 0, qtySold = 0;

  // One pass over the lines.
  lines.forEach(x => {
    if (!byProduct[x.productId]) byProduct[x.productId] = {name:x.productName, qty:0, revenue:0, profit:0};
    const p = byProduct[x.productId];
    p.qty += x.qty; p.revenue += x.lineTotal; p.profit += x.profit;

    const d = iso_(x.timestamp);
    dailyMap[d] = (dailyMap[d] || 0) + x.lineTotal;

    saleIds[x.saleId] = true;
    revenue += x.lineTotal; profit += x.profit; qtySold += x.qty;
  });

  return {
    rangeStart:range.startLabel,
    rangeEnd:range.endLabel,
    revenue:revenue,
    profit:profit,
    qtySold:qtySold,
    salesCount:Object.keys(saleIds).length,
    daily:Object.keys(dailyMap).sort().map(date => ({date:date,revenue:dailyMap[date]})),
    // Best sellers first (by revenue) instead of by internal product ID.
    byProduct:Object.keys(byProduct).map(k => byProduct[k]).sort((a,b) => b.revenue - a.revenue)
  };
}

function readSales_() {
  const sh = salesSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(r => normalize_(r[0]))
    .map(r => ({
      saleId: normalize_(r[0]),
      timestamp: r[1] instanceof Date ? new Date(r[1]) : new Date(String(r[1])),
      staff: normalize_(r[2]),
      total: num_(r[3]),
      qtyUnits: num_(r[4])
    }))
    .filter(x => !isNaN(x.timestamp.getTime()));
}

function dashboard_(customerType) {
  const products = listProducts_();
  const lowStock = products.filter(p => p.stock < 5);
  const lines = readLines_();
  const sales = readSales_();
  const now = new Date();
  const todayStart = new Date(now); todayStart.setHours(0,0,0,0);
  const weekStart = new Date(todayStart); weekStart.setDate(todayStart.getDate()-todayStart.getDay());
  const monthStart = new Date(todayStart.getFullYear(),todayStart.getMonth(),1);

  function summary(start) {
    const rows = lines.filter(x => x.timestamp >= start && x.timestamp <= now);
    return {
      revenue:rows.reduce((a,x)=>a+x.lineTotal,0),
      profit:rows.reduce((a,x)=>a+x.profit,0),
      qtySold:rows.reduce((a,x)=>a+x.qty,0)
    };
  }

  // Lifetime totals come directly from the permanent SaleLines sheet.
  const lifetime = {
    revenue: lines.reduce((a,x)=>a+x.lineTotal,0),
    profit: lines.reduce((a,x)=>a+x.profit,0),
    qtySold: lines.reduce((a,x)=>a+x.qty,0),
    salesCount: sales.length
  };

  // Return recent completed sales for the dashboard. The full history remains
  // permanently stored in Sales + SaleLines; this is only the display window.
  const recentSales = sales
    .sort((a,b)=>b.timestamp.getTime()-a.timestamp.getTime())
    .slice(0,50)
    .map(s => {
      const saleLines = lines.filter(x => x.saleId === s.saleId);
      return {
        saleId:s.saleId,
        timestamp:s.timestamp,
        staff:s.staff,
        total:s.total,
        qtyUnits:s.qtyUnits,
        profit:saleLines.reduce((a,x)=>a+x.profit,0),
        items:saleLines.map(x=>({productName:x.productName,qty:x.qty,unit:x.unit,unitPrice:x.unitPrice,lineTotal:x.lineTotal,profit:x.profit}))
      };
    });

  return {
    products:products,
    lowStock:lowStock,
    today:summary(todayStart),
    week:summary(weekStart),
    month:summary(monthStart),
    lifetime:lifetime,
    recentSales:recentSales
  };
}