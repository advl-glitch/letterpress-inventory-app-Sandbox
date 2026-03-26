// =============================================================================
// CONFIGURATION
// =============================================================================

const SPREADSHEET = SpreadsheetApp.openById('1FiDZXPV6aimKpKUvzDCQczq01nCdvMZLzRhWq-DB50U');

// Virginia sales tax rate for Staunton VA
const VA_SALES_TAX = 0.053;

// Wholesale prices by product type (used for wholesale order estimates)
const WHOLESALE_PRICES = {
  'Folded Card': 3.00,
  'Folded':      3.00,
  'Flat':        2.00,
  'Note Card':   2.00,
  'Postcard':    2.00,
  '2-Notecard Set': 4.00,
  'Set':         4.00,
};

// =============================================================================
// ROUTER — handles all incoming GET and POST requests
// =============================================================================

function doGet(e) {
  const action = e.parameter.action;
  let result;

  switch (action) {
    case 'getItems':             result = getItems(); break;
    case 'getProductTypes':      result = getProductTypes(); break;
    case 'getNextItemId':        result = getNextItemId(); break;
    case 'getRetailPartners':    result = getRetailPartners(); break;
    case 'getRetailerAuth':      result = getRetailerAuth(e.parameter.locationId); break;
    case 'getPartnerInventory':  result = getPartnerInventory(e.parameter.partnerId); break;
    case 'getPartnerSalesHistory': result = getPartnerSalesHistory(e.parameter.partnerId); break;
    case 'getVendingMachines':   result = getVendingMachines(); break;
    case 'getTags':              result = getTags(); break;
    case 'getOrders':            result = getOrders(); break;
    case 'getPublicStock':       result = getPublicStock(); break;
    case 'searchRetailers':      result = searchRetailers(e.parameter.query); break;
    case 'getDashboardStats':    result = getDashboardStats(); break;
    case 'getConsignmentTotals': result = getConsignmentTotals(); break;
    case 'getPrintRunTotals':    result = getPrintRunTotals(); break;
    case 'getMarketSales':       result = getMarketSales(); break;
    case 'getSalesReportData':   result = getSalesReportData(e.parameter); break;
    case 'migrateToPartnerStock': result = migrateToPartnerStock(); break;
    case 'getLatestVisit':       result = getLatestVisit(e.parameter.partnerId); break;
    default:
      result = { success: false, error: 'Invalid action: ' + action };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Invalid JSON payload.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = payload.action;
  let result;

  switch (action) {
    case 'addItem':                  result = addItem(payload.itemData, payload.printRunData); break;
    case 'updateItem':               result = updateItem(payload.itemData); break;
    case 'addPrintRun':              result = addPrintRun(payload.itemId, payload.quantity, payload.date); break;
    case 'addVendingMachine':        result = addVendingMachine(payload.machineData); break;
    case 'updateVendingMachine':     result = updateVendingMachine(payload.machineData); break;
    case 'addRetailPartner':         result = addRetailPartner(payload.partnerData); break;
    case 'updateRetailPartner':      result = updateRetailPartner(payload.partnerData); break;
    case 'addTag':                   result = addTag(payload.tagData); break;
    case 'deleteTag':                result = deleteTag(payload.tagId); break;
    case 'submitOrder':              result = submitOrder(payload.orderData); break;
    case 'fulfillOrder':             result = fulfillOrder(payload.orderId, payload.adjustments, payload.notifData); break;
    case 'unfulfillOrder':           result = unfulfillOrder(payload.orderId); break;
    case 'sendRestockNotification':  result = sendRestockNotification(payload.notifData); break;
    case 'submitPartnerRequest':     result = submitPartnerRequest(payload.requestData); break;
    case 'verifyRetailer':           result = verifyRetailer(payload.locationId, payload.email, payload.phone); break;
    case 'uploadPhoto':              result = uploadPhoto(payload); break;
    case 'addProductType':           result = addProductType(payload.typeData); break;
    case 'updateItemStatus':         result = updateItemStatus(payload); break;
    case 'updatePartnerInventory':   result = updatePartnerInventory(payload); break;
    case 'logActualSale':            result = logActualSale(payload); break;
    case 'logMarketSale':            result = logMarketSale(payload); break;
    case 'sendVisitReport':          result = sendVisitReport(payload); break;
    default:
      result = { success: false, error: 'Invalid action: ' + action };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// =============================================================================
// ITEMS
// =============================================================================

function getItems() {
  try {
    const sheet = SPREADSHEET.getSheetByName('Items');
    if (!sheet) return { success: false, error: 'Sheet "Items" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const tagMap = getItemTagsMap();

    const items = data.map(row => {
      const item = {};
      headers.forEach((header, i) => { item[header] = row[i]; });
      // ItemIDs in sheet are plain integers (no leading zeros) — convert to string
      if (item.ItemID !== undefined && item.ItemID !== '') {
        item.ItemID = String(item.ItemID);
      }
      item._tagIds = tagMap[item.ItemID] || [];
      return item;
    }).filter(item => item.ItemID && item.ItemID.trim() !== '');

    return { success: true, items };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getNextItemId() {
  try {
    const sheet = SPREADSHEET.getSheetByName('Items');
    if (!sheet) return { success: false, error: 'Sheet "Items" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const idIndex = headers.indexOf('ItemID');
    if (idIndex === -1) return { success: false, error: 'ItemID column not found.' };

    let maxId = 0;
    data.forEach(row => {
      const id = parseInt(row[idIndex], 10);
      if (!isNaN(id) && id > maxId) maxId = id;
    });

    const nextId = String(maxId + 1); // plain integer — JS display layer adds padding
    return { success: true, nextId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getProductTypes() {
  try {
    const sheet = SPREADSHEET.getSheetByName('ProductType');
    if (!sheet) return { success: false, error: 'Sheet "ProductType" not found.' };
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const types = data.map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    }).filter(t => t.TypeName || t.Name);
    return { success: true, types };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addProductType(typeData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('ProductType');
    if (!sheet) return { success: false, error: 'Sheet "ProductType" not found.' };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    // Generate TypeID by removing spaces (e.g. "Dog House" → "DogHouse")
    const typeId = typeData.typeName.replace(/\s+/g, '');
    const fieldMap = {
      'TypeID': typeId,
      'TypeName': typeData.typeName,
      'Name': typeData.typeName,
      'DefaultRetailPrice': typeData.retailPrice,
      'RetailPrice': typeData.retailPrice,
      'Price': typeData.retailPrice
    };
    Object.entries(fieldMap).forEach(([col, val]) => {
      const idx = headers.indexOf(col);
      if (idx !== -1) newRow[idx] = val;
    });
    sheet.appendRow(newRow);
    return { success: true, message: 'Product type added.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addItem(itemData, printRunData) {
  try {
    const itemsSheet = SPREADSHEET.getSheetByName('Items');
    if (!itemsSheet) return { success: false, error: 'Sheet "Items" not found.' };

    const headers = itemsSheet.getRange(1, 1, 1, itemsSheet.getLastColumn()).getValues()[0];
    const now = new Date().toISOString();
    const itemIdDisplay = String(itemData.itemId).padStart(3, '0');
    const displayName = itemIdDisplay + ' — ' + itemData.designName;

    // Build row array based on headers so column order doesn't matter
    const newRow = new Array(headers.length).fill('');
    const fieldMap = {
      'ItemID': itemData.itemId,
      'Name': itemData.designName,
      'DisplayName': displayName,
      'Photo': itemData.photo || '',
      'ProductType': itemData.itemType,
      'UnitPrice': itemData.unitPrice || '',
      'Status': itemData.status || 'Open',
      'Notes': itemData.notes || '',
      'CreatedAt': now,
      'StartingAtHome': 0,
    };
    Object.entries(fieldMap).forEach(([col, val]) => {
      const idx = headers.indexOf(col);
      if (idx !== -1) newRow[idx] = val;
    });
    itemsSheet.appendRow(newRow);

    if (itemData.tags && itemData.tags.length > 0) {
      saveItemTags(itemData.itemId, itemData.tags);
    }

    if (printRunData && printRunData.quantity && parseInt(printRunData.quantity) > 0) {
      const printResult = addPrintRun(itemData.itemId, printRunData.quantity, printRunData.date);
      if (!printResult.success) {
        return { success: false, error: 'Item added but print run failed: ' + printResult.error };
      }
    }

    return { success: true, message: 'Item "' + itemData.designName + '" added successfully.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateItem(itemData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('Items');
    if (!sheet) return { success: false, error: 'Sheet "Items" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('ItemID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(itemData.itemId)) {
        const rowNum = i + 1;
        const fieldMap = {
          'Name': itemData.designName,
          'DisplayName': itemData.designName ? String(parseInt(itemData.itemId,10)).padStart(3,'0') + ' — ' + itemData.designName : undefined,
          'Photo': itemData.photo,
          'ProductType': itemData.itemType,
          'UnitPrice': itemData.unitPrice,
          'Notes': itemData.notes,
          'Status': itemData.status
        };
        // Update home stock if newStock is provided (from audit page)
        if (itemData.newStock !== undefined) {
          const stockCol = headers.indexOf('StartingAtHome');
          if (stockCol !== -1) sheet.getRange(rowNum, stockCol + 1).setValue(itemData.newStock);
        }
        Object.entries(fieldMap).forEach(([field, val]) => {
          const colIndex = headers.indexOf(field);
          if (colIndex !== -1 && val !== undefined) {
            sheet.getRange(rowNum, colIndex + 1).setValue(val);
          }
        });

        if (itemData.tags) saveItemTags(itemData.itemId, itemData.tags);
        return { success: true, message: 'Item updated.' };
      }
    }
    return { success: false, error: 'Item not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// PRINT RUNS
// =============================================================================

function addPrintRun(itemId, quantity, date) {
  try {
    const sheet = SPREADSHEET.getSheetByName('PrintRuns');
    if (!sheet) return { success: false, error: 'Sheet "PrintRuns" not found.' };

    const now = new Date().toISOString();
    const runDate = new Date(date);
    const dateStr = date.replace(/-/g, '');
    const runId = 'PR-' + dateStr + '-' + itemId;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    const fieldMap = {
      'RunID': runId,
      'RunDate': runDate,
      'ItemID': itemId,
      'QtyPrinted': parseInt(quantity),
      'Quantity': parseInt(quantity),
      'CreatedAt': now
    };
    Object.entries(fieldMap).forEach(([col, val]) => {
      const idx = headers.indexOf(col);
      if (idx !== -1) newRow[idx] = val;
    });
    sheet.appendRow(newRow);

    // Update StartingAtHome on Items sheet
    const itemsSheet = SPREADSHEET.getSheetByName('Items');
    if (itemsSheet) {
      const data = itemsSheet.getDataRange().getValues();
      const headers = data[0];
      const idIdx = headers.indexOf('ItemID');
      const stockIdx = headers.indexOf('StartingAtHome');
      if (idIdx !== -1 && stockIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idIdx]) === String(itemId)) {
            const current = parseInt(data[i][stockIdx]) || 0;
            itemsSheet.getRange(i + 1, stockIdx + 1).setValue(current + parseInt(quantity));
            break;
          }
        }
      }
    }

    return { success: true, message: 'Print run added for Item ' + itemId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// TAGS
// =============================================================================

function getTags() {
  try {
    let sheet = SPREADSHEET.getSheetByName('Tags');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('Tags');
      sheet.appendRow(['TagID', 'TagName', 'Category', 'Active']);
      const defaultTags = [
        ['TAG001','For Women','Audience',true],
        ['TAG002','For Men','Audience',true],
        ['TAG003','For Kids','Audience',true],
        ['TAG004','Latino / Spanish','Audience',true],
        ['TAG005','Black Culture','Audience',true],
        ['TAG006','Queer Culture','Audience',true],
        ['TAG007','Romantic','Theme',true],
        ['TAG008','Humor / Funny','Theme',true],
        ['TAG009','Political','Theme',true],
        ['TAG010','Famous Quotes','Theme',true],
        ['TAG011','Song Lyrics','Theme',true],
        ['TAG012','Movie / TV Quotes','Theme',true],
        ['TAG013','Outdoors / Nature','Theme',true],
        ['TAG014','Virginia','Theme',true],
        ['TAG015','Shenandoah Valley','Theme',true],
        ['TAG016','Winter Holiday','Theme',true],
        ['TAG017','General / All Ages','Rating',true],
        ['TAG018','Rated R','Rating',true],
        ['TAG019','Rated XXX','Rating',true],
        ['TAG020','Postcard','Format',true],
        ['TAG021','Folded Card','Format',true],
        ['TAG022','Note Card','Format',true],
        ['TAG023','2-Notecard Set','Format',true],
        ['TAG024','Card Club — Limited Edition','Card Club',true],
      ];
      defaultTags.forEach(tag => sheet.appendRow(tag));
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const tags = data.map(row => {
      const t = {};
      headers.forEach((h, i) => { t[h] = row[i]; });
      return t;
    }).filter(t => t.TagID && t.Active !== false);

    return { success: true, tags };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addTag(tagData) {
  try {
    let sheet = SPREADSHEET.getSheetByName('Tags');
    if (!sheet) {
      getTags();
      sheet = SPREADSHEET.getSheetByName('Tags');
    }

    const data = sheet.getDataRange().getValues();
    const ids = data.slice(1).map(r => r[0]).filter(Boolean);
    const nums = ids.map(id => parseInt(id.replace('TAG',''),10)).filter(n => !isNaN(n));
    const nextNum = nums.length > 0 ? Math.max(...nums) + 1 : 1;
    const tagId = 'TAG' + String(nextNum).padStart(3,'0');

    sheet.appendRow([tagId, tagData.tagName, tagData.category || 'Theme', true]);
    return { success: true, tagId, message: 'Tag added.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function deleteTag(tagId) {
  try {
    // Remove from Tags sheet
    const tagsSheet = SPREADSHEET.getSheetByName('Tags');
    if (tagsSheet) {
      const data = tagsSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(tagId)) {
          tagsSheet.deleteRow(i + 1);
          break;
        }
      }
    }
    // Remove all ItemTags references
    const itemTagsSheet = SPREADSHEET.getSheetByName('ItemTags');
    if (itemTagsSheet) {
      const data = itemTagsSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][1]) === String(tagId)) {
          itemTagsSheet.deleteRow(i + 1);
        }
      }
    }
    return { success: true, message: 'Tag deleted.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveItemTags(itemId, tags) {
  try {
    let sheet = SPREADSHEET.getSheetByName('ItemTags');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('ItemTags');
      sheet.appendRow(['ItemID', 'TagID']);
    }
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(itemId)) rowsToDelete.push(i + 1);
    }
    rowsToDelete.forEach(r => sheet.deleteRow(r));
    tags.forEach(tagId => sheet.appendRow([itemId, tagId]));
    return true;
  } catch (e) {
    return false;
  }
}

function getItemTagsMap() {
  try {
    const sheet = SPREADSHEET.getSheetByName('ItemTags');
    if (!sheet) return {};
    const data = sheet.getDataRange().getValues();
    data.shift();
    const map = {};
    data.forEach(row => {
      const itemId = String(row[0]);
      if (!map[itemId]) map[itemId] = [];
      map[itemId].push(row[1]);
    });
    return map;
  } catch (e) {
    return {};
  }
}


// =============================================================================
// RETAIL PARTNERS
// =============================================================================

function getRetailPartners() {
  try {
    const sheet = SPREADSHEET.getSheetByName('Locations');
    if (!sheet) return { success: false, error: 'Sheet "Locations" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idx = {};
    headers.forEach((h, i) => { idx[h] = i; });

    const partners = data
      .map(row => ({
        locationId: row[idx['LocationID']],
        displayName: row[idx['DisplayName']],
        locationType: row[idx['LocationType']],
        active: row[idx['Active']],
        city: row[idx['City']],
        split: row[idx['SplitToYou']],
        address: row[idx['Address']],
        phone: row[idx['Phone']],
        contactName: row[idx['ContactName']],
        contactPhone: row[idx['ContactPhone']],
        contactEmail: row[idx['ContactEmail']],
        storePhoto: row[idx['StorePhotoURL']],
        notes: row[idx['Notes']],
        retailItemCodes: row[idx['RetailItemCodes']] || '',
      }))
      .filter(p => p.locationType === 'retail_partner' && (p.active === true || String(p.active).toUpperCase() === 'TRUE'));

    const choices = partners.map(p => ({
      value: p.locationId,
      label: p.displayName,
      city: p.city,
      split: p.split,
      address: p.address,
      phone: p.phone,
      contactName: p.contactName,
      contactPhone: p.contactPhone,
      contactEmail: p.contactEmail,
      storePhoto: p.storePhoto,
      notes: p.notes,
      retailItemCodes: p.retailItemCodes,
    }));

    return { success: true, partners: choices };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addRetailPartner(partnerData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('Locations');
    if (!sheet) return { success: false, error: 'Sheet "Locations" not found.' };

    const data = sheet.getDataRange().getValues();
    const ids = data.slice(1).map(r => r[0]).filter(Boolean);
    const nums = ids.map(id => parseInt(String(id).replace(/\D/g,''),10)).filter(n => !isNaN(n));
    const nextNum = nums.length > 0 ? Math.max(...nums) + 1 : 1;
    const locationId = 'LOC' + String(nextNum).padStart(3,'0');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    const fieldMap = {
      'LocationID': locationId,
      'DisplayName': partnerData.storeName,
      'SplitToYou': partnerData.split ? parseFloat(partnerData.split) / 100 : '',
      'LocationType': partnerData.partnerType || 'retail_partner',
      'Active': true,
      'Notes': partnerData.notes || '',
      'Address': partnerData.address || '',
      'City': partnerData.city || '',
      'Phone': partnerData.phone || '',
      'ContactName': partnerData.contactName || '',
      'ContactPhone': partnerData.contactPhone || '',
      'ContactEmail': partnerData.contactEmail || '',
      'RetailItemCodes': partnerData.retailItemCodes || '',
    };
    Object.entries(fieldMap).forEach(([col, val]) => {
      const idx = headers.indexOf(col);
      if (idx !== -1) newRow[idx] = val;
    });
    sheet.appendRow(newRow);

    saveRetailerVerificationInfo(locationId, partnerData.ownerEmail, partnerData.ownerPhone);

    return { success: true, locationId, message: 'Partner added.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateRetailPartner(partnerData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('Locations');
    if (!sheet) return { success: false, error: 'Sheet "Locations" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('LocationID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(partnerData.locationId)) {
        const rowNum = i + 1;
        const fieldMap = {
          'DisplayName': partnerData.storeName,
          'SplitToYou': partnerData.split ? parseFloat(partnerData.split) / 100 : undefined,
          'Notes': partnerData.notes,
          'Address': partnerData.address,
          'City': partnerData.city,
          'Phone': partnerData.phone,
          'ContactName': partnerData.contactName,
          'ContactPhone': partnerData.contactPhone,
          'ContactEmail': partnerData.contactEmail,
          'Active': partnerData.active,
          'RetailItemCodes': partnerData.retailItemCodes,
        };
        Object.entries(fieldMap).forEach(([field, val]) => {
          const colIndex = headers.indexOf(field);
          if (colIndex !== -1 && val !== undefined) {
            sheet.getRange(rowNum, colIndex + 1).setValue(val);
          }
        });
        if (partnerData.ownerEmail !== undefined || partnerData.ownerPhone !== undefined) {
          saveRetailerVerificationInfo(partnerData.locationId, partnerData.ownerEmail, partnerData.ownerPhone);
        }
        return { success: true, message: 'Partner updated.' };
      }
    }
    return { success: false, error: 'Partner not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getRetailerAuth(locationId) {
  try {
    const sheet = SPREADSHEET.getSheetByName('RetailerAuth');
    if (!sheet) return { success: true, ownerEmail: '', ownerPhone: '' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(locationId)) {
        return { success: true, ownerEmail: data[i][1] || '', ownerPhone: data[i][2] || '' };
      }
    }
    return { success: true, ownerEmail: '', ownerPhone: '' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveRetailerVerificationInfo(locationId, ownerEmail, ownerPhone) {
  try {
    let sheet = SPREADSHEET.getSheetByName('RetailerAuth');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('RetailerAuth');
      sheet.appendRow(['LocationID', 'OwnerEmail', 'OwnerPhone']);
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(locationId)) {
        sheet.getRange(i + 1, 2).setValue(ownerEmail || '');
        sheet.getRange(i + 1, 3).setValue(ownerPhone || '');
        return;
      }
    }
    sheet.appendRow([locationId, ownerEmail || '', ownerPhone || '']);
  } catch (e) {}
}

function verifyRetailer(locationId, email, phone) {
  try {
    const sheet = SPREADSHEET.getSheetByName('RetailerAuth');
    if (!sheet) return { success: false, verified: false };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(locationId)) {
        const storedEmail = String(data[i][1]).trim().toLowerCase();
        const storedPhone = String(data[i][2]).trim().replace(/\D/g,'');
        const inputEmail = String(email || '').trim().toLowerCase();
        const inputPhone = String(phone || '').trim().replace(/\D/g,'');

        const emailMatch = inputEmail && storedEmail && inputEmail === storedEmail;
        const phoneMatch = inputPhone && storedPhone && inputPhone === storedPhone;

        return { success: true, verified: emailMatch || phoneMatch };
      }
    }
    return { success: true, verified: false };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function searchRetailers(query) {
  try {
    if (!query || query.length < 2) return { success: true, partners: [] };
    const result = getRetailPartners();
    if (!result.success) return result;
    const q = query.toLowerCase();
    const matches = result.partners.filter(p =>
      p.label.toLowerCase().includes(q) || (p.city || '').toLowerCase().includes(q)
    );
    return { success: true, partners: matches.slice(0, 5) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getPartnerInventory(locationId) {
  try {
    const locationsSheet = SPREADSHEET.getSheetByName('Locations');
    if (!locationsSheet) return { success: false, error: 'Locations sheet not found.' };

    // Get partner name
    const locationsData = locationsSheet.getDataRange().getValues();
    const locationsHeaders = locationsData.shift();
    const locIdIndex = locationsHeaders.indexOf('LocationID');
    const locNameIndex = locationsHeaders.indexOf('DisplayName');
    const partnerRow = locationsData.find(row => String(row[locIdIndex]) === String(locationId));
    const partnerName = partnerRow ? partnerRow[locNameIndex] : 'Unknown Partner';

    // Read from PartnerStock (current-state table)
    let stockSheet = SPREADSHEET.getSheetByName('PartnerStock');
    if (!stockSheet) {
      // Auto-migrate from old retailInventory if it exists
      const oldSheet = SPREADSHEET.getSheetByName('retailInventory');
      if (oldSheet && oldSheet.getLastRow() > 1) {
        migrateToPartnerStock();
        stockSheet = SPREADSHEET.getSheetByName('PartnerStock');
      }
      if (!stockSheet) {
        return { success: true, data: { name: partnerName, lastVisit: 'N/A', inventory: [] } };
      }
    }

    const stockData = stockSheet.getDataRange().getValues();
    if (stockData.length < 2) {
      return { success: true, data: { name: partnerName, lastVisit: 'N/A', inventory: [] } };
    }

    const headers = stockData.shift();
    const locCol = headers.indexOf('LocationID');
    const itemCol = headers.indexOf('ItemID');
    const nameCol = headers.indexOf('DesignName');
    const stockCol = headers.indexOf('CurrentStock');
    const priceCol = headers.indexOf('UnitPrice');
    const updatedCol = headers.indexOf('LastUpdated');

    // Get fresh item names/prices from Items sheet
    const allItems = getItems().items;
    const itemNames = {};
    const itemPrices = {};
    allItems.forEach(item => {
      itemNames[String(item.ItemID)] = item.DisplayName || item.Name;
      itemPrices[String(item.ItemID)] = parseFloat(item.UnitPrice) || 0;
    });

    let lastUpdated = null;
    const inventory = [];

    stockData.forEach(row => {
      if (String(row[locCol]) === String(locationId)) {
        const itemId = String(row[itemCol]);
        const stock = parseInt(row[stockCol]) || 0;
        if (stock > 0) {
          inventory.push({
            designId: row[itemCol],
            designName: itemNames[itemId] || row[nameCol] || 'Unknown Design',
            unitPrice: itemPrices[itemId] || parseFloat(row[priceCol]) || 0,
            currentStock: stock
          });
        }
        // Track most recent update
        const updated = row[updatedCol];
        if (updated) {
          const d = new Date(updated);
          if (!lastUpdated || d > lastUpdated) lastUpdated = d;
        }
      }
    });

    return {
      success: true,
      data: {
        name: partnerName,
        lastVisit: lastUpdated ? lastUpdated.toLocaleDateString() : 'N/A',
        inventory: inventory
      }
    };
  } catch (e) {
    return { success: false, error: e.message, stack: e.stack };
  }
}


// =============================================================================
// PARTNER INVENTORY UPDATE (from Retail Stock & Sales page)
// =============================================================================

function updatePartnerInventory(payload) {
  try {
    const { partnerId, partnerName, visitDate, updates } = payload;
    const today = visitDate || new Date().toLocaleDateString('en-CA');

    // ── 1. Upsert PartnerStock (current-state table) ──
    let stockSheet = SPREADSHEET.getSheetByName('PartnerStock');
    if (!stockSheet) {
      stockSheet = SPREADSHEET.insertSheet('PartnerStock');
      stockSheet.appendRow(['LocationID', 'ItemID', 'DesignName', 'CurrentStock', 'UnitPrice', 'LastUpdated']);
    }

    const stockData = stockSheet.getDataRange().getValues();
    const stockHeaders = stockData[0];
    const sLocCol = stockHeaders.indexOf('LocationID');
    const sItemCol = stockHeaders.indexOf('ItemID');
    const sNameCol = stockHeaders.indexOf('DesignName');
    const sStockCol = stockHeaders.indexOf('CurrentStock');
    const sPriceCol = stockHeaders.indexOf('UnitPrice');
    const sUpdatedCol = stockHeaders.indexOf('LastUpdated');

    // Build index of existing rows: key = "partnerId|itemId" → row number
    const existingRows = {};
    for (let i = 1; i < stockData.length; i++) {
      const key = String(stockData[i][sLocCol]) + '|' + String(stockData[i][sItemCol]);
      existingRows[key] = i + 1; // 1-based sheet row
    }

    // Track rows to delete (stock went to 0)
    const rowsToDelete = [];

    updates.forEach(u => {
      const key = String(partnerId) + '|' + String(u.designId);
      const existingRow = existingRows[key];

      if (u.newStock <= 0) {
        // Remove from current stock if exists
        if (existingRow) rowsToDelete.push(existingRow);
      } else if (existingRow) {
        // Update existing row
        stockSheet.getRange(existingRow, sNameCol + 1).setValue(u.designName);
        stockSheet.getRange(existingRow, sStockCol + 1).setValue(u.newStock);
        stockSheet.getRange(existingRow, sPriceCol + 1).setValue(u.unitPrice || 0);
        stockSheet.getRange(existingRow, sUpdatedCol + 1).setValue(today);
      } else {
        // Insert new row
        stockSheet.appendRow([partnerId, u.designId, u.designName, u.newStock, u.unitPrice || 0, today]);
      }
    });

    // Delete zero-stock rows (bottom-up to avoid index shifting)
    rowsToDelete.sort((a, b) => b - a).forEach(row => stockSheet.deleteRow(row));

    // ── 2. Append visit log to retailInventory (audit trail) ──
    let logSheet = SPREADSHEET.getSheetByName('retailInventory');
    if (!logSheet) {
      logSheet = SPREADSHEET.insertSheet('retailInventory');
      logSheet.appendRow([
        'LocationID', 'PartnerName', 'VisitDate', 'ItemID', 'DesignName',
        'StartOnShelf', 'EndOnShelf', 'EstimatedSold', 'Added', 'Pulled', 'UnitPrice', 'EntryType'
      ]);
    }

    // Use header-based mapping to match existing column order
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];

    updates.forEach(u => {
      const logRow = new Array(logHeaders.length).fill('');
      const logFields = {
        'LocationID': partnerId,
        'PartnerName': partnerName,
        'VisitDate': today,
        'ItemID': u.designId,
        'DesignName': u.designName,
        'StartOnShelf': u.previousStock,
        'EndOnShelf': u.newStock,
        'EstimatedSold': u.estimatedSold || 0,
        'Added': u.added || 0,
        'Pulled': u.pulled || 0,
        'UnitPrice': u.unitPrice || 0,
        'EntryType': u.isNew ? 'New' : 'Update'
      };
      Object.entries(logFields).forEach(([col, val]) => {
        const idx = logHeaders.indexOf(col);
        if (idx !== -1) logRow[idx] = val;
      });
      logSheet.appendRow(logRow);
    });

    // ── 3. Update Items.StartingAtHome (home stock moves with store stock) ──
    const itemsSheet = SPREADSHEET.getSheetByName('Items');
    if (itemsSheet) {
      const itemsData = itemsSheet.getDataRange().getValues();
      const itemsHeaders = itemsData[0];
      const iIdCol = itemsHeaders.indexOf('ItemID');
      const iStockCol = itemsHeaders.indexOf('StartingAtHome');

      if (iIdCol !== -1 && iStockCol !== -1) {
        updates.forEach(u => {
          const added = parseInt(u.added) || 0;
          const pulled = parseInt(u.pulled) || 0;
          if (added === 0 && pulled === 0) return;

          for (let i = 1; i < itemsData.length; i++) {
            if (String(itemsData[i][iIdCol]) === String(u.designId)) {
              const current = parseInt(itemsData[i][iStockCol]) || 0;
              // Added to store = leaves home; Pulled from store = returns home
              const newHome = Math.max(0, current - added + pulled);
              itemsSheet.getRange(i + 1, iStockCol + 1).setValue(newHome);
              itemsData[i][iStockCol] = newHome; // update local cache for subsequent items
              break;
            }
          }
        });
      }
    }

    // ── 4. Roll up estimated sales into RetailSales for this visit ──
    const totalEstimatedRevenue = updates.reduce((sum, u) => {
      const estSold = parseInt(u.estimatedSold) || 0;
      return sum + (estSold * (parseFloat(u.unitPrice) || 0));
    }, 0);
    const totalEstimatedCards = updates.reduce((sum, u) => sum + (parseInt(u.estimatedSold) || 0), 0);

    if (totalEstimatedRevenue > 0 || totalEstimatedCards > 0) {
      let salesSheet = SPREADSHEET.getSheetByName('RetailSales');
      if (!salesSheet) {
        salesSheet = SPREADSHEET.insertSheet('RetailSales');
        salesSheet.appendRow(['PartnerID', 'PartnerName', 'Month', 'ActualSales', 'EstimatedSales', 'CardsSold', 'LoggedAt']);
      }
      const visitMonth = today.substring(0, 7); // "YYYY-MM"
      const salesData = salesSheet.getDataRange().getValues();
      const salesHeaders = salesData[0];
      const spIdx = salesHeaders.indexOf('PartnerID');
      const smIdx = salesHeaders.indexOf('Month');
      const seIdx = salesHeaders.indexOf('EstimatedSales');

      let found = false;
      for (let i = 1; i < salesData.length; i++) {
        if (String(salesData[i][spIdx]) === String(partnerId) && String(salesData[i][smIdx]) === String(visitMonth)) {
          // Update estimated sales for existing row
          const existing = parseFloat(salesData[i][seIdx]) || 0;
          salesSheet.getRange(i + 1, seIdx + 1).setValue(existing + totalEstimatedRevenue);
          found = true;
          break;
        }
      }
      if (!found) {
        salesSheet.appendRow([partnerId, partnerName, visitMonth, 0, totalEstimatedRevenue, 0, new Date().toISOString()]);
      }
    }

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getPartnerSalesHistory(partnerId) {
  try {
    const sheet = SPREADSHEET.getSheetByName('RetailSales');
    if (!sheet) return { success: true, history: [] }; // sheet may not exist yet

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, history: [] };

    const headers = data.shift();
    const pIdx    = headers.indexOf('PartnerID');
    const mIdx    = headers.indexOf('Month');
    const aIdx    = headers.indexOf('ActualSales');
    const eIdx    = headers.indexOf('EstimatedSales');
    const cIdx    = headers.indexOf('CardsSold');

    const history = data
      .filter(row => String(row[pIdx]) === String(partnerId))
      .map(row => ({
        month:          row[mIdx],
        actualSales:    parseFloat(row[aIdx]) || 0,
        estimatedSales: parseFloat(row[eIdx]) || 0,
        cardsSold:      parseInt(row[cIdx]) || 0,
      }))
      .sort((a, b) => b.month.localeCompare(a.month));

    return { success: true, history };
  } catch (e) {
    return { success: true, history: [] };
  }
}

function logActualSale(payload) {
  try {
    const { partnerId, partnerName, month, actualSales, cardsSold } = payload;

    // Get or create RetailSales sheet
    let sheet = SPREADSHEET.getSheetByName('RetailSales');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('RetailSales');
      sheet.appendRow(['PartnerID', 'PartnerName', 'Month', 'ActualSales', 'EstimatedSales', 'CardsSold', 'LoggedAt']);
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const pIdx    = headers.indexOf('PartnerID');
    const mIdx    = headers.indexOf('Month');

    // Check if row already exists for this partner+month (update it)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][pIdx]) === String(partnerId) && String(data[i][mIdx]) === String(month)) {
        sheet.getRange(i + 1, headers.indexOf('ActualSales') + 1).setValue(actualSales);
        if (cardsSold !== undefined) sheet.getRange(i + 1, headers.indexOf('CardsSold') + 1).setValue(cardsSold);
        sheet.getRange(i + 1, headers.indexOf('LoggedAt') + 1).setValue(new Date().toISOString());
        return { success: true };
      }
    }

    // Append new row
    sheet.appendRow([partnerId, partnerName, month, actualSales, 0, cardsSold || 0, new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================================
// MARKET SALES
// =============================================================================

function logMarketSale(payload) {
  try {
    const { date, marketName, totalSales, misprintSales, cardsSold } = payload;

    let sheet = SPREADSHEET.getSheetByName('MarketSales');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('MarketSales');
      sheet.appendRow(['Date', 'MarketName', 'TotalSales', 'MisprintSales', 'CardsSold', 'LoggedAt']);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    const fieldMap = {
      'Date': date,
      'MarketName': marketName || '',
      'TotalSales': parseFloat(totalSales) || 0,
      'MisprintSales': parseFloat(misprintSales) || 0,
      'CardsSold': parseInt(cardsSold) || 0,
      'LoggedAt': new Date().toISOString()
    };
    Object.entries(fieldMap).forEach(([col, val]) => {
      const idx = headers.indexOf(col);
      if (idx !== -1) newRow[idx] = val;
    });
    sheet.appendRow(newRow);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getMarketSales() {
  try {
    const sheet = SPREADSHEET.getSheetByName('MarketSales');
    if (!sheet) return { success: true, sales: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, sales: [] };

    const headers = data.shift();
    const sales = data.map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    }).filter(s => s.Date);

    return { success: true, sales };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// =============================================================================
// VENDING MACHINES
// =============================================================================

function getVendingMachines() {
  try {
    const sheet = SPREADSHEET.getSheetByName('VendingMachines');
    if (!sheet) return { success: false, error: 'Sheet "VendingMachines" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const machines = data.map(row => {
      const m = {};
      headers.forEach((h, i) => { m[h] = row[i]; });
      return m;
    }).filter(m => m.MachineID);

    return { success: true, machines };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function addVendingMachine(machineData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('VendingMachines');
    if (!sheet) return { success: false, error: 'Sheet "VendingMachines" not found.' };

    const data = sheet.getDataRange().getValues();
    const ids = data.slice(1).map(r => r[0]).filter(Boolean);
    const nums = ids.map(id => parseInt(String(id).replace(/\D/g,''),10)).filter(n => !isNaN(n));
    const nextNum = nums.length > 0 ? Math.max(...nums) + 1 : 1;
    const machineId = 'VM' + String(nextNum).padStart(3,'0');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill('');
    const fieldMap = {
      'MachineID': machineId,
      'MachineName': machineData.machineName,
      'MachineType': machineData.machineType || 'Standard',
      'Status': machineData.status || 'In Storage',
      'VenueName': machineData.venueName || '',
      'VenueType': machineData.venueType || '',
      'VenueAddress': machineData.venueAddress || '',
      'VenueCity': machineData.venueCity || '',
      'ContactName': machineData.contactName || '',
      'ContactPhone': machineData.contactPhone || '',
      'ContactEmail': machineData.contactEmail || '',
      'InstallDate': machineData.installDate || '',
      'RemovalDate': machineData.removalDate || '',
      'ResidencyNotes': machineData.residencyNotes || '',
      'NotifyEmail': machineData.notifyEmail || '',
      'NotifyPhone': machineData.notifyPhone || '',
      'NotifyActive': machineData.notifyActive !== undefined ? machineData.notifyActive : true,
      'Photo': machineData.photo || '',
    };
    Object.entries(fieldMap).forEach(([col, val]) => {
      const idx = headers.indexOf(col);
      if (idx !== -1) newRow[idx] = val;
    });
    sheet.appendRow(newRow);

    return { success: true, machineId, message: 'Machine "' + machineData.machineName + '" added.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateVendingMachine(machineData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('VendingMachines');
    if (!sheet) return { success: false, error: 'Sheet "VendingMachines" not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('MachineID');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIndex]) === String(machineData.machineId)) {
        const rowNum = i + 1;
        const fieldMap = {
          'MachineName':    machineData.machineName,
          'MachineType':    machineData.machineType,
          'Status':         machineData.status,
          'VenueName':      machineData.venueName,
          'VenueType':      machineData.venueType,
          'VenueAddress':   machineData.venueAddress,
          'VenueCity':      machineData.venueCity,
          'ContactName':    machineData.contactName,
          'ContactPhone':   machineData.contactPhone,
          'ContactEmail':   machineData.contactEmail,
          'InstallDate':    machineData.installDate,
          'RemovalDate':    machineData.removalDate,
          'ResidencyNotes': machineData.residencyNotes,
          'NotifyEmail':    machineData.notifyEmail,
          'NotifyPhone':    machineData.notifyPhone,
          'NotifyActive':   machineData.notifyActive,
          'Photo':          machineData.photo,
          'RemovedDate':    machineData.removedDate,
        };
        Object.entries(fieldMap).forEach(([field, val]) => {
          const colIndex = headers.indexOf(field);
          if (colIndex !== -1 && val !== undefined) {
            sheet.getRange(rowNum, colIndex + 1).setValue(val);
          }
        });
        return { success: true, message: 'Machine updated.' };
      }
    }
    return { success: false, error: 'Machine not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// PUBLIC STOCK (Retailer-facing page)
// =============================================================================

function getPublicStock() {
  try {
    const itemsResult = getItems();
    if (!itemsResult.success) return itemsResult;

    const tagMap = getItemTagsMap();
    const pendingMap = getPendingStockMap();

    const stockItems = itemsResult.items
      .filter(item => item.Status !== 'Retired')
      .map(item => {
        const pending = pendingMap[String(item.ItemID)] || 0;
        const available = Math.max(0, (parseInt(item.StartingAtHome) || 0) - pending);
        return {
          itemId: item.ItemID,
          displayName: item.DisplayName || item.Name,
          photo: item.Photo || '',
          productType: item.ProductType || '',
          unitPrice: item.UnitPrice || 0,
          totalStock: parseInt(item.StartingAtHome) || 0,
          pending: pending,
          available: available,
          tags: tagMap[String(item.ItemID)] || [],
          wholesalePrice: getWholesalePrice(item.ProductType),
        };
      })
      .filter(item => item.totalStock > 0);

    return { success: true, items: stockItems };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getWholesalePrice(productType) {
  if (!productType) return 2.00;
  for (const [key, price] of Object.entries(WHOLESALE_PRICES)) {
    if (productType.toLowerCase().includes(key.toLowerCase())) return price;
  }
  return 2.00;
}

function getPendingStockMap() {
  try {
    const sheet = SPREADSHEET.getSheetByName('Orders');
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const statusIdx = headers.indexOf('Status');
    const itemsIdx = headers.indexOf('ItemsJSON');

    const pendingMap = {};
    data.forEach(row => {
      const status = row[statusIdx];
      if (status === 'Pending' || status === 'In Progress') {
        try {
          const items = JSON.parse(row[itemsIdx] || '[]');
          items.forEach(item => {
            const id = String(item.itemId);
            pendingMap[id] = (pendingMap[id] || 0) + parseInt(item.qty || 0);
          });
        } catch (e) {}
      }
    });
    return pendingMap;
  } catch (e) {
    return {};
  }
}


// =============================================================================
// ORDERS
// =============================================================================

function getOrders() {
  try {
    let sheet = SPREADSHEET.getSheetByName('Orders');
    if (!sheet) return { success: true, orders: [] };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const orders = data.map(row => {
      const o = {};
      headers.forEach((h, i) => { o[h] = row[i]; });
      try { o.items = JSON.parse(o.ItemsJSON || '[]'); } catch(e) { o.items = []; }
      return o;
    }).filter(o => o.OrderID);

    return { success: true, orders };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function submitOrder(orderData) {
  try {
    let sheet = SPREADSHEET.getSheetByName('Orders');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('Orders');
      sheet.appendRow([
        'OrderID','SubmittedAt','LocationID','PartnerName','PartnerType',
        'SubmitterName','SubmitterEmail','ItemsJSON','SubTotal','TaxAmount',
        'EstTotal','Status','Notes','FulfilledAt'
      ]);
    }

    const now = new Date();
    const orderId = 'ORD-' + now.getTime();
    const isWholesale = orderData.partnerType === 'wholesale';

    let subTotal = 0;
    if (isWholesale) {
      orderData.items.forEach(item => {
        subTotal += (item.wholesalePrice || 2) * item.qty;
      });
    }
    const taxAmount = isWholesale ? subTotal * VA_SALES_TAX : 0;
    const estTotal = subTotal + taxAmount;

    sheet.appendRow([
      orderId,
      now.toISOString(),
      orderData.locationId,
      orderData.partnerName,
      orderData.partnerType,
      orderData.submitterName,
      orderData.submitterEmail,
      JSON.stringify(orderData.items),
      isWholesale ? subTotal.toFixed(2) : '',
      isWholesale ? taxAmount.toFixed(2) : '',
      isWholesale ? estTotal.toFixed(2) : '',
      'Pending',
      orderData.notes || '',
      ''
    ]);

    sendOrderNotification(orderId, orderData, subTotal, taxAmount, estTotal);
    sendRetailerConfirmation(orderId, orderData, subTotal, taxAmount, estTotal);

    return { success: true, orderId, message: 'Order submitted successfully.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// fulfillOrder — marks order fulfilled, deducts stock, and optionally sends restock notification
// notifData is optional — only passed for consignment orders when Angel fills out the notification form
function fulfillOrder(orderId, adjustments, notifData) {
  try {
    const sheet = SPREADSHEET.getSheetByName('Orders');
    if (!sheet) return { success: false, error: 'Orders sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx      = headers.indexOf('OrderID');
    const statusIdx  = headers.indexOf('Status');
    const itemsIdx   = headers.indexOf('ItemsJSON');
    const notesIdx   = headers.indexOf('Notes');
    const fulfilledIdx = headers.indexOf('FulfilledAt');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(orderId)) {
        const rowNum = i + 1;

        if (adjustments) {
          sheet.getRange(rowNum, itemsIdx + 1).setValue(JSON.stringify(adjustments));
        }
        sheet.getRange(rowNum, statusIdx + 1).setValue('Fulfilled');
        sheet.getRange(rowNum, fulfilledIdx + 1).setValue(new Date().toISOString());

        // Deduct stock
        const items = adjustments || JSON.parse(data[i][itemsIdx] || '[]');
        updateStockAfterFulfillment(items);

        // Send restock notification if notifData provided (consignment orders)
        if (notifData) {
          try {
            sendRestockNotification(notifData);
            // Log that notification was sent
            if (notesIdx !== -1) {
              const existingNotes = data[i][notesIdx] || '';
              const noteAppend = existingNotes
                ? existingNotes + ' | Restock notification sent ' + new Date().toLocaleDateString()
                : 'Restock notification sent ' + new Date().toLocaleDateString();
              sheet.getRange(rowNum, notesIdx + 1).setValue(noteAppend);
            }
          } catch(e) {}
        }

        return { success: true, message: 'Order fulfilled.' };
      }
    }
    return { success: false, error: 'Order not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// unfulfillOrder — sets a fulfilled order back to Pending (does NOT restore stock)
function unfulfillOrder(orderId) {
  try {
    const sheet = SPREADSHEET.getSheetByName('Orders');
    if (!sheet) return { success: false, error: 'Orders sheet not found.' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx        = headers.indexOf('OrderID');
    const statusIdx    = headers.indexOf('Status');
    const fulfilledIdx = headers.indexOf('FulfilledAt');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(orderId)) {
        const rowNum = i + 1;
        sheet.getRange(rowNum, statusIdx + 1).setValue('Pending');
        if (fulfilledIdx !== -1) sheet.getRange(rowNum, fulfilledIdx + 1).setValue('');
        return { success: true, message: 'Order set back to Pending.' };
      }
    }
    return { success: false, error: 'Order not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function updateStockAfterFulfillment(items) {
  try {
    const sheet = SPREADSHEET.getSheetByName('Items');
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx    = headers.indexOf('ItemID');
    const stockIdx = headers.indexOf('StartingAtHome');

    items.forEach(orderItem => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idIdx]) === String(orderItem.itemId)) {
          const current = parseInt(data[i][stockIdx]) || 0;
          const newStock = Math.max(0, current - parseInt(orderItem.qty || 0));
          sheet.getRange(i + 1, stockIdx + 1).setValue(newStock);
          break;
        }
      }
    });
  } catch (e) {}
}


// =============================================================================
// RESTOCK NOTIFICATION — sends approval email to retailer after fulfillment
// =============================================================================
//
// notifData shape:
// {
//   partnerName:    'Store Name',
//   partnerEmail:   'contact@store.com',   // store contact email
//   visitDateFrom:  '2026-03-15',
//   visitDateTo:    '2026-03-20',
//   approvedItems:  [{ itemId, designName, approvedQty, notes }],
//   declinedItems:  [{ itemId, designName, reason }],          // optional
//   adminNote:      'See you soon!',                           // optional
// }
//
function sendRestockNotification(notifData) {
  try {
    if (!notifData || !notifData.partnerEmail) {
      return { success: false, error: 'No partner email provided.' };
    }

    const approvedRows = (notifData.approvedItems || []).map(item => `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-family:Georgia,serif;color:#4AABAB;font-weight:700;font-size:0.8rem;">#${item.itemId}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;color:#3D2B1F;">${item.designName}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;text-align:center;font-weight:700;color:#5A9E6F;font-size:1.1rem;">${item.approvedQty}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-size:0.82rem;color:#6B4C3B;">${item.notes || '—'}</td>
      </tr>`).join('');

    const declinedRows = (notifData.declinedItems || []).map(item => `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-family:Georgia,serif;color:#4AABAB;font-weight:700;font-size:0.8rem;">#${item.itemId}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;color:#3D2B1F;">${item.designName}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-size:0.82rem;color:#E05C45;">${item.reason || 'Out of stock'}</td>
      </tr>`).join('');

    const visitRange = notifData.visitDateFrom && notifData.visitDateTo
      ? `${new Date(notifData.visitDateFrom).toLocaleDateString('en-US', {month:'long',day:'numeric'})} – ${new Date(notifData.visitDateTo).toLocaleDateString('en-US', {month:'long',day:'numeric',year:'numeric'})}`
      : (notifData.visitDateFrom ? new Date(notifData.visitDateFrom).toLocaleDateString() : 'TBD');

    const body = `
      <div style="font-family:Georgia,serif;max-width:620px;margin:0 auto;color:#3D2B1F;background:#F5F0E4;">

        <!-- Header -->
        <div style="background:#2C1F17;padding:28px 32px;text-align:center;border-radius:12px 12px 0 0;">
          <div style="font-size:1.5rem;color:#F0E6D3;font-family:Georgia,serif;letter-spacing:0.04em;">✦ Prints by Angel</div>
          <div style="color:#C4A882;font-size:0.8rem;margin-top:4px;text-transform:uppercase;letter-spacing:0.12em;">Consignment Restock Approval</div>
        </div>

        <!-- Body -->
        <div style="padding:28px 32px;background:#F5F0E4;">
          <p style="font-size:1rem;color:#3D2B1F;margin:0 0 8px;">Hi ${notifData.partnerName},</p>
          <p style="font-size:0.9rem;color:#6B4C3B;margin:0 0 20px;line-height:1.6;">
            Great news — your restock request has been reviewed and approved! Here's a summary of what's coming your way.
          </p>

          <!-- Visit window -->
          <div style="background:white;border-radius:10px;padding:14px 18px;margin-bottom:24px;border-left:4px solid #4AABAB;">
            <div style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#A07860;font-weight:700;margin-bottom:4px;">Estimated Visit Window</div>
            <div style="font-size:1.1rem;color:#3D2B1F;font-weight:600;">${visitRange}</div>
          </div>

          <!-- Approved items -->
          <div style="font-family:Georgia,serif;font-size:1rem;color:#3D2B1F;margin-bottom:10px;font-weight:600;">✅ Approved Items</div>
          <table style="width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;margin-bottom:${declinedRows ? '20px' : '24px'};">
            <thead>
              <tr style="background:#4AABAB;color:white;">
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Design #</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Name</th>
                <th style="padding:10px 12px;text-align:center;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Qty</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Notes</th>
              </tr>
            </thead>
            <tbody>${approvedRows || '<tr><td colspan="4" style="padding:12px;color:#A07860;font-style:italic;">No items approved.</td></tr>'}</tbody>
          </table>

          ${declinedRows ? `
          <!-- Declined items -->
          <div style="font-family:Georgia,serif;font-size:1rem;color:#3D2B1F;margin-bottom:10px;font-weight:600;">❌ Unavailable Items</div>
          <table style="width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;margin-bottom:24px;">
            <thead>
              <tr style="background:#E05C45;color:white;">
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Design #</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Name</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Reason</th>
              </tr>
            </thead>
            <tbody>${declinedRows}</tbody>
          </table>` : ''}

          ${notifData.adminNote ? `
          <!-- Admin note -->
          <div style="background:white;border-radius:10px;padding:14px 18px;margin-bottom:24px;border-left:4px solid #E8933A;">
            <div style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#A07860;font-weight:700;margin-bottom:6px;">A Note from Angel</div>
            <div style="font-size:0.9rem;color:#3D2B1F;line-height:1.6;">${notifData.adminNote}</div>
          </div>` : ''}

          <p style="font-size:0.9rem;color:#6B4C3B;margin-top:8px;line-height:1.6;">
            I'll be in touch to confirm the exact visit date. Thank you so much for carrying Prints by Angel! 🖤
          </p>
          <p style="font-size:0.9rem;color:#3D2B1F;margin-top:16px;">— Angel</p>
        </div>

        <!-- Footer -->
        <div style="background:#2C1F17;padding:16px 32px;text-align:center;border-radius:0 0 12px 12px;">
          <div style="font-size:0.75rem;color:#C4A882;">✦ Prints by Angel · Handmade Letterpress Cards</div>
        </div>
      </div>`;

    MailApp.sendEmail({
      to: notifData.partnerEmail,
      subject: `✦ Your Prints by Angel restock is approved — ${visitRange}`,
      htmlBody: body
    });

    return { success: true, message: 'Restock notification sent to ' + notifData.partnerEmail };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// STORE VISIT DATA & REPORT EMAIL
// =============================================================================

function getLatestVisit(partnerId) {
  try {
    const sheet = SPREADSHEET.getSheetByName('retailInventory');
    if (!sheet) return { success: true, entries: [], visitDate: null };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, entries: [], visitDate: null };

    const headers = data[0];
    const idx = {};
    headers.forEach((h, i) => { idx[h] = i; });

    // Find the latest visit date for this partner
    let latestDate = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx['LocationID']]) === String(partnerId)) {
        const vd = String(data[i][idx['VisitDate']]);
        if (!latestDate || vd > latestDate) latestDate = vd;
      }
    }
    if (!latestDate) return { success: true, entries: [], visitDate: null };

    // Get all entries for that visit
    const entries = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx['LocationID']]) === String(partnerId) && String(data[i][idx['VisitDate']]) === latestDate) {
        entries.push({
          itemId: data[i][idx['ItemID']],
          designName: data[i][idx['DesignName']],
          startOnShelf: parseInt(data[i][idx['StartOnShelf']]) || 0,
          endOnShelf: parseInt(data[i][idx['EndOnShelf']]) || 0,
          estimatedSold: parseInt(data[i][idx['EstimatedSold']]) || 0,
          added: parseInt(data[i][idx['Added']]) || 0,
          pulled: parseInt(data[i][idx['Pulled']]) || 0,
          unitPrice: parseFloat(data[i][idx['UnitPrice']]) || 0,
          entryType: data[i][idx['EntryType']],
        });
      }
    }

    return { success: true, entries, visitDate: latestDate };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function sendVisitReport(payload) {
  try {
    const { recipientEmail, partnerId, partnerName, note } = payload;
    if (!recipientEmail) return { success: false, error: 'No email provided.' };

    // Pull latest visit data from sheet
    const visitResult = getLatestVisit(partnerId);
    const entries = visitResult.entries || [];
    const visitDate = visitResult.visitDate || new Date().toLocaleDateString('en-CA');
    const pulledItems = entries.filter(e => (e.pulled || 0) > 0).map(e => ({ designId: e.itemId, designName: e.designName, qty: e.pulled }));
    const soldItems = entries.filter(e => (e.estimatedSold || 0) > 0).map(e => ({ designId: e.itemId, designName: e.designName, qty: e.estimatedSold, revenue: e.estimatedSold * (e.unitPrice || 0) }));
    const addedItems = entries.filter(e => (e.added || 0) > 0).map(e => ({ designId: e.itemId, designName: e.designName, qty: e.added }));

    const addedRows = (addedItems || []).map(item => `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-family:Georgia,serif;color:#4AABAB;font-weight:700;font-size:0.8rem;">#${item.designId}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;color:#3D2B1F;">${item.designName}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;text-align:center;font-weight:700;color:#4AABAB;font-size:1.1rem;">${item.qty}</td>
      </tr>`).join('');

    const pulledRows = (pulledItems || []).map(item => `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-family:Georgia,serif;color:#4AABAB;font-weight:700;font-size:0.8rem;">#${item.designId}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;color:#3D2B1F;">${item.designName}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;text-align:center;font-weight:700;color:#E05C45;font-size:1.1rem;">${item.qty}</td>
      </tr>`).join('');

    const soldRows = (soldItems || []).map(item => `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;font-family:Georgia,serif;color:#4AABAB;font-weight:700;font-size:0.8rem;">#${item.designId}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;color:#3D2B1F;">${item.designName}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;text-align:center;font-weight:700;color:#5A9E6F;font-size:1.1rem;">${item.qty}</td>
        <td style="padding:10px 12px;border-bottom:1px solid #EDE7D6;text-align:right;color:#3D2B1F;">$${(item.revenue || 0).toFixed(2)}</td>
      </tr>`).join('');

    const formattedDate = new Date(visitDate).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

    const body = `
      <div style="font-family:Georgia,serif;max-width:620px;margin:0 auto;color:#3D2B1F;background:#F5F0E4;">
        <div style="background:#2C1F17;padding:28px 32px;text-align:center;border-radius:12px 12px 0 0;">
          <div style="font-size:1.5rem;color:#F0E6D3;font-family:Georgia,serif;letter-spacing:0.04em;">✦ Prints by Angel</div>
          <div style="color:#C4A882;font-size:0.8rem;margin-top:4px;text-transform:uppercase;letter-spacing:0.12em;">Store Visit Report</div>
        </div>

        <div style="padding:28px 32px;background:#F5F0E4;">
          <p style="font-size:1rem;color:#3D2B1F;margin:0 0 8px;">Hi ${partnerName},</p>
          <p style="font-size:0.9rem;color:#6B4C3B;margin:0 0 20px;line-height:1.6;">
            Here's a summary of today's store visit. Please keep this for your records.
          </p>

          <div style="background:white;border-radius:10px;padding:14px 18px;margin-bottom:24px;border-left:4px solid #4AABAB;">
            <div style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#A07860;font-weight:700;margin-bottom:4px;">Visit Date</div>
            <div style="font-size:1.1rem;color:#3D2B1F;font-weight:600;">${formattedDate}</div>
          </div>

          ${addedRows ? `
          <div style="font-family:Georgia,serif;font-size:1rem;color:#3D2B1F;margin-bottom:10px;font-weight:600;">🆕 Added to Shelf</div>
          <table style="width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;margin-bottom:20px;">
            <thead>
              <tr style="background:#4AABAB;color:white;">
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Design #</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Name</th>
                <th style="padding:10px 12px;text-align:center;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Qty Added</th>
              </tr>
            </thead>
            <tbody>${addedRows}</tbody>
          </table>` : ''}

          ${pulledRows ? `
          <div style="font-family:Georgia,serif;font-size:1rem;color:#3D2B1F;margin-bottom:10px;font-weight:600;">📦 Items Pulled</div>
          <table style="width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;margin-bottom:20px;">
            <thead>
              <tr style="background:#E05C45;color:white;">
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Design #</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Name</th>
                <th style="padding:10px 12px;text-align:center;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Qty Pulled</th>
              </tr>
            </thead>
            <tbody>${pulledRows}</tbody>
          </table>` : ''}

          ${soldRows ? `
          <div style="font-family:Georgia,serif;font-size:1rem;color:#3D2B1F;margin-bottom:10px;font-weight:600;">💰 Items Sold Out (Removed from Shelf)</div>
          <table style="width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;margin-bottom:24px;">
            <thead>
              <tr style="background:#5A9E6F;color:white;">
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Design #</th>
                <th style="padding:10px 12px;text-align:left;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Name</th>
                <th style="padding:10px 12px;text-align:center;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Qty Sold</th>
                <th style="padding:10px 12px;text-align:right;font-size:0.75rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;">Revenue</th>
              </tr>
            </thead>
            <tbody>${soldRows}</tbody>
          </table>` : ''}

          ${!addedRows && !pulledRows && !soldRows ? `
          <p style="font-size:0.9rem;color:#6B4C3B;font-style:italic;">No inventory changes during this visit.</p>` : ''}

          ${note ? `
          <div style="background:white;border-radius:10px;padding:14px 18px;margin-bottom:24px;border-left:4px solid #E8933A;">
            <div style="font-size:0.7rem;text-transform:uppercase;letter-spacing:0.1em;color:#A07860;font-weight:700;margin-bottom:6px;">Note</div>
            <div style="font-size:0.9rem;color:#3D2B1F;line-height:1.6;">${note}</div>
          </div>` : ''}

          <p style="font-size:0.9rem;color:#6B4C3B;margin-top:8px;line-height:1.6;">
            Thank you for carrying Prints by Angel! 🖤
          </p>
          <p style="font-size:0.9rem;color:#3D2B1F;margin-top:16px;">— Angel</p>
        </div>

        <div style="background:#2C1F17;padding:16px 32px;text-align:center;border-radius:0 0 12px 12px;">
          <div style="font-size:0.75rem;color:#C4A882;">✦ Prints by Angel · Handmade Letterpress Cards</div>
        </div>
      </div>`;

    MailApp.sendEmail({
      to: recipientEmail,
      subject: `✦ Store Visit Report — ${partnerName} · ${formattedDate}`,
      htmlBody: body
    });

    return { success: true, message: 'Visit report sent to ' + recipientEmail };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// PARTNER ACCOUNT REQUESTS
// =============================================================================

function submitPartnerRequest(requestData) {
  try {
    let sheet = SPREADSHEET.getSheetByName('PartnerRequests');
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet('PartnerRequests');
      sheet.appendRow([
        'RequestID','SubmittedAt','PersonName','StoreName','Address',
        'AccountType','Region','Email','Phone','Status','Notes'
      ]);
    }

    const now = new Date();
    const requestId = 'REQ-' + now.getTime();

    sheet.appendRow([
      requestId,
      now.toISOString(),
      requestData.personName,
      requestData.storeName,
      requestData.address,
      requestData.accountType,
      requestData.region || '',
      requestData.email,
      requestData.phone,
      'New',
      ''
    ]);

    try {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: '🆕 New Partner Request — ' + requestData.storeName,
        htmlBody: `
          <h2>New Partner Account Request</h2>
          <p><strong>Store:</strong> ${requestData.storeName}</p>
          <p><strong>Contact:</strong> ${requestData.personName}</p>
          <p><strong>Type:</strong> ${requestData.accountType}</p>
          <p><strong>Region:</strong> ${requestData.region || 'N/A'}</p>
          <p><strong>Email:</strong> ${requestData.email}</p>
          <p><strong>Phone:</strong> ${requestData.phone}</p>
          <p><strong>Address:</strong> ${requestData.address}</p>
        `
      });
    } catch(e) {}

    return { success: true, requestId, message: 'Request submitted.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// NOTIFICATIONS
// =============================================================================

function sendOrderNotification(orderId, orderData, subTotal, taxAmount, estTotal) {
  try {
    const isWholesale = orderData.partnerType === 'wholesale';
    const itemsHtml = orderData.items.map(item =>
      `<tr>
        <td>${item.itemId}</td>
        <td>${item.designName}</td>
        <td>${item.qty}</td>
        ${isWholesale ? `<td>$${(item.wholesalePrice || 2).toFixed(2)}</td><td>$${((item.wholesalePrice || 2) * item.qty).toFixed(2)}</td>` : ''}
      </tr>`
    ).join('');

    const subject = isWholesale
      ? `🛒 New Wholesale Order — ${orderData.partnerName}`
      : `📦 New Restock Request — ${orderData.partnerName}`;

    const body = `
      <h2>${isWholesale ? 'Wholesale Order' : 'Consignment Restock Request'}</h2>
      <p><strong>Order ID:</strong> ${orderId}</p>
      <p><strong>Partner:</strong> ${orderData.partnerName}</p>
      <p><strong>Submitted by:</strong> ${orderData.submitterName} (${orderData.submitterEmail})</p>
      <table border="1" cellpadding="6" cellspacing="0">
        <tr><th>ID</th><th>Design</th><th>Qty</th>${isWholesale ? '<th>Unit Price</th><th>Line Total</th>' : ''}</tr>
        ${itemsHtml}
      </table>
      ${isWholesale ? `
        <p><strong>Subtotal:</strong> $${subTotal.toFixed(2)}</p>
        <p><strong>Tax (5.3%):</strong> $${taxAmount.toFixed(2)}</p>
        <p><strong>Est. Total:</strong> $${estTotal.toFixed(2)}</p>
        <p><em>This is an estimate. Adjust in the app before fulfilling.</em></p>
      ` : ''}
    `;

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: subject,
      htmlBody: body
    });
  } catch (e) {}
}

function sendRetailerConfirmation(orderId, orderData, subTotal, taxAmount, estTotal) {
  try {
    if (!orderData.submitterEmail) return;

    const isWholesale = orderData.partnerType === 'wholesale';
    const itemsHtml = orderData.items.map(item =>
      `<tr>
        <td style="padding:8px;border-bottom:1px solid #eee;">${item.itemId}</td>
        <td style="padding:8px;border-bottom:1px solid #eee;">${item.designName}</td>
        <td style="padding:8px;border-bottom:1px solid #eee;text-align:center;">${item.qty}</td>
        ${isWholesale ? `<td style="padding:8px;border-bottom:1px solid #eee;text-align:right;">$${((item.wholesalePrice || 2) * item.qty).toFixed(2)}</td>` : ''}
      </tr>`
    ).join('');

    const body = `
      <div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;color:#3D2B1F;">
        <div style="background:#2C1F17;padding:24px;text-align:center;">
          <h1 style="color:#F0E6D3;font-size:24px;margin:0;">✦ Prints by Angel</h1>
          <p style="color:#C4A882;margin:4px 0 0;font-size:14px;">
            ${isWholesale ? 'Wholesale Order Confirmation' : 'Restock Request Received'}
          </p>
        </div>
        <div style="padding:24px;background:#F5F0E4;">
          <p>Hi ${orderData.submitterName},</p>
          <p>We've received your ${isWholesale ? 'wholesale order' : 'restock request'} for <strong>${orderData.partnerName}</strong>.
          We'll review it and be in touch soon!</p>
          <p style="font-size:12px;color:#A07860;"><strong>Order ID:</strong> ${orderId}</p>
          <table style="width:100%;border-collapse:collapse;background:white;border-radius:8px;overflow:hidden;">
            <thead>
              <tr style="background:#4AABAB;color:white;">
                <th style="padding:10px;text-align:left;">Design #</th>
                <th style="padding:10px;text-align:left;">Name</th>
                <th style="padding:10px;text-align:center;">Qty</th>
                ${isWholesale ? '<th style="padding:10px;text-align:right;">Total</th>' : ''}
              </tr>
            </thead>
            <tbody>${itemsHtml}</tbody>
          </table>
          ${isWholesale ? `
            <div style="margin-top:16px;text-align:right;">
              <p>Subtotal: <strong>$${subTotal.toFixed(2)}</strong></p>
              <p>Tax (5.3%): <strong>$${taxAmount.toFixed(2)}</strong></p>
              <p style="font-size:18px;">Est. Total: <strong>$${estTotal.toFixed(2)}</strong></p>
              <p style="font-size:11px;color:#A07860;">*This is an estimate. Final invoice will reflect actual fulfilled quantities.</p>
            </div>
          ` : '<p><em>This is a consignment restock request — no payment required at this time.</em></p>'}
          <p style="margin-top:24px;">Thanks for partnering with us! 🖤</p>
          <p>— Angel, Prints by Angel</p>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: orderData.submitterEmail,
      subject: `Your ${isWholesale ? 'wholesale order' : 'restock request'} — Prints by Angel`,
      htmlBody: body
    });
  } catch (e) {}
}


// =============================================================================
// DASHBOARD STATS
// =============================================================================

function getDashboardStats() {
  try {
    const itemsResult = getItems();
    const partnersResult = getRetailPartners();
    const machinesResult = getVendingMachines();

    const totalDesigns = itemsResult.success ? itemsResult.items.length : 0;
    const totalPartners = partnersResult.success ? partnersResult.partners.length : 0;
    const totalMachines = machinesResult.success ? machinesResult.machines.filter(m => m.Status !== 'In Storage').length : 0;

    let totalInStock = 0;
    if (itemsResult.success) {
      itemsResult.items.forEach(item => {
        totalInStock += parseInt(item.StartingAtHome) || 0;
      });
    }

    // Sum total printed from PrintRuns sheet (the real source of truth)
    let totalPrinted = 0;
    try {
      const prSheet = SPREADSHEET.getSheetByName('PrintRuns');
      if (prSheet) {
        const prData = prSheet.getDataRange().getValues();
        const prHeaders = prData.shift();
        const qtyIdx = prHeaders.indexOf('QtyPrinted') !== -1 ? prHeaders.indexOf('QtyPrinted') : (prHeaders.indexOf('Quantity') !== -1 ? prHeaders.indexOf('Quantity') : 3);
        prData.forEach(row => {
          totalPrinted += parseInt(row[qtyIdx]) || 0;
        });
      }
    } catch (e) {}

    // Sum actual revenue from RetailSales sheet
    let totalRevenue = 0;
    try {
      const salesSheet = SPREADSHEET.getSheetByName('RetailSales');
      if (salesSheet) {
        const salesData = salesSheet.getDataRange().getValues();
        const salesHeaders = salesData.shift();
        const actualIdx = salesHeaders.indexOf('ActualSales');
        if (actualIdx !== -1) salesData.forEach(row => { totalRevenue += parseFloat(row[actualIdx]) || 0; });
      }
    } catch (e) {}

    // Sum market sales revenue
    let marketRevenue = 0;
    try {
      const mktSheet = SPREADSHEET.getSheetByName('MarketSales');
      if (mktSheet) {
        const mktData = mktSheet.getDataRange().getValues();
        const mktHeaders = mktData.shift();
        const tsIdx = mktHeaders.indexOf('TotalSales');
        if (tsIdx !== -1) mktData.forEach(row => { marketRevenue += parseFloat(row[tsIdx]) || 0; });
      }
    } catch (e) {}

    totalRevenue += marketRevenue;

    // Get consignment totals once
    let totalOnConsignment = 0;
    try { const c = getConsignmentTotals(); if (c.success) totalOnConsignment = c.grandTotal; } catch(e) {}

    return {
      success: true,
      stats: {
        totalDesigns,
        totalPartners,
        totalMachines,
        totalPrinted,
        totalInStock,
        totalRevenue,
        totalOnConsignment,
        totalSold: Math.max(0, totalPrinted - totalInStock - totalOnConsignment),
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// PRINT RUN TOTALS — total ever printed per item
// =============================================================================

function getPrintRunTotals() {
  try {
    const sheet = SPREADSHEET.getSheetByName('PrintRuns');
    if (!sheet) return { success: true, totals: {} };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, totals: {} };

    const headers = data.shift();
    const itemIdx = headers.indexOf('ItemID') !== -1 ? headers.indexOf('ItemID') : 2;
    const qtyIdx = headers.indexOf('QtyPrinted') !== -1 ? headers.indexOf('QtyPrinted') : (headers.indexOf('Quantity') !== -1 ? headers.indexOf('Quantity') : 3);

    const totals = {};
    data.forEach(row => {
      const itemId = String(row[itemIdx]);
      totals[itemId] = (totals[itemId] || 0) + (parseInt(row[qtyIdx]) || 0);
    });

    return { success: true, totals };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// CONSIGNMENT TOTALS — sum of latest EndOnShelf per partner per item
// =============================================================================

function getConsignmentTotals() {
  try {
    const sheet = SPREADSHEET.getSheetByName('PartnerStock');
    if (!sheet) return { success: true, totals: {}, grandTotal: 0 };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, totals: {}, grandTotal: 0 };

    const headers = data.shift();
    const itemIdx = headers.indexOf('ItemID');
    const stockIdx = headers.indexOf('CurrentStock');

    const totals = {};
    let grandTotal = 0;
    data.forEach(row => {
      const itemId = String(row[itemIdx]);
      const onShelf = parseInt(row[stockIdx]) || 0;
      if (onShelf > 0) {
        totals[itemId] = (totals[itemId] || 0) + onShelf;
        grandTotal += onShelf;
      }
    });

    return { success: true, totals, grandTotal };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// PHOTO UPLOAD
// =============================================================================

function uploadPhoto(payload) {
  try {
    const { base64, filename } = payload;
    if (!base64) return { success: false, error: 'No image data' };

    let folder;
    const folders = DriveApp.getFoldersByName('PrintsByAngel_Photos');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('PrintsByAngel_Photos');
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg',
      (filename || 'photo') + '_' + Date.now() + '.jpg');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();
    const url = 'https://lh3.googleusercontent.com/d/' + fileId;

    return { success: true, url };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// UPDATE ITEM STATUS (audit retire toggle)
// =============================================================================

function updateItemStatus(payload) {
  try {
    const { itemId, status } = payload;
    const sheet = SPREADSHEET.getSheetByName('Items');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ItemID');
    const statusCol = headers.indexOf('Status');

    if (statusCol === -1) return { success: false, error: 'No Status column in Items sheet' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(itemId)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(status);
        return { success: true };
      }
    }
    return { success: false, error: 'Item not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// MIGRATE retailInventory → PartnerStock (run once)
// =============================================================================

function migrateToPartnerStock() {
  try {
    const oldSheet = SPREADSHEET.getSheetByName('retailInventory');
    if (!oldSheet) return { success: true, message: 'No retailInventory sheet to migrate.' };

    // Create PartnerStock if it doesn't exist
    let stockSheet = SPREADSHEET.getSheetByName('PartnerStock');
    if (!stockSheet) {
      stockSheet = SPREADSHEET.insertSheet('PartnerStock');
      stockSheet.appendRow(['LocationID', 'ItemID', 'DesignName', 'CurrentStock', 'UnitPrice', 'LastUpdated']);
    }

    const oldData = oldSheet.getDataRange().getValues();
    if (oldData.length < 2) return { success: true, message: 'No data to migrate.' };

    const headers = oldData.shift();
    const locIdx = headers.indexOf('LocationID');
    const dateIdx = headers.indexOf('VisitDate');
    const itemIdx = headers.indexOf('ItemID');
    const nameIdx = headers.indexOf('DesignName');
    const shelfIdx = headers.indexOf('EndOnShelf');
    const priceIdx = headers.indexOf('UnitPrice');

    // Find most recent visit date per partner
    const latestVisit = {};
    oldData.forEach(row => {
      const loc = String(row[locIdx]);
      const d = new Date(row[dateIdx]);
      if (!latestVisit[loc] || d > latestVisit[loc]) latestVisit[loc] = d;
    });

    // Extract latest snapshot per partner+item
    const stockMap = {}; // key: "loc|item" → {name, stock, price, date}
    oldData.forEach(row => {
      const loc = String(row[locIdx]);
      const d = new Date(row[dateIdx]);
      if (d.toDateString() === latestVisit[loc].toDateString()) {
        const itemId = String(row[itemIdx]);
        const key = loc + '|' + itemId;
        const stock = parseInt(row[shelfIdx]) || 0;
        if (stock > 0) {
          stockMap[key] = {
            locationId: loc,
            itemId: row[itemIdx],
            designName: row[nameIdx] || '',
            currentStock: stock,
            unitPrice: parseFloat(row[priceIdx]) || 0,
            lastUpdated: d.toLocaleDateString('en-CA')
          };
        }
      }
    });

    // Clear existing PartnerStock data (keep headers)
    if (stockSheet.getLastRow() > 1) {
      stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, stockSheet.getLastColumn()).clearContent();
    }

    // Write migrated data
    const entries = Object.values(stockMap);
    entries.forEach(e => {
      stockSheet.appendRow([e.locationId, e.itemId, e.designName, e.currentStock, e.unitPrice, e.lastUpdated]);
    });

    return { success: true, message: 'Migrated ' + entries.length + ' stock entries to PartnerStock.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// SALES REPORT DATA — aggregated data for the Sales Reports page
// =============================================================================

function getSalesReportData(params) {
  try {
    const partnerId = params ? params.partnerId : null;
    const dateFrom = params && params.dateFrom ? params.dateFrom : null;
    const dateTo = params && params.dateTo ? params.dateTo : null;

    // ── Retail Sales ──
    let retailSales = [];
    const retailSheet = SPREADSHEET.getSheetByName('RetailSales');
    if (retailSheet) {
      const rData = retailSheet.getDataRange().getValues();
      if (rData.length > 1) {
        const rHeaders = rData.shift();
        const rPIdx = rHeaders.indexOf('PartnerID');
        const rNameIdx = rHeaders.indexOf('PartnerName');
        const rMonthIdx = rHeaders.indexOf('Month');
        const rActualIdx = rHeaders.indexOf('ActualSales');
        const rCardsIdx = rHeaders.indexOf('CardsSold');

        rData.forEach(row => {
          if (partnerId && String(row[rPIdx]) !== String(partnerId)) return;
          const month = String(row[rMonthIdx]);
          if (dateFrom && month < dateFrom.substring(0, 7)) return;
          if (dateTo && month > dateTo.substring(0, 7)) return;
          retailSales.push({
            partnerId: String(row[rPIdx]),
            partnerName: row[rNameIdx] || '',
            month: month,
            revenue: parseFloat(row[rActualIdx]) || 0,
            cardsSold: parseInt(row[rCardsIdx]) || 0
          });
        });
      }
    }

    // ── Market Sales ──
    let marketSales = [];
    if (!partnerId) { // market sales aren't per-partner
      const mktSheet = SPREADSHEET.getSheetByName('MarketSales');
      if (mktSheet) {
        const mData = mktSheet.getDataRange().getValues();
        if (mData.length > 1) {
          const mHeaders = mData.shift();
          const mDateIdx = mHeaders.indexOf('Date');
          const mNameIdx = mHeaders.indexOf('MarketName');
          const mTotalIdx = mHeaders.indexOf('TotalSales');
          const mMisprintIdx = mHeaders.indexOf('MisprintSales');
          const mCardsIdx = mHeaders.indexOf('CardsSold');

          mData.forEach(row => {
            const saleDate = row[mDateIdx] instanceof Date
              ? row[mDateIdx].toLocaleDateString('en-CA')
              : String(row[mDateIdx]);
            if (dateFrom && saleDate < dateFrom) return;
            if (dateTo && saleDate > dateTo) return;
            marketSales.push({
              date: saleDate,
              marketName: row[mNameIdx] || '',
              revenue: parseFloat(row[mTotalIdx]) || 0,
              misprintRevenue: parseFloat(row[mMisprintIdx]) || 0,
              cardsSold: parseInt(row[mCardsIdx]) || 0
            });
          });
        }
      }
    }

    // ── Consignment stock data (per partner) ──
    let partnerStockSummary = [];
    const stockSheet = SPREADSHEET.getSheetByName('PartnerStock');
    if (stockSheet) {
      const sData = stockSheet.getDataRange().getValues();
      if (sData.length > 1) {
        const sHeaders = sData.shift();
        const sLocIdx = sHeaders.indexOf('LocationID');
        const sItemIdx = sHeaders.indexOf('ItemID');
        const sNameIdx = sHeaders.indexOf('DesignName');
        const sStockIdx = sHeaders.indexOf('CurrentStock');
        const sPriceIdx = sHeaders.indexOf('UnitPrice');

        const byPartner = {};
        sData.forEach(row => {
          if (partnerId && String(row[sLocIdx]) !== String(partnerId)) return;
          const loc = String(row[sLocIdx]);
          if (!byPartner[loc]) byPartner[loc] = { totalCards: 0, totalValue: 0 };
          const stock = parseInt(row[sStockIdx]) || 0;
          const price = parseFloat(row[sPriceIdx]) || 0;
          byPartner[loc].totalCards += stock;
          byPartner[loc].totalValue += stock * price;
        });
        // Get partner names
        const locSheet = SPREADSHEET.getSheetByName('Locations');
        const partnerNames = {};
        if (locSheet) {
          const lData = locSheet.getDataRange().getValues();
          const lHeaders = lData.shift();
          const lIdIdx = lHeaders.indexOf('LocationID');
          const lNameIdx = lHeaders.indexOf('DisplayName');
          lData.forEach(row => { partnerNames[String(row[lIdIdx])] = row[lNameIdx]; });
        }

        Object.entries(byPartner).forEach(([loc, data]) => {
          partnerStockSummary.push({
            partnerId: loc,
            partnerName: partnerNames[loc] || loc,
            cardsOnShelf: data.totalCards,
            shelfValue: data.totalValue
          });
        });
      }
    }

    // ── Aggregate stats ──
    let totalRetailRevenue = 0, totalRetailCards = 0;
    retailSales.forEach(s => { totalRetailRevenue += s.revenue; totalRetailCards += s.cardsSold; });

    let totalMarketRevenue = 0, totalMarketCards = 0;
    marketSales.forEach(s => { totalMarketRevenue += s.revenue; totalMarketCards += s.cardsSold; });

    // Top store by revenue
    const storeRevenue = {};
    retailSales.forEach(s => {
      storeRevenue[s.partnerName] = (storeRevenue[s.partnerName] || 0) + s.revenue;
    });
    const topStore = Object.entries(storeRevenue).sort((a, b) => b[1] - a[1])[0];

    // Revenue by month (for chart)
    const monthlyRevenue = {};
    retailSales.forEach(s => {
      monthlyRevenue[s.month] = (monthlyRevenue[s.month] || 0) + s.revenue;
    });
    marketSales.forEach(s => {
      const month = s.date ? s.date.substring(0, 7) : '';
      if (month) monthlyRevenue[month] = (monthlyRevenue[month] || 0) + s.revenue;
    });

    // Sales by store (for chart)
    const salesByStore = Object.entries(storeRevenue)
      .sort((a, b) => b[1] - a[1])
      .map(([name, revenue]) => ({ name, revenue }));

    // Top designs sold (from visit log estimated sold)
    let designSales = {};
    const logSheet = SPREADSHEET.getSheetByName('retailInventory');
    if (logSheet) {
      const logData = logSheet.getDataRange().getValues();
      if (logData.length > 1) {
        const logHeaders = logData.shift();
        const logNameIdx = logHeaders.indexOf('DesignName');
        const logSoldIdx = logHeaders.indexOf('EstimatedSold');
        const logDateIdx = logHeaders.indexOf('VisitDate');

        logData.forEach(row => {
          const saleDate = row[logDateIdx] instanceof Date
            ? row[logDateIdx].toLocaleDateString('en-CA')
            : String(row[logDateIdx]);
          if (dateFrom && saleDate < dateFrom) return;
          if (dateTo && saleDate > dateTo) return;
          const name = row[logNameIdx] || '';
          const sold = parseInt(row[logSoldIdx]) || 0;
          if (name && sold > 0) {
            designSales[name] = (designSales[name] || 0) + sold;
          }
        });
      }
    }
    const topDesigns = Object.entries(designSales)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, sold]) => ({ name, sold }));

    // Partner list for dropdown
    let partnerList = [];
    const locSheetForList = SPREADSHEET.getSheetByName('Locations');
    if (locSheetForList) {
      const lData = locSheetForList.getDataRange().getValues();
      const lHeaders = lData.shift();
      const lIdIdx = lHeaders.indexOf('LocationID');
      const lNameIdx = lHeaders.indexOf('DisplayName');
      const lTypeIdx = lHeaders.indexOf('LocationType');
      const lActiveIdx = lHeaders.indexOf('Active');
      lData.forEach(row => {
        if (row[lTypeIdx] === 'retail_partner' && (row[lActiveIdx] === true || String(row[lActiveIdx]).toUpperCase() === 'TRUE')) {
          partnerList.push({ id: String(row[lIdIdx]), name: row[lNameIdx] });
        }
      });
    }

    return {
      success: true,
      data: {
        retailSales,
        marketSales,
        partnerStockSummary,
        topDesigns,
        partnerList,
        stats: {
          totalRevenue: totalRetailRevenue + totalMarketRevenue,
          totalRetailRevenue,
          totalMarketRevenue,
          totalCardsSold: totalRetailCards + totalMarketCards,
          topStore: topStore ? topStore[0] : '—',
          topDesign: topDesigns.length > 0 ? topDesigns[0].name : '—'
        },
        charts: {
          monthlyRevenue: Object.entries(monthlyRevenue)
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([month, revenue]) => ({ month, revenue })),
          salesByStore,
          topDesigns
        }
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

