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
      'Location': 'HOME',
      'CreatedAt': now,
      'StartingAtHome': printRunData ? parseInt(printRunData.quantity) || 0 : 0,
      'StartingApproxTotal': 0,
      'StartingOutConsignmentEst': 0,
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

    sheet.appendRow([runId, runDate, itemId, parseInt(quantity), '', '', '', now]);

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
    const inventorySheet = SPREADSHEET.getSheetByName('retailInventory');
    const locationsSheet = SPREADSHEET.getSheetByName('Locations');
    if (!inventorySheet || !locationsSheet) {
      return { success: false, error: 'Required sheet not found.' };
    }

    const allItems = getItems().items;
    const itemNames = allItems.reduce((acc, item) => {
      acc[parseInt(item.ItemID)] = item.DisplayName || item.Name;
      return acc;
    }, {});
    const itemPrices = allItems.reduce((acc, item) => {
      acc[parseInt(item.ItemID)] = parseFloat(item.UnitPrice) || 0;
      return acc;
    }, {});

    const locationsData = locationsSheet.getDataRange().getValues();
    const locationsHeaders = locationsData.shift();
    const locIdIndex = locationsHeaders.indexOf('LocationID');
    const locNameIndex = locationsHeaders.indexOf('DisplayName');
    const partnerRow = locationsData.find(row => row[locIdIndex] === locationId);
    const partnerName = partnerRow ? partnerRow[locNameIndex] : 'Unknown Partner';

    const inventoryData = inventorySheet.getDataRange().getValues();
    const inventoryHeaders = inventoryData.shift();
    const invLocationIdIndex = inventoryHeaders.indexOf('LocationID');
    const visitDateIndex = inventoryHeaders.indexOf('VisitDate');
    const itemIdIndex = inventoryHeaders.indexOf('ItemID');
    const endOnShelfIndex = inventoryHeaders.indexOf('EndOnShelf');

    let mostRecentDate = new Date(0);
    inventoryData.forEach(row => {
      if (String(row[invLocationIdIndex]) === locationId) {
        const visitDate = new Date(row[visitDateIndex]);
        if (visitDate > mostRecentDate) mostRecentDate = visitDate;
      }
    });

    if (mostRecentDate.getTime() === 0) {
      return { success: true, data: { name: partnerName, lastVisit: 'N/A', inventory: [] } };
    }

    const latestInventory = inventoryData
      .filter(row => {
        const rowDate = new Date(row[visitDateIndex]);
        return String(row[invLocationIdIndex]) === locationId &&
               rowDate.toDateString() === mostRecentDate.toDateString();
      })
      .map(row => ({
        designId:   row[itemIdIndex],
        designName: itemNames[parseInt(row[itemIdIndex])] || 'Unknown Design',
        unitPrice:  itemPrices[parseInt(row[itemIdIndex])] || 0,
        currentStock: row[endOnShelfIndex] || 0
      }));

    return {
      success: true,
      data: { name: partnerName, lastVisit: mostRecentDate.toLocaleDateString(), inventory: latestInventory }
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
    const inventorySheet = SPREADSHEET.getSheetByName('retailInventory');
    if (!inventorySheet) return { success: false, error: 'retailInventory sheet not found.' };

    const today = visitDate || new Date().toLocaleDateString('en-CA');

    // Ensure headers exist — getPartnerInventory reads by header name
    const firstRow = inventorySheet.getRange(1, 1, 1, inventorySheet.getLastColumn()).getValues()[0];
    const hasHeaders = firstRow[0] === 'LocationID';
    if (!hasHeaders || inventorySheet.getLastColumn() === 0) {
      inventorySheet.getRange(1, 1, 1, 12).setValues([[
        'LocationID', 'PartnerName', 'VisitDate', 'ItemID', 'DesignName',
        'StartOnShelf', 'Added', 'Pulled', 'EndOnShelf', 'EstimatedSold', 'UnitPrice', 'EntryType'
      ]]);
    }

    updates.forEach(u => {
      inventorySheet.appendRow([
        partnerId,        // LocationID
        partnerName,      // PartnerName
        today,            // VisitDate
        u.designId,       // ItemID
        u.designName,     // DesignName
        u.previousStock,  // StartOnShelf
        u.added || 0,     // Added
        u.pulled || 0,    // Pulled
        u.newStock,       // EndOnShelf
        u.estimatedSold || 0,
        u.unitPrice || 0,
        u.isNew ? 'New' : 'Update',
      ]);
    });

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
      .filter(item => item.Active === true || item.Active === 'TRUE')
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

    let totalPrinted = 0;
    let totalInStock = 0;
    if (itemsResult.success) {
      itemsResult.items.forEach(item => {
        totalPrinted += parseInt(item.InitialPrintQty) || 0;
        totalInStock += parseInt(item.StartingAtHome) || 0;
      });
    }

    // Sum actual revenue and cards sold from RetailSales sheet
    let totalRevenue = 0;
    let totalSold = 0;
    try {
      const salesSheet = SPREADSHEET.getSheetByName('RetailSales');
      if (salesSheet) {
        const salesData = salesSheet.getDataRange().getValues();
        const salesHeaders = salesData.shift();
        const actualIdx = salesHeaders.indexOf('ActualSales');
        const cardsIdx = salesHeaders.indexOf('CardsSold');
        salesData.forEach(row => {
          if (actualIdx !== -1) totalRevenue += parseFloat(row[actualIdx]) || 0;
          if (cardsIdx !== -1) totalSold += parseInt(row[cardsIdx]) || 0;
        });
      }
    } catch (e) {}

    return {
      success: true,
      stats: {
        totalDesigns,
        totalPartners,
        totalMachines,
        totalPrinted,
        totalInStock,
        totalRevenue,
        totalSold,
        totalOnConsignment: (() => { try { const c = getConsignmentTotals(); return c.success ? c.grandTotal : 0; } catch(e) { return 0; } })(),
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// =============================================================================
// CONSIGNMENT TOTALS — sum of latest EndOnShelf per partner per item
// =============================================================================

function getConsignmentTotals() {
  try {
    const sheet = SPREADSHEET.getSheetByName('retailInventory');
    if (!sheet) return { success: true, totals: {} };

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, totals: {} };

    const headers = data.shift();
    const locIdx = headers.indexOf('LocationID');
    const dateIdx = headers.indexOf('VisitDate');
    const itemIdx = headers.indexOf('ItemID');
    const shelfIdx = headers.indexOf('EndOnShelf');

    // Find the most recent visit date per partner
    const latestVisit = {};
    data.forEach(row => {
      const loc = String(row[locIdx]);
      const d = new Date(row[dateIdx]);
      if (!latestVisit[loc] || d > latestVisit[loc]) latestVisit[loc] = d;
    });

    // Sum EndOnShelf for each item across all partners (latest visit only)
    const totals = {};
    let grandTotal = 0;
    data.forEach(row => {
      const loc = String(row[locIdx]);
      const d = new Date(row[dateIdx]);
      if (d.toDateString() === latestVisit[loc].toDateString()) {
        const itemId = String(row[itemIdx]);
        const onShelf = parseInt(row[shelfIdx]) || 0;
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

