// ============================================================
// SPLITWISE BILL IMPORTER
// Paste this into Tools > Script editor in your Google Sheet
// ============================================================

// --- CONFIGURATION: fill these in before running ---
const SPLITWISE_API_KEY = "YOUR_API_KEY_HERE";
const GROUP_ID          = "YOUR_GROUP_ID_HERE";   // numeric, from get_groups

const USER_IDS = {
  me:   "YOUR_USER_ID_HERE",   // numeric
  adam: "ADAM_USER_ID_HERE",
  ivy:  "IVY_USER_ID_HERE",
};

// Column indices (1-based, matching your sheet layout)
const COL = {
  supplyPeriod:  1,   // A - e.g. "August"
  dueDate:       2,   // B - e.g. "24/09/2024"
  type:          3,   // C - "Electricity" or "NBN"
  adamShare:     4,   // D
  ivyShare:      5,   // E
  myShare:       6,   // F
  total:         7,   // G
  imported:      11,  // K - script writes "Imported ✓" here when done
};

// Row where data starts (1-based; assumes row 1 is a header)
const DATA_START_ROW = 2;

// ============================================================

function importToSplitwise() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No data rows found.");
    return;
  }

  const range = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 11);
  const rows  = range.getValues();

  let successCount = 0;
  let skipCount    = 0;
  const errors     = [];

  rows.forEach((row, i) => {
    const sheetRow = DATA_START_ROW + i;

    // Skip if already imported
    if (row[COL.imported - 1] && row[COL.imported - 1].toString().includes("Imported")) {
      skipCount++;
      return;
    }

    const supplyPeriod = row[COL.supplyPeriod - 1].toString().trim();
    const dueDateRaw   = row[COL.dueDate - 1];
    const type         = row[COL.type - 1].toString().trim();
    const total        = Math.round(parseFloat(row[COL.total - 1]) * 100) / 100;

    // Skip rows with no usable data
    if (!type || total === 0) {
      skipCount++;
      return;
    }

    // Round two shares from the sheet; derive the third as the remainder
    // so the three values always sum to exactly the total, avoiding the
    // "owed shares don't match total cost" error from floating point drift.
    const adamShare = Math.round(parseFloat(row[COL.adamShare - 1]) * 100) / 100;
    const ivyShare  = Math.round(parseFloat(row[COL.ivyShare - 1])  * 100) / 100;
    const myShare   = Math.round((total - adamShare - ivyShare) * 100) / 100;

    // Parse due date — handles Date objects and "DD/MM/YYYY" strings
    const dueDate = parseDueDate(dueDateRaw);

    const description = `${type} - ${capitalise(supplyPeriod)}`;

    const CATEGORY_IDS = {
      "Electricity": 5,
      "NBN":         8,
    };
    const categoryId = CATEGORY_IDS[type] || 18;  // 18 = General fallback

    const payload = {
      cost:          total.toFixed(2),
      description:   description,
      date:          dueDate,          // ISO 8601: "YYYY-MM-DD"
      // No currency_code — lets Splitwise use the group's default currency
      // rather than triggering a conversion to USD.
      category_id:   categoryId,
      group_id:      GROUP_ID,
      split_equally: false,

      // Splitwise requires paid_share + owed_share for each user.
      // You paid the full bill; Adam and Ivy owe their shares back to you.
      users__0__user_id:    USER_IDS.me,
      users__0__paid_share: total.toFixed(2),
      users__0__owed_share: myShare.toFixed(2),

      users__1__user_id:    USER_IDS.adam,
      users__1__paid_share: "0.00",
      users__1__owed_share: adamShare.toFixed(2),

      users__2__user_id:    USER_IDS.ivy,
      users__2__paid_share: "0.00",
      users__2__owed_share: ivyShare.toFixed(2),
    };

    try {
      const response = UrlFetchApp.fetch("https://secure.splitwise.com/api/v3.0/create_expense", {
        method:  "POST",
        headers: {
          "Authorization": `Bearer ${SPLITWISE_API_KEY}`,
          "Content-Type":  "application/x-www-form-urlencoded",
        },
        payload:            encodePayload(payload),
        muteHttpExceptions: true,
      });

      const result = JSON.parse(response.getContentText());

      // The API returns "expenses" (plural array) on success, not "expense".
      // An empty errors object ({}) means success; a populated one means failure.
      const hasErrors = result.errors && Object.keys(result.errors).length > 0;
      const expense   = result.expenses && result.expenses[0];

      if (hasErrors) {
        errors.push(`Row ${sheetRow} (${description}): ${JSON.stringify(result.errors)}`);
      } else if (expense && expense.id) {
        sheet.getRange(sheetRow, COL.imported).setValue("Imported ✓");
        successCount++;
      } else {
        errors.push(`Row ${sheetRow} (${description}): unexpected response — ${response.getContentText()}`);
      }

    } catch (e) {
      errors.push(`Row ${sheetRow} (${description}): ${e.message}`);
    }

    // Brief pause to avoid hammering the API
    Utilities.sleep(300);
  });

  // Summary alert
  let summary = `Done.\n✓ ${successCount} imported, ⏭ ${skipCount} skipped.`;
  if (errors.length > 0) {
    summary += `\n\n⚠️ ${errors.length} error(s):\n` + errors.join("\n");
  }
  SpreadsheetApp.getUi().alert(summary);
}

// ---- Helpers ----

/**
 * Accepts a JS Date object (from Sheets) or a "DD/MM/YYYY" string,
 * returns "YYYY-MM-DD" for the Splitwise API.
 */
function parseDueDate(raw) {
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  // Handle string "DD/MM/YYYY"
  const parts = raw.toString().trim().split("/");
  if (parts.length === 3) {
    return `${parts[2]}-${parts[1].padStart(2, "0")}-${parts[0].padStart(2, "0")}`;
  }
  // Fallback: today
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/** Capitalises first letter of a string. */
function capitalise(str) {
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

/**
 * Encodes a flat object as application/x-www-form-urlencoded.
 * The Splitwise v3 API requires this format (not JSON) for expense creation.
 */
function encodePayload(obj) {
  return Object.entries(obj)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join("&");
}

// ---- Adds menu items in the sheet UI ----
// onOpen must only be declared once across the whole project.
// The test menu item calls importToSplitwiseTest() defined in
// splitwise_import_test.gs.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Splitwise")
    .addItem("Import outstanding bills", "importToSplitwise")
    .addItem("TEST: Import to solo group", "importToSplitwiseTest")
    .addToUi();
}
