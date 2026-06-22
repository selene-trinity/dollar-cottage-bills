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
    const adamShare    = parseFloat(row[COL.adamShare - 1]) || 0;
    const ivyShare     = parseFloat(row[COL.ivyShare - 1]) || 0;
    const myShare      = parseFloat(row[COL.myShare - 1]) || 0;
    const total        = parseFloat(row[COL.total - 1]) || 0;

    // Skip rows with no usable data
    if (!type || total === 0) {
      skipCount++;
      return;
    }

    // Parse due date — handles Date objects and "DD/MM/YYYY" strings
    const dueDate = parseDueDate(dueDateRaw);

    const description = `${type} - ${capitalise(supplyPeriod)}`;

    const payload = {
      cost:              total.toFixed(2),
      description:       description,
      date:              dueDate,               // ISO 8601: "YYYY-MM-DD"
      currency_code:     "AUD",
      group_id:          GROUP_ID,
      split_equally:     false,

      // Who paid — adjust if the bill payer rotates
      // Splitwise requires paid_by_<n> + owed_share_<n> for each user
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

      if (result.errors && Object.keys(result.errors).length > 0) {
        errors.push(`Row ${sheetRow} (${description}): ${JSON.stringify(result.errors)}`);
      } else if (result.expense && result.expense.id) {
        // Mark as imported in column K
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

// ---- Optional: adds a menu item in the sheet UI ----
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Splitwise")
    .addItem("Import outstanding bills", "importToSplitwise")
    .addToUi();
}
