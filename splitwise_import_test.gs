// ============================================================
// SPLITWISE BILL IMPORTER — TEST VERSION
// Uses a solo test group (just you) to verify expenses look
// correct before running against the real shared group.
//
// Paste into Tools > Script editor in your Google Sheet.
// ============================================================

// --- CONFIGURATION ---
const SPLITWISE_API_KEY  = "YOUR_API_KEY_HERE";
const TEST_GROUP_ID      = "YOUR_TEST_GROUP_ID_HERE";  // group with only you in it
const MY_USER_ID         = "YOUR_USER_ID_HERE";

// Column indices (1-based)
const COL = {
  supplyPeriod: 1,   // A
  dueDate:      2,   // B
  type:         3,   // C
  adamShare:    4,   // D
  ivyShare:     5,   // E
  myShare:      6,   // F
  total:        7,   // G
  imported:     12,  // L — separate column from production (K) so markers don't clash
};

const DATA_START_ROW = 2;

// ============================================================

function importToSplitwiseTest() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No data rows found.");
    return;
  }

  const rows = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();

  let successCount = 0;
  let skipCount    = 0;
  const errors     = [];

  rows.forEach((row, i) => {
    const sheetRow = DATA_START_ROW + i;

    // Skip if already test-imported
    if (row[COL.imported - 1] && row[COL.imported - 1].toString().includes("Test ✓")) {
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

    if (!type || total === 0) {
      skipCount++;
      return;
    }

    const dueDate     = parseDueDate(dueDateRaw);
    const description = `[TEST] ${type} - ${capitalise(supplyPeriod)}`;

    // In a solo group, you are both the payer and the only person who owes.
    // We record the full total as paid and owed by you, so the expense
    // still shows the real amount without involving anyone else.
    const payload = {
      cost:             total.toFixed(2),
      description:      description,
      date:             dueDate,
      currency_code:    "AUD",
      group_id:         TEST_GROUP_ID,
      split_equally:    false,

      users__0__user_id:    MY_USER_ID,
      users__0__paid_share: total.toFixed(2),
      users__0__owed_share: total.toFixed(2),
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
        sheet.getRange(sheetRow, COL.imported).setValue("Test ✓");
        successCount++;
      } else {
        errors.push(`Row ${sheetRow} (${description}): unexpected response — ${response.getContentText()}`);
      }

    } catch (e) {
      errors.push(`Row ${sheetRow} (${description}): ${e.message}`);
    }

    Utilities.sleep(300);
  });

  let summary = `Test run done.\n✓ ${successCount} imported to test group, ⏭ ${skipCount} skipped.`;
  if (errors.length > 0) {
    summary += `\n\n⚠️ ${errors.length} error(s):\n` + errors.join("\n");
  }
  SpreadsheetApp.getUi().alert(summary);
}

// ---- Helpers ----

function parseDueDate(raw) {
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const parts = raw.toString().trim().split("/");
  if (parts.length === 3) {
    return `${parts[2]}-${parts[1].padStart(2, "0")}-${parts[0].padStart(2, "0")}`;
  }
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function capitalise(str) {
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

function encodePayload(obj) {
  return Object.entries(obj)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join("&");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Splitwise")
    .addItem("Import outstanding bills", "importToSplitwise")          // production
    .addItem("TEST: Import to solo group", "importToSplitwiseTest")    // test
    .addToUi();
}
