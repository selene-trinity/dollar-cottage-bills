// ============================================================
// SPLITWISE BILL IMPORTER — TEST VERSION
// Uses a solo test group (just you) to verify expenses look
// correct before running against the real shared group.
//
// IMPORTANT: Keep this file in the same Apps Script project
// as splitwise_import.gs. It intentionally re-uses the
// constants declared there (SPLITWISE_API_KEY, COL,
// DATA_START_ROW) — do not redeclare them here.
// ============================================================

// --- TEST-SPECIFIC CONFIG ---
const TEST_GROUP_ID = "YOUR_TEST_GROUP_ID_HERE";  // group with only you in it

// Column L for the test import marker, separate from
// production's column K so they don't interfere.
const TEST_IMPORTED_COL = 12;  // L

// ============================================================

function importToSplitwiseTest() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("No data rows found.");
    return;
  }

  const rows = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, TEST_IMPORTED_COL).getValues();

  let successCount = 0;
  let skipCount    = 0;
  const errors     = [];

  rows.forEach((row, i) => {
    const sheetRow = DATA_START_ROW + i;

    // Skip if already test-imported
    if (row[TEST_IMPORTED_COL - 1] && row[TEST_IMPORTED_COL - 1].toString().includes("Test ✓")) {
      skipCount++;
      return;
    }

    const supplyPeriod = row[COL.supplyPeriod - 1].toString().trim();
    const dueDateRaw   = row[COL.dueDate - 1];
    const type         = row[COL.type - 1].toString().trim();
    const total        = Math.round(parseFloat(row[COL.total - 1]) * 100) / 100;

    if (!type || total === 0) {
      skipCount++;
      return;
    }

    const dueDate     = parseDueDate(dueDateRaw);
    const description = `[TEST] ${type} - ${capitalise(supplyPeriod)}`;

    const CATEGORY_IDS = {
      "Electricity": 5,
      "NBN":         8,
    };
    const categoryId = CATEGORY_IDS[type] || 18;  // 18 = General fallback

    // In a solo group, you are both the payer and the only person who owes.
    // The full total is recorded as paid and owed by you so the expense
    // shows the real amount without involving anyone else.
    const payload = {
      cost:          total.toFixed(2),
      description:   description,
      date:          dueDate,
      // No currency_code — use the group's default currency.
      category_id:   categoryId,
      group_id:      TEST_GROUP_ID,
      split_equally: false,

      users__0__user_id:    USER_IDS.me,
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

      const result    = JSON.parse(response.getContentText());
      const hasErrors = result.errors && Object.keys(result.errors).length > 0;
      const expense   = result.expenses && result.expenses[0];

      if (hasErrors) {
        errors.push(`Row ${sheetRow} (${description}): ${JSON.stringify(result.errors)}`);
      } else if (expense && expense.id) {
        sheet.getRange(sheetRow, TEST_IMPORTED_COL).setValue("Test ✓");
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
