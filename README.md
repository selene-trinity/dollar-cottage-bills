# Dollar Cottage Bills

A Google Apps Script that imports outstanding household utility bills from a Google Sheet into a Splitwise group, split evenly between three housemates.

## What it does

- Reads electricity and NBN bill rows from a filtered Google Sheet
- Creates one Splitwise expense per row, dated to the bill due date
- Marks each row as imported to prevent duplicates on re-runs
- Adds a **Splitwise** menu item to the sheet for one-click importing

A test variant (`splitwise_import_test.gs`) imports into a solo group first so you can verify bills look correct before they hit the shared group.

## Setup

### 1. Splitwise API key

Register a new application at [secure.splitwise.com/apps](https://secure.splitwise.com/apps) to get an API key.

### 2. Get your IDs

With your API key in hand, call these endpoints in a browser (while logged into Splitwise):

```
https://secure.splitwise.com/api/v3.0/get_current_user   → your user ID
https://secure.splitwise.com/api/v3.0/get_friends         → housemates' user IDs
https://secure.splitwise.com/api/v3.0/get_groups          → group ID
```

### 3. Configure the script

Open `splitwise_import.gs` and fill in the config block at the top:

```javascript
const SPLITWISE_API_KEY = "your_api_key";
const GROUP_ID          = "your_group_id";

const USER_IDS = {
  me:   "your_user_id",
  adam: "adam_user_id",
  ivy:  "ivy_user_id",
};
```

### 4. Add to Google Sheets

1. Open your bill tracking Google Sheet
2. Go to **Extensions → Apps Script**
3. Paste in the contents of `splitwise_import.gs` (and optionally `splitwise_import_test.gs`)
4. Save and reload the sheet — a **Splitwise** menu will appear in the toolbar

### 5. Run

Use the **Splitwise → Import outstanding bills** menu item. A summary alert will show how many bills were imported and flag any errors.

## Sheet structure

The script expects columns in this order:

| Col | Field |
|-----|-------|
| A | Supply period (e.g. "August") |
| B | Due date (DD/MM/YYYY) |
| C | Type ("Electricity" or "NBN") |
| D | Adam's share |
| E | Ivy's share |
| F | My share |
| G | Total |
| H | Outstanding |
| I | Days |
| J | $ per day |
| K | Import marker (written by script) |

## Files

| File | Purpose |
|------|---------|
| `splitwise_import.gs` | Production script — imports to the shared housemate group |
| `splitwise_import_test.gs` | Test script — imports to a solo group for verification |

## Notes

- Currency is set to AUD — change `currency_code` in the script if needed
- The script assumes you paid each bill in full; Adam and Ivy owe their shares back to you
- `Utilities.sleep(300)` adds a small delay between API calls to avoid rate limiting
