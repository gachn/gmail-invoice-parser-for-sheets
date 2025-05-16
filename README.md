# Google Sheets Expense Analyzer from Gmail

## Description

This Google Apps Script helps you automatically track and analyze your monthly expenses by reading invoice and order confirmation emails from your Gmail account. It categorizes spending from various platforms (e.g., e-commerce, food delivery, subscriptions) and generates detailed monthly reports and an overall summary directly within a Google Sheet.

The script is highly configurable, allowing you to define which platforms to track, the keywords and sender emails to identify relevant messages, and the specific patterns (using regular expressions) to extract the transaction amounts.

## Features

* **Automated Email Parsing:** Scans your Gmail for expense-related emails within a specified date range.
* **Configurable Platforms:** Easily define new platforms, keywords, sender emails, and amount extraction patterns via a `config.json.html` file.
* **Detailed Monthly Reports:** Generates individual sheets for each month, including:
    * A summary of spending per platform.
    * A pie chart visualizing platform spending distribution.
    * A line graph showing daily spending trends for the month.
    * A detailed list of all transactions extracted for that month with links to the original emails.
* **Overall Summary Report:** Creates a consolidated summary sheet with:
    * Total overall expense for the period.
    * Average monthly expense.
    * Trend of total monthly expenses (column chart).
    * Overall spending breakdown by platform (pie chart and table).
    * Average monthly spending per platform.
* **Flexible Report Generation:**
    * **Incremental Updates:** Processes only missing months within the selected date range for faster updates.
    * **Full Regeneration:** Option to delete all previous report sheets and rebuild everything from scratch.
* **Sheet Organization:**
    * "Master Config" sheet is always the first tab.
    * "Overall Summary" sheet is always the second tab.
    * Monthly report sheets are sorted in descending chronological order (newest first).
* **User-Friendly Interface:**
    * Custom menu within Google Sheets for easy operation.
    * Status updates on the "Master Config" sheet regarding recent report generation.
    * Notifications for missing reports for the current/previous month on sheet open.
* **Active Sheet Navigation:** Lands on the current calendar month's report (if available) after script completion.

## Prerequisites

* A Google Account (with Gmail and Google Sheets).
* Basic familiarity with Google Sheets.
* Ability to copy/paste code into the Google Apps Script editor.

## Setup Instructions

1.  **Create a New Google Sheet:**
    * Go to [sheet.new](https://sheet.new) to create a new spreadsheet. Name it whatever you like (e.g., "My Expense Tracker").

2.  **Open the Apps Script Editor:**
    * In your new Google Sheet, click on `Extensions > Apps Script`. A new browser tab will open with the Apps Script editor.
    * Delete any boilerplate code in the default `Code.gs` file.

3.  **Copy `Code.gs` Script:**
    * Copy the entire content of the `Code.gs` file from this GitHub repository.
    * Paste it into the `Code.gs` file in your Apps Script editor.

4.  **Create `config.json.html` File:**
    * In the Apps Script editor, click the `+` icon next to "Files."
    * Choose "HTML."
    * Name the file `config.json` (it will be automatically saved as `config.json.html`).
    * Delete any default HTML content in this new file.
    * Copy the example `config.json.html` content provided below (or from the `config.json.html` file in this repository) and paste it into your newly created `config.json.html` file.

    **Example `config.json.html` Structure:**
    ```json
    [
      {
        "name": "Swiggy",
        "keywords": ["Your Swiggy order has been placed", "Order confirmation from Swiggy", "Swiggy order delivered"],
        "sender": "noreply@swiggy.in",
        "extractionPatterns": [
          "/Order Total\\s*:\\s*₹\\s*([\\d,]+\\.?\\d*)/i",
          "/Paid\\s*(?:Via[^:]*:\\s*)?₹\\s*([\\d,]+\\.?\\d*)/i",
          "/Total Bill\\s*₹\\s*([\\d,]+\\.?\\d*)/i",
          "/Grand Total\\s*₹\\s*([\\d,]+\\.?\\d*)/i"
        ]
      },
      {
        "name": "Amazon",
        "keywords": ["Your Amazon.in order of", "Amazon.in order confirmation"],
        "sender": "order-update@amazon.in",
        "extractionPatterns": [
          "/Order Total:\\s*₹\\s*([\\d,]+\\.?\\d*)/i",
          "/Order Total:\\s*Rs\\.\\s*([\\d,]+\\.?\\d*)/i",
          "/Grand Total:\\s*₹\\s*([\\d,]+\\.?\\d*)/i"
        ]
      }
      // Add more platforms here
    ]
    ```

5.  **IMPORTANT: Customize `config.json.html`:**
    * **`name`**: The display name for the platform in your reports.
    * **`keywords`**: An array of strings. These are terms the script will search for in email subjects or bodies to identify relevant emails.
    * **`sender`**: (Optional but highly recommended) The specific email address from which the platform sends its invoices/confirmations.
    * **`extractionPatterns`**: This is an array of regular expression **strings** used to find and extract the transaction amount from the email body. The script tries them in the order they are listed.
        * **Format:** Each pattern must be a string starting and ending with `/`, followed by flags (e.g., `i` for case-insensitive). Example: `"/Total Amount:\\s*₹\\s*([\\d,]+\\.?\\d*)/i"`
        * **Escaping Backslashes:** Since this is a JSON string, any backslash `\` within your regex pattern **must be escaped with another backslash**.
            * `\d` (digit) becomes `\\d`
            * `\s` (whitespace) becomes `\\s`
            * `\.` (literal dot) becomes `\\.`
        * **Capture Group:** Your regex must have at least one capture group `(...)` that specifically captures the numerical amount (which can include commas and a decimal point). This captured group is what the script will parse as the transaction value.
        * **Example Breakdown:** For `/Order Total\\s*:\\s*₹\\s*([\\d,]+\\.?\\d*)/i`
            * `Order Total\\s*:\\s*₹\\s*`: Matches "Order Total", optional spaces, a colon, optional spaces, the Rupee symbol, optional spaces.
            * `([\\d,]+\\.?\\d*)`: This is the capture group.
                * `[\\d,]+`: Matches one or more digits or commas.
                * `\\.?`: Matches an optional literal decimal point.
                * `\\d*`: Matches zero or more digits after the decimal point.
            * `/i`: Case-insensitive flag.

6.  **Save All Files:**
    * In the Apps Script editor, click the save icon (floppy disk) or `File > Save all`.

7.  **Run Initial Setup:**
    * In the Apps Script editor, select the function `setupMasterSheetAndMenu` from the function dropdown list (next to the "Debug" and "Run" buttons).
    * Click the "Run" (play icon) button.
    * **Authorization:** You will be prompted to authorize the script.
        * Click "Review permissions."
        * Choose your Google account.
        * You might see a "Google hasn't verified this app" screen. If so, click "Advanced" and then "Go to [Your Script Name] (unsafe)." (This is normal for personal scripts).
        * Review the permissions the script needs (access to Gmail to read emails, access to Spreadsheets to create and manage sheets) and click "Allow."
    * This will create/update a "Master Config" sheet in your spreadsheet and add the "Expense Reports" custom menu.

8.  **Configure "Master Config" Sheet:**
    * Go to the "Master Config" sheet in your Google Sheet.
    * Enter the **Start Date** and **End Date** (in `YYYY-MM-DD` format) for the period you want to analyze.
    * Optionally, set the "Max Expected Run Time (minutes)" - this is informational for you.

## How to Use

1.  **Fill in "Master Config":** Ensure your desired "Start Date" and "End Date" are set.
2.  **Use the "Expense Reports" Menu:**
    * **Generate/Update Report (Incremental):** This is the recommended option for regular use. It processes only those months within your selected date range for which a report sheet doesn't already exist. This is much faster for ongoing tracking. The "Overall Summary" is always rebuilt.
    * **Regenerate Full Report (Deletes Old):** Use this if you want a complete refresh. It will ask for confirmation, then delete all existing monthly report sheets and the "Overall Summary" sheet before processing all months in your selected date range from scratch.
3.  **View Reports:**
    * **Monthly Sheets:** (e.g., "May 2025", "April 2025" - sorted newest first after the Summary sheet) will contain the platform spending summary, pie chart, daily spending line graph, and detailed transactions for that month.
    * **"Overall Summary" Sheet:** (Second sheet tab) provides a consolidated view of your expenses across the selected period.
    * The script will attempt to activate the current calendar month's sheet (if generated) upon completion.

## Troubleshooting

* **"No transactions found":**
    * Double-check your `keywords`, `sender`, and especially `extractionPatterns` in `config.json.html` for the relevant platforms. Ensure they exactly match the content of your emails.
    * Verify the date range in "Master Config" is correct and contains emails you expect to be processed.
    * Manually search in Gmail using similar criteria (sender, keywords, dates) to see if emails are actually present.
* **Incorrect amounts extracted:**
    * The `extractionPatterns` (regex) for that platform in `config.json.html` likely need adjustment. Remember to escape backslashes (`\\d`, `\\s`). Test your regex patterns with an online regex tester using sample text from your email bodies.
* **Errors during script execution:**
    * Open the Apps Script editor (`Extensions > Apps Script`).
    * Click on "Executions" in the left sidebar.
    * Find the failed execution and review the logs and error messages. This often provides clues.
* **`TypeError: ss.newChart is not a function` or similar:** Ensure you have copied the latest `Code.gs` correctly. This error usually means `.newChart()` is being called on the wrong object.
* **`config.json.html` parsing errors:** Ensure your `config.json.html` is valid JSON. Use an online JSON validator if you suspect syntax errors.
* **Script running too long / timeouts:**
    * Google Apps Script has execution limits (approx. 6 minutes for consumer accounts).
    * Use the "Generate/Update Report (Incremental)" option for regular updates.
    * For initial setup or very large date ranges, process in smaller date chunks if necessary.
    * The "Max Expected Run Time (minutes)" in Master Config is informational and does not stop the script.

## License

*(Choose a license like MIT, Apache 2.0, or specify your terms here. If you don't add one, standard copyright laws apply, meaning others cannot distribute or modify your work without permission. An open-source license like MIT is common for sharing.)*

**Example: MIT License**
