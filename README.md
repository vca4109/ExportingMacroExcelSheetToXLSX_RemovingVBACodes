📤 VBA Macro: ExportSheet1ToXLSX_RemoveVBA
🔍 Purpose
This macro exports only Sheet1 from the current Excel workbook to a new .xlsx file — completely stripping out any VBA code or modules in the process. It's ideal for sharing a clean, macro-free version of your sheet.

⚙️ How It Works
The user is prompted to choose a save location via a Save As dialog box.

A new blank workbook is created.

Sheet1 from the current workbook is copied into this new workbook.

Any default sheets (like Sheet1, Sheet2, etc.) that come with the new workbook are automatically deleted — leaving only the copied Sheet1.

The workbook is saved as a .xlsx file, which does not support VBA (ensuring all macros are removed).

A confirmation message box notifies the user of success.

💡 Key Features
Ensures sensitive VBA code is not included when sharing workbooks.

Keeps your shared files lightweight and compliant with macro-free environments.

Automates a common process for Document Controllers, Engineers, Admins, or QA/QC staff.

📄 Output Example
Let’s say your original workbook has:

Macros

Sheet1, Sheet2, Sheet3

After running the macro:

A new .xlsx file is saved with only Sheet1, and no VBA code included.

📌 Use Case
Use this when:

Sending reports to clients who cannot open macro-enabled files.

Submitting documentation to platforms like Aconex that don't allow macros.

Creating clean backup copies of your data without embedded code.

