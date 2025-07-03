# VBA PDF Manager Tool

This folder contains a set of VBA modules that create a userform for basic PDF management tasks.  The code is designed for Excel and automates building the form and wiring up the button handlers.

## Files
- `PM_SetupUserForm.bas` – Programmatically creates the `frmPM_PDFManager` userform, its controls and event code.
- `frmPM_PDFManager.frm` – Standalone userform file (optional if you generate the form via the setup macro).
- `PM_PDFOperations.bas` – Procedures for exporting sheets, merging PDFs, splitting, and deleting pages using Adobe Acrobat.
- `clsPM_PDFManager.cls` – Example class encapsulating PDF path handling.

## Usage
1. Open your workbook in the VBA editor.
2. Import the above modules or run `PM_CreatePDFManagerUserForm` from `PM_SetupUserForm.bas` to generate the form automatically.
3. Ensure **Trust access to the VBA project object model** is enabled under Macro Settings.
4. Add references to **Microsoft Visual Basic for Applications Extensibility** and **Adobe Acrobat xx.x Type Library** (for merge/split/delete features).
5. Run the form (`frmPM_PDFManager.Show`) to manage PDFs directly from Excel.

Advanced operations rely on Adobe Acrobat being installed.  If Acrobat is not available, only the sheet export functionality will work.
