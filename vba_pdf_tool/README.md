# VBA PDFX Manager Tool

This folder contains a set of VBA modules that create a userform for advanced PDF management tasks in Excel. The code programmatically builds the form and wires up button handlers using unique naming conventions.

## Files
- `PXFormBuilder.bas` – Builds the `frmPXManager` userform and injects its event code.
- `frmPXManager.frm` – Optional pre-built userform if you prefer to import it directly.
- `PXPDFOperations.bas` – Procedures for exporting sheets, merging PDFs, splitting PDFs and deleting pages using Adobe Acrobat.
- `clsPXPDFManager.cls` – Class that encapsulates PDF file handling with object-oriented methods.

## Usage
1. Open your workbook in the VBA editor.
2. Import the above modules or run `PX_BuildManagerForm` from `PXFormBuilder.bas` to generate the form automatically.
3. Ensure **Trust access to the VBA project object model** is enabled under Macro Settings.
4. Add references to **Microsoft Visual Basic for Applications Extensibility** and **Adobe Acrobat xx.x Type Library** (for merge/split/delete features).
5. Show the form with `frmPXManager.Show` to manage PDFs directly from Excel.

Advanced operations rely on Adobe Acrobat being installed. If Acrobat is not available, only the sheet export functionality will work.
