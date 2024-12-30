# Excel Add-In: Export Data to Word

This repository contains a VBA-based Excel add-in that automates the process of exporting data from Excel to Word documents. The add-in uses markers in Word templates to insert data at specified locations.

<!--toc:start-->

- [Excel Add-In: Export Data to Word](#excel-add-in-export-data-to-word)
  - [Directory Structure](#directory-structure)
  - [How to Use the Add-In](#how-to-use-the-add-in)
  - [Testing the Add-In](#testing-the-add-in)
  - [Notes](#notes)
  - [License](#license)
  <!--toc:end-->

## Directory Structure

- **`src/`**: Contains the source code files (`Config`, `Main`, `Util`) used for the add-in.
- **`dist/`**: Contains the compiled Excel add-in file `auto-export.xlam`.
- **`test/`**: Contains test files to show the functionality:
  - `test.xlsx`: An Excel workbook that uses the add-in.
  - `test.docx`: A Word document with markers to receive the exported data.

## How to Use the Add-In

1. **Download the Add-In**
   Download the `auto-export.xlam` file from the `dist/` directory.

2. **Add the Add-In to Excel**

   - Open Excel.
   - Go to `File > Options > Add-Ins`.
   - In the "Manage" dropdown at the bottom, select `Excel Add-ins` and click `Go`.
   - Click `Browse`, locate the downloaded `auto-export.xlam` file, and click `OK`.

3. **Run the Add-In**

   - Open the VBA editor in Excel (`Alt + F11`).
   - In the VBA editor, locate the `auto-export.xlam` project under `Project Explorer`.
   - Expand `Modules` and select `Main`.
   - With the `Main` module open, click the green "Run" triangle in the editor toolbar.
   - When prompted, confirm to run the `ExportData` macro.

4. **Follow the Instructions**
   - The add-in will prompt you to confirm closing all Word documents. Save your work and accept to proceed.
   - The macro will export the data from the specified Excel ranges to the Word document(s) based on the configuration.

## Testing the Add-In

Use the files in the `test/` directory to verify functionality:

1. **Open `test.xlsx`**:
   This file includes sample data to export.

2. **Run the Add-In**:
   Follow the steps above to run the `ExportData` macro.

3. **Check the Results**:
   The data from `test.xlsx` is copied into `test.docx` at the specified markers.

## Notes

- Ensure all Word documents are closed before running the add-in.
- Modify the source code in the `dist/auto-export.xlam` directory if customization is needed.

## License

This project is open-source and available under the MIT License.
