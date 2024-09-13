Here is a `README.md` file that you can use for your Excel VBA macros project:

---

# Excel VBA Macros for Invoice Management and Automation

## Overview

This Excel project is designed to automate various tasks related to invoice creation, ledger management, and communication via WhatsApp or email. The project consists of multiple VBA macros distributed across 7 modules to perform the following key functions:

1. **Automatic PDF Creation**  
   Creates a PDF invoice with the format `<Bill No & Party Name & ".pdf">` by retrieving data from specific cells in `Sheet1`. The PDF is saved with a dynamically generated file name based on the bill number and party name.

2. **Automated File Sending via WhatsApp**  
   Sends the created PDF to a predefined or dynamic WhatsApp number retrieved from the Excel sheet. This functionality is powered by a Python script that uses Selenium to automate the file sending process. The same logic can be applied to send emails with attachments.

3. **Automatic Bill Printing**  
   Allows the printing of bills, with the option to automatically update the right header with labels such as "Original", "Duplicate", and more as required (e.g., "Triplicate").

4. **Convert Numbers to Words (SpellNumbers)**  
   A module that converts numeric amounts into words, which is particularly useful for displaying amounts on invoices in written form.

5. **Automatic Bill Number Generation**  
   Automatically determines the next available bill number when the workbook is opened. It checks the directory where all bills are saved, extracts the bill numbers using RegEx, identifies the highest value, adds 1 to it, and sets the new bill number in the appropriate cell (`Sheet1.G2`).

6. **Ledger Management**  
   Adds invoice details to a ledger in two ways:
   - A common ledger file (`ledger.xlsx`) for all parties.
   - A dedicated file for each party. If a file for the party does not exist, it will be created. This feature uses an if-else statement to check whether the file exists.
   - A message box (`MsgBox`) confirms the successful addition of the data.

7. **Error Handling Using VLOOKUP with IFERROR**  
   Uses the formula `=IFERROR(VLOOKUP(C7,Sheet2!A:I,2,0),"")` to avoid displaying errors when data (like Address Line 3 or GST number) is missing for a specific party. Instead, it returns a blank cell.

8. **Additional Functionalities**  
   - **Automatic Date:** The function `=TODAY()` is used for auto-populating the date.
   - **Automatic GST Calculation:** Automatically calculates GST based on the invoice amount.
   - **Automatic Round-Off:** Rounds off the total amount as per standard accounting rules.

## Installation

1. Open the Excel file that contains the macros.
2. Make sure that macros are enabled in your Excel environment.
3. Ensure that the Python script (for WhatsApp and email automation) is set up correctly, and Selenium is installed and configured.
4. All necessary Excel modules should already be in the workbook, ready for execution.

## How to Use

### 1. **Create a PDF Invoice**
   - Ensure that `Sheet1` contains the required data (e.g., Bill Number, Party Name, etc.).
   - Run the macro to generate the PDF invoice, which will be saved in the specified format.

### 2. **Send the PDF via WhatsApp**
   - Input the phone number in the designated cell in `Sheet1`, or use a predefined number.
   - The macro will trigger a Python script via a Shell command, which automates sending the file to WhatsApp using Selenium.

### 3. **Print Bills**
   - Run the macro to print the invoice.
   - The header will automatically change to "Original", "Duplicate", or other specified values for multiple copies.

### 4. **Add Bill to Ledger**
   - Run the macro to append the invoice details to a common ledger file (`ledger.xlsx`) or create a dedicated file for each party.
   - A message box will confirm when the data is successfully added to the ledger.

### 5. **Automatic Bill Numbering**
   - On opening the workbook, the macro will scan the directory for existing bills, identify the highest bill number, and increment it to set the next bill number.

### 6. **Error Handling with IFERROR**
   - VLOOKUP is used to retrieve party-specific data like address, GST number, etc., from `Sheet2`. If any data is missing, the IFERROR function ensures that no error message is displayed; instead, a blank is shown.

## Python Script Setup (for WhatsApp and Email Automation)

1. Install Python and ensure that Selenium is installed:
   ```
   pip install selenium
   ```
2. Download the appropriate WebDriver for your browser (e.g., ChromeDriver for Chrome).
3. The Python script should be set up to receive the bill file name and phone number via command-line arguments.
4. Make sure the script is running when the VBA macro calls it via a Shell command to send the file.

## Important Notes

- Ensure that the directory where bills are saved is correctly specified in the macros.
- The Python script for sending files should be tested separately to confirm that it works before integrating with the Excel macros.
- Customize the paths and dynamic ranges (`ThisWorkbook.Sheets("Sheet1").Range(" ")`) in the macros as per your specific file structure.

## Troubleshooting

- **Macros Disabled:** Ensure that macros are enabled in Excel.
- **Python Script Fails:** Check for errors in the Selenium script and ensure the correct WebDriver is installed.
- **Bill Number Generation Fails:** Ensure the directory where bills are stored is accessible and the file names follow the correct pattern.

---

This `README.md` will guide users through the functionalities and setup of your Excel VBA macros project. Feel free to modify it to suit additional features or changes in your project.
