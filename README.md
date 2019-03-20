# EpXL
Simple yet powerful macro-enabled Microsoft Excel spreadsheet user-interface to the EnergyPlus™ whole-building energy simulation program (https://energyplus.net/). It can import both new .epJSON and legacy .IDF input data files. The input data is tabulated in a very compact form (one row per object; one column per input field) in sheet 'Input', with popup field descriptions from the epJSON schema, thus giving a good overview for editing. EpXL can optionally automate multipe-case simulations (discrete multidimensional parametric/sensitivity analysis, and quasi- or pseudorandom Monte Carlo simulation).

### Functionality
- EpXL is distributed with the input data schema (earlier called Input Data Dictionary, IDD) for EnergyPlus v8.9.0 embedded. Whenever a new version of Energy Plus is relased, you can update EpXL to the latest schema by viewing sheet 'Schema', which automatically imports the latest schema (Energy+.schema.epJSON file), or an earlier version of you choose. Save the EpXL workbook after you have done this.
- Key CTRL+I to import an EnergyPlus data file. Both legacy Input Data Files (.IDF) and new .epJSON files can be imported. Immediately after importing the file into sheet 'Input', EpXL vets the data against the JSON schema. After this you can edit the data just like any spreadsheet by inserting or deleting rows and editing cells. Each row is an object type, and each column is an input field. The object types can appear in any order, but fields must appear in the order specified by the schema.
- Key F1 (Help-key) to show popup-comments describing all the input data fields for the current row (object) in the 'Input' sheet.
- EpXL is distributed with example input data in the 'Input' sheet, which is a variation of the EnergyPlus example buildings. You can safely clear this data (CTRL+A selects the whole sheet, then DELETE key). The import function (CTRL+I) also clears sheet 'Input' before importing the file.
- Key CTRL+Q (for Quality check) any time, to vet the input data in sheet 'Input' against the schema. This checks for missing or superfluous object types, missing required field values, out-of-range values, and case-sensitive text options. Incorrect case is corrected automatically for case-sensitive options. This checking is also conducted automatically when you click on the 'Start simulation' button.
- To run a simulation, click on button "Run simulation" on sheet 'Simulate'. For "single-case" simulation, EpXL shows a list of hyperlinks to important output files. For multiple-case simulations, EpXL generates a list inputs and outputs for each case in sheet 'Output'.
- For multiple-case simulation, you specify both input parameters and output parameters (or formulae) in a table at the bottom of sheet sheet 'Simulate'.
- EpXL is written in Visual Basic for Applications (VBA), which you can be easily modify to automate your own simulation tasks.
- EpXL checks GitHub for updates when you open the workbook (only if the last time it checked was 2+ days ago). An update message is shown in the splash-screen window. 

### Summary of the main macros
- CTRL+I: Import an input data file (.IDF or .epJSON) file to sheet 'Input'.
- CTRL+Q: Quality check of the input data in sheet 'Input'.
- CTRL+E: Export the input data to a .IDF file or .epJSON*. This is normally not necessary, as the 'Start simulation' button does this anyway.
- F1 (Help button): Show pop-up comments describing input data fields on selected row of sheet 'Input'.
- The 'Start simulation' button on sheet 'Simulate' automatically exports the input data file, runs EnergyPlus, reports any errors, and compiles list of  list of output files (with hyperlinks). For example, clicking on the hypetlinks for .csv files opens then directly in Excel.
- Viewing the 'Schema' sheet automatically imports the latest schema "Energy+.schema.epJSON", which and is found in the EnergyPlus root directory on your PC.

(* Outputs epJSON by default. Conditional compilation argument outputJSON is defined in Tools > VBAProjectProperties: =1 to output new .epJSON, =0 for legacy .IDF file)

### Installation and activation
- Simply download the spreadsheet file and open it in Microsoft Excel. No installation or registration is needed.
- However, this is a Visual Basic macro-enabled spreadhseet. You must activate macros for it to function: 
  - When you open the file for the first time in Excel, you will see a yellow bar at the top of the window, with the message *"PROTECTED VIEW Be careful... [Enable Editing]"*. Click on the 'Enable Editing' button. 
  - Next, depending on the security settings on your installation of Microsoft Excel, a red bar may appear at the top of the window, with the message *"BLOCKED CONTENT Macros in this document have been disabled..."*. This can be solved by moving the file to a directory on your PC that you designate for files that you trust, then open the file in Excel. To designate a 'trusted directory', click on **File > Options > Trust Center > Trust Center Settings > Trusted Locations > Add new location**, then browse to a directory, e.g. C:\TEMP\. Finally check that **Trust Center Settings > Trusted Documents > Disable Trusted Documents**  is not ticked.
  - When macros are properly activated, **EpXL** shows a small square splash-screen when you open the file. This splash screen shows the licence info, and advises you if an update is available for download from GitHub. Simply press 'Close' button to close the splash-screen. 
  - If still no splash-screen appears, then you might be able to fix it by menu option **File > Options > Trust Center > Trust Center Settings > Macro Settings > Disable all macros with notification**, which is a suitable level of security.

### License & warranty
- Distributed under the GLP v3 lisence (https://www.gnu.org/licenses/gpl-3.0.en.html). Users of this software shall attribute its use in reports/publications with an appropriate autohor citation and URL to this site.
- Provided without warranty of any kind.

### Author & copyright
© Peter.Schild@OsloMet.no
