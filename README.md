# SAP Refresh

A python package to automate the extraction of data using SAP Analysis fo Office plugin inside Excel. It opens an Excel workbook and use the API provided by SAP [(link)](https://help.sap.com/viewer/ca9c58444d64420d99d6c136a3207632/2.6.1.0/en-US/f270fd456c9b1014bf2c9a7eb0e91070.html) to control the workflow of the data extraction. 

This project was possible due to the work of:
- Regan Macdonald
    - [Automated updating of BW data in Excel files (BEX & AO) via VBA/VBscript](https://blogs.sap.com/2016/12/18/automated-updating-of-data-in-excel-files-bex-ao-via-vbavbscript/)
    - [Analysis for Office Variables and Filters via VBA](https://blogs.sap.com/2017/02/03/analysis-for-office-variables-and-filters-via-vba/)
- Ivan Bondarenko
    - [SAP BO Analysis for Office (BOAO) Automation](https://github.com/IvanBond/SAP-BOA-Automation)

All the solutions of them were implemented using VBA and/or VB Script. This project aims to implement the same idea using Python Programming. This language was chosen because it has:
1) A library that can make COM interface with windows applications  (such as Excel). That makes easy to convert VBA script to Python.
2) Python has Pandas, witch makes the data transformation workflow a very easy task to be done.
3) Using Pandas you can capture the data refreshed from the Excel workbook and use Pandas Method named to_sql(), to create automatically the SQL insert instruction to ingest data into SQL server.
