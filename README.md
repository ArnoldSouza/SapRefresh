### TODO List:
---------------------------
- **Handling multiple connections:**
    - The algorithm need improvements to handle more than one object (connections or crosstabs)
- **Handling Function Returns:**
    - The Analysis Functions can return different types of objects. Sometimes it returns a `single value`, a list (`one-dimension array`) or an array (`two dimensions array`). It is necessary to build a method to treat those returns. 
- **Check connection:**
    - Implement attempt method: to refresh in case of problems until certain number of tries
    - Create a Python Decorator Method to check if connection is still alive: 
        - Example: `Application.Run("SAPGetProperty", "IsDataSourceActive", "DS_1")`
- **Handling exceptions:**
    - If SAP AfO Crash (for any reason), force quit Excel and start attempt program
    - Sometimes the VBA API cannot be callable. I believe it is related to Excel Instances running out of control. Those instances must be killed in the task manager. Need to find WHEN it happens and implement a method to automatically detect and kill it. 
- **Developing the Command Line Interface (CLI):**
    - Refine the spinning cursor
- **Debugging the application:**
    - Create a log system to the application. The LOG must have:
        - Report ID | Start Time | Process ID | End Time | Result
    - Design a method to apply filters and variables. During this step, remember to do:
        - SAPSetRefreshBehaviour to OFF
        - PauseVariableSubmit to ON
    - After Logon - Initial Refresh (for SAPSetFilter).



### Done List: 
-------------------------------
- **Authentication Methods**:
    - Using Obfuscation Methods
        - Using Encryption + Config Files + Environmental Variables
- Implement a way to check if the execution of `SAP Logon` and `SAP RefreshData` was successful 
- Get more variables from the VBA API: `sheet, Crosstab, DataSourceName, Query, System, Variables, Filters, Dimensions`
- Implement a `CONFIG.INI` file to handle the variables that shouldn't be present in the source code
- Create a timeit decorator to get the execution time of methods
- Doubts 
    - What's the difference between `Refresh` and `RefreshData` VBA command 
        - Apparently there is no difference
    - Understand what callback means in SAP VBA API
        - callbacks are routines that are executed with certain events. In this project it is used when Variables (Filters) are Set.
        - The project uses the callback `BeforeFirstPromptsDisplay` to better handle dynamic values to variables.
        - More info [here](https://help.sap.com/viewer/ca9c58444d64420d99d6c136a3207632/2.6.1.0/en-US/f26f56246c9b1014bf2c9a7eb0e91070.html)
- Check connection:
    - Ping to server to check if the application is running inside Hydro
- Developing the Command Line Interface (CLI):
    - Create a Waiting Cursor while waiting for time-consuming methods
- **Debugging the application:**
    -If SAPSetFilter is used then we need "Initial Refresh" before ApplyScopeFilters
    - `Application.Run("SAPExecuteCommand", "Refresh", "DS_1")`
        - Use this command to initially refresh the data in the workbook.
