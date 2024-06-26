# VBA-Automation-Scripts
Collection of VBA scripts for Automating Tasks, Designed to Boost Productivity!
# 1. Set Up SAP GUI Scripting
Firstly, ensure that SAP GUI Scripting is enabled on your SAP system and client machine. This involves:
- Enabling scripting in SAP GUI via transaction `RZ11`.
- Configuring SAP GUI settings to allow scripting (usually found under `Customize Local Layout` -> `Options` -> `Accessibility & Scripting`).
# 2. Develop the SAP GUI Script
Write a SAP GUI Script that navigates through the SAP transaction `ZS86` (assuming it's a custom transaction for sales order extraction) and extracts the data.
```vbscript
' SAP GUI Script
Set SapGuiAuto = GetObject("SAPGUI")
Set SAPApp = SapGuiAuto.GetScriptingEngine
Set SAPCon = SAPApp.Children(0)
Set session = SAPCon.Children(0)

' Open transaction ZS86
session.StartTransaction "ZS86"

' Navigate to required fields and execute extraction
session.findById("wnd[0]/usr/ctxtSOME_FIELD").Text = "Value"
session.findById("wnd[0]/tbar[1]/btn[8]").Press ' Execute button

' Get the extracted data
Dim data
data = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(1, "FIELD_NAME")

' Close SAP session
session.findById("wnd[0]").Close
```

Replace `"ZS86"`, `"SOME_FIELD"`, and `"FIELD_NAME"` with actual values from your SAP environment.

# 3. Develop VBA Script for Automation
Create a VBA script in Excel (assuming you want to automate from Excel) to trigger the SAP GUI Script and process the extracted data:

```vba
Sub RunSAPScript()
    Dim sapgui As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object

    ' Connect to SAP GUI
    Set sapgui = GetObject("SAPGUI")
    Set SAPApp = sapgui.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)

    ' Execute SAP GUI Script
    session.StartTransaction "ZS86"

    ' Code to fill in and execute SAP GUI actions

    ' Close SAP session
    session.findById("wnd[0]").Close
End Sub
```

# 4. Publishing to GitHub
To publish your scripts on GitHub:
- Create a GitHub repository for your project.
- Organize your scripts into appropriate folders (`SAP_GUI_Scripts` and `VBA_Scripts`, for example).
- Write a `README.md` file explaining the purpose, setup instructions, and usage of your scripts.
- Ensure any sensitive information (like SAP credentials) is not hardcoded into the scripts.

# 5. Documentation and Instructions
Include detailed instructions in your `README.md` file on how to:
- Enable SAP GUI Scripting.
- Install and use the VBA script from Excel.
- Customize the scripts for different environments or scenarios.

### Additional Tips
- Test your scripts thoroughly in a development or sandbox SAP environment before using them in production.
- Consider security implications and access controls when sharing or using scripts that interact with sensitive data.
