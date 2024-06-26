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

# README.md
Document explaining the purpose and basic instructions for using the scripts.

### Additional Tips
- Test your scripts thoroughly in a development or sandbox SAP environment before using them in production.
- Consider security implications and access controls when sharing or using scripts that interact with sensitive data.
