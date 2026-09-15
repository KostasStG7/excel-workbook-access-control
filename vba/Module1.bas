Attribute VB_Name = "Module1"
Sub InitDRMStructure()
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    Dim pwd As String: pwd = "REPLACE_WITH_SHEET_PWD" ' default sheet protection pwd from project

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- OPEN_ME ----
    If Not SheetExists(wb, "OPEN_ME") Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        sh.name = "OPEN_ME"
        sh.Range("A1").Value = "Re-open this file with macros enabled." & vbCrLf & _
                              "If Windows blocked the file: Right-click the .xlsm -> Properties -> check 'Unblock'."
        sh.Columns(1).ColumnWidth = 80
        sh.Rows(1).RowHeight = 60
        sh.Range("A1").WrapText = True
        sh.Protect Password:=pwd, UserInterfaceOnly:=True
    End If

    ' ---- DRM_CFG ----
    If Not SheetExists(wb, "DRM_CFG") Then
        Set sh = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        sh.name = "DRM_CFG"
        sh.Range("A1").Value = "Machines"
        sh.Range("B1").Value = "Paths"
        sh.Range("C1").Value = "Secret"
        sh.Range("D1").Value = "LastChallengeDate"
        sh.Range("E1").Value = "LastChallengeSerialUTC"
        sh.Range("F1").Value = "AdminMachines"
        sh.Columns("A:F").AutoFit
        sh.Protect Password:=pwd, UserInterfaceOnly:=True
        sh.Visible = xlSheetVeryHidden
    End If

    ' ---- OWNER_TOOLS ----
    If Not SheetExists(wb, "OWNER_TOOLS") Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(1))
        sh.name = "OWNER_TOOLS"
        With sh
            .Range("A1").Value = "Paste Challenge (from user):"
            .Range("B1").Value = "Response (auto):"
            .Range("A2").Value = ""
            .Range("B2").Value = ""
            .Columns("A:B").AutoFit
            .Range("A4").Value = "Remove Machine/Path:"
            .Hyperlinks.Add Anchor:=.Range("B4"), Address:="", SubAddress:="'OWNER_TOOLS'!A4", TextToDisplay:="Run Remove Tool"
            .Range("B4").Font.Color = vbBlue
            .Range("B4").Font.Underline = xlUnderlineStyleSingle
            .Cells.Locked = False
            .Protect Password:=pwd, UserInterfaceOnly:=True
            .Visible = xlSheetVeryHidden
        End With
    End If

    ' ---- DRM_LOG ----
    If Not SheetExists(wb, "DRM_LOG") Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        sh.name = "DRM_LOG"
        With sh
            .Range("A1:G1").Value = Array("UTCDateTime", "Machine", "Path", "Challenge", "Response", "AdminUser", "Signature")
            .Columns("A:G").AutoFit
            .Protect Password:=pwd, UserInterfaceOnly:=True
            .Visible = xlSheetVeryHidden
        End With
    End If

    ' ---- Final: show OPEN_ME and hide others (sealed) ----
    wb.Worksheets("OPEN_ME").Visible = xlSheetVisible
    wb.Worksheets("OPEN_ME").Activate
    If SheetExists(wb, "DRM_CFG") Then wb.Worksheets("DRM_CFG").Visible = xlSheetVeryHidden
    If SheetExists(wb, "OWNER_TOOLS") Then wb.Worksheets("OWNER_TOOLS").Visible = xlSheetVeryHidden
    If SheetExists(wb, "DRM_LOG") Then wb.Worksheets("DRM_LOG").Visible = xlSheetVeryHidden

    MsgBox "DRM structure initialized (OPEN_ME, DRM_CFG, OWNER_TOOLS, DRM_LOG).", vbInformation

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error initializing DRM structure: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub

Function SheetExists(wb As Workbook, name As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = wb.Worksheets(name)
    SheetExists = (Not s Is Nothing)
    On Error GoTo 0
End Function


