Attribute VB_Name = "Module2"
Public Sub SealForDistribution()
    Dim pwd As String
    Dim ws As Object                ' µp??e? ?a e??a? Worksheet ? Chart
    Dim openMe As Worksheet
    Dim failed As String
    Dim visCount As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' --- ???e?d?p???s? ??a Shared Workbook (de? ?p?st????e? VeryHidden) ---
    If ThisWorkbook.MultiUserEditing Then
        MsgBox "?? workbook e??a? se Shared mode. ?pe?e???p???s? t? (Review > Share Workbook (Legacy) > Untick).", vbCritical
        GoTo Cleanup
    End If

    ' --- ??e? / f??e µp??st? t? OPEN_ME ---
    On Error Resume Next
    Set openMe = ThisWorkbook.Worksheets("OPEN_ME")
    On Error GoTo Fail
    If openMe Is Nothing Then
        MsgBox "Sheet 'OPEN_ME' not found. Aborting.", vbCritical
        GoTo Cleanup
    End If

    ' --- ?e??e?d?se t? d?µ? a? ??e???eta? ---
    If ThisWorkbook.ProtectStructure Then
        pwd = InputBox("Workbook structure is protected." & vbCrLf & _
                       "Enter password to seal for distribution:", "Seal workbook")
        If Len(pwd) = 0 Then GoTo Cleanup
        ThisWorkbook.Unprotect Password:=pwd
        If ThisWorkbook.ProtectStructure Then
            MsgBox "Wrong password. Cannot unprotect structure. Aborting.", vbCritical
            GoTo Cleanup
        End If
    End If

    ' --- ???e t? OPEN_ME ??at? ?a? e?e??? p??? ????e?? ?t?d?p?te ---
    If openMe.Visible <> xlSheetVisible Then openMe.Visible = xlSheetVisible
    openMe.Activate

    ' --- 1? p??asµa: ????e ??a ta ???a (µ??? Worksheets) se Hidden ---
    On Error Resume Next
    For Each ws In ThisWorkbook.Sheets
        If TypeName(ws) = "Worksheet" Then
            If ws.name <> openMe.name Then
                ws.Visible = xlSheetHidden
                If Err.Number <> 0 Then
                    failed = failed & vbCrLf & "- " & ws.name & " (Hidden)"
                    Err.Clear
                End If
            End If
        End If
    Next ws
    On Error GoTo Fail

    ' ?sf??e?a: ﬂeﬂa??s?? ?t? ?p???e? t??????st?? 1 ??at? (OPEN_ME)
    visCount = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then visCount = visCount + 1
    Next ws
    If visCount = 0 Then openMe.Visible = xlSheetVisible

    ' --- 2? p??asµa: ﬂ??e VeryHidden sta e?a?s??ta ---
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> openMe.name Then
            ws.Visible = xlSheetVeryHidden
            If Err.Number <> 0 Then
                failed = failed & vbCrLf & "- " & ws.name & " (VeryHidden)"
                Err.Clear
            End If
        End If
    Next ws
    On Error GoTo Fail

    ' --- ???stas?a d?µ?? ?a? ap????e?s? ---
    If Len(pwd) = 0 Then pwd = "REPLACE_WITH_SHEET_PWD"
    ThisWorkbook.Protect Password:=pwd, Structure:=True, Windows:=False

    ThisWorkbook.Save

    If Len(failed) = 0 Then
        MsgBox "Sealed OK. After reopening (macros off) ?a fa??eta? µ??? t? OPEN_ME.", vbInformation
    Else
        MsgBox "Sealed µe p??e?d?p???se??. ??p??a f???a de? ???ft??a?:" & failed, vbExclamation
    End If

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Unexpected error while sealing:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
End Sub


