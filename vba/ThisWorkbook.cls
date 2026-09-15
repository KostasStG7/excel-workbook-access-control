VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Αυτό_το_βιβλίο_εργασίας"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    On Error GoTo Fail

    ' ========= PROFILE SETUP (TEMP) =========
    Dim t0 As Double, t As Double, prevCalcMode As XlCalculation
    t0 = Timer: t = t0

    ' --- PERFORMANCE: temporarily disable UI updates, events and auto-calc ---
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    prevCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    Debug.Print "OPEN: start @"; Format$(t0, "0.000")

    ' 1) Infrastructure
    UnprotectStructureIfNeeded
    EnsureCfg
    EnsureOpenMeSheetExists
    EnsureOwnerTools
    SafeHideSheet "OWNER_TOOLS"          ' keep admin sheet hidden by default
    Debug.Print "OPEN: infra done +"; Format$(Timer - t, "0.000s"): t = Timer

    ' === START SEALED: show only OPEN_ME before any prompts ===
HideAllSensitiveForRest
SafeHideSheet "OWNER_TOOLS"   ' extra safety (admin sheet always VeryHidden)
SafeHideSheet "DRM_LOG"
SafeHideSheet "DRM_CFG"
DoEvents


    ' 2) Must be saved to disk
    If Len(ThisWorkbook.path) = 0 Then
        MsgBox "Please save the workbook to a folder first.", vbInformation
        HideAllSensitiveForRest

        ' --- restore environment before leaving ---
        On Error Resume Next
        Application.Calculation = prevCalcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        On Error GoTo 0

        Debug.Print "OPEN: exit (not saved) T="; Format$(Timer - t0, "0.000s")
        Exit Sub
    End If
    Debug.Print "OPEN: path check +"; Format$(Timer - t, "0.000s"): t = Timer

    ' 3) Expiry
    If IsExpired() Then
        MsgBox "This workbook has expired. Please contact the owner.", vbCritical

        ' --- restore environment before leaving ---
        On Error Resume Next
        Application.Calculation = prevCalcMode
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        On Error GoTo 0

        Debug.Print "OPEN: exit (expired) T="; Format$(Timer - t0, "0.000s")
        ThisWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    Debug.Print "OPEN: expiry check +"; Format$(Timer - t, "0.000s"): t = Timer

    ' 4) Admin choice
    Dim adminChoice As VbMsgBoxResult

    ' --- ensure UI is paintable & focused before the prompt ---
    Dim prevSU As Boolean: prevSU = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = True
    AppActivate Application.Caption    ' focus Excel window (no error if already active)
    DoEvents                           ' let Windows repaint the dialog
    On Error GoTo Fail

    adminChoice = MsgBox("Do you want to open the workbook as ADMIN?" & vbCrLf & _
                     "(If NO, normal user flow will continue.)", _
                     vbYesNo + vbQuestion, "Open as Admin?")

    ' restore prior ScreenUpdating state
    On Error Resume Next
    Application.ScreenUpdating = prevSU
    On Error GoTo 0

    Debug.Print "OPEN: admin prompt +"; Format$(Timer - t, "0.000s"): t = Timer


    If adminChoice = vbYes Then
        If IsAdminEnrolledForThisPC() Then
            EnsureOwnerTools
            SafeShowSheet "OWNER_TOOLS"
            ThisWorkbook.Worksheets("OWNER_TOOLS").Activate
            RevealForSession
            UnlockForEditing
            OwnerTools_AdminFullyEditable
            UnhideFormulasForAdmin
            ' Hotkeys
            With Application
            .OnKey "^s"        ' Ctrl+S default
            .OnKey "+^{S}"     ' Shift+Ctrl+S default
            .OnKey "^c"        ' Ctrl+C default
            .OnKey "^v"        ' Ctrl+V default
            .OnKey "^x"        ' Ctrl+X default
            End With
            
            If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:="REPLACE_WITH_SHEET_PWD"


            ' --- restore environment before leaving ---
            On Error Resume Next
            Application.Calculation = prevCalcMode
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            On Error GoTo 0

            Debug.Print "OPEN: admin-enrolled branch T="; Format$(Timer - t0, "0.000s")

            Exit Sub
        End If

        Dim pwd As String
        pwd = Trim$(InputBox("Enter ADMIN password:", "Admin Login"))
        Debug.Print "OPEN: admin password prompt +"; Format$(Timer - t, "0.000s"): t = Timer

        If Len(pwd) = 0 Then GoTo NormalFlow

        If AdminPasswordIsValid(pwd) Then
            EnrollAdminForThisPC

            EnsureOwnerTools
            SafeShowSheet "OWNER_TOOLS"
            ThisWorkbook.Worksheets("OWNER_TOOLS").Activate
            RevealForSession
            UnlockForEditing
            OwnerTools_AdminFullyEditable
            UnhideFormulasForAdmin
            ' Hotkeys
            With Application
            .OnKey "^s"        ' Ctrl+S default
            .OnKey "+^{S}"     ' Shift+Ctrl+S default
            .OnKey "^c"        ' Ctrl+C default
            .OnKey "^v"        ' Ctrl+V default
            .OnKey "^x"        ' Ctrl+X default
            End With

            If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:="REPLACE_WITH_SHEET_PWD"


            ' --- restore environment before leaving ---
            On Error Resume Next
            Application.Calculation = prevCalcMode
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            On Error GoTo 0

            Debug.Print "OPEN: admin-enroll branch T="; Format$(Timer - t0, "0.000s")

            Exit Sub
        Else
            MsgBox "Wrong admin password. Proceeding with normal flow.", vbExclamation
        End If
        Debug.Print "OPEN: wrong admin password +"; Format$(Timer - t, "0.000s"): t = Timer
    End If

NormalFlow:
    ' Enrollment wizard if not authorized
    If Not IsAuthorizedMachine() Then ' Or Not IsAllowedPath(ThisWorkbook.path) Then
        Debug.Print "OPEN: entering ShowEnrollWizard..."
        If Not ShowEnrollWizard() Then

            ' --- restore environment before leaving ---
            On Error Resume Next
            Application.Calculation = prevCalcMode
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            On Error GoTo 0

            Debug.Print "OPEN: exit (wizard cancelled) T="; Format$(Timer - t0, "0.000s")
            ThisWorkbook.Close SaveChanges:=False
            Exit Sub
        End If
        Debug.Print "OPEN: ShowEnrollWizard +"; Format$(Timer - t, "0.000s"): t = Timer
    End If

    ' Hotkeys for normal session
    Application.OnKey "^s", "DoSave"
    Application.OnKey "^c", "DisableCopy"
   ' Application.OnKey "^v", "DisableCopy"
    Application.OnKey "^x", "DisableCopy"
    Debug.Print "OPEN: hotkeys set +"; Format$(Timer - t, "0.000s"): t = Timer

    ' Reveal/Seal per authorization
    If IsAuthorizedMachine() Then 'And IsAllowedPath(ThisWorkbook.path) Then
        RevealForSession
        UnlockForEditing
        SafeHideSheet "OWNER_TOOLS"
        SafeHideSheet "DRM_LOG"
        SafeHideSheet "DRM_CFG"
        HideOnlyFormulasForUser
        Debug.Print "OPEN: reveal-for-session +"; Format$(Timer - t, "0.000s"): t = Timer
    Else
        HideAllSensitiveForRest
        Debug.Print "OPEN: sealed (unauthorized) +"; Format$(Timer - t, "0.000s"): t = Timer
    End If

    ' --- restore environment before leaving normally ---
    On Error Resume Next
    Application.Calculation = prevCalcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    Debug.Print "OPEN: done T="; Format$(Timer - t0, "0.000s")
    Exit Sub

Fail:
    ' Restore environment on error
    On Error Resume Next
    Application.Calculation = prevCalcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Open error: " & Err.Description, vbCritical
    HideAllSensitiveForRest
    Debug.Print "OPEN: FAIL T="; Format$(Timer - t0, "0.000s")
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not IsAdminSession() Then
        HideAllSensitiveForRest
        SafeHideSheet "OWNER_TOOLS"
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   ' If Not IsAdminSession() Then
        HideAllSensitiveForRest
        SafeHideSheet "OWNER_TOOLS"
   ' End If
End Sub

