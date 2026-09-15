Attribute VB_Name = "DRM_Core"
Option Private Module
Option Explicit

' =====================================================
' DRM_Core - Access control, enrollment, visibility
' NOTE: Only comments were translated; code unchanged.
'       Functions were grouped for clarity.
' =====================================================

' -------------------------
' CONFIGURATION CONSTANTS
' -------------------------
Private Const SHEET_PWD As String = "REPLACE_WITH_SHEET_PWD"                 ' Sheet/workbook protection password
Private Const ADMIN_PASSWORD As String = "REPLACE_WITH_ADMIN_PASSWORD"            ' Admin password
Private Const EXPIRY_DATE As String = "2026-12-31"         ' Expiry date "YYYY-MM-DD"; blank = none
Private Const RESP_LEN As Long = 24     ' 24 hex chars (3?8 μπλοκ FNV-1a32)
Private Const RESP_ROUNDS As Long = 2048
Private Const CH_TTL_MINUTES As Long = 10                  ' Challenge Time-To-Live (minutes)
Private Const WEBHOOK_URL As String = "http://127.0.0.1:5000/v1/challenge"
'"http://127.0.0.1:5000/v1/challenge"
Private Const SHARED_SECRET As String = "REPLACE_WITH_SHARED_SECRET" ' == DRM_SHARED_SECRET


' -------------------------
' WINDOWS API FOR TIMEZONE
' -------------------------
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#Else
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
#End If

' -------------------------
' ADMIN PASSWORD CHECK
' -------------------------
Public Function AdminPasswordIsValid(ByVal candidate As String) As Boolean
    ' Centralized, safe, case-insensitive check (no secret leakage)
    AdminPasswordIsValid = (StrComp(Trim$(CStr(candidate)), Trim$(ADMIN_PASSWORD), vbTextCompare) = 0)
End Function

' -------------------------
' BASIC UTILITIES
' -------------------------
Public Function DeviceName() As String
    DeviceName = UCase$(Environ$("USERNAME"))
End Function

Public Function CurrentPath() As String
    CurrentPath = ThisWorkbook.path
End Function

Private Function Nz(ByVal v, Optional ByVal dflt As String = "") As String
    If IsError(v) Or IsNull(v) Or Len(Trim$(v & "")) = 0 Then
        Nz = dflt
    Else
        Nz = CStr(v)
    End If
End Function

Private Function NormPath(ByVal p As String) As String
    Dim s As String: s = LCase$(Trim$(Replace(p, "/", "\")))
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    NormPath = s
End Function

Private Function TimeZoneInformationBiasMinutes() As Long
    Dim Bias As Long, tzi As TIME_ZONE_INFORMATION, rc As Long
    rc = GetTimeZoneInformation(tzi)
    If rc = 2 Then
        Bias = tzi.Bias + tzi.DaylightBias
    Else
        Bias = tzi.Bias + tzi.StandardBias
    End If
    TimeZoneInformationBiasMinutes = Bias
End Function

Private Function UtcNowSerial() As Date
    ' SOS??: UTC = Local + BiasMinutes/1440
    UtcNowSerial = Now + (TimeZoneInformationBiasMinutes() / 1440#)
End Function


Private Function RandomHex(ByVal nBytes As Long) As String
    Dim i As Long, s As String
    Randomize
    For i = 1 To nBytes
        s = s & Right$("00" & Hex$(Int(Rnd() * 256)), 2)
    Next
    RandomHex = UCase$(s)
End Function

' -------------------------
' FNV-1a 32-bit HASH
' -------------------------
Private Function FNV1a32Hex(ByVal s As String) As String
    Dim h As Long, i As Long, b As Long
    h = &H811C9DC5
    For i = 1 To Len(s)
        b = Asc(mid$(s, i, 1)) And &HFF
        h = h Xor b
        h = Mul32D(h, 16777619)
    Next
    FNV1a32Hex = UCase$(Right$("00000000" & Hex$(h), 8))
End Function

Public Function Mul32D(ByVal a As Long, ByVal b As Long) As Long
    Dim p As Double
    p = CDbl(a) * CDbl(b)
    p = p - 4294967296# * Fix(p / 4294967296#)
    If p > 2147483647# Then p = p - 4294967296#
    If p < -2147483648# Then p = p + 4294967296#
    Mul32D = CLng(p)
End Function

' -------------------------
' CONFIGURATION SHEET
' -------------------------
Public Sub EnsureCfg()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("DRM_CFG")
    On Error GoTo 0

    If sh Is Nothing Then
        If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD
        Set sh = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        sh.name = "DRM_CFG"
        sh.Range("A1").Value = "Machines"
        sh.Range("B1").Value = "Paths"
        sh.Range("C1").Value = "Secret"
        sh.Range("D1").Value = "LastChallengeDate"
        sh.Range("E1").Value = "LastChallengeSerialUTC"
        sh.Range("F1").Value = "AdminMachines"
    End If

    SafeHideSheet "DRM_CFG"
End Sub

Private Function CfgSheet() As Worksheet
    Call EnsureCfg
    Set CfgSheet = ThisWorkbook.Worksheets("DRM_CFG")
End Function

' -------------------------
' SECRET KEY MANAGEMENT
' -------------------------
Private Function GetSecretKey() As String
    Dim sh As Worksheet: Set sh = CfgSheet()
    Dim key As String: key = Nz(sh.Range("C2").Value, "")
    If Len(key) = 0 Then
        key = RandomHex(16)
        sh.Unprotect Password:=SHEET_PWD
        sh.Range("C2").Value = key
        sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    End If
    GetSecretKey = key
End Function

' -------------------------
' AUTHORIZATION CHECKS
' -------------------------
Public Function IsAuthorizedMachine() As Boolean
    Dim sh As Worksheet, last As Long, i As Long, meName As String
    Set sh = CfgSheet()
    meName = UCase$(Trim$(DeviceName))
    last = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    For i = 2 To last
        If UCase$(Trim$(sh.Cells(i, "A").Value)) = meName Then
            IsAuthorizedMachine = True
            Exit Function
        End If
    Next
End Function

Public Function IsAllowedPath(ByVal cur As String) As Boolean
    Dim sh As Worksheet, last As Long, i As Long, base As String
    Set sh = CfgSheet()
    cur = NormPath(cur)
    If Len(cur) = 0 Then Exit Function
    last = sh.Cells(sh.Rows.count, "B").End(xlUp).Row
    For i = 2 To last
        base = NormPath(Nz(sh.Cells(i, "B").Value, ""))
        If Len(base) > 0 Then
            If Left$(cur, Len(base)) = base Then
                IsAllowedPath = True
                Exit Function
            End If
        End If
    Next
End Function

' -------------------------
' CHALLENGE / RESPONSE LOGIC
' -------------------------
Public Function GenerateChallenge() As String
    On Error GoTo Fail
    Dim userName As String, compName As String, curPath As String
    Dim tsDate As String, tsTime As String, nonce As String
    Dim payload As String, secret As String
    Dim payloadBytes() As Byte, keyBytes() As Byte, encBytes() As Byte
    Dim token As String
    Dim cfg As Worksheet
    Dim nowUTC As Date

    nowUTC = UtcNowSerial    ' <-- ???S? UTC

    userName = UCase$(Environ$("USERNAME"))
    compName = UCase$(Environ$("COMPUTERNAME"))
    curPath = NormPath(CurrentPath())
    tsDate = Format$(nowUTC, "yyyymmdd")   ' <-- ap? UTC
    tsTime = Format$(nowUTC, "hhnnss")     ' n=minutes, s=seconds (ap? UTC)
    Randomize
    nonce = Right$("000000" & Hex$(CLng(Rnd() * 16777215)), 6)

    payload = userName & "|" & compName & "|" & curPath & "|" & tsDate & "|" & tsTime & "|" & nonce

    secret = GetSecretKey()
    If Len(secret) = 0 Then secret = "DEFAULT_SECRET"

    payloadBytes = Utf8BytesFromString(payload)
    keyBytes = Utf8BytesFromString(secret)
    encBytes = XorBytes(payloadBytes, keyBytes)
    token = Base64Encode(encBytes)

    On Error Resume Next
    Set cfg = ThisWorkbook.Worksheets("DRM_CFG")
    On Error GoTo Fail
    If Not cfg Is Nothing Then
        cfg.Range("D2").Value = token
        cfg.Range("H2").Value = payload
        cfg.Range("E2").Value = UtcNowSerial
    End If


'gia http





'telos gia http

    ' --- SEND TO WEBHOOK (simple verify) ---
Call PostChallengeWebhook(token, payload)

        ' --- Send silently to admin via email ---
    'Call SendChallengeEmail(token, payload)

    GenerateChallenge = token
    Exit Function
Fail:
    GenerateChallenge = vbNullString
End Function

' -------------------------
' ENROLLMENT (ADMIN RESPONSE VALIDATION)
' -------------------------


' --- Verify challenge against stored token (cell D2) ---
' --- Compute expected response from payload ---
' --- Compare with user-provided response ---
' === Update DRM_CFG (Machine & Path) ===
' 2) Path (col B) – exact
' --- FORCE LOG WRITE (if DRM_LOG is missing, create it & append) ---
Public Function Admin_ValidateResponseAndEnroll(ByVal challenge As String, ByVal response As String) As Boolean
    On Error GoTo Fail

    Debug.Print "ENTER Admin_ValidateResponseAndEnroll"
    Debug.Print " challenge(in)='" & Trim$(challenge) & "'  response(in)='" & Trim$(response) & "'"

    Dim sh As Worksheet
    Set sh = CfgSheet()   ' DRM_CFG

    Dim payload As String
    payload = Trim$(Nz(sh.Range("H2").Value, ""))   ' USER|COMPUTER|PATH|YYYYMMDD|HHNNSS|NONCE
    If Len(payload) = 0 Then
        Debug.Print "EXIT @PAYLOAD_MISSING(H2)"
        MsgBox "EXIT: Missing payload in DRM_CFG!H2. Generate a new challenge.", vbExclamation
        Admin_ValidateResponseAndEnroll = False
        Exit Function
    End If


    Dim storedToken As String
    storedToken = Trim$(Nz(sh.Range("D2").Value, ""))
    If Len(storedToken) > 0 And Len(Trim$(challenge)) > 0 Then
        If StrComp(storedToken, Trim$(challenge), vbBinaryCompare) <> 0 Then
            Debug.Print "EXIT @TOKEN_MISMATCH  D2='" & storedToken & "'  in='" & Trim$(challenge) & "'"
            MsgBox "EXIT: Challenge token mismatch. Ask the user to resend the latest token.", vbExclamation
            Admin_ValidateResponseAndEnroll = False
            Exit Function
        End If
    End If

    ' --- TTL ---
   
   ' --- TTL ---
Dim chAt As Date
chAt = Nz(sh.Range("E2").Value, 0)   ' UTC timestamp t?? challenge ap? GenerateChallenge

If chAt = 0 Then
    Debug.Print "EXIT @NO_TS(E2)"
    MsgBox "EXIT: Missing challenge timestamp. Generate a new challenge.", vbExclamation
    Admin_ValidateResponseAndEnroll = False
    Exit Function
End If

' ?p?????se p?sa ?ept? ????? pe??se? ap? t?te p?? e?d????e t? challenge
Dim ageMin As Long
ageMin = DateDiff("n", chAt, UtcNowSerial)

If ageMin < 0 Then ageMin = 0  ' p??stas?a ap? clock skew

If ageMin > CH_TTL_MINUTES Then
    Debug.Print "EXIT @TTL_EXPIRED ageMin=" & ageMin
    MsgBox "Challenge expired (" & ageMin & " minutes old). Please request a new challenge.", vbExclamation
    Admin_ValidateResponseAndEnroll = False
    Exit Function
End If

   
   
   ' --- END TTL ---

    Dim expected As String
    expected = Left$(StrongResponseHex(payload), RESP_LEN)

    If StrComp(expected, Trim$(response), vbTextCompare) <> 0 Then
        Debug.Print "EXIT @RESP_MISMATCH expected='" & expected & "'  got='" & Trim$(response) & "'"
        MsgBox "EXIT: Invalid response.", vbCritical
        Admin_ValidateResponseAndEnroll = False
        Exit Function
    End If

    ' DRM_CFG (Machine & Path) ===
    sh.Unprotect Password:=SHEET_PWD

    ' 1) Machine (col A)
    If Not IsAuthorizedMachine() Then
        Dim meName As String, lastA As Long, i As Long, existsA As Boolean
        meName = UCase$(Trim$(DeviceName))
        lastA = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
        For i = 2 To lastA
            If UCase$(Trim$(sh.Cells(i, "A").Value)) = meName Then existsA = True: Exit For
        Next
        If Not existsA Then sh.Cells(lastA + 1, "A").Value = meName
    End If

    ' 2) Path (col B) – exact
    If Not IsAllowedPath(CurrentPath) Then
        Dim cur As String, lastB As Long, base As String, existsB As Boolean
        cur = NormPath(CurrentPath)
        lastB = sh.Cells(sh.Rows.count, "B").End(xlUp).Row
        For i = 2 To lastB
            base = NormPath(Nz(sh.Cells(i, "B").Value, ""))
            If Len(base) > 0 Then If base = cur Then existsB = True: Exit For
        Next
        If Not existsB Then sh.Cells(lastB + 1, "B").Value = cur
    End If

    sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    SafeHideSheet "DRM_CFG"

    
    Admin_ValidateResponseAndEnroll = True
    Debug.Print "ENROLL=TRUE – appending log…"

    ' --- FORCE LOG WRITE  ---
    Dim logSh As Worksheet
    Dim r As Long

    On Error Resume Next
    Set logSh = ThisWorkbook.Worksheets("DRM_LOG")
    On Error GoTo 0

    If logSh Is Nothing Then
        If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD
        Set logSh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        logSh.name = "DRM_LOG"
        With logSh
            .Range("A1:G1").Value = Array("UTCDateTime", "Machine", "Path", "Challenge", "Response", "AdminUser", "Signature")
            .Columns("A:G").AutoFit
        End With
        logSh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    End If

    On Error Resume Next
    AppendEnrollmentLog DeviceName, CurrentPath, challenge, response, Application.userName
    If Err.Number <> 0 Then
        Debug.Print "LOG append ERROR: " & Err.Number & " - " & Err.Description
        Err.Clear
    Else
        Debug.Print "LOG appended OK."
    End If
    On Error GoTo 0

    Debug.Print "EXIT Admin_ValidateResponseAndEnroll(TRUE)"
    Exit Function

Fail:
    Debug.Print "ERROR Admin_ValidateResponseAndEnroll: " & Err.Number & " - " & Err.Description
    MsgBox "Error during enrollment: " & Err.Description, vbCritical
    Admin_ValidateResponseAndEnroll = False
End Function


' -------------------------
' OWNER TOOLS (ADMIN SHEET)
' -------------------------

Public Sub EnsureOwnerTools()
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets("OWNER_TOOLS")
    On Error GoTo 0
    If s Is Nothing Then
        UnprotectStructureIfNeeded
        Set s = ThisWorkbook.Worksheets.Add
        With s
            .name = "OWNER_TOOLS"
            .Range("A1").Value = "Paste Challenge (from user):"
            .Range("B1").Value = "Response (auto):"
            
            ' A2: admin pastes the encrypted token from the user
            ' B2: will be filled by Owner_ProcessEncryptedChallenge
            .Range("A2").Value = ""
            .Range("B2").Value = ""

            .Columns("A:B").AutoFit
            
            
            .Range("A4").Value = "Remove Machine/Path:"
            .Hyperlinks.Add Anchor:=.Range("B4"), _
                Address:="", _
                SubAddress:="'OWNER_TOOLS'!A4", _
                TextToDisplay:="Run Remove Tool"
            .Range("B4").Font.Color = vbBlue
            .Range("B4").Font.Underline = xlUnderlineStyleSingle
            .Range("B4").EntireRow.AutoFit

        End With
        ProtectStructureIfNeeded
    End If
End Sub

Public Sub Owner_ShowTools()
    Dim pwd As String
    pwd = Trim$(InputBox("Enter ADMIN password:", "Admin Login"))
    If Len(pwd) = 0 Then Exit Sub
    If Not AdminPasswordIsValid(pwd) Then
        MsgBox "Invalid admin password.", vbExclamation
        Exit Sub
    End If

    EnsureOwnerTools
    SafeShowSheet "OWNER_TOOLS"
    ThisWorkbook.Worksheets("OWNER_TOOLS").Activate
    RevealForSession
    UnlockForEditing
    
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:="REPLACE_WITH_SHEET_PWD"

    MsgBox "Owner Tools are now visible.", vbInformation
End Sub

Public Sub Owner_HideTools()
    Dim pwd As String
    pwd = Trim$(InputBox("Enter ADMIN password:", "Admin Login"))
    If Len(pwd) = 0 Then Exit Sub
    If Not AdminPasswordIsValid(pwd) Then
        MsgBox "Invalid admin password.", vbExclamation
        Exit Sub
    End If
    
    If Not ThisWorkbook.ProtectStructure Then _
    ThisWorkbook.Protect Password:="REPLACE_WITH_SHEET_PWD", Structure:=True, Windows:=False

    SafeHideSheet "OWNER_TOOLS"
    MsgBox "Owner Tools hidden.", vbInformation
End Sub

Public Sub Owner_SealForDistribution()
    If MsgBox("Seal this workbook for distribution? " & vbCrLf & _
              "Only OPEN_ME will remain visible on disk; OWNER_TOOLS+others go VeryHidden.", _
              vbYesNo + vbQuestion, "Seal & Save") <> vbYes Then Exit Sub

    HideAllSensitiveForRest
    SafeHideSheet "OWNER_TOOLS"
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    MsgBox "Sealed. If opened with macros disabled, only OPEN_ME will be visible.", vbInformation
End Sub

' -------------------------
' ENROLLMENT WIZARD UI
' -------------------------
Public Function ShowEnrollWizard() As Boolean
    On Error GoTo Fail

    Dim ch As String
    ch = GenerateChallenge()
    
    Dim frm As EnrollForm
    Set frm = New EnrollForm

    frm.InitWithChallenge ch
    frm.Show

    ' Returns True if authorization is achieved after wizard
    ShowEnrollWizard = (IsAuthorizedMachine())

    Unload frm
    Exit Function

Fail:
    MsgBox "Enrollment wizard error: " & Err.Description, vbCritical
    ShowEnrollWizard = False
End Function

' -------------------------
' WORKBOOK/WORKSHEET STRUCTURE HELPERS
' -------------------------
Public Sub UnprotectStructureIfNeeded()
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Unprotect Password:=SHEET_PWD
    End If
End Sub

Public Sub ProtectStructureIfNeeded()
    On Error Resume Next
    If Not ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Protect Password:=SHEET_PWD, Structure:=True, Windows:=False
    End If
End Sub

' OPEN_ME helper (visible info page for blocked/no-macro scenarios)
Public Sub EnsureOpenMeSheetExists()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("OPEN_ME")
    On Error GoTo 0

    If sh Is Nothing Then
        UnprotectStructureIfNeeded
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        With sh
            .name = "OPEN_ME"
            .Range("A1").Value = "Re-open this file with macros enabled." & vbCrLf & _
                                 "If Windows blocked the file: Right-click the .xlsm -> Properties -> check ""Unblock""."
            .Columns(1).ColumnWidth = 80
            .Rows(1).RowHeight = 60
            .Range("A1").WrapText = True
        End With
        ProtectStructureIfNeeded
    End If
End Sub

' Seal everything on disk except OPEN_ME; keep CFG/ADMIN VeryHidden
Public Sub HideAllSensitiveForRest()
    Dim sh As Worksheet

    UnprotectStructureIfNeeded
    EnsureOpenMeSheetExists
    SafeShowSheet "OPEN_ME"
    ThisWorkbook.Worksheets("OPEN_ME").Activate

    For Each sh In ThisWorkbook.Worksheets
        Select Case sh.name
            Case "OPEN_ME"
                On Error Resume Next
                sh.Unprotect Password:=SHEET_PWD
                sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
                sh.EnableSelection = xlNoSelection
                On Error GoTo 0
            Case "DRM_CFG", "OWNER_TOOLS"
                SafeHideSheet sh.name
            Case Else
                On Error Resume Next
                sh.Unprotect Password:=SHEET_PWD
                On Error GoTo 0
                SafeHideSheet sh.name
                On Error Resume Next
                sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
                sh.EnableSelection = xlNoSelection
                On Error GoTo 0
        End Select
    Next sh
    Application.OnKey "^+r"
    Application.OnKey "^+p"

    ProtectStructureIfNeeded
End Sub

' Reveal all user sheets for current session (CFG always hidden; OWNER_TOOLS admin-only)
Public Sub RevealForSession()
    Dim sh As Worksheet
    UnprotectStructureIfNeeded

    For Each sh In ThisWorkbook.Worksheets
        If sh.name <> "DRM_CFG" And sh.name <> "OWNER_TOOLS" Then
            SafeShowSheet sh.name
            On Error Resume Next
            sh.EnableSelection = xlNoRestrictions
            On Error GoTo 0
        End If
    Next sh

    ' Keep helper page hidden while editing
    SafeHideSheet "OPEN_ME"

    ProtectStructureIfNeeded
End Sub

' Temporarily remove sheet protection for editing in-session
Public Sub UnlockForEditing()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        sh.Unprotect Password:=SHEET_PWD
        On Error GoTo 0
    Next sh
End Sub

' Re-apply sheet protection (UserInterfaceOnly) to all worksheets
Public Sub LockSheets()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
        On Error GoTo 0
    Next sh
End Sub

' -------------------------
' EXPIRY CHECK
' -------------------------
Public Function IsExpired() As Boolean
    If Len(Trim$(EXPIRY_DATE)) = 0 Then
        IsExpired = False
    Else
        Dim y As Integer, m As Integer, d As Integer
        y = CInt(Left$(EXPIRY_DATE, 4))
        m = CInt(mid$(EXPIRY_DATE, 6, 2))
        d = CInt(Right$(EXPIRY_DATE, 2))
        IsExpired = (Date > DateSerial(y, m, d))
    End If
End Function

' -------------------------
' HOTKEY ACTIONS
' -------------------------
Public Sub DoSave()
    Application.CutCopyMode = False
    ThisWorkbook.Save
End Sub

Public Sub DisableSave()
    MsgBox "Save As / Save Copy As is disabled for this workbook.", vbExclamation
End Sub

Public Sub DisableCopy()
    Application.CutCopyMode = False
    MsgBox "Copy / Cut / Paste is disabled in this workbook.", vbExclamation
End Sub

' -------------------------
' SAFE SHOW/HIDE HELPERS
' -------------------------
Public Sub SafeShowSheet(ByVal sheetName As String)
    On Error GoTo Done
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD
    ThisWorkbook.Worksheets(sheetName).Visible = xlSheetVisible
Done:
    If Not ThisWorkbook.ProtectStructure Then _
        ThisWorkbook.Protect Password:=SHEET_PWD, Structure:=True, Windows:=False
End Sub

Public Sub SafeHideSheet(ByVal sheetName As String, Optional ByVal fallbackVisible As String = "OPEN_ME")
    On Error GoTo Done
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD

    ' Ensure fallback exists so we can activate it if needed
    EnsureOpenMeSheetExists

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(sheetName)
    If sh.name = ActiveSheet.name Then
        ThisWorkbook.Worksheets(fallbackVisible).Activate
    End If
    sh.Visible = xlSheetVeryHidden
Done:
    If Not ThisWorkbook.ProtectStructure Then _
        ThisWorkbook.Protect Password:=SHEET_PWD, Structure:=True, Windows:=False
End Sub

' -------------------------
' USER VISIBILITY (FORMULAS) – USER VS ADMIN
' -------------------------
Public Sub HideOnlyFormulasForUser()
    If IsAdminSession() Then Exit Sub
    Dim sh As Worksheet, rngF As Range
    On Error Resume Next

    Application.ShowFormulas = False

    For Each sh In ThisWorkbook.Worksheets
        If sh.name <> "DRM_CFG" And sh.name <> "OWNER_TOOLS" Then
            sh.Unprotect Password:=SHEET_PWD

            ' 1) Unlock everything and unhide formulas by default
            sh.Cells.Locked = False
            sh.Cells.FormulaHidden = False

            ' 2) Find formula cells and lock/hide them
            Set rngF = Nothing
            Set rngF = sh.UsedRange.SpecialCells(xlCellTypeFormulas)
            If Not rngF Is Nothing Then
                rngF.Locked = True
                rngF.FormulaHidden = True
            End If

            ' 3) Re-protect so user can interact but not see formulas
            sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True, _
                       AllowFormattingCells:=True, AllowFiltering:=True
        End If
    Next sh
End Sub

Public Sub UnhideFormulasForAdmin()
    Dim sh As Worksheet, rng As Range
    On Error Resume Next

    For Each sh In ThisWorkbook.Worksheets
        If sh.name <> "DRM_CFG" Then
            sh.Unprotect Password:=SHEET_PWD

            Set rng = Nothing
            Set rng = sh.UsedRange.SpecialCells(xlCellTypeFormulas)

            If Not rng Is Nothing Then
                rng.FormulaHidden = False   ' show formulas to admin
                rng.Locked = False          ' allow editing for admin
            End If

            sh.EnableSelection = xlNoRestrictions  ' admin can select anything

            ' re-protect only for non-admin sessions
            If Not IsAdminSession() Then
                sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True, _
                           AllowFormattingCells:=True, AllowFiltering:=True
            End If
        End If
    Next sh
End Sub


Public Sub OwnerTools_AdminFullyEditable()
    With ThisWorkbook.Worksheets("OWNER_TOOLS")
        On Error Resume Next
        .Unprotect Password:=SHEET_PWD
        .Cells.Locked = False
        .Cells.FormulaHidden = False
        
        Application.OnKey "^l", "Owner_ShowFullLogs"   ' Ctrl+L -> show full logs
        Application.OnKey "+^c", "Owner_ClearAllLogs"   ' Ctrl+Shift+C -> clear all logs


        On Error GoTo 0
    End With
End Sub

' -------------------------
' ADMIN ENROLLMENT (THIS PC)
' -------------------------
Public Function IsAdminEnrolledForThisPC() As Boolean
    On Error Resume Next
    Dim sh As Worksheet: Set sh = CfgSheet()
    Dim meName As String: meName = UCase$(Trim$(DeviceName))
    Dim bag As String: bag = Nz(sh.Range("F2").Value, "")
    If Len(bag) = 0 Then Exit Function

    ' Stored as "PC1;PC2;PC3"
    Dim items As Variant, it As Variant
    items = Split(bag, ";")
    For Each it In items
        If UCase$(Trim$(it)) = meName Then
            IsAdminEnrolledForThisPC = True
            Exit Function
        End If
    Next it
End Function

Public Sub EnrollAdminForThisPC()
    On Error GoTo Fail
    Dim sh As Worksheet: Set sh = CfgSheet()
    Dim meName As String: meName = UCase$(Trim$(DeviceName))

    Dim bag As String: bag = Nz(sh.Range("F2").Value, "")
    ' If empty, initialize with this PC
    If Len(bag) = 0 Then
        sh.Unprotect Password:=SHEET_PWD
        sh.Range("F2").Value = meName
        sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
        Exit Sub
    End If

    ' If not present, append to list
    If InStr(1, ";" & bag & ";", ";" & meName & ";", vbTextCompare) = 0 Then
        sh.Unprotect Password:=SHEET_PWD
        sh.Range("F2").Value = bag & ";" & meName
        sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    End If
    Exit Sub

Fail:
    MsgBox "EnrollAdminForThisPC error: " & Err.Description, vbExclamation
End Sub

' -------------------------
' SESSION ROLE HELPER
' -------------------------
Public Function IsAdminSession() As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("OWNER_TOOLS")
    ' If OWNER_TOOLS is visible, treat session as admin
    If Not ws Is Nothing Then
        IsAdminSession = (ws.Visible = xlSheetVisible)
    End If
End Function



' -------------------------
' LOGGING: Append enrollment event to DRM_LOG (VeryHidden)
' -------------------------
Private Sub AppendEnrollmentLog(ByVal machine As String, ByVal path As String, _
                                ByVal challenge As String, ByVal response As String, _
                                ByVal adminUser As String)
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("DRM_LOG")
    On Error GoTo CreateSheet
    GoTo WriteRow
CreateSheet:
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD
    Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    sh.name = "DRM_LOG"
    With sh
        .Range("A1").Value = "UTCDateTime"
        .Range("B1").Value = "Machine"
        .Range("C1").Value = "Path"
        .Range("D1").Value = "Challenge"
        .Range("E1").Value = "Response"
        .Range("F1").Value = "AdminUser"
        .Range("G1").Value = "Signature"
        .Columns("A:G").AutoFit
    End With
    sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
WriteRow:
    sh.Unprotect Password:=SHEET_PWD
    Dim r As Long: r = sh.Cells(sh.Rows.count, "A").End(xlUp).Row + 1
    sh.Cells(r, "A").Value = Format$(UtcNowSerial, "yyyy-mm-dd HH:nn:ss")
    sh.Cells(r, "B").Value = Nz(machine, "")
    sh.Cells(r, "C").Value = Nz(path, "")
    sh.Cells(r, "D").Value = Nz(challenge, "")
    sh.Cells(r, "E").Value = Nz(response, "")
    sh.Cells(r, "F").Value = Nz(adminUser, Application.userName)
    ' simple signature (not crypto-strong) — keeps tamper-evidence if secret unchanged
    sh.Cells(r, "G").Value = Left$(FNV1a32Hex(GetSecretKey() & "|" & sh.Cells(r, "A").Value & "|" & machine & "|" & path), RESP_LEN)
    sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    SafeHideSheet "DRM_LOG"
End Sub




Public Sub Owner_ShowFullLogs()
    On Error GoTo Fail
    Dim sh As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim cnt As Long

    If Not IsAdminSession() Then
        MsgBox "This option is available only in an ADMIN session.", vbExclamation
        Exit Sub
    End If

    ' Ensure log sheet exists (if not, inform)
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("DRM_LOG")
    On Error GoTo Fail

    If sh Is Nothing Then
        MsgBox "No DRM_LOG sheet found. No enrollments recorded yet.", vbInformation
        Exit Sub
    End If

    ' Temporarily unprotect workbook structure so we can show the sheet
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD

    ' Unhide, unprotect and activate the log sheet for admin inspection
    sh.Visible = xlSheetVisible
    sh.Unprotect Password:=SHEET_PWD
    sh.Activate

    ' Find used area and format neatly
    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    lastCol = sh.Cells(1, sh.Columns.count).End(xlToLeft).Column
    If lastRow < 2 Then
        MsgBox "DRM_LOG exists but contains no entries.", vbInformation
        GoTo ProtectAndDone
    End If

    ' Autofit & freeze header row
    sh.Columns("A:" & ColLetter(lastCol)).AutoFit
    With sh
        .Rows(1).Font.Bold = True
        .Range("A1").Resize(1, lastCol).Interior.Color = RGB(240, 240, 240)
        .Range("A1").Resize(lastRow, lastCol).Borders.LineStyle = xlContinuous
        .Range("A1").Select
        .Application.ActiveWindow.SplitRow = 1
        .Application.ActiveWindow.FreezePanes = True
    End With

    ' Count entries
    cnt = lastRow - 1
    MsgBox "DRM_LOG shown. Total enrollments: " & cnt & vbCrLf & _
           "Columns: " & sh.Cells(1, 1).Resize(1, lastCol).Address(False, False), vbInformation, "DRM Log"

ProtectAndDone:
    ' Keep sheet visible while admin inspects; optionally hide with SafeHideSheet after review
    ' Re-protect sheet for safety (UserInterfaceOnly preserves VBA access)
    sh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    If Not ThisWorkbook.ProtectStructure Then ThisWorkbook.Protect Password:=SHEET_PWD, Structure:=True, Windows:=False
    Exit Sub

Fail:
    MsgBox "Error showing DRM_LOG: " & Err.Description, vbCritical
End Sub

' Helper: returns column letter from number
Private Function ColLetter(ByVal colNum As Long) As String
    Dim v As Long, s As String
    v = colNum
    s = ""
    Do While v > 0
        Dim r As Long
        r = ((v - 1) Mod 26)
        s = Chr$(65 + r) & s
        v = Int((v - 1) / 26)
    Loop
    ColLetter = s
End Function


Public Sub Owner_ProcessEncryptedChallenge()
    On Error GoTo Fail
    Dim ownerSh As Worksheet, encToken As String
    Dim secret As String, payload As String
    Dim encBytes() As Byte, payloadBytes() As Byte
    Dim parts() As String
    Dim u As String, c As String, p As String, d As String, t As String, nonce As String
    Dim fullHex As String, expected As String
    Dim logSh As Worksheet, r As Long
    Dim cfg As Worksheet

    Set ownerSh = ThisWorkbook.Worksheets("OWNER_TOOLS")
    encToken = Trim$(ownerSh.Range("A2").Value)
    If Len(encToken) = 0 Then MsgBox "OWNER_TOOLS!A2 is empty.", vbExclamation: Exit Sub

    secret = GetSecretKey()
    If Len(secret) = 0 Then MsgBox "Secret missing", vbCritical: Exit Sub

    ' Decrypt: Base64 -> XOR -> payload string
    encBytes = Base64Decode(encToken)
    payloadBytes = XorBytes(encBytes, Utf8BytesFromString(secret))
    payload = StringFromUtf8Bytes(payloadBytes)

    ' Parse payload: user|comp|path|YYYYMMDD|HHNNSS|nonce
    parts = Split(payload, "|")
    If UBound(parts) < 5 Then
        MsgBox "Invalid decrypted payload.", vbCritical: Exit Sub
    End If
    u = parts(0): c = parts(1): p = parts(2): d = parts(3): t = parts(4): nonce = parts(5)

    ' Compute the response (this is what the client must paste during enrollment)
    fullHex = StrongResponseHex(payload)
    If RESP_LEN > 0 Then expected = Left$(fullHex, RESP_LEN) Else expected = fullHex


    ' Write the response to OWNER_TOOLS!B2 for user copy-paste
    ownerSh.Range("B2").Value = expected

    ' === LOG immediately (admin computed) ===
    On Error Resume Next
    Set logSh = ThisWorkbook.Worksheets("DRM_LOG")
    On Error GoTo 0
    If logSh Is Nothing Then
        If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:=SHEET_PWD
        Set logSh = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.count))
        logSh.name = "DRM_LOG"
        With logSh
            .Range("A1:G1").Value = Array("UTCDateTime", "Machine", "Path", "Event", "Response", "AdminUser", "Signature")
            .Columns("A:G").AutoFit
        End With
        logSh.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    End If

    r = logSh.Cells(logSh.Rows.count, "A").End(xlUp).Row + 1
    logSh.Cells(r, "A").Value = Format$(UtcNowSerial, "yyyy-mm-ddThh:nn:ss")
    logSh.Cells(r, "B").Value = UCase$(u)
    logSh.Cells(r, "C").Value = p
    logSh.Cells(r, "D").Value = "<ADMIN_COMPUTED_RESPONSE>"
    logSh.Cells(r, "E").Value = expected
    logSh.Cells(r, "F").Value = Environ$("USERNAME")
    logSh.Cells(r, "G").Value = Left$(fullHex, 16)

    ' (p??a??et???) ???ta t? decrypted payload ?a? st? DRM_CFG ??a audit
    On Error Resume Next
    Set cfg = ThisWorkbook.Worksheets("DRM_CFG")
    On Error GoTo 0
    If Not cfg Is Nothing Then
    cfg.Range("H2").Value = payload
    cfg.Range("E2").Value = UtcNowSerial
    cfg.Range("D2").Value = encToken
    End If
    
    MsgBox "OK. Response in B2. Log updated.", vbInformation
    Exit Sub
Fail:
    MsgBox "Owner_ProcessEncryptedChallenge failed: " & Err.Description, vbExclamation
End Sub



' === Base64 encode/decode using MSXML ===
Private Function Base64Encode(bytes() As Byte) As String
    Dim xml As Object, el As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = xml.createElement("b64")
    el.DataType = "bin.base64"
    el.nodeTypedValue = bytes
    Base64Encode = Replace(el.Text, vbLf, "")
End Function

Private Function Base64Decode(s As String) As Byte()
    Dim xml As Object, el As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = xml.createElement("b64")
    el.DataType = "bin.base64"
    el.Text = s
    Base64Decode = el.nodeTypedValue
End Function

' === UTF-8 <-> String using ADODB.Stream ===
Public Function Utf8BytesFromString(txt As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText txt
    st.Position = 0: st.Type = 1
    Utf8BytesFromString = st.Read
    st.Close
End Function

Private Function StringFromUtf8Bytes(bytes() As Byte) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    StringFromUtf8Bytes = st.ReadText
    st.Close
End Function

' === XOR (repeating key) ===
Private Function XorBytes(data() As Byte, key() As Byte) As Byte()
    Dim i As Long, k As Long, out() As Byte
    ReDim out(LBound(data) To UBound(data))
    For i = LBound(data) To UBound(data)
        k = LBound(key) + ((i - LBound(data)) Mod (UBound(key) - LBound(key) + 1))
        out(i) = data(i) Xor key(k)
    Next i
    XorBytes = out
End Function



Public Sub Owner_ClearAllLogs()
    If Not IsAdminSession() Then
        MsgBox "Admin only.", vbExclamation: Exit Sub
    End If

    Dim sh As Worksheet, lastRow As Long
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("DRM_LOG")
    On Error GoTo 0

    If sh Is Nothing Then
        MsgBox "No DRM_LOG sheet found.", vbInformation
        Exit Sub
    End If

    If MsgBox("Delete ALL log entries from DRM_LOG?", vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    ' Unprotect & clear (keep header row 1)
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect Password:="REPLACE_WITH_SHEET_PWD"
    sh.Visible = xlSheetVisible
    sh.Unprotect Password:="REPLACE_WITH_SHEET_PWD"

    lastRow = sh.Cells(sh.Rows.count, "A").End(xlUp).Row
    If lastRow >= 2 Then sh.Rows("2:" & lastRow).Delete

    ' Re-protect & re-hide
    sh.Protect Password:="REPLACE_WITH_SHEET_PWD", UserInterfaceOnly:=True
    If Not ThisWorkbook.ProtectStructure Then _
        ThisWorkbook.Protect Password:="REPLACE_WITH_SHEET_PWD", Structure:=True, Windows:=False
    SafeHideSheet "DRM_LOG"

    MsgBox "All logs cleared.", vbInformation
End Sub

'gia http

' HTTP helpers
' FNV32 over token|payload|ts (no secret)
' Hex dump for debug
Private Sub PostChallengeWebhook(ByVal token As String, ByVal payloadRaw As String)
    On Error Resume Next

    Dim ts As String, wbId As String
    ts = Format$(UtcNowSerial, "yyyy-mm-dd HH:nn:ss")
    wbId = ThisWorkbook.name

    
    Dim payloadJson As String, safeWb As String
    payloadJson = Replace(Replace(payloadRaw, "\", "\\"), """", "\""")
    safeWb = Replace(wbId, """", "\""")
    
    
    Dim json As String
    json = "{""token"":""" & token & """," & _
           """payload"":""" & payloadJson & """," & _
           """ts_utc"":""" & ts & """," & _
           """workbook_id"":""" & safeWb & """}"

    Debug.Print "JSON:", json
    Debug.Print "DATASTR:", token & "|" & payloadRaw & "|" & ts

    Dim ok As Boolean
    ok = HttpPostJson(WEBHOOK_URL, json)
    If Not ok Then Debug.Print "Webhook POST failed"
End Sub





Private Function HttpPostJson(ByVal url As String, ByVal bodyJson As String) As Boolean
    On Error GoTo Fail
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "User-Agent", "Excel-DRM/1.0"
    http.SetRequestHeader "X-Sig-Alg", "fnv32"   ' (p??a??et??? a? t? ???s?µ?p??e??)
    http.SetRequestHeader "X-Webhook-Key", SHARED_SECRET
    http.SetTimeouts 5000, 5000, 10000, 10000

    http.Send bodyJson   ' <-- string, ??? bytes

    Debug.Print "HTTP STATUS:", http.Status
    Debug.Print "HTTP BODY:", http.ResponseText
    HttpPostJson = (http.Status >= 200 And http.Status < 300)
    Exit Function
Fail:
    Debug.Print "HttpPostJson error:", Err.Description
    HttpPostJson = False
End Function



' JSON-escape ???? ??a t? s?µa t?? JSON (??? ??a t?? ?p???af?)
Private Function JsonEscape(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    JsonEscape = t
End Function




' Hex dump ??a debug
Private Function Hexify(ByVal s As String) As String
    Dim i As Long, out As String
    For i = 1 To Len(s)
        out = out & Right$("00" & Hex$(Asc(mid$(s, i, 1)) And &HFF), 2)
    Next
    Hexify = out
End Function



' --- Strong response: 3?FNV1a32 + light stretching ---
Private Function StrongResponseHex(ByVal payload As String) As String
    ' Χτίζουμε ισχυρότερο digest χωρίς εξαρτήσεις από εξωτερικές βιβλιοθήκες.
    ' Συνδυάζει:
    '  - βάση: secret|payload (secret από DRM_CFG!C2)
    '  - ελαφρύ key-stretching: RESP_ROUNDS φορές
    '  - 3 ανεξάρτητα μπλοκ FNV1a32 που συνενώνονται (24 hex chars)
    Dim secret As String, base As String
    Dim h1 As String, h2 As String, h3 As String
    Dim mix As String
    Dim i As Long

    secret = GetSecretKey()
    base = secret & "|" & payload

    ' 1) Αρχικό μπλοκ
    h1 = FNV1a32Hex(base)

    ' 2) Light key-stretching: «δένει» το h1 με τη βάση για RESP_ROUNDS γύρους
    mix = h1
    For i = 1 To RESP_ROUNDS
        mix = FNV1a32Hex(mix & "|" & base & "|" & CStr(i))
    Next i

    ' 3) Δεύτερο & τρίτο μπλοκ με διαφορετική μίξη/διάκριση domain
    h2 = FNV1a32Hex(base & "|A|" & mix)
    h3 = FNV1a32Hex(StrReverse(base) & "|B|" & mix)

    ' 4) Τελικό response: 8+8+8 = 24 hex
    StrongResponseHex = Left$(h1, 8) & Left$(h2, 8) & Left$(h3, 8)
End Function


Private Sub CopyToClipboard(ByVal s As String)
    On Error GoTo Fallback
    Dim o As Object
    Set o = CreateObject("MSForms.DataObject") ' late-binding, δεν θέλει reference
    o.SetText s
    o.PutInClipboard
    Exit Sub
Fallback:
    ' Fallback αν κάτι μπλοκάρει το clipboard: δείξε μόνο μήνυμα.
    MsgBox "Δεν μπόρεσα να γράψω στο clipboard. " & _
           "Αντέγραψε χειροκίνητα τον κώδικα από το μήνυμα.", vbExclamation
End Sub


