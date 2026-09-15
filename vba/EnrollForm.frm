VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnrollForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "EnrollForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnrollForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================
' EnrollForm – UI for challenge/response enrollment
' NOTE: Comments translated; code unchanged.
' =====================================================

Private challengeValue As String
Private Const CH_TTL_MINUTES As Long = 10 ' must match DRM_Core constant

' Initialize form fields with a new challenge
Public Sub InitWithChallenge(ByVal ch As String)
    challengeValue = ch
    Me.txtResponse.Text = ""
    Me.lblInfo.Caption = "New PC/path detected. Send the challenge code to your admin within " & CH_TTL_MINUTES & " minutes."
    Me.txtChallenge.Text = ch
End Sub

Private Sub cmdValidate_Click()
    On Error GoTo Fail
    Dim resp As String
    resp = Trim$(Me.txtResponse.Text)
    If Len(resp) = 0 Then
        Me.lblStatus.ForeColor = vbRed
        Me.lblStatus.Caption = "Please paste the Response code."
        Exit Sub
    End If

    ' Validate response and enroll machine/path
    If Admin_ValidateResponseAndEnroll(challengeValue, resp) Then
        Me.lblStatus.ForeColor = vbGreen
        Me.lblStatus.Caption = "Enrollment successful. You can close this window."
        Me.Hide
    Else
        Me.lblStatus.ForeColor = vbRed
        Me.lblStatus.Caption = "Invalid or expired Response. Please request a new Challenge."
    End If
    Exit Sub
Fail:
    Me.lblStatus.ForeColor = vbRed
    Me.lblStatus.Caption = "Error: " & Err.Description
End Sub

Private Sub lblInfo_Click()

End Sub

Private Sub UserForm_Click()
    ' No-op
End Sub


