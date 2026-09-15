VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Φύλλο5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Done
    If Intersect(Target, Me.Range("A2")) Is Nothing Then Exit Sub
    If Len(Trim$(Me.Range("A2").Value)) = 0 Then Exit Sub

    Application.EnableEvents = False
    Owner_ProcessEncryptedChallenge   ' <-- Will read A2, compute B2, and append to the log
Done:
    Application.EnableEvents = True
End Sub


