Attribute VB_Name = "modExamine"
Public Sub ExObj(txtComm As String)
DeclareMobs
Dim a As Integer, lDesc As String
Dim Y As Integer, R As Byte
lDesc = Right(txtComm, Len(txtComm) - InStr(txtComm, " "))
For Y = 0 To 100
If Obj(Y).Name <> lDesc Then GoTo NextY
If Ploc = Obj(Y).Location Then
BPrintF Obj(Y).Examine
Exit Sub
End If
NextY:
Next Y
For R = 0 To 100
If IsNumeric(Mob(R).Location) Then
If Ploc = CInt(Mob(R).Location) And Mob(R).Location <> "" And LCase(lDesc) = LCase(Mob(R).MobName) Then
BPrintF (Mob(R).Examine)
Exit Sub
End If
End If
Next R
BPrintF (lDesc & " is not here, you bum.")
End Sub
Public Sub LiSto(txtType As String)
Dim a As Byte, B As Byte, C As Byte
For a = 0 To MaxObj
BPrintF Obj(a).Pname & vbTab & "  |" & Obj(a).Location
Next a
For a = 0 To MaxMob
BPrintF Mob(a).Pname & vbTab & "  |" & Mob(a).Location
Next a
End Sub
