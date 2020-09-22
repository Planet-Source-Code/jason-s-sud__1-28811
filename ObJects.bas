Attribute VB_Name = "ObJects"
Type OBJ_TYPE
StartLoc As String
Location As String
Desc As String
Name As String
Pname As String
Carr As Boolean
Examine As String
Damage As String
Armor As String
Lightable As Boolean
Wear(10) As Boolean
WearOn(10) As String
Nogeto As String
NoGet As Boolean
End Type
Global Obj(100) As OBJ_TYPE

Global MaxObj As Long


Public Sub saveObj()
Dim X As Byte, Y As String
For X = 0 To MaxMob
Y = X
Call WriteToINI(Y, "Name", Obj(Y).Name, App.Path + "\objsys.sud")
Call WriteToINI(Y, "Pname", Obj(Y).Pname, App.Path + "\objsys.sud")
Call WriteToINI(Y, "Desc", Obj(Y).Desc, App.Path + "\objsys.sud")
Call WriteToINI(Y, "Location", Obj(Y).StartLoc, App.Path + "\objsys.sud")
Call WriteToINI(Y, "Examine", Obj(Y).Examine, App.Path + "\objsys.sud")
Call WriteToINI(Y, "Noget", Obj(Y).Nogeto, App.Path + "\objsys.sud")
If LCase(Obj(X).Nogeto) = "true" Then
Obj(X).NoGet = True
Else
Obj(X).NoGet = False
End If
Next X

End Sub
Public Sub DecObj()
Dim X As Integer, Y As String, A As Byte
MaxObj = -1
For X = 0 To 100
Y = X
        Obj(X).Name = GetFromINI(Y, "Name", App.Path + "\objsys.sud")
  Obj(X).Pname = GetFromINI(Y, "Pname", App.Path + "\objsys.sud")
  Obj(X).Desc = GetFromINI(Y, "Desc", App.Path + "\objsys.sud")
    Obj(X).StartLoc = GetFromINI(Y, "Location", App.Path + "\objsys.sud")
    Obj(X).Examine = GetFromINI(Y, "Examine", App.Path + "\objsys.sud")
    Obj(X).Nogeto = GetFromINI(Y, "Noget", App.Path + "\objsys.sud")
  If Obj(X).Name = "" Then Exit Sub
  Obj(X).Location = Obj(X).StartLoc
  If LCase(Obj(X).Nogeto) = "true" Then
Obj(X).NoGet = True
Else
Obj(X).NoGet = False
End If
MaxObj = X
Next X
For A = (MaxObj + 1) To 100
Obj(A).StartLoc = -1
Next A
If MaxObj = -1 Then
End If
End Sub
Public Sub TaKe(txt As String)
Dim ItEm As String, A As Byte, B As Byte, Proof As Integer
ItEm = Right(txt, Len(txt) - InStr(txt, " "))
If ItEm = "all" Then
For B = 0 To MaxObj
If Obj(B).Location = Ploc And Obj(B).NoGet = False And Obj(B).Carr = False Then
BPrintF "You take the " & Obj(A).Name
Obj(B).Carr = True
End If
Next B
Else
For A = 0 To MaxObj
If Obj(A).Location = Ploc And Obj(A).Name = ItEm And Obj(A).NoGet = False And Obj(A).Carr = False Then
Obj(A).Carr = True
BPrintF "You take the " & Obj(A).Name
End If
Next A
End If
End Sub
Public Sub DrOp(txt)
Dim PeE As String
PeE = Right(txt, Len(txt) - InStr(txt, " "))
Dim A As Byte, B As Byte

If PeE = "all" Then
    For A = 0 To MaxObj
    If Obj(A).Carr = True Then
        Obj(A).Carr = False
        BPrintF "You drop the " & Obj(A).Name & "."
        End If
    Next A
    ElseIf PeE = "" Then
        BPrintF "What do you want to drop?"
    Else
        For A = 0 To MaxMob
        If PeE = Obj(A).Name And Obj(A).Carr = True Then
        Obj(A).Carr = False
        BPrintF "You drop the " & Obj(A).Name & "."
        Exit Sub
        End If
        Next A
End If
End Sub

Public Sub InventOry()
Dim A As Byte, LisTer As String
BPrintF "You are currently carrying up your ass:"
For A = 0 To MaxObj
    If Obj(A).Carr = True Then LisTer = LisTer & " " & Obj(A).Name
Next A
BPrintF LisTer
End Sub
