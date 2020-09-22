Attribute VB_Name = "Locs"
Option Explicit
Option Compare Text

Type ROOM_TYPE
    Title As String
    Exits(5) As Long
    Desc As String
    IsDark As Boolean
End Type
Global Exifs(5) As String
Global Loc(100) As ROOM_TYPE
Global MaxLoc As Long
Global Ploc As Long
Public Sub MoveDir(Direction As Byte)
Dim Y As Byte
Dim X As Byte, Z As Byte
X = 6
For Y = 0 To 5
If Direction = Y Then X = Y
Next Y
If X = 6 Then
BPrintF ("An error has occurred!")
Exit Sub
End If
If Loc(Ploc).Exits(X) <> -1 Then
Ploc = Loc(Ploc).Exits(X)
For Z = 0 To MaxObj
If Obj(Z).Carr = True Then Obj(Z).Location = Ploc
Next Z
Look (Ploc)
Else
BPrintF ("You can't go that way.")
End If
End Sub
Public Sub DecLoc()
Dim ABC As Long, CBA As Byte
On Error GoTo uhOh
Exifs(0) = "North"
Exifs(1) = "South"
Exifs(2) = "East"
Exifs(3) = "West"
Exifs(4) = "Up"
Exifs(5) = "Down"

Exifs2.Num(0) = "north."
Exifs2.Num(1) = "south."
Exifs2.Num(2) = "east."
Exifs2.Num(3) = "west."
Exifs2.Num(4) = "lands above."
Exifs2.Num(5) = "lands below."

Exo.Num(1) = "north."
Exo.Num(0) = "south."
Exo.Num(3) = "east."
Exo.Num(2) = "west."
Exo.Num(5) = "lands above."
Exo.Num(4) = "lands below."

Dim L As Byte
Dim M As Byte
Dim P As Byte
Dim P2 As Byte
Dim X As String
MaxLoc = -1
For L = 0 To 100
X = L
Loc(L).Desc = WordWrap(GetFromINI(X, "Desc", App.Path & "\locsys.sud"))
For M = 0 To 5
If IsNumeric(GetFromINI(X, Exifs(M), App.Path & "\locsys.sud")) Then Loc(L).Exits(M) = CInt(GetFromINI(X, Exifs(M), App.Path & "\locsys.sud"))
Next M
Loc(L).Title = GetFromINI(X, "Title", App.Path & "\locsys.sud")
If Loc(L).Title = "" And Loc(L).Desc = "" Then GoTo mano
MaxLoc = L
Next L
mano:
For P = (MaxLoc + 1) To 100
Loc(P).Title = "Empty Void..."
Loc(P).Desc = "You are stuck in an empty void..."
For P2 = 0 To 5
Loc(P).Exits(P2) = -1
Next P2
Next P
For CBA = 0 To MaxLoc
For ABC = 0 To 32768
If ABC <> 0 And ABC Mod 80 = 0 And ABC <= Len(Loc(CBA).Desc) Then
Mid(Loc(CBA).Desc, ABC, 2) = NL
End If
Next ABC
Next CBA
Exit Sub
''''''''''''''''''''''''''''''''''END''''''''''''''''''''''''''''

uhOh:
If MaxLoc = -1 Then
Call WriteToINI("0", "title", "Temple of Paradise", App.Path & "\locsys.sud")
Call WriteToINI("0", "North", "-1", App.Path & "\locsys.sud")
Call WriteToINI("0", "South", "-1", App.Path & "\locsys.sud")
Call WriteToINI("0", "East", "-1", App.Path & "\locsys.sud")
Call WriteToINI("0", "West", "-1", App.Path & "\locsys.sud")
Call WriteToINI("0", "Up", "-1", App.Path & "\locsys.sud")
Call WriteToINI("0", "Down", "-1", App.Path & "\locsys.sud")
Call WriteToINI("0", "desc", "You're in the Temple of Paradise.", App.Path & "\locsys.sud")
DecLoc
End If
Exit Sub
End Sub
Public Sub SaveLoc()
Dim L As Byte
Dim M As Byte
Dim S As String
For L = 0 To MaxLoc
S = L
Call WriteToINI(S, "Desc", Loc(L).Desc, App.Path & "\locsys.sud")
Call WriteToINI(S, "Title", Loc(L).Title, App.Path & "\locsys.sud")
For M = 0 To 5
Call WriteToINI(S, Exifs(M), CInt(Loc(L).Exits(M)), App.Path & "\locsys.sud")
Next M
Next L
End Sub
Public Sub newroom(roomName As String)
Dim tWo As String
MaxLoc = MaxLoc + 1
tWo = Right(roomName, Len(roomName) - InStr(roomName, " "))
If tWo = Left("newroom", Len(tWo)) Then tWo = "Voidless Void..."
Loc(MaxLoc).Title = tWo
SaveLoc
DecLoc
Ploc = MaxLoc
End Sub
Public Sub Look(roomNum As Long)
Dim txt As String
Dim S As Byte, D As Byte, O As Byte
Dim Proof As Byte

txt = txt + NL & Loc(roomNum).Title & NL & Loc(roomNum).Desc & NL
For D = 0 To 100
If IsNumeric(Mob(D).Location) Then
If Mob(D).Location = Ploc And Mob(D).Desc <> "" Then
txt = txt & Mob(D).Desc & NL
End If
End If
Next D
For O = 0 To 100
If IsNumeric(Obj(O).Location) Then
If Obj(O).Location = Ploc And Obj(O).Desc <> "" And Obj(O).Carr = False Then
txt = txt & Obj(O).Desc & NL
End If
End If
Next O
txt = txt & NL & "Obvious Exits are:"
For S = 0 To 5
If Loc(Ploc).Exits(S) <> -1 Then
txt = txt & NL & Exifs(S) & ":" & vbTab & Loc(Loc(Ploc).Exits(S)).Title
Proof = 44
End If
Next S
If Proof <> 44 Then txt = txt & NL & "None!"
BPrintF txt
End Sub
Public Sub GoToLoc(roomNum As String)
Dim S As String, a As Integer
S = Right(roomNum, Len(roomNum) - InStr(roomNum, " "))
If IsNumeric(S) Then
a = S
Ploc = a
Look (Ploc)
End If
End Sub
Public Function WordWrap(txt As String) As String
    Dim J As String, I, K As String
    
    For I = 1 To Len(txt)
        J = Mid(txt, I, 1)
        If I Mod 80 <> 0 Then
        K = K + J
        Else
        K = K + J + J + NL
        End If
    Next I
    
    WordWrap = K
End Function

