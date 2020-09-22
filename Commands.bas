Attribute VB_Name = "Commands"
Type Exifs2Type
    Num(5) As String
End Type
Global Exifs2 As Exifs2Type
Type ExoT
   Num(5) As String
 End Type
 Global Exo As ExoT
Global YouF(4) As String
Global Const NL = vbCrLf
Public Sub BPrintF(txt As String)
Dim Y As Long, Z As Long, M As Long
Form1!Text1 = (Right(Form1!Text1, 30000) & NL & txt & LN)
Form1!Text1.SelStart = Len(Form1!Text1)
End Sub
Public Sub rcvStr(StrBuf As String)
Dim Verb As String
Dim Y As Integer
Dim F As Byte
Dim X As Integer
X = InStr(StrBuf, " ")
If X = 0 Then
    Verb = StrBuf
Else
    Verb = Trim(Left(StrBuf, X - 1))
End If
Y = Len(Trim(Verb))
If Y = 0 Then Exit Sub
Select Case Verb
    Case Left("north", Y)
    MoveDir 0
    Case Left("south", Y)
    MoveDir 1
       Case Left("east", Y)
    MoveDir 2
       Case Left("west", Y)
    MoveDir 3
       Case Left("up", Y)
    MoveDir 4
       Case Left("down", Y)
    MoveDir (5)
    Case Left("drop", Y)
    DrOp StrBuf
    Case Left("look", Y)
    Look Ploc
    Case Left("examine", Y)
    ExObj StrBuf
    Case Left("kill", Y)
    KILL StrBuf
    Case Left("newroom", Y)
    newroom StrBuf
    Case Left("msg", Y)
    MSG
    Case Left("quit", Y)
    End
    Case Left("acct", Y)
    BPrintF "There are " & MaxLoc & " locations, " & MaxObj & " objects, and " & MaxMob & " mobiles."
    Case Left("save", Y)
    SaveLoc
    BPrintF "Locations saved."
    saveObj
    BPrintF "Objects saved."
    saveMobs
    BPrintF "Mobiles Saved."
    SavePlayer
    BPrintF "Player Files Saved."
    BPrintF "Save completed sucessfully ^_^"
    Case Left("summon", Y)
    Summon StrBuf
    Case Left("list", Y)
    LiSto StrBuf
    Case Left("take", Y)
    TaKe StrBuf
    Case Left("inventory", Y)
    InventOry
    Case Left("goto", Y)
    GoToLoc StrBuf
    Case Else
    BPrintF (NL & "I can't find " & Verb & " in my tree of verbs ^_^")
End Select
End Sub
