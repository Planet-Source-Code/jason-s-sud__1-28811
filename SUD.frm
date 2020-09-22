VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JasOn's SUD"
   ClientHeight    =   7335
   ClientLeft      =   2820
   ClientTop       =   2400
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   10425
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9000
      Top             =   0
   End
   Begin VB.Timer MobFight 
      Interval        =   1500
      Left            =   9000
      Top             =   0
   End
   Begin VB.Timer moveMobs 
      Interval        =   1
      Left            =   9000
      Top             =   0
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6960
      Width           =   10455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   6975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "SUD.frx":0000
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
DecLoc
DecObj
DeclareMobs
DecPlayer
You.Hp = ((You.MaxHp) * 3 / 4)
You.Mana = ((You.MaxMana * 3) / 4)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text2.Top = (Form1.Height - (Text2.Height * 2))
Text1.Left = 0
Text1.Width = Form1.Left + Form1.Width
Text1.Height = Text2.Top
Text2.Width = Form1.Width
End Sub
Private Sub MobFight_Timer()
Dim a As Integer, B As Integer, MaxMoBy As Integer, MaxMoBo As Integer, moBo(100) As Integer, Rand1 As Byte, Rand2 As Byte
Dim RandMob As Integer, Proof As Long

Randomize
Rand1 = Int(Rnd * 5)
Randomize
Rand2 = Int(Rnd * 5)

For a = 0 To MaxMob
If Mob(a).FightingYou = True Then
If a < MaxMob Then MaxMoBy = MaxMoBy + 1
moBo(MaxMoBy) = a
Proof = 5
End If
Randomize

Next a
Randomize
RandMob = Int(Rnd * MaxMoBy + 1)
If Proof <> 5 Then You.AreFighting = False
If You.AreFighting = True Then
YouF(0) = "You sock " & Mob(moBo(RandMob)).Pname & " in the balls! Oooh..."
YouF(1) = "You deliver " & Mob(moBo(RandMob)).Pname & " a powerful uppercut!"
YouF(2) = "You jab " & Mob(moBo(RandMob)).Pname & " in the shoulder."
YouF(3) = "You hit " & Mob(moBo(RandMob)).Pname & " hard!"
YouF(4) = "You knock " & Mob(moBo(RandMob)).Pname & "'s lights out!"
End If
Randomize
If Mob(moBo(RandMob)).Location = Ploc And Mob(moBo(RandMob)).Hp > 0 And Proof = 5 Then
Randomize
BPrintF YouF(Int(Rnd * 5))
If Proof = 5 Then
BPrintF NL & "(" & Ploc & ")" & You.Hp & "/" & You.MaxHp & " " & You.Mana & "/" & You.MaxMana
End If
Randomize
Mob(moBo(RandMob)).Hp = Mob(moBo(RandMob)).Hp - Round(((Int(Rnd * You.Damage) + You.Damage / 10)), 0)
ElseIf Mob(moBo(RandMob)).Hp <= 0 And Mob(moBo(RandMob)).FightingYou = True Then
Mob(moBo(RandMob)).FightingYou = False
BPrintF Mob(moBo(RandMob)).Pname & " has died."
Mob(moBo(RandMob)).Location = 100
Mob(moBo(RandMob)).Hp = Mob(moBo(RandMob)).MaxHp
End If
For B = 0 To MaxMoBo
If Mob(moBo(B)).Location = Ploc And Mob(moBo(B)).Hp > 0 And Mob(moBo(B)).FightingYou = True Then
BPrintF Mob(moBo(B)).Pname & Fmsg(Int(Rnd * 5))
You.Hp = You.Hp - Round((Int((Rnd * Mob(moBo(B)).Damage) + (Mob(moBo(B)).Damage / 10)) * (You.Armor / 100)), 0)
End If

Randomize
Next B
Randomize
If Proof = 5 Then
BPrintF NL & "(" & Ploc & ")" & You.Hp & "/" & You.MaxHp & " " & You.Mana & "/" & You.MaxMana
End If
blimpies:
End Sub

Private Sub moveMobs_Timer()

Dim RandExit As Byte, Gothrough As Long

For Gothrough = 0 To MaxMob
Mob(Gothrough).Speed = Mob(Gothrough).Speed - 1
poople: If Mob(Gothrough).Speed = 0 And Mob(Gothrough).FightingYou = False Then
    Randomize
    RandExit = Int(Rnd * 6)
    If Loc(Mob(Gothrough).Location).Exits(RandExit) <> -1 Then
        If Loc(Mob(Gothrough).Location).Exits(RandExit) = Ploc Then
        BPrintF (Mob(Gothrough).Pname & " wanders in from the " & Exo.Num(RandExit))
        BPrintF NL & "(" & Ploc & ")" & You.Hp & "/" & You.MaxHp & " " & You.Mana & "/" & You.MaxMana
        End If
        If Mob(Gothrough).Location = Ploc Then
        BPrintF (Mob(Gothrough).Pname & " wanders to the " & Exifs2.Num(RandExit))
        BPrintF NL & "(" & Ploc & ")" & You.Hp & "/" & You.MaxHp & " " & You.Mana & "/" & You.MaxMana
        End If
        Mob(Gothrough).Location = Loc(Mob(Gothrough).Location).Exits(RandExit)
        Mob(Gothrough).Speed = Mob(Gothrough).SSpeed
        ElseIf Loc(Mob(Gothrough).Location).Exits(RandExit) = -1 Then GoTo poople
    End If
 
End If

Next Gothrough

End Sub

Private Sub Text1_GotFocus()
SendKeys ("{Tab}")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Exit Sub
Else

BPrintF Text2
rcvStr Text2
Text2 = ""
BPrintF NL & "(" & Ploc & ")" & You.Hp & "/" & You.MaxHp & "~" & You.Mana & "/" & You.MaxMana
End If
End Sub

Private Sub Timer1_Timer()
If You.AreFighting = False Then
If You.Hp <= You.MaxHp Then
You.Hp = You.Hp + 1
End If
If You.Mana <= You.MaxMana Then
You.Mana = You.Mana + 1
End If
End If
If You.Hp > You.MaxHp Then You.Hp = You.MaxHp
If You.Mana > You.MaxMana Then You.Mana = You.MaxMana
End Sub

Private Sub Txtview1_GotFocus()

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
