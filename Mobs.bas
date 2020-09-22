Attribute VB_Name = "Mobs"
Type MobTypes
    MobName As String
    Desc As String
    Location As String
    Examine As String
    Pname As String
    Speed As String
    Damage As String
    Armor As String
    Hp As String
    MaxHp As String
    SSpeed As String
    FightingYou As Boolean
End Type
Global Mob(100) As MobTypes
Global Fmsg(4) As String
Global FWmsg(4) As String
Global MaxMob As Long
Global MaxMob2 As Long

Public Sub saveMobs()
Dim X As Integer, Y As String, Name As String, a As String, Location As String
Dim B As String, C As String, D As String, E As String
For X = 0 To 100
Y = X
Call WriteToINI("Mob" & Y, "Name", Mob(Y).MobName, App.Path + "\mobsys.sud")
Call WriteToINI("Mob" & Y, "Pname", Mob(Y).Pname, App.Path + "\mobsys.sud")
Call WriteToINI("Mob" & Y, "Desc", Mob(Y).Desc, App.Path + "\mobsys.sud")
Call WriteToINI("Mob" & Y, "Location", Mob(Y).Location, App.Path + "\mobsys.sud")
Call WriteToINI("Mob" & Y, "Examine", Mob(Y).Examine, App.Path + "\mobsys.sud")
Call WriteToINI("Mob" & Y, "Speed", Mob(Y).SSpeed, App.Path + "\mobsys.sud")
Call WriteToINI("Mob" & Y, "Hp", Mob(Y).MaxHp, App.Path + "\mobsys.sud")
Next X
End Sub
Public Sub DeclareMobs()
Dim X As Integer, Y As String
Fmsg(0) = " socks you in the balls! Oooh..."
Fmsg(1) = " delivers you a powerful uppercut!"
Fmsg(2) = " jabs you in the shoulder."
Fmsg(3) = " hits you hard!"
Fmsg(4) = " knocks your lights out!"
FWmsg(0) = " TOTALLY MASSACRES you with the "
FWmsg(1) = " ANIHILATES you with the "
FWmsg(2) = " CUTS you with the "
FWmsg(3) = " beats you with the "
FWmsg(4) = " NOONINS you with the "
MaxMob = -1
For X = 0 To 100
Y = X
Mob(X).MobName = GetFromINI("Mob" + Y, "Name", App.Path + "\mobsys.sud")
Mob(X).Desc = GetFromINI("Mob" + Y, "Desc", App.Path + "\mobsys.sud")
Mob(X).Location = GetFromINI("Mob" + Y, "Location", App.Path + "\mobsys.sud")
Mob(X).Pname = GetFromINI("Mob" & Y, "Pname", App.Path + "\mobsys.sud")
Mob(X).Examine = GetFromINI("Mob" + Y, "Examine", App.Path + "\mobsys.sud")
Mob(X).SSpeed = GetFromINI("Mob" + Y, "Speed", App.Path + "\mobsys.sud")
Mob(X).MaxHp = GetFromINI("Mob" + Y, "Hp", App.Path & "\mobsys.sud")
Mob(X).Damage = GetFromINI("Mob" + Y, "Damage", App.Path & "\mobsys.sud")
Mob(X).Armor = GetFromINI("Mob" + Y, "Armor", App.Path & "\mobsys.sud")
If Mob(X).MobName = "" Then Exit Sub
If SSpeed = "" Then SSpeed = "0"
Mob(X).Speed = Mob(X).SSpeed
If Mob(X).MaxHp = "" Then Mob(X).MaxHp = "80"
Mob(X).Hp = Mob(X).MaxHp
If Mob(X).Damage = "" Then Mob(X).Damage = "8"
If Mob(X).Armor = "" Then Mob(X).Armor = "0"
If Mob(X).Location = "" Then Mob(X).Location = "-1"
MaxMob = X
Next X
If MaxMob = -1 Then
BPrintF ("ERRORRRRRRRRR")
End If
End Sub
Public Sub Summon(txt As String)
Dim namo As String
Dim Looper As Byte
Dim Looper2 As Byte
namo = Right(txt, Len(txt) - InStr(txt, " "))
For Looper = 0 To MaxObj
If Obj(Looper).Name = namo Then
Obj(Looper).Location = Ploc
Obj(Looper).Carr = True
BPrintF ("You fetch something from another dimension.")
Exit Sub
ElseIf namo = Mob(Looper).MobName Then
Mob(Looper).Location = Ploc
BPrintF ("You fetch something from another dimension.")
End If
Next Looper
End Sub
Public Sub KILL(txt As String)
Dim MoBiLe As String, a As Byte
MoBiLe = Right(txt, Len(txt) - InStr(txt, " "))
For a = 0 To MaxMob
If MoBiLe = Mob(a).MobName And Mob(a).Location = Ploc And Mob(a).FightingYou = False Then
You.AreFighting = True
Mob(a).FightingYou = True
ElseIf Mob(a).MobName = MoBiLe And Mob(a).Location <> Ploc Then
BPrintF Mob(a).Pname & " isn't in this room, please leave a message!"
Exit Sub
End If
Next a
End Sub
