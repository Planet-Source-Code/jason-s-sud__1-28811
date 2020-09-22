Attribute VB_Name = "modPlayer"
Type PLAYER_REC
Hp As String
MaxHp As String
Mana As String
MaxMana As String
Damage As String
Armor As String
AreFighting As Boolean
End Type
Global You As PLAYER_REC
Global Const PlS = "Player Stats"
Public Sub DecPlayer()
You.MaxHp = GetFromINI(PlS, "Hp", App.Path & "\pfiles.sud")
You.MaxMana = GetFromINI(PlS, "Mana", App.Path & "\pfiles.sud")
You.Damage = GetFromINI(PlS, "Damage", App.Path & "\pfiles.sud")
You.Armor = GetFromINI(PlS, "Armor", App.Path & "\pfiles.sud")
You.Hp = You.MaxHp
You.Mana = You.MaxMana
End Sub
Public Sub SavePlayer()
Call WriteToINI(PlS, "Hp", You.MaxHp, App.Path & "\pfiles.sud")
Call WriteToINI(PlS, "Mana", You.MaxMana, App.Path & "\pfiles.sud")
Call WriteToINI(PlS, "Armor", You.Armor, App.Path & "\pfiles.sud")
Call WriteToINI(PlS, "Damage", You.Damage, App.Path & "\pfiles.sud")
End Sub
Public Sub MSG()
Dim A As Integer
For A = 0 To 4
BPrintF YouF(A)
Next A
End Sub
