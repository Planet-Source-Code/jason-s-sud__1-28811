VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000010&
   Caption         =   "SUD"
   ClientHeight    =   7035
   ClientLeft      =   195
   ClientTop       =   1020
   ClientWidth     =   9585
   Icon            =   "Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu Menu1 
      Caption         =   "Access"
      Begin VB.Menu SubMenu1 
         Caption         =   "SUD &Editor"
         Shortcut        =   ^E
      End
      Begin VB.Menu SubMenu2 
         Caption         =   "SUD &List"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()

End Sub

Private Sub MDIForm_Load()
Unload Form1
End Sub

Private Sub SubMenu1_Click()
frmEditor.Show
End Sub

Private Sub SubMenu2_Click()
Form2.Show
End Sub
