VERSION 5.00
Begin VB.Form frmEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SUD Editor"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9390
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "writer.frx":0000
      Left            =   120
      List            =   "writer.frx":0010
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   7200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Private Sub Command1_Click()

Select Case List1.Text
Case "Mobiles"
Open App.Path & "\mobsys.sud" For Output As #1 Len = FileLen(App.Path & "\mobsys.sud")
Print #1, Text1
Close #1
Case "Objects"
Open App.Path & "\objsys.sud" For Output As #1 Len = FileLen(App.Path & "\objsys.sud")
Print #1, Text1
Close #1
Case "Locations"
Open App.Path & "\locsys.sud" For Output As #1 Len = FileLen(App.Path & "\locsys.sud")
Print #1, Text1
Close #1
Case "PlayerFiles"
Open App.Path & "\pfiles.sud" For Output As #1 Len = FileLen(App.Path & "\pfiles.sud")
Print #1, Text1
Close #1
Case Else
Text1 = "???"
End Select

End Sub
Private Sub List1_Click()
On Error GoTo endop


Select Case List1.Text
Case "Mobiles"
Open App.Path & "\mobsys.sud" For Input As #1 Len = FileLen(App.Path & "\mobsys.sud")
Text1 = Input$(FileLen(App.Path & "\mobsys.sud"), #1)
Close #1
Case "Objects"
Open App.Path & "\objsys.sud" For Input As #1 Len = FileLen(App.Path & "\objsys.sud")
Text1 = Input$(FileLen(App.Path & "\objsys.sud"), #1)
Close #1
Case "Locations"
Open App.Path & "\locsys.sud" For Input As #1 Len = FileLen(App.Path & "\locsys.sud")
Text1 = Input$(FileLen(App.Path & "\locsys.sud"), #1)
Close #1
Case "PlayerFiles"
Open App.Path & "\pfiles.sud" For Input As #1 Len = FileLen(App.Path & "\pfiles.sud")
Text1 = Input$(FileLen(App.Path & "\pfiles.sud"), #1)
Close #1
Case Else
Text1 = "???"
End Select
endop:
End Sub
