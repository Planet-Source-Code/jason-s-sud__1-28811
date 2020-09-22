VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lisr"
   ClientHeight    =   1770
   ClientLeft      =   5730
   ClientTop       =   2775
   ClientWidth     =   1620
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   1620
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Listero.frx":0000
      Left            =   0
      List            =   "Listero.frx":000D
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()
If List1.List(0) = "My SUD" Then
Form1.Show
Form1.WindowState = 2
End If
End Sub
