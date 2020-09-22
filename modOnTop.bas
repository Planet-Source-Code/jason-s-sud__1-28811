Attribute VB_Name = "modOnTop"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global SnakeSpeed As Long
Global IncreaseSpeed As Boolean
Global LastDir As String
Global LastMove As String

Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormCentre(FormName As Form)
    FormName.Left = (Screen.Width / 2) - (FormName.Width / 2)
    FormName.Top = (Screen.Height / 2) - (FormName.Height / 2)
End Sub
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Function LoadText(Path As String) As String
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    LoadText = TextString$
End Function
Sub SaveText(txtSave As String, Path As String)
    Dim TextString As String, OldText As String
    On Error Resume Next
    TextString$ = txtSave
    OldText = LoadText(Path)
    Open Path$ For Output As #1
    Print #1, OldText & TextString$
    Close #1
End Sub
Function Round(NumberThatYouWantToRound As Double, DigitThatYouWantToRoundTheNumberTo As Integer) As Double
    Round = Int(NumberThatYouWantToRound * (10 ^ DigitThatYouWantToRoundTheNumberTo) + 0.5) / (10 ^ DigitThatYouWantToRoundTheNumberTo)
End Function
