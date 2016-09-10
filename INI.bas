Attribute VB_Name = "Module2"
Option Explicit

#If Win16 Then


Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer


Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
    ' NOTE: The lpKeyName argument for GetPr
    '     ofileString, WriteProfileString,
    'GetPrivateProfileString, and WritePriva
    '     teProfileString can be either
    'a string or NULL. This is why the argum
    '     ent is defined as "As Any".
    ' For example, to pass a string specifyB
    '     yVal "wallpaper"
    ' To pass NULL specifyByVal 0&
    'You can also pass NULL for the lpString
    '     argument for WriteProfileString
    'and WritePrivateProfileString
    ' Below it has been changed to a string
    '     due to the ability to use vbNullString


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal HwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Type Client
    Serving As Boolean
    IP As String
    First As Boolean
    Name As String
    User As Boolean
    ListIndex As Integer
    LastSeen As Integer
    TimedOut As Boolean
    Version As String
End Type

Function ReadINI(Section, KeyName, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Function writeini(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Private Sub UpdateIt(ByVal Hwnd As Long, ByVal Top As Long, ByVal lFlags As Long)
  Call SetWindowPos(Hwnd, Top, 0, 0, 0, 0, lFlags)
End Sub

Private Sub SetRemoveFrame(ByVal Hwnd As Long, ByVal lStyle As Long)
    Call SetWindowLong(Hwnd, (-16), lStyle)
    Call UpdateIt(Hwnd, 0, &H10 + &H4 + _
        &H20 + &H1 + &H2)
End Sub

Public Sub RemoveFrame(Hwnd As Long)
SetRemoveFrame Hwnd, GetWindowLong(Hwnd, (-16)) + &H400000
End Sub

Public Sub Pause(Length As Long)

    Dim n As Long
    n = Timer + Abs(Length)


    Do While Timer <= n


        DoEvents
        Loop

    End Sub
