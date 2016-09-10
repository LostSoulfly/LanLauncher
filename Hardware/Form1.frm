VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UNIQUE HARDWARE TEST"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "SECRET"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Click Me"
      Height          =   855
      Left            =   960
      TabIndex        =   7
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "^ XOR"
      Height          =   615
      Left            =   5040
      TabIndex        =   5
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "XOR v"
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3360
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long
Private Declare Function GetTickCount Lib "Kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

Private Function CreateUID() As String
Dim bytID(1 To 16)      As Byte
Dim intIndex            As Integer
Dim strUID              As String
    Call CoCreateGuid(bytID(1))
    For intIndex = 1 To 16
        If bytID(intIndex) < CByte(16) Then
            strUID = strUID & "0"
        End If
        strUID = strUID & Hex$(bytID(intIndex))
        Select Case intIndex
            Case 4, 6, 8, 10
                strUID = strUID & "-"
        End Select
    Next intIndex
    CreateUID = strUID
End Function

Private Sub Command1_Click()
Text2.Text = Encrypt(txtKey.Text, Text1.Text)
End Sub

Private Sub Command2_Click()
Text1.Text = Encrypt(txtKey.Text, Text2.Text)
End Sub

Private Sub Command3_Click()
Dim temp As String
temp = "HDD: " & GetSerialNumber & vbCrLf & "CPU: " & GetProc
'temp = Encrypt("SECRET", temp)
Clipboard.Clear
DoEvents
Clipboard.SetText temp
DoEvents
MsgBox "Thanks bro, paste it back to me over xfire please."
Text1.Text = Clipboard.GetText
End Sub

Private Sub Form_DblClick()
MsgBox Encrypt("SECRET", Clipboard.GetText)
End Sub

Private Sub Form_Load()
Label1_Click
Label2_Click
Label3_Click
Command3_Click
End Sub

Private Sub Label1_Click()
Dim strOld      As String
Dim strNew      As String
Dim intIndex    As Long
Dim lngStart    As Long
    lngStart = GetTickCount
    For intIndex = 1 To 10000
        strNew = CreateUID
        If strNew = strOld Then
            MsgBox "No unique!", vbCritical
        End If
        strOld = strNew
    Next intIndex
    Form1.Caption = GetTickCount - lngStart & "ms - UNIQUE HARDWARE TESTS"
    Label1.Caption = strNew
End Sub

Private Sub Label2_Click()
Label2.Caption = GetSerialNumber
End Sub

Public Function GetSerialNumber() As Long

    Dim strVolumeBuffer As String
    Dim strSysName As String
    Dim lngSerialNumber As Long
    Dim lngSysFlags As Long
    Dim lngComponentLen As Long
    Dim lngResult As Long
    
    strVolumeBuffer$ = String$(256, 0)
    strSysName$ = String$(256, 0)
    lngResult = GetVolumeInformation("c:\", strVolumeBuffer$, 255, lngSerialNumber, _
            lngComponentLen, lngSysFlags, strSysName$, 255)
                 
    GetSerialNumber = lngSerialNumber
    
End Function

Public Function Encrypt(CodeKey As String, DataIn As String) As String
Dim lonDataPtr As Long
Dim strDataOut As String
Dim intXOrValue1 As Integer, intXOrValue2 As Integer
For lonDataPtr = 1 To Len(DataIn)
 'The first value to be XOr-ed comes from the data to be encrypted
 intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
 'The second value comes from the code key
 intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
 strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
Next lonDataPtr
Encrypt = strDataOut
End Function

Public Function GetProc() As String
Dim objwmiservice As Object
Dim colitems
Dim objitem
On Error Resume Next
Set objwmiservice = GetObject("winmgmts:\\.\root\cimv2")
Set colitems = objwmiservice.ExecQuery("Select * from Win32_Processor", , 48)
For Each objitem In colitems
    GetProc = objitem.ProcessorId
Next
End Function


Private Sub Label3_Click()
Label3.Caption = GetProc
End Sub
