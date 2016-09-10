VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "LanLauncher"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12330
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   12330
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtxtName 
      Height          =   375
      Left            =   2520
      TabIndex        =   45
      Top             =   4680
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   4210752
      Enabled         =   0   'False
      MultiLine       =   0   'False
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Form1.frx":1CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Frame frmColor 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   0
      TabIndex        =   46
      Top             =   4800
      Visible         =   0   'False
      Width           =   5535
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2040
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   56
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   55
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1080
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   54
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   600
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   53
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   52
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   51
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1080
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   50
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   49
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2040
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   48
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   360
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   47
         ToolTipText     =   "Select a position in your Nickname and then click a color to have all the text after that appear as the color you clicked."
         Top             =   840
         Width           =   375
      End
      Begin LanLauncher.jcbutton jcbutton2 
         Height          =   615
         Left            =   2880
         TabIndex        =   57
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Hide Nick Editor"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         PicturePushOnHover=   -1  'True
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
   End
   Begin LanLauncher.jcbutton cmdReset 
      Height          =   255
      Left            =   4440
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "AlterIW.net"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.PictureBox pNormal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4800
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   43
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pOffline 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4800
      Picture         =   "Form1.frx":1D4E
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   42
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pOnline 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4800
      Picture         =   "Form1.frx":5A31
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame frmUpdate 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   5535
      Begin LanLauncher.jcbutton btnUClose 
         Height          =   255
         Left            =   4200
         TabIndex        =   40
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Close"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin LanLauncher.jcbutton jcbutton1 
         Height          =   255
         Left            =   4200
         TabIndex        =   44
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16765357
         Caption         =   "Ignore Version"
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.Label lblURL 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label lblUpdate 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Available"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.CheckBox chkUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      ToolTipText     =   "Check for updates upon program launch."
      Top             =   5640
      Width           =   175
   End
   Begin VB.Timer btnTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5280
      Top             =   5400
   End
   Begin LanLauncher.jcbutton btnScroll 
      Height          =   615
      Left            =   7800
      TabIndex        =   32
      Top             =   3480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "<"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin LanLauncher.jcbutton cmdStart 
      Height          =   615
      Left            =   5760
      TabIndex        =   25
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Start MW2"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin LanLauncher.jcbutton cmdSave 
      Height          =   375
      Left            =   5760
      TabIndex        =   29
      Top             =   5160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Save Config"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin LanLauncher.jcbutton cmdRld 
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Reload Config"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin LanLauncher.jcbutton cmdServer 
      Height          =   495
      Left            =   5760
      TabIndex        =   27
      Top             =   4680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Start IWnet Server"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin LanLauncher.jcbutton cmdSettings 
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   4080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Settings"
      PictureEffectOnOver=   0
      PictureEffectOnDown=   0
      PicturePushOnHover=   -1  'True
      CaptionEffects  =   0
      TooltipBackColor=   0
   End
   Begin VB.HScrollBar hsbOpacity 
      Height          =   255
      Left            =   5760
      Max             =   255
      Min             =   20
      TabIndex        =   10
      Top             =   5880
      Value           =   255
      Width           =   2295
   End
   Begin VB.CheckBox chkNoIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "Shows or hides user's IP addresses in the list. (Now instant!)"
      Top             =   5400
      Width           =   175
   End
   Begin VB.TextBox txtLog2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtServer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtOver 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   5400
      TabIndex        =   6
      Text            =   "127.0.0.1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CheckBox chkSave 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   $"Form1.frx":9791
      Top             =   5400
      Width           =   175
   End
   Begin VB.CheckBox chkName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "This will toggle the option to show either Names or IPs in the Chat window."
      Top             =   5160
      Width           =   175
   End
   Begin VB.CheckBox chkAuto 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Auto-Select Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "This will choose an IP that is hosting a server."
      Top             =   5160
      Value           =   1  'Checked
      Width           =   175
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox txtMSG 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      MaxLength       =   1024
      TabIndex        =   0
      Text            =   "Type here and press enter to send message.."
      Top             =   4080
      Width           =   5535
   End
   Begin VB.ListBox lstIPs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   5760
      TabIndex        =   14
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtLog1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   -120
      Picture         =   "Form1.frx":9831
      ScaleHeight     =   1800
      ScaleWidth      =   9750
      TabIndex        =   11
      Top             =   0
      Width           =   9750
      Begin VB.Timer tmrBroadcast 
         Interval        =   15000
         Left            =   3360
         Top             =   1200
      End
      Begin VB.Timer tmrSpam 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   4320
         Top             =   1200
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   3840
         Top             =   1200
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Timer tmrName 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2760
         Top             =   1200
      End
      Begin VB.CommandButton cmdMinimize 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         MouseIcon       =   "Form1.frx":D815
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Minimize"
         Top             =   0
         Width           =   255
      End
      Begin VB.Timer tmrStatus 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2280
         Top             =   1200
      End
      Begin VB.Timer tmrCheck 
         Enabled         =   0   'False
         Interval        =   8000
         Left            =   1800
         Top             =   1200
      End
      Begin VB.Timer tmrPing 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1320
         Top             =   1200
      End
      Begin MSWinsockLib.Winsock sckStatus 
         Left            =   840
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock sckUDP2 
         Left            =   840
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock sckTCP 
         Left            =   360
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock sckUDP 
         Left            =   360
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   43210
      End
      Begin VB.CommandButton cmdX 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8050
         MouseIcon       =   "Form1.frx":D967
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label lblIP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Top             =   1480
         Width           =   3015
      End
      Begin VB.Label lblVer 
         BackStyle       =   0  'Transparent
         Caption         =   "ver. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   2295
      End
   End
   Begin VB.CheckBox chkOffline 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   58
      ToolTipText     =   "Immediately removes users if they go offline."
      Top             =   5640
      Width           =   175
   End
   Begin VB.CheckBox chkExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   60
      ToolTipText     =   "Automatically exit the LanLauncher after starting MW2."
      Top             =   5880
      Width           =   175
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00000000&
      Caption         =   "Exit Automatically"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   61
      ToolTipText     =   "Automatically exit the LanLauncher after starting MW2."
      Top             =   5910
      Width           =   1695
   End
   Begin VB.Label lblOffline 
      BackColor       =   &H00000000&
      Caption         =   "Remove Offline Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   59
      ToolTipText     =   "Immediately removes users if they go offline."
      Top             =   5665
      Width           =   1695
   End
   Begin VB.Label lblUpdates 
      BackColor       =   &H00000000&
      Caption         =   "Check for Updates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   34
      ToolTipText     =   "Check for updates upon program launch."
      Top             =   5670
      Width           =   1695
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label lblNoIP 
      BackColor       =   &H00000000&
      Caption         =   "Show IPs in List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   24
      ToolTipText     =   "Shows or hides user's IP addresses in the list. (Now instant!)"
      Top             =   5430
      Width           =   1695
   End
   Begin VB.Label lblSave 
      BackColor       =   &H00000000&
      Caption         =   "Auto-Save Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   $"Form1.frx":DAB9
      Top             =   5430
      Width           =   1695
   End
   Begin VB.Label lblNames 
      BackColor       =   &H00000000&
      Caption         =   "Show Names in Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "This will toggle the option to show either Names or IPs in the Chat window."
      Top             =   5190
      Width           =   1695
   End
   Begin VB.Label lblAuto 
      BackColor       =   &H00000000&
      Caption         =   "Auto-Select Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "This will choose an IP that is hosting a server."
      Top             =   5190
      Width           =   1695
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2640
      TabIndex        =   17
      ToolTipText     =   $"Form1.frx":DB59
      Top             =   4440
      Width           =   705
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "options"
      Visible         =   0   'False
      Begin VB.Menu mnuServer 
         Caption         =   "Connect to Server"
      End
      Begin VB.Menu mnuLobby 
         Caption         =   "Connect to Lobby"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hWnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Dim SplitChar As String
Public LocalIP As String
Dim Client() As Client
Dim Last() As Client
Dim Saved As Boolean
Dim Server As Boolean
Dim ClientName As String
Dim Startup As Boolean
Dim Max As String
Dim Min As String
Dim FWidth As String
Dim ShowLast As Boolean
Dim LastMessage As String
Dim msgTotal As Long
Dim cUpdate As Boolean
Dim daNum As Integer
Dim IsOnline As Boolean
Dim blDebug As Boolean

Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Dim OldX As Single
Dim OldY As Single

Public Function GetSerialNumber() As Long

    Dim strVolumeBuffer As String
    Dim strSysName As String
    Dim lngSerialNumber As Long
    Dim lngSysFlags As Long
    Dim lngComponentLen As Long
    Dim lngResult As Long
    
    strVolumeBuffer$ = String$(256, 0)
    strSysName$ = String$(256, 0)
    'lngResult = GetVolumeInformation("c:\", strVolumeBuffer$, 255, lngSerialNumber, _
            lngComponentLen, lngSysFlags, strSysName$, 255)
                 
    'GetSerialNumber = "NOP"
    
    
End Function

Private Function Opacity(Value As Byte, Frm As Form)
On Error GoTo ErrorHandler

Dim MaxVal As Byte, MinVal As Byte
    
MinVal = 20: MaxVal = 255
    
If Value > MaxVal Then Value = MaxVal
If Value < MinVal Then Value = MinVal
    
SetWindowLongA Frm.hWnd, GWL_EXSTYLE, GetWindowLongA(Frm.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
SetLayeredWindowAttributes Frm.hWnd, 0, Value, LWA_ALPHA
    
ErrorHandler:   Exit Function
End Function

Private Sub btnScroll_Click()
On Error Resume Next
btnScroll_MouseLeave
If Saved = False Then
    a = MsgBox("Your settings are unsaved, would you like to save them before starting?", vbYesNoCancel, "Settings")
        If a = vbYes Then
        cmdSave_Click
        StartIt "iw4mp", False
        Exit Sub
        ElseIf a = vbCancel Then
        Exit Sub
        End If
End If
'Send "SYN" & SplitChar & LocalIP
If StartIt("iw4mp", False) = True Then If chkExit.Value = vbChecked Then Send "STARTING", "MW2 and his LanLauncher is closing automatically": DoEvents: End
End Sub

Private Sub btnScroll_MouseEnter()
If btnScroll.Caption = "<" Then
        btnScroll.Left = 5760
        btnScroll.Width = 2295
        btnScroll.Caption = "Launch w/ Updater"
        
        
End If
End Sub

Private Sub btnScroll_MouseLeave()
        btnScroll.Left = 7800
        btnScroll.Width = 255
        btnScroll.Caption = "<"
End Sub

Private Sub btnUClose_Click()
frmUpdate.Visible = False
End Sub

Private Sub chkExit_Click()
cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
End Sub

Private Sub chkNoIP_Click()
cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
UpdateList
End Sub

Private Sub chkOffline_Click()
cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
DoEvents
RemoveOffline
End Sub

Private Sub chkSave_Click()
cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
End Sub


Private Sub chkUpdates_Click()
cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
Update
End Sub

Private Sub cmdMinimize_Click()
    Me.WindowState = 1
End Sub

Private Sub cmdSettings_Click()
On Error Resume Next

If cmdSettings.Caption = "Settings" Then
cmdSettings.Caption = "Hide Settings"
Else
cmdSettings.Caption = "Settings"
End If

If cmdSettings.Caption = "Hide Settings" Then
    frmMain.Height = Max
    Exit Sub
Else
    frmMain.Height = Min
End If
End Sub

'255 to 615
'4080 to 3840
Private Sub cmdSettings_MouseEnter()
cmdStart.Height = 255
btnScroll.Height = 255
cmdSettings.Height = 615
cmdSettings.Top = 3720
End Sub

Private Sub cmdSettings_MouseLeave()
cmdStart.Height = 615
btnScroll.Height = 615
cmdSettings.Height = 255
cmdSettings.Top = 4080
End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
cmdStart.Height = cmdStart.Height * 2
cmdSettings.Height = cmdSettings.Height / 2
End Sub

Private Sub cmdX_Click()
    Unload Me
    End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
If FWidth = "" Then Exit Sub
    If Me.WindowState <> 1 Then
        If Not Me.Height = Min Then
            If Not Me.Height = Max Then
                Me.Height = Min
                cmdSettings.Caption = "Settings"
            End If
        End If
    If Me.Width < FWidth Then Me.Width = FWidth
    If Me.Width > FWidth Then Me.Width = FWidth
    'If Me.Height < min Then Me.Height = min
    'If Me.Height > max Then Me.Height = max
  End If
End Sub

Private Sub frmColor_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub frmcolor_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Sub frmUpdate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub frmUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Sub hsbOpacity_Change()
On Error Resume Next

hsbOpacity_Scroll
If chkSave.Value = 1 Then cmdSave_Click
End Sub

Private Sub hsbOpacity_Scroll()
On Error Resume Next

Saved = False
cmdSave.Enabled = True
'If chkSave.Value = 1 Then cmdSave_Click
Opacity hsbOpacity.Value, Me
End Sub

Private Sub chkAuto_Click()
If chkAuto.Value = Checked Then ChooseServer
'cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
End Sub

Private Function StartIt(EXE As String, UseDAT As Boolean) As Boolean
On Error GoTo OHSHT
If UseDAT = False Then
    If Dir(App.Path & "\" & EXE & ".exe") <> "" Then
        Shell App.Path & "\" & EXE & ".exe"
        'SendToAll "STARTING", EXE
        StartIt = True
    Else
        MsgBox EXE & ".exe could not be found. Please place the LanLauncher in the MW2 folder."
        StartIt = False
    End If
Exit Function
End If

If UseDAT = True Then
    If Dir(App.Path & "\" & EXE & ".dat") <> "" Then
        Shell App.Path & "\" & EXE & ".dat"
        'SendToAll "STARTING", EXE
        StartIt = True
    Else
        MsgBox EXE & ".dat could not be found. Please place the LanLauncher in the MW2 folder."
        StartIt = False
    End If
Exit Function
End If

OHSHT:
If Not Error = "" Then MsgBox Err.Number & " " & Err.Description
Exit Function
End Function

Private Sub chkName_Click()
cmdSave.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
End Sub

Private Sub cmdReset_Click()
txtServer.Text = "master.alterrev.net"
End Sub

Private Sub cmdSave_Click()
If Startup = True Then Exit Sub
writeini "Configuration", "Server", txtServer.Text, App.Path & "\iw4nick.ini"
writeini "Configuration", "Nickname", Trim(txtName.Text), App.Path & "\iw4nick.ini"
writeini "LanLauncher", "Auto", chkAuto.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "Save", chkSave.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "NoIP", chkNoIP.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "Update", chkUpdates.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "AutoExit", chkExit.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "RemOffline", chkOffline.Value, App.Path & "\iw4nick.ini"
'writeini "LanLauncher", "Bar", chkBar.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "ShowNames", chkName.Value, App.Path & "\iw4nick.ini"
writeini "LanLauncher", "Opacity", hsbOpacity.Value, App.Path & "\iw4nick.ini"
If ReadINI("Configuration", "WebHost", App.Path & "\iw4nick.ini") = "" Then _
writeini "Configuration", "WebHost", "auto", App.Path & "\iw4nick.ini"
Saved = True
cmdSave.Enabled = False
Client(0).Name = NormalizeName(txtName.Text): ClientName = Client(0).Name
End Sub

Public Function NormalizeName(Name As String) As String
NormalizeName = Trim(txtName.Text)
NormalizeName = Replace(NormalizeName, "^0", "")
NormalizeName = Replace(NormalizeName, "^1", "")
NormalizeName = Replace(NormalizeName, "^2", "")
NormalizeName = Replace(NormalizeName, "^3", "")
NormalizeName = Replace(NormalizeName, "^4", "")
NormalizeName = Replace(NormalizeName, "^5", "")
NormalizeName = Replace(NormalizeName, "^6", "")
NormalizeName = Replace(NormalizeName, "^7", "")
NormalizeName = Replace(NormalizeName, "^8", "")
NormalizeName = Replace(NormalizeName, "^9", "")
'NormalizeName = Replace(NormalizeName, ":", "")
End Function

Private Sub cmdServer_Click()
Call StartIt("IWNetServer", False)
End Sub

Private Sub cmdStart_Click()
On Error Resume Next

If IsOnline = False Then
    a = MsgBox("The server you are trying to connect to does not appear to be online." & vbCrLf & vbCrLf & "Are you sure you want to attempt to connect?", vbYesNo, "Server Appears Offline")
    If a = vbNo Then Exit Sub
End If

If Saved = False Then
    a = MsgBox("Your settings are unsaved, would you like to save them before starting?", vbYesNoCancel, "Settings")
        If a = vbYes Then
        cmdSave_Click
        StartIt "iw4mp", True
        Exit Sub
        ElseIf a = vbCancel Then
        Exit Sub
        End If
End If
'Send "SYN" & SplitChar & LocalIP
If StartIt("iw4mp", True) = True Then If chkExit.Value = vbChecked Then Send "STARTING", "MW2 and his LanLauncher is closing automatically": DoEvents: End
End Sub

Private Sub cmdRld_Click()
x = MsgBox("Are you sure you want to reload settings? You'll lose what you have not saved.", vbYesNo, "Reload Settings")
If x = vbNo Then Exit Sub
LoadINI
cmdSave.Enabled = False
End Sub

Private Sub RemoveOld()
On Error Resume Next
Kill App.Path & "\LanLaunch.exe"

End Sub

Private Sub Form_Initialize()
InitCommonControls
RemoveFrame Me.hWnd
'ShowBar False
RemoveOld

End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then
MsgBox "There is an instance already running."
End
End If

pNormal.Picture = Picture1.Picture
If InStr(1, LCase(Environ("ALLUSERSPROFILE")), "programdata") > 0 Then
Max = 6400 '6550 '6350
Min = 4700
FWidth = 8400
Else
Max = 6400
Min = 4550
FWidth = 8250
End If

ClientName = "Player"
Startup = True
sckUDP.Protocol = sckUDPProtocol
sckUDP.RemoteHost = "255.255.255.255"
sckUDP.RemotePort = "43210"
sckUDP.Bind 43210
sckUDP2.Protocol = sckUDPProtocol
sckUDP2.RemoteHost = "255.255.255.255"
sckUDP2.RemotePort = "43210"
frmMain.Height = Min
'frmMain.Height = 4950
'sckUDP2.Bind 43210

lblVer.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
SplitChar = "/n/"
LocalIP = sckTCP.LocalIP
ReDim Client(0)
ReDim Last(0)

Client(0).Version = App.Major & App.Minor & App.Revision
Client(0).ListIndex = 0
'txtLog.Text = "Goodbye, alterIW. You've been a large part of my life for so long and I've met a lot of great people. It was so much fun! You will be missed. I am LostSoulFly/Dragoonadept. Xfire: dragoonadept"
txtLog.Text = "Hello TrollParty goers, this is a special version I've cooked up just for us! <3"
Log "~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
Call Colors
Call LoadINI

Call AddIP(LocalIP, True, ClientName)

txtLog1.Text = "Network Traffic"
'Send "DISCOVER" & SplitChar & LocalIP
Send "PING1" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "0"
DoEvents
Send "PING1" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "0"
DoEvents
CheckSrv

'Call UpdateList
'Call AskUpdate
Startup = False
DoEvents
tmrName.Enabled = True
tmrStatus.Enabled = True
tmrCheck.Enabled = True
tmrPing.Enabled = True
End Sub

Private Sub AskUpdate()
If ReadINI("LanLauncher", "Update", App.Path & "\iw4nick.ini") = "" Then
    Call writeini("LanLauncher", "#UPDATE1", "dl.dropbox.com", App.Path & "\iw4nick.ini")
    Call writeini("LanLauncher", "#UPDATE2", "http://dl.dropbox.com/u/4275989/Version.txt", App.Path & "\iw4nick.ini")
    If MsgBox("Would you like the LanLauncher to check for updates when you run it?" & vbCrLf & "(You can disable this under Settings at any time.)", vbYesNo, "Check for Updates") = vbYes Then
    chkUpdates.Value = 1
    Else
    chkUpdates.Value = 0
    End If
    Call writeini("LanLauncher", "Update", chkUpdates.Value, App.Path & "\iw4nick.ini")
ElseIf ReadINI("LanLauncher", "Update", App.Path & "\iw4nick.ini") = "1" Then
    chkUpdates.Value = 1
Else
    chkUpdates.Value = 0
End If
DoEvents
'Call Update
DoEvents
End Sub

Private Sub LoadINI()
'On Error Resume Next

txtServer.Text = ReadINI("Configuration", "Server", App.Path & "\iw4nick.ini")
txtName.Text = ReadINI("Configuration", "Nickname", App.Path & "\iw4nick.ini")

If Not ReadINI("LanLauncher", "LASTPC", App.Path & "\iw4nick.ini") = Environ("COMPUTERNAME") & "|" & Environ("USERNAME") Then
    If Dir(App.Path & "\LANConnecter.exe") <> "" Then
    Call Shell(App.Path & "\lanconnecter.exe /sinstall", vbNormalFocus)
    Call writeini("LanLauncher", "LASTPC", Environ("COMPUTERNAME") & "|" & Environ("USERNAME"), App.Path & "\iw4nick.ini")
    'Call writeini("Dynamic", "ServerWarningShown", "True", App.Path & "\iw4nick.ini")
    Log "## LANConnecter was found, you can now connect to MW2 lobbies directly from the LanLauncher!"
    Else
    Log "## LANConnecter was not found! (You won't be able to join lobbies.)"
    End If
End If

    

If ReadINI("LanLauncher", "ShowNames", App.Path & "\iw4nick.ini") = "" Then
chkName.Value = 1
ElseIf ReadINI("LanLauncher", "ShowNames", App.Path & "\iw4nick.ini") = "1" Then
chkName.Value = 1
Else
chkName.Value = 0
End If

If ReadINI("LanLauncher", "Auto", App.Path & "\iw4nick.ini") = "" Then
chkAuto.Value = 1
ElseIf ReadINI("LanLauncher", "Auto", App.Path & "\iw4nick.ini") = "1" Then
chkAuto.Value = 1
Else
chkAuto.Value = 0
End If

If ReadINI("LanLauncher", "NoIP", App.Path & "\iw4nick.ini") = "" Then
chkNoIP.Value = 0
ElseIf ReadINI("LanLauncher", "NoIP", App.Path & "\iw4nick.ini") = "1" Then
chkNoIP.Value = 1
Else
chkNoIP.Value = 0
End If

If ReadINI("LanLauncher", "Save", App.Path & "\iw4nick.ini") = "" Then
chkSave.Value = 1
ElseIf ReadINI("LanLauncher", "Save", App.Path & "\iw4nick.ini") = "1" Then
chkSave.Value = 1
Else
chkSave.Value = 0
End If

If ReadINI("LanLauncher", "AutoExit", App.Path & "\iw4nick.ini") = "" Then
chkExit.Value = 0
ElseIf ReadINI("LanLauncher", "AutoExit", App.Path & "\iw4nick.ini") = "1" Then
chkExit.Value = 1
Else
chkExit.Value = 0
End If

If ReadINI("LanLauncher", "RemOffline", App.Path & "\iw4nick.ini") = "" Then
chkOffline.Value = 0
ElseIf ReadINI("LanLauncher", "RemOffline", App.Path & "\iw4nick.ini") = "1" Then
chkOffline.Value = 1
Else
chkOffline.Value = 0
End If

'If ReadINI("LanLauncher", "Bar", App.Path & "\iw4nick.ini") = "" Then
'chkSave.Value = 1
'ElseIf ReadINI("LanLauncher", "Bar", App.Path & "\iw4nick.ini") = "1" Then
'chkBar.Value = 1
'Else
'chkBar.Value = 0
'End If

If ReadINI("LanLauncher", "Opacity", App.Path & "\iw4nick.ini") = "" Then
hsbOpacity.Value = 255
Else
hsbOpacity.Value = ReadINI("LanLauncher", "Opacity", App.Path & "\iw4nick.ini")
End If
 
If txtServer.Text = "" Then
txtServer.Text = "lulzy"
'Log ""
Log "## You have no server set, defaulting to Brad's LAN server!"
End If

If txtName.Text = "" Then txtName.Text = "Player"
If txtName.Text = "Player" Then
    'Log ""
    Log "## Your name is not set! Click the Settings button to change it!"
End If

'End If
Saved = True
cmdSave.Enabled = False
If ReadINI("LanLauncher", "Auto", App.Path & "\iw4nick.ini") = "" Then cmdSave_Click
Client(0).Name = NormalizeName(txtName.Text): ClientName = Client(0).Name
End Sub

Private Sub Update()
Dim IP As String
If Not chkUpdates.Value = vbChecked Then Exit Sub
'Log "## Starting update check.."
IP = ReadINI("LanLauncher", "UPDATE1", App.Path & "\iw4nick.ini")
If IP = "" Then IP = "www.tehwez.com"
cUpdate = True
tmrCheck.Enabled = False
If CheckFail(False) = True Then
    Log "## Update failed 3 times in a row. Trying alt update server.."
    IP = "dl.dropbox.com"
End If
With sckTCP
.Close 'closes possible open winsock
.Protocol = sckTCPProtocol
.RemotePort = 80 'sets the port winsock will use, 80 is default for http
.RemoteHost = IP 'sets the host winsock will request the page from, also knows as the domain
.Connect 'connects winsock
End With

End Sub


'dl.dropbox.com/u/4275989/Version.txt
'tehwez.com/Dragoon/Version.txt

'Private Sub SizeMe()
'If IsShown = True Then max = 8035 Else max = 6350
'If IsShown = True Then min = 5040 Else min = 4700

'max = 6350
'max = 6650
'min = 4700
'min = 5060

'max = 6650
'min = 5060


'If frmMain.Height > max Then frmMain.Height = max: Exit Sub
'If frmMain.Height < min Then frmMain.Height = min

'End Sub

'Private Sub ShowBar(Show As Boolean)
'SizeMe
'End Sub

Private Sub Form_Unload(Cancel As Integer)
'Send "BYE" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar
SendToAll "BYE", ""
End
End Sub

Private Sub jcbutton1_Click()
writeini "LanLauncher", "IGNOREVER", Replace(lblVersion.Caption, ".", ""), App.Path & "\iw4nick.ini"
Log "## Version " & lblVersion.Caption & " ignored. Type /update to review it."
btnUClose_Click
End Sub

Private Sub jcbutton2_Click()
ShowRich False
End Sub

Private Sub lblAuto_Click()
If chkAuto.Value = 1 Then chkAuto.Value = 0 Else chkAuto.Value = 1
End Sub


Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
End Sub

Private Sub lblExit_Click()
If chkExit.Value = 1 Then chkExit.Value = 0 Else chkExit.Value = 1

End Sub

Private Sub lblinfo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Sub lblip_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub lblip_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtOver.Visible = True
'txtOver.SetFocus
If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
End Sub

Private Sub lblNames_Click()
If chkName.Value = 1 Then chkName.Value = 0 Else chkName.Value = 1
End Sub

Private Sub lblNoIP_Click()
If chkNoIP.Value = 1 Then chkNoIP.Value = 0 Else chkNoIP.Value = 1
End Sub

Private Sub lblOffline_Click()
If chkOffline.Value = 1 Then chkOffline.Value = 0 Else chkOffline.Value = 1
End Sub

Private Sub lblSave_Click()
If chkSave.Value = 1 Then chkSave.Value = 0 Else chkSave.Value = 1
End Sub

Private Sub lblupdate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub lblupdate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Sub lblUpdates_Click()
If chkUpdates.Value = 1 Then chkUpdates.Value = 0 Else chkUpdates.Value = 1
End Sub

Private Sub lblURL_Click()
If lblURL.Caption = "No Update Link" Then Exit Sub
ShellExecute 0, "open", lblURL.Caption, 0, 0, 1
'DoEvents
'SendToAll "ME", "is downloading a new LanLauncher update and will be right back! The new Version is from:" & lblURL.Caption
'DoEvents
'End
End Sub

Private Sub lblVer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub lblver_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
End Sub

Private Sub lblVer_Click()
Dim strVer As String
strVer = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
If lblVer.Caption = strVer Then
lblVer.Caption = "I Love You, Andrea <3"
Else
lblVer.Caption = strVer
End If
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub lblversion_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Function GetIndex(Lindex As Integer) As Integer
    For i = 0 To UBound(Client)
            If LCase(Client(i).ListIndex) = Lindex Then
                GetIndex = i
                Exit Function
                Else
            End If
        Next
    GetIndex = "-1"
End Function

Private Sub lstIPs_DblClick()
If Client(GetIndex(lstIPs.ListIndex)).Serving Then
    If Not txtServer.Text = Client(GetIndex(lstIPs.ListIndex)).IP Then
            txtServer.Text = ""
            txtServer.Text = Client(GetIndex(lstIPs.ListIndex)).IP
        Exit Sub
    ElseIf Client(GetIndex(lstIPs.ListIndex)).User = True Then
        If MsgBox("Would you like to connect to " & Client(GetIndex(lstIPs.ListIndex)).Name & "'s lobby?", vbYesNo, "Connect to lobby") = vbYes Then
            ConnectLobby Client(GetIndex(lstIPs.ListIndex)).IP
            Exit Sub
        End If
    End If
End If

If Client(GetIndex(lstIPs.ListIndex)).Serving Then
    If Client(GetIndex(lstIPs.ListIndex)).User = False Then
        txtServer.Text = ""
        txtServer.Text = Client(GetIndex(lstIPs.ListIndex)).IP
        Exit Sub
    End If
End If

If Client(GetIndex(lstIPs.ListIndex)).Serving Then
    If Not txtServer.Text = Client(GetIndex(lstIPs.ListIndex)).IP Then
            txtServer.Text = ""
            txtServer.Text = Client(GetIndex(lstIPs.ListIndex)).IP
        Exit Sub
    Else
        If MsgBox("Would you like to connect to " & Client(GetIndex(lstIPs.ListIndex)).Name & "'s lobby?", vbYesNo, "Connect to lobby") = vbYes Then
            ConnectLobby Client(GetIndex(lstIPs.ListIndex)).IP
            Exit Sub
        End If
    End If
Else
    If MsgBox("Would you like to connect to " & Client(GetIndex(lstIPs.ListIndex)).Name & "'s lobby?", vbYesNo, "Connect to lobby") = vbYes Then
    ConnectLobby Client(GetIndex(lstIPs.ListIndex)).IP
    Exit Sub
    End If
End If

End Sub

Private Sub lstIPs_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
     Dim nCurRow As Integer
     
     If Button = 2 Then
          'Determine if mouse is on an item
          nCurRow = Int(Y / Picture1.TextHeight("A")) + lstIPs.TopIndex
          If nCurRow <= lstIPs.ListCount - 1 Then
               'Mouse is over an existing row in the listbox
               lstIPs.ListIndex = nCurRow
               mnuLobby.Visible = True
               mnuServer.Visible = True
               If Not Client(GetIndex(lstIPs.ListIndex)).Serving Then mnuServer.Visible = False Else mnuServer.Visible = True
               If Not Client(GetIndex(lstIPs.ListIndex)).User Then mnuLobby.Visible = False Else mnuLobby.Visible = True
          End If
     PopupMenu mnuOptions
     End If
     
     
     
End Sub

Private Sub lstIPs_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtOver.Visible = False
End Sub

Private Sub mnuLobby_Click()

ConnectLobby Client(GetIndex(lstIPs.ListIndex)).IP

End Sub

Private Sub ConnectLobby(IP As String)
If IP = "" Then Exit Sub
If Dir(App.Path & "\LANConnecter.exe") <> "" Then
    If Dir(App.Path & "\iw4mp.dat") <> "" Then
        Log "Attempting to connect to lobby " & IP & "..."
        'Call Shell("lan://connect/" & IP & ":28960")
        If Not Shell(App.Path & "\LANConnecter.exe" & " lan://connect/" & IP & ":28960", vbNormalFocus) > 0 Then
            Log "## Error! LANConnecter did not start successfully!"
        End If
    Else
        Log "## LANConnecter found, but no iw4mp.dat!"
    End If
Else
    Log "## LANConnecter was not found in this directory! Please redownload it."
End If
    
End Sub

Private Sub mnuServer_Click()
txtServer.Text = ""
txtServer.Text = Client(GetIndex(lstIPs.ListIndex)).IP

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        OldX = x
        OldY = Y
    End If
End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Me.Move Left + (x - OldX), Top + (Y - OldY)
txtOver.Visible = False
End Sub

Private Sub Picture2_Click(Index As Integer)
'ShowRich True
Select Case Index
Case Is = "0"
txtName.SelText = "^" & Index & txtName.SelText
Case Is = "1"
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 2
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 3
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 4
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 5
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 6
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 7
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 8
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 9
txtName.SelText = "^" & Index & txtName.SelText
Case Is = 0
txtName.SelText = "^" & Index & txtName.SelText
End Select
txtName.SetFocus
End Sub

Private Sub sckStatus_Close()
On Error Resume Next
SetServer False, False, True
tmrStatus.Enabled = True
cmdStart.Caption = "Connect to " & sckStatus.RemoteHost
End Sub

Private Sub sckStatus_Connect()
On Error Resume Next
SetServer True

If Not sckStatus.RemoteHost = LocalIP Then
    If Not sckStatus.RemoteHost = "127.0.0.1" Then
        If Not LCase(sckStatus.RemoteHost) = "localhost" Then
        Call AddIP(sckStatus.RemoteHost, False, "Shared IP")
        SendToAll "OSERVER", sckStatus.RemoteHost
        'Log sckStatus.RemoteHost & "OSERVER1"
        End If
    End If
End If
CheckList sckStatus.RemoteHost, 1
cmdStart.Caption = "Connect to " & sckStatus.RemoteHost
End Sub

Private Sub sckStatus_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
SetServer False, False, True
CheckSrv
tmrStatus.Enabled = True
cmdStart.Caption = "Connect to " & sckStatus.RemoteHost

ChooseServer

End Sub

Private Sub sckTCP_Close()
'On Error Resume Next
If cUpdate = True Then ParseUpdate True
Server = False
'Send "SERVER" & SplitChar & LocalIP & SplitChar & "0"
End Sub

Private Sub sckTCP_Connect()
'On Error Resume Next
If cUpdate = True Then

Dim URL As String
Dim IP As String
If CheckFail(False) = True Then
    IP = "dl.dropbox.com"
    URL = "http://dl.dropbox.com/u/4275989/Version.txt"
Else
    IP = ReadINI("LanLauncher", "UPDATE1", App.Path & "\iw4nick.ini")
    URL = ReadINI("LanLauncher", "UPDATE2", App.Path & "\iw4nick.ini")
End If


If URL = "" Then URL = "http://tehwez.com/Dragoon/version.php?ver=" & App.Major & "." & App.Minor & "." & App.Revision & "&rnd=" & GetSerialNumber
If IP = "" Then IP = "www.tehwez.com"
    sckTCP.SendData "GET " & URL & " HTTP/1.1" & vbCrLf & "Host: " & IP & vbCrLf & vbCrLf
    Exit Sub
End If

Server = True
Send "SERVER" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "1"
CheckList LocalIP, "1"
'CheckSrv
End Sub

Private Sub sckTCP_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String 'dims data as string
Dim strVersion As String
Dim strURL As String
Dim strInfo As String
sckTCP.GetData data 'gets the data as it arrived and sets "data" as the string
'If cUpdate = False Then MsgBox data 'Exit Sub

Dim strSplit() As String
If InStr(1, UCase(data), "404 NOT FOUND") > 0 Then
    ParseUpdate True
   Exit Sub
End If
If Not InStr(1, data, "|-|") > 2 Then Exit Sub
strSplit = Split(data, "|-|")
strVersion = strSplit(1)
strURL = strSplit(2)
strInfo = strSplit(3)
ParseUpdate False, strVersion, strURL, strInfo
End Sub

Private Sub sckTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

If cUpdate = True Then ParseUpdate (True): Exit Sub
Server = False
'Send "SERVER" & SplitChar & LocalIP & SplitChar & "0"
CheckList LocalIP, "0"
sckTCP.Close
End Sub

Public Sub ParseUpdate(Error As Boolean, Optional Version As String, Optional URL As String, Optional Info As String)
If cUpdate = False Then Exit Sub
If Error = True Then
    cUpdate = False
    Log vbCrLf & "## Error checking for updates.. The LanLauncher is either several years old, you've got no internet, or there is something wrong with the server."
    CheckFail (True)
    sckTCP.Close
    tmrCheck.Enabled = True
    sckTCP.RemotePort = 13000
    Exit Sub
End If
lblVersion.Caption = Version
If Not URL = "" Then lblURL.Caption = URL Else lblURL.Caption = "No Update Link"

lblInfo.Caption = Replace(Info, "/$n/", vbCrLf)
Version = Replace(Version, ".", "")
If Not IsNumeric(Version) Then
Log "## Update error: Numeric field invalid, cancelling.."
Else
Dim Ver2 As String
If App.Minor < 10 Then Ver2 = App.Major & "0" & App.Minor & App.Revision Else Ver2 = App.Major & App.Minor & App.Revision
If Ver2 < Version Then
        If Not ReadINI("LanLauncher", "IGNOREVER", App.Path & "/iw4nick.ini") = Version Then
                'Log vbCrLf & "## There is an update! The URL is: " & URL
                frmUpdate.Visible = True
            Else
                Log "## Version " & Version & " ignored. Type /update to review it."
        End If
        'Log "The update reason is: " & Info
    ElseIf Ver2 > Version Then
        Log "## Hey guy, you're using BETA version " & Ver2 & " while the public still suffers with version " & Version & "! lulz"
    Else
        Log "## You have the newest version.  (" & Version & ")"
        'Log Info
    End If
End If

cUpdate = False
sckTCP.Close
sckTCP.RemotePort = 13000
tmrCheck.Enabled = True
End Sub

Private Function CheckFail(HasFailed As Boolean) As Boolean
Dim intFails As String
    On Error Resume Next

intFails = ReadINI("LanLauncher", "FAILS", App.Path & "/iw4nick.ini")

If intFails = "" Then intFails = "0"

If HasFailed = True Then intFails = intFails + 1

If intFails >= 3 Then CheckFail = True Else CheckFail = False

writeini "LanLauncher", "FAILS", intFails, App.Path & "/iw4nick.ini"

    
End Function

Private Sub sckUDP_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strData As String
sckUDP.GetData strData
If Len(strData) <= 3 Then Exit Sub
Log2 "<< " & strData
Call Parse(strData)

End Sub

Private Sub ChangeName(IPData As String, UName As String)
Dim NoIP As String
Dim Srv As String
Dim cIndex As Integer

If chkNoIP.Value = 1 Then Srv = Srv Else Srv = "Server: "
                i = SearchArray(IPData)
                cIndex = Client(i).ListIndex
                If ShowLast = True Then LSeen = " (" & Client(i).LastSeen & ")"
                    If i >= 0 Then
                        If Client(i).Name = "Shared IP" Then
                            Client(i).Name = UName
                        End If
                            If chkNoIP.Value = 0 Then NoIP = "" Else NoIP = "@" & IPData
                                If Client(i).Serving = True Then
                                    lstIPs.List(cIndex) = Srv & UName & NoIP & LSeen
                                Else
                                    lstIPs.List(cIndex) = UName & NoIP & LSeen
                                End If
                        Client(i).Name = UName
                        Client(i).User = True
                        Client(i).LastSeen = 0
                        Client(i).TimedOut = False
                    End If
        UpdateList
End Sub

Private Sub UpdateList()
Dim CurrIP As String
Dim Srv As String
Dim cIndex As Integer
Dim Remove As Boolean
If chkNoIP.Value = 1 Then Srv = "Srv: " Else Srv = "Server: "
Client(0).LastSeen = 0
Client(0).TimedOut = False
For i = 0 To UBound(Client)
If ShowLast = True Then LSeen = " (" & Client(i).LastSeen & ")"
CurrIP = "@" & Client(i).IP
If chkNoIP.Value = 0 Then CurrIP = ""
    If Client(i).User = True Then
    cIndex = Client(i).ListIndex
        'If Not Client(I).IP = LocalIP Then
            If Client(i).TimedOut = True Then
                    lstIPs.List(cIndex) = "Offline: " & Client(i).Name & CurrIP
                    Remove = True
            Else
                If Client(i).Serving = True Then
                    If Not lstIPs.List(cIndex) = Srv & Client(i).Name & CurrIP & LSeen Then
                    lstIPs.List(cIndex) = Srv & Client(i).Name & CurrIP & LSeen
    
                    End If
                Else
                    If Not lstIPs.List(cIndex) = Client(i).Name & CurrIP & LSeen Then
                    lstIPs.List(cIndex) = Client(i).Name & CurrIP & LSeen
    
                    End If
                End If
            End If
        'End If
    End If
Next

If chkOffline.Value = vbChecked And Remove = True Then RemoveOffline

End Sub

Private Sub CheckList(IP As String, Status As String)
On Error Resume Next
    Dim NoIP As String
    Dim i As Integer
    Dim Srv As String
    Dim cIndex As Integer
            i = SearchArray(IP)
            cIndex = Client(i).ListIndex
            If ShowLast = True Then LSeen = " (" & Client(i).LastSeen & ")"
            Client(i).LastSeen = 0
            If chkNoIP.Value = 0 Then NoIP = "" Else NoIP = "@" & IP
            If chkNoIP.Value = 1 Then Srv = "Srv: " Else Srv = "Server: "
            If Not i = "-1" Then
                If Status = "1" Then Client(i).Serving = True Else Client(i).Serving = False
                If Client(i).Name = "Shared IP" Then Exit Sub
                
                If Client(i).First = False Then
                'lblCon.Caption = "Detected Users: " & UBound(Client) + 1
                    If Status = 1 Then
                        lstIPs.List(cIndex) = Srv & Client(i).Name & NoIP & LSeen
                        Log Client(i).Name & " is now hosting a server."
                        ChooseServer
                        Else
                        lstIPs.List(cIndex) = Client(i).Name & NoIP & LSeen
                    End If
                    Client(i).First = True
                    Last(i).Serving = Client(i).Serving
                End If
                If Last(i).Serving = Client(i).Serving Then
                    Exit Sub
                Else
                    If Status = 1 Then
                        Log Client(i).Name & " is now hosting a server."
                        lstIPs.List(cIndex) = Srv & Client(i).Name & NoIP & LSeen
                        ChooseServer
                        'If chkAuto.Value = Checked Then txtServer.Text = IP
                        Else
                        lstIPs.List(i) = Client(i).Name & NoIP & LSeen
                    End If
                End If
            'If Status = "1" Then Client(i).Serving = True Else Client(i).Serving = False
            Last(i).Serving = Client(i).Serving
'            lblCon.Caption = "Detected Users: " & UBound(Client)
            End If
            
            Client(i).LastSeen = 0
            
End Sub

Private Sub ChooseServer()
On Error Resume Next
If chkAuto.Value = Unchecked Then Exit Sub

Dim i As Integer
Dim Top, IP, data() As String
For i = 0 To UBound(Client)
    If Client(i).Serving Then
        If IsNumeric(Replace(Client(i).IP, ".", "")) Then
        data = Split(Client(i).IP, ".")
        If data(3) > 150 Then data(3) = data(3) / 7
        If data(3) > 75 Then data(3) = data(3) / 5
            If Top <= data(3) Then
            Top = data(3)
            IP = Client(i).IP
            End If
        ElseIf Top = "" Then
        Top = Len(Client(i).IP)
        IP = Client(i).IP
        Else
            If Top <= Len(Client(i).IP) Then
            Top = Len(Client(i).IP)
            IP = Client(i).IP
            End If
        End If
    End If
Next
If Not IP = "" Then txtServer.Text = IP
End Sub

Public Function CheckSrv() As Boolean
On Error Resume Next
If sckTCP.State = sckConnected Then
    CheckSrv = True
    Exit Function
        ElseIf sckTCP.State = sckConnecting Then
        CheckSrv = False
        Exit Function
Else
    sckTCP.Close
    'sckTCP.LocalPort = 43211
    sckTCP.RemotePort = "13000"
    sckTCP.RemoteHost = LocalIP
    sckTCP.Connect
    CheckSrv = False
    Exit Function
End If
End Function

    Private Function SearchArray(IP As String) As Integer
        If Len(IP) < 1 Then
        SearchArray = "-1"
        Exit Function
        End If

    If Client(0).IP = "" Then
        SearchArray = "-1"
        Exit Function
        Else
        For i = 0 To UBound(Client)
        If Not Client(i).IP = "" Then
            If LCase(Client(i).IP) = LCase(IP) Then
                SearchArray = i
                'Client(I).LastSeen = 1
                Exit Function
                'Exit For
                Else
            End If
        End If
        Next
    End If
    SearchArray = "-1"
    End Function

Private Sub AddNonUser(IP As String, User As Boolean, Name As String, Optional i As Integer)
Client(i).IP = IP
Client(i).User = User
Client(i).Name = Name
Client(i).Serving = True
lstIPs.AddItem Name & ": " & IP
Client(i).ListIndex = lstIPs.NewIndex
ChooseServer
End Sub

Private Function CheckVersion(IP As String, Version As String) As Boolean
Dim i As Integer
Dim cVer As String
cVer = App.Major & App.Minor & App.Revision
i = SearchArray(IP)
    If i >= 1 Then
    If IsNumeric(Version) = False Then Client(i).Version = "h4x0r": Exit Function
    If Client(i).Version = "" Or Client(i).Version = "0" Then
        If Version < cVer Then
            'Log "## " & Client(i).Name & "'s version is old!"
            Client(i).Version = Version
        End If
        If Version > cVer Then
            If Client(0).First = False Then
                Log "## Someone has a newer version than you! Please enable update checking or download here: http://dl.dropbox.com/u/4275989/LAN.html"
                Client(0).First = True
            End If
                    Client(i).Version = Version
                    
        End If
        If Version = cVer Then
            'Log "## You have the same version as " & Client(i).Name
            Client(i).Version = Version
        End If
    End If
    
End If
End Function

Private Function AddIP(IP As String, User As Boolean, Optional Name As String) As Boolean
Dim ii As Integer
Dim NoIP As String
Dim Srv As String
ii = SearchArray(IP)
        If ii = "-1" Then
            Dim i As Integer
            For i = 0 To UBound(Client)
                
                If Client(i).IP = "" Then
                        If Name = "Shared IP" Then
                            Call AddNonUser(IP, User, Name, i)
                            AddIP = True
                            Exit Function
                        End If
                    Client(i).IP = IP
                    Client(i).Name = Name
                    Last(i).IP = IP
                    Client(i).User = User
                    
                    If chkNoIP.Value = 0 Then NoIP = "" Else NoIP = "@" & Client(i).IP
                    If chkNoIP.Value = 1 Then Srv = "Srv: " Else Srv = "Server: "
                    If User = True Then
                        If Client(i).Serving = True Then
                        lstIPs.AddItem Srv & Client(i).Name & NoIP & LSeen
                        Else
                        lstIPs.AddItem Client(i).Name & NoIP & LSeen
                        End If
                        Client(i).ListIndex = lstIPs.NewIndex
                    End If
                    If User = True Then Client(i).LastSeen = 0
                    
                    AddIP = True
                    'client(i).
                    Exit Function
                    'Exit For
                    
                End If
            Next
                ReDim Preserve Client(UBound(Client) + 1)
                ReDim Preserve Last(UBound(Last) + 1)
                Call AddIP(IP, User, Name)
                AddIP = True
                Exit Function
            Else
            If Client(ii).Name = "Shared IP" Then
                If Not Name = "" Then
                    Client(ii).Name = Name
                End If
            End If
            Client(ii).User = User
            Client(ii).LastSeen = "0"
            Client(ii).TimedOut = False
        End If
AddIP = False
    End Function

Private Sub Send(Text As String, Optional IP As String)
On Error GoTo OHSNAP
If LocalIP = "127.0.0.1" Then LocalIP = sckTCP.LocalIP
Log1 ">>" & IP & " " & Text
If Not IP = "" Then
    sckUDP.RemoteHost = IP
    sckUDP.SendData Text
    sckUDP.RemoteHost = "255.255.255.255"
    Exit Sub
End If
sckUDP.SendData Text
Exit Sub
OHSNAP:
If Not Error = "" Then
    If Err.Number = "10048" Then
        Log "## Port in use! Do you have multiple clients open? Type /reset to try again."
    Else
    Log Err.Number & " - " & Err.Description
    End If
Resume Next
End If
End Sub

Private Sub SendToAll(Header As String, Text As String)
On Error GoTo SckErr
If LocalIP = "127.0.0.1" Then LocalIP = sckTCP.LocalIP
Log1 ">> " & Header & SplitChar & LocalIP & SplitChar & Text
For i = 0 To UBound(Client)
If Not Client(i).User = False Then
    If Not Client(i).IP = "" Then
        If Not Client(i).IP = LocalIP Then
            If IsNumeric(Replace(Client(i).IP, ".", "")) Then
                sckUDP2.RemoteHost = Client(i).IP
                sckUDP2.SendData Header & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & Text & SplitChar
            End If
        End If
    End If
End If
Next
SckErr:
If Err.Number = 10065 Then
Log "Error 10065, No route to host. Connectivity problems. Check your cables!"
Exit Sub
End If
End Sub

Private Sub sckUDP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Log "## Error: " & Number & ": " & Description & " - " & Source
End Sub

Private Sub Timer1_Timer()
UpdateList

End Sub

Private Sub tmrBroadcast_Timer()
Send "PING1" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "0"
End Sub

Private Sub tmrName_Timer()
If chkSave.Value = 1 Then
cmdSave_Click
Else
Client(0).Name = NormalizeName(txtName.Text): ClientName = Client(0).Name
End If
If Len(Trim(NormalizeName(txtName.Text))) < 3 Then tmrName.Enabled = False: Exit Sub
SendName
ChangeName LocalIP, ClientName
tmrName.Enabled = False
End Sub

Private Sub SendName()
SendToAll "NCHANGE", ClientName
End Sub

Private Sub tmrCheck_Timer()
If CheckSrv = True Then
SendToAll "SERVER", "1"
'Send "SERVER" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "1"
CheckList LocalIP, "1"
Else
SendToAll "SERVER", "0"
'Send "SERVER" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "0"
CheckList LocalIP, "0"
End If
End Sub

Private Sub tmrSpam_Timer()
tmrSpam.Enabled = False

End Sub

Private Sub tmrStatus_Timer()
Status
tmrStatus.Enabled = False
End Sub

Private Sub Status()
'On Error Resume Next
With sckStatus
.Close
'SetServer False, True
'.LocalPort = 43215
.RemotePort = "13000"
.RemoteHost = Trim(txtServer.Text)
.Connect
End With

End Sub

Private Sub txtLog_Click()
'txtMSG.SetFocus
End Sub

Private Sub txtLog_KeyDown(KeyCode As Integer, Shift As Integer)
txtMSG.SetFocus
DoEvents
'txtMSG.Text = txtMSG.Text & Chr(KeyCode)
txtMSG.SelStart = 65535
End Sub

Private Sub txtLog_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
txtOver.Visible = False
End Sub

Private Sub txtLog2_Change()
If Not blDebug = True Then Exit Sub
txtLog2.SelStart = 65535
If Len(txtLog2.Text) > 65534 Then txtLog2.Text = ""
End Sub

Private Sub txtMSG_GotFocus()
If Trim(txtMSG.Text) = "Type here and press enter to send message.." Then txtMSG.Text = ""
End Sub

Public Sub ResetPort()
On Error GoTo OHMAN
sckUDP.Close
sckUDP2.Close
DoEvents
sckUDP.Protocol = sckUDPProtocol
sckUDP.RemoteHost = "255.255.255.255"
sckUDP.RemotePort = "43210"
sckUDP.Bind 43210
sckUDP2.Protocol = sckUDPProtocol
sckUDP2.RemoteHost = "255.255.255.255"
sckUDP2.RemotePort = "43210"
Send "PING1" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & "0"
DoEvents
'Log "## Success! Port re-opened!"
Exit Sub

OHMAN:
If Not Error = "" Then
    If Err.Number = "10048" Then
        Log "## Port still in use! Please close the other program occupying this port!"
    Else
    Log Err.Number & " - " & Err.Description
    End If
Exit Sub
End If
End Sub

Private Sub ListVersion()
Dim ii As Integer
'Log "Usercount: " & i
Log "User Name" & vbTab & "IP Address" & vbTab & "Version" & vbTab & "Hosting"
For i = 0 To UBound(Client)
'    if not client(i).
    If Client(i).User = True Then
        If Len(Client(i).Name) < 9 Then
        Log Client(i).Name & vbTab & vbTab & Client(i).IP & vbTab & Client(i).Version & vbTab & Client(i).Serving
        Else
        Log Client(i).Name & vbTab & Client(i).IP & vbTab & Client(i).Version & vbTab & Client(i).Serving
        End If
        ii = ii + 1
    Else
        If Not Client(i).Name = "" Then Log "Shared IP: " & vbTab & Client(i).IP
    End If
Next
Log "Usercount: " & ii
'Log ""
'Log "Shared IPs:"
'For i = 0 To UBound(Client)
'    If Client(i).User = False Then If Not Client(i).Name = "" Then Log "Shared IP:" & vbTab & Client(i).IP
'    End If
'Next


End Sub

Private Sub TxtMSG_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 38 Then
txtMSG.Text = LastMessage
txtMSG.SelStart = Len(LastMessage)
Exit Sub
End If

Dim strSpam As String
If KeyCode = 13 Then
If Trim(txtMSG.Text) = "" Then Exit Sub
LastMessage = txtMSG.Text


If Left(LCase(txtMSG.Text), 6) = "/reset" Then
ResetPort
txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 5) = "/open" Then
Dim x As Integer
On Error Resume Next
x = Shell("EXPLORER " & App.Path, vbNormalFocus)
If Not x = 0 Then Log "## Opening folder.." Else: Log ("Errar, brah")
txtMSG.Text = ""
Exit Sub
End If


If Left(LCase(txtMSG.Text), 4) = "/cls" Then
txtLog.Text = ""
txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 7) = "/clear2" Then
        RemoveNonUser
        txtMSG.Text = ""
Exit Sub
        'lstIPs.Clear
        'lstIPs.AddItem ClientName
        'Log "## Witing for discovery responses from others.."
End If

If Left(LCase(txtMSG.Text), 6) = "/clear" Then
        RemoveOffline
        txtMSG.Text = ""
Exit Sub
        'lstIPs.Clear
        'lstIPs.AddItem ClientName
        'Log "## Witing for discovery responses from others.."
Exit Sub
End If



If Left(LCase(txtMSG.Text), 8) = "/version" Then
ListVersion
txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 5) = "/list" Then
ListVersion
txtMSG.Text = ""
Exit Sub
End If


If Left(LCase(txtMSG.Text), 6) = "/debug" Then
    If InputBox("", "", "") = "poop" Then
            blDebug = True
            FWidth = 12570
            Max = Max + 500
            Min = Min + 500
            txtLog1.Visible = True
            txtLog2.Visible = True
            Log "## I herd u liekd dbug."
            frmMain.Caption = "DEBUG, BRO"
            cmdSettings.Caption = "Settings"
            cmdSettings_Click
        Else
            blDebug = False
            Log "## Whoops, you don't need that anyway =)"
    End If
txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 5) = "/help" Then
Log "Right click or double click on a user/server in the list to connect to as a server or to join their lobby. Check 'Auto-Select Server' in the Settings to automatically set the server when one is found." & vbCrLf & "Type /help for a short list of commands."
    Log "Help list:"
    Log "/me - Emote your words."
    Log "/cls - Clear the chat text."
    Log "/list - Print User and IP list to the Chat."
    Log "/to - Display the time since last ping from users."
    Log "/time - display local machine time."
    Log "/clear - Removes offline users from the list."
    Log "/clear2 - Removes all Shared IPs from the list."
    txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 5) = "/time" Then
If tmrSpam.Enabled = True Then
txtMSG.Text = ""
Log "Too soon!"
Exit Sub
Else
tmrSpam.Enabled = True
End If
SendToAll "TIME", Now
Log "It is " & Now
txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 7) = "/update" Then
    writeini "LanLauncher", "IGNOREVER", "", App.Path & "\iw4nick.ini"
    Update
    DoEvents
    txtMSG.Text = ""
    Exit Sub
End If


If Left(LCase(txtMSG.Text), 3) = "/to" Then
ShowLast = True
Timer1.Enabled = True
Log "TO enabled: There will be a number next to users names in the list on the right. 0 means everything is GOOD."
txtMSG.Text = ""
Exit Sub
End If

If Left(LCase(txtMSG.Text), 4) = "/me " Then
    
    SendToAll "ME", Mid(txtMSG.Text, 4)
    
        If chkName.Value = Checked Then
        Log ClientName & " " & Mid(txtMSG.Text, 5)
    Else
        Log LocalIP & " " & Mid(txtMSG.Text, 5)
    End If
Else
    If Left(LCase(txtMSG.Text), 1) = "/" Then
    Log ("Command not recognized!")
    txtMSG.Text = ""
    Exit Sub
    End If
    SendToAll "MSG", txtMSG.Text
    
    If chkName.Value = Checked Then
        Log ClientName & ": " & txtMSG.Text
    Else
        Log LocalIP & ": " & txtMSG.Text
    End If
End If
txtMSG.Text = ""
End If
End Sub

Private Sub tmrPing_Timer()
Dim cVer As String
SendToAll "PING2", App.Major & App.Minor & App.Revision '"PING" 'Send "DISCOVER" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar
LastSeen
'tmrStatus.Enabled = True
End Sub

Private Sub LastSeen()
Dim cIndex As Integer
Dim i As Integer
Dim Remove As Boolean
Client(0).LastSeen = 0
Client(0).TimedOut = False
For i = 0 To UBound(Client)
cIndex = Client(i).ListIndex
If Client(i).User = True Then

    'If Not Client(I).IP = LocalIP Then
        If Client(i).LastSeen >= 3 Then
            If Client(i).TimedOut = False Then
                Log Client(i).Name & " has timed out."
                lstIPs.List(cIndex) = GetName(i)
                Client(i).TimedOut = True
                Remove = True
            End If
        Else
            If Client(i).Serving = True Then
                If Not lstIPs.List(cIndex) = GetName(i) Then
                lstIPs.List(cIndex) = GetName(i)
                'Client(I).LastSeen = Client(I).LastSeen + 1
                'Client(I).TimedOut = False
                End If
            Else
                If Not lstIPs.List(cIndex) = GetName(i) Then
                lstIPs.List(cIndex) = GetName(i)
                'Client(I).LastSeen = Client(I).LastSeen + 1
                'Client(I).TimedOut = False
                End If
            End If
        'Else
            Client(i).LastSeen = Client(i).LastSeen + 1
            Client(i).TimedOut = False
        End If
    'End If
End If
Next

If chkOffline.Value = vbChecked And Remove = True Then RemoveOffline

End Sub

Public Function GetName(i As Integer) As String
Dim CurrIP As String
Dim Srv As String
Dim cIndex As Integer
cIndex = Client(i).ListIndex
If chkNoIP.Value = 1 Then Srv = "Srv: " Else Srv = "Server: "
If ShowLast = True Then LSeen = " (" & Client(i).LastSeen & ")"
CurrIP = "@" & Client(i).IP
If chkNoIP.Value = 0 Then CurrIP = ""
If blDebug = True Then Log "DEB u:" & Client(i).User & " n:" & Client(i).Name & " s:" & Client(i).Serving & " v:" & Client(i).Version
    If Client(i).User = False Then
        If Not Client(i).Name = "" Then
            GetName = Client(i).Name & ": " & Client(i).IP
            Exit Function
        Else
            Exit Function
        End If
    End If

If Client(i).LastSeen >= 3 Then Client(i).TimedOut = True

If Client(i).TimedOut = True Then
    GetName = "Offline: " & Client(i).Name & CurrIP
Else
    If Client(i).Serving = True Then
        GetName = Srv & Client(i).Name & CurrIP & LSeen
    Else
        GetName = Client(i).Name & CurrIP & LSeen
    End If
End If
    
End Function

Private Sub Colors()
Picture2(0).BackColor = vbBlack
Picture2(1).BackColor = vbRed
Picture2(2).BackColor = vbGreen
Picture2(3).BackColor = vbYellow
Picture2(4).BackColor = vbBlue
Picture2(5).BackColor = vbCyan
Picture2(6).BackColor = &HFF00FF
Picture2(7).BackColor = vbWhite
Picture2(8).BackColor = &HC0C0C0
Picture2(9).BackColor = &HE0E0E0

lblErr.ForeColor = RGB(51, 153, 255)
lblServer.ForeColor = RGB(51, 153, 255)
lblInfo.ForeColor = RGB(51, 153, 255)
lblURL.ForeColor = RGB(51, 153, 255)
lblUpdate.ForeColor = RGB(51, 153, 255)
lblVersion.ForeColor = RGB(51, 153, 255)


End Sub

Private Sub Log(Text As String)
msgTotal = msgTotal + Len(Text)
If msgTotal > 60000 Then
txtLog.Text = ""
msgTotal = 0
End If
txtLog.Text = txtLog.Text & vbCrLf & Text
End Sub

Private Sub Log1(Text As String)
If Not blDebug = True Then Exit Sub
txtLog1.Text = txtLog1.Text & vbCrLf & Text
End Sub

Private Sub Log2(Text As String)
If Not blDebug = True Then Exit Sub
txtLog2.Text = txtLog2.Text & vbCrLf & Text
End Sub

Private Sub txtLog_Change()
txtLog.SelStart = 65535

End Sub

Private Sub txtLog1_Change()
If Not blDebug = True Then Exit Sub
txtLog1.SelStart = 65535

If Len(txtLog1.Text) > 65534 Then txtLog1.Text = ""
End Sub

Private Sub txtMSG_LostFocus()
If txtMSG.Text = "Type here and press enter to send message.." Or txtMSG.Text = "" Then
txtMSG.Text = "Type here and press enter to send message.."
End If
End Sub
Private Sub txtName_Change()
Saved = False
cmdSave.Enabled = True
If Trim(Len(NormalizeName(txtName.Text))) < 3 Then
tmrName.Enabled = False
lblErr.ForeColor = vbRed
lblErr.Alignment = 2
lblErr.Caption = "Too short!"
tmrName.Enabled = False
Exit Sub
End If
lblErr.ForeColor = RGB(51, 153, 255)
lblErr.Alignment = 0
lblErr.FontBold = True
lblErr.Caption = "Nickname"
tmrName.Enabled = False
tmrName.Enabled = True
Code2Color txtName, rtxtName
End Sub

Private Sub SetCaption(Caption As String)
If isshown = True Then
frmMain.Caption = Caption
Else
lblCaption.Caption = Caption
End If
End Sub

Private Sub SetServer(Optional Connected As Boolean = False, Optional Connecting As Boolean = False, Optional Offline As Boolean = True)
If Connected = True Then
IsOnline = True
'imgOnline.Visible = True
'imgOffline.Visible = False
'imgTrying.Visible = False
SetCaption "LanLauncher - Server Online"
If Not Picture1.Picture = pOnline.Picture Then Picture1.Picture = pOnline.Picture
lblIP.ForeColor = vbGreen
lblCaption.ForeColor = lblIP.ForeColor
lblVer.ForeColor = lblIP.ForeColor
Exit Sub
End If

If Connecting = True Then
IsOnline = False
'imgOffline.Visible = False
'imgOnline.Visible = False
'imgTrying.Visible = True
SetCaption "LanLauncher - Testing Server.."
If Not Picture1.Picture = pNormal.Picture Then Picture1.Picture = pNormal.Picture
lblIP.ForeColor = RGB(51, 153, 255)
'cmdStart.BackColor = &HFFD1AD
lblCaption.ForeColor = lblIP.ForeColor
lblVer.ForeColor = lblIP.ForeColor
Exit Sub
End If

If Offline = True Then
IsOnline = False
'imgOnline.Visible = False
'imgTrying.Visible = False
'imgOffline.Visible = True
If Not Picture1.Picture = pOffline.Picture Then Picture1.Picture = pOffline.Picture
lblIP.ForeColor = vbRed
lblCaption.ForeColor = lblIP.ForeColor
lblVer.ForeColor = lblIP.ForeColor


SetCaption "LanLauncher - Server Offline"
End If


End Sub

Public Function ShowRich(blRich As Boolean)
If blRich = True Then lblServer.Caption = "Example Nick:" Else lblServer.Caption = "Server"
frmColor.Visible = blRich
    'For i = 0 To Picture2.UBound
    '    Picture2(i).Visible = blRich
    'Next

rtxtName.Visible = blRich
If blRich = False Then txtServer.Visible = True Else txtServer.Visible = False
End Function

Private Sub txtName_GotFocus()
ShowRich True
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ShowRich False
If KeyAscii = 27 Then ShowRich False
End Sub

Private Sub txtName_LostFocus()
'ShowRich False
End Sub

Private Sub txtOver_Change()
tmrStatus.Enabled = False
sckStatus.Close
txtServer.Text = txtOver.Text
End Sub

Private Sub txtOver_GotFocus()
txtOver.Visible = True
End Sub

Private Sub txtServer_Change()
txtOver.Text = txtServer.Text
lblIP.Caption = txtServer.Text
If Trim(txtServer.Text) = "" Then
tmrStatus.Enabled = False
SetServer False, False, True
Exit Sub
End If
cmdSave.Enabled = True
SetServer False, True, False
Saved = False
tmrStatus.Enabled = True
If chkSave.Value = 1 Then cmdSave_Click
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
chkAuto.Value = 0
End Sub

Private Sub FixIndexes()
Dim i As Integer
Dim tUser As String
Dim newi As Integer
lstIPs.Clear
lstIPs.AddItem GetName(0)
For i = 1 To UBound(Client)
tUser = GetName(i)

If Not tUser = "" Then
    lstIPs.AddItem tUser
    Client(i).ListIndex = lstIPs.NewIndex
End If
Next




'If index > lstIPs.ListCount - 1 Then Exit Sub
   ' For i = Index + 1 To lstIPs.ListCount - 1
   '
        'Client(i).ListIndex = Client(i).ListIndex - 1

    'Next
    'lstIPs.RemoveItem Client(Index).ListIndex
    'Client(Index).ListIndex = "-1"
    
End Sub


Private Sub RemoveClient(Index As Integer)

With Client(Index)
.First = False
.IP = ""
.LastSeen = 10
.ListIndex = -1
.Serving = False
.TimedOut = False
.User = False
.Version = ""
.Name = ""
End With


'ii = SearchArray(IP)
'        If ii = "-1" Then

End Sub

Private Sub RemoveOffline()
Dim i As Integer
Dim ii As Integer
Dim DidWork As Boolean
'Log "Usercount: " & i
ii = UBound(Client)
For i = 0 To ii
'    if not client(i).
    If Client(i).User = True Then
                    
        If Client(i).TimedOut = True Then
            RemoveClient i
            DidWork = True
            'ii = ii - 1
        End If
    End If

Next
If DidWork Then Call FixIndexes
End Sub

Private Sub RemoveNonUser()
Dim i As Integer
'Log "Usercount: " & i
For i = 0 To UBound(Client)
'    if not client(i).
    If Client(i).User = False Then
            RemoveClient i
    End If
Next
Call FixIndexes
End Sub

Private Sub Parse(strData As String)
On Error Resume Next
        Dim data() As String
        Dim IPData As String
        Dim UName As String
        Dim i As Integer
        Dim cIndex As Integer
        If Len(strData) <= 9 Then Exit Sub
        data = Split(strData, SplitChar)
        IPData = data(1)
        UName = data(2)
        If IPData = LocalIP Then Exit Sub
        If IPData = "" Then Log "## Corrupt packet/hacker faggot on your network.": Exit Sub
        If UName = "" Then Log "## Corrupt packet detected, dropped."
        
        Select Case data(0)

            Case Is = "BYE"
            i = SearchArray(IPData)
                If i >= 0 Then
                cIndex = Client(i).ListIndex
                    If chkNoIP.Value = 1 Then
                        lstIPs.List(cIndex) = "Offline: " & Client(i).Name & "@" & Client(i).IP
                    Else
                        lstIPs.List(cIndex) = "Offline: " & Client(i).Name '& "@" & Client(I).IP
                    End If
                Client(i).TimedOut = True
                Client(i).LastSeen = "10"
                End If
                If chkOffline.Value = vbChecked Then RemoveOffline: Exit Sub

            Case Is = "STARTING"
            If chkName.Value = Checked Then
                Log UName & " is starting " & data(3)
            Else
                Log IPData & " is starting " & data(3)
            End If
            
            Case Is = "CLEAR"
                RemoveOffline
                'lstIPs.Clear
                'lstIPs.AddItem ClientName
                
            Case Is = "OSERVER"
                Call AddIP(data(3), False, "Shared IP")
                'Call CheckList(Data(3), 1)
            
            Case Is = "SERVER"
                Call CheckList(IPData, data(3))
                'Log "SERVER: " & strData
                'Call CheckUser(IPData, UName)
                
            Case Is = "MSG"
            If IPData = LocalIP Then Exit Sub
            If chkName.Value = Checked Then
                Log UName & ": " & data(3)
            Else
                Log IPData & ": " & data(3)
            End If
            
            Case Is = "TIME"
            'If IPData = LocalIP Then Exit Sub
            If chkName.Value = Checked Then
                Log UName & "'s time is " & data(3)
            Else
                Log IPData & "'s time is " & data(3)
            End If
            SendToAll "TIME2", Now
            Log "## You responded to all with your time."
            tmrSpam.Enabled = True
            
            Case Is = "TIME2"
            'If IPData = LocalIP Then Exit Sub
            If chkName.Value = Checked Then
                Log UName & "'s time is " & data(3)
            Else
                Log IPData & "'s time is " & data(3)
            End If
            tmrSpam.Enabled = True
            
            Case Is = "ME"
            If chkName.Value = Checked Then
                Log UName & " " & Trim(data(3))
            Else
                Log IPData & " " & Trim(data(3))
            End If
            
            Case Is = "NCHANGE"
                Call AddIP(IPData, True, UName)
                Call ChangeName(IPData, UName)
            
            Case Is = "PING2"
            Call CheckVersion(IPData, data(3))
            Call AddIP(IPData, True, UName)
            
            Case Is = "PING1"
            Call AddIP(IPData, True, UName)
            Send "PING2" & SplitChar & LocalIP & SplitChar & ClientName & SplitChar & App.Major & App.Minor & App.Revision & SplitChar
            'AddIP (IPData)
                'Call Send("ADD" & SplitChar & LocalIP & SplitChar, IPData)
                'If AddIP(IPData) = False Then Send ("ADD" & SplitChar & LocalIP)
            Case Else

        End Select
        'Call UpdateList
End Sub
