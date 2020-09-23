VERSION 5.00
Object = "{24365B29-A3B5-11D1-B8B0-444553540000}#1.0#0"; "XFXFORMS.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00F4926C&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "                                                   About TwinBeat SE"
   ClientHeight    =   3348
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   6600
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3348
   ScaleWidth      =   6600
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4926C&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Desdemona"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2400
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   2640
      Width           =   1812
   End
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   480
      Top             =   4320
      _ExtentX        =   1863
      _ExtentY        =   1291
   End
   Begin VB.Shape Shape2 
      Height          =   3252
      Left            =   240
      Shape           =   2  'Oval
      Top             =   0
      Width           =   6252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "oli-t@topmail.de"
      BeginProperty Font 
         Name            =   "Ultra Shadow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   372
      Left            =   2040
      MouseIcon       =   "frmAboutTwinBeatSE.frx":0000
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   1
      ToolTipText     =   "Send me a Mail!"
      Top             =   2160
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Any Questions??? Send an e-Mail to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   3852
   End
   Begin VB.Image Image1 
      Height          =   2088
      Left            =   360
      Picture         =   "frmAboutTwinBeatSE.frx":0CCA
      Top             =   120
      Width           =   5832
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    
    Dim x As Long
Dim inc As Long
inc = 40


For x = Me.Height To 300 Step -inc


    DoEvents
        Me.Move Me.Left, Me.Top + (inc \ 2), Me.Width, x
    Next x


    'This is the width part of the same sequence above


    For x = Me.Width To 2000 Step -inc


        DoEvents
            Me.Move Me.Left + (inc \ 2), Me.Top, x, Me.Height
        Next x
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    FormShaper1.ShapeIt
    
End Sub

Private Sub Label2_Click()
    
    ShellExecute 0, "Open", "mailto: oli-t@topmail.de?subject=TwinBeat SE", "", "", vbNormalFocus
  
End Sub
