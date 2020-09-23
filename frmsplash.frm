VERSION 5.00
Object = "{24365B29-A3B5-11D1-B8B0-444553540000}#1.0#0"; "XFXFORMS.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00F4926C&
   BorderStyle     =   0  'Kein
   Caption         =   "Form2"
   ClientHeight    =   2208
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6996
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2208
   ScaleWidth      =   6996
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   360
      Top             =   3000
      _ExtentX        =   1863
      _ExtentY        =   1291
   End
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   5520
      Top             =   1800
   End
   Begin VB.Shape Shape1 
      Height          =   2412
      Left            =   0
      Shape           =   2  'Oval
      Top             =   -120
      Width           =   6852
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING..."
      BeginProperty Font 
         Name            =   "Bedrock"
         Size            =   22.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0034D5FC&
      Height          =   732
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   4812
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5520
      TabIndex        =   0
      Top             =   1080
      Width           =   2172
   End
   Begin VB.Image Image1 
      Height          =   2088
      Left            =   -120
      Picture         =   "frmsplash.frx":0000
      Top             =   0
      Width           =   5832
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    '.Caption = Version
    Label1.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'Shape the form in the shape of alle shapes
    FormShaper1.ShapeIt

End Sub
Private Sub Timer2_Timer()
    
    'Show main form
    Form1.Show
    'Interval = 0
    Timer2.Interval = 0
    'Unload me
    Unload Me
    
End Sub
