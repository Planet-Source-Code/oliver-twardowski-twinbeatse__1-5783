VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{24365B29-A3B5-11D1-B8B0-444553540000}#1.0#0"; "XFXFORMS.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TwinBeat SE"
   ClientHeight    =   744
   ClientLeft      =   120
   ClientTop       =   600
   ClientWidth     =   1812
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   744
   ScaleWidth      =   1812
   StartUpPosition =   3  'Windows-Standard
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   2400
      Top             =   480
      _ExtentX        =   1863
      _ExtentY        =   1291
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   960
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   0
   End
   Begin VB.Shape Shape2 
      Height          =   2292
      Left            =   0
      Top             =   -360
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Shape Shape1 
      Height          =   1812
      Left            =   -120
      Top             =   240
      Width           =   2172
   End
   Begin VB.Label lblBeat 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1812
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1812
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuoptionsbackcolor 
         Caption         =   "&Backcolour"
      End
      Begin VB.Menu mnuoptionstextcolor 
         Caption         =   "&Textcolour"
      End
      Begin VB.Menu mnuoptionsfont 
         Caption         =   "&Font"
      End
      Begin VB.Menu mnuoptionsline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&?"
      Begin VB.Menu mnuhelpabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuhelphelp 
         Caption         =   "&Help"
      End
   End
   Begin VB.Menu mnuMini 
      Caption         =   "            _"
   End
   Begin VB.Menu mnuX 
      Caption         =   " X"
      NegotiatePosition=   2  'Mitte
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare Variables
Dim GetStatus, GetStatus1, GetStatus2, GetStatus3, GetStatus4 As Long
Dim lonStatus, lonStatus1, lonStatus2, lonStatus3, lonStatus4 As String
Dim Vol, File As String
 

'Declare the shell function for sending mails etc...
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Declare the function for let the form stay on top
Private Declare Function SetWindowPos Lib "user32" _
 (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x _
 As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As _
 Long, ByVal wFlags As Long) As Long

'Declare the variables for let the form stay on top
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2


Public Function GetTimeZone(Optional ByRef strTZName As String) As Long

'Get timezone
    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    lngResult = GetTimeZoneInformation&(objTimeZone)


    Select Case lngResult
        Case 0&, 1& 'use standard time
        GetTimeZone = -(objTimeZone.Bias + objTimeZone.StandardBias) 'into minutes


        For i = 0 To 31
            If objTimeZone.StandardName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.StandardName(i))
        Next

        Case 2& 'use daylight savings time
        GetTimeZone = -(objTimeZone.Bias + objTimeZone.DaylightBias) 'into minutes


        For i = 0 To 31
            If objTimeZone.DaylightName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
        Next

    End Select

End Function


Public Function InternetTime()

    Dim tmpH                '\
    Dim tmpS                ' \
    Dim tmpM                '  \
    Dim itime               '   \
    Dim tmpZ                '    \
    Dim testtemp As String  '====|> Declare variables for the interntetime
    tmpH = Hour(Time)       '    /
    tmpM = Minute(Time)     '   /
    tmpS = Second(Time)     '  /
    tmpZ = GetTimeZone      '_/
    
    'calculate internettime
    itime = ((tmpH * 3600 + ((tmpM - tmpZ + 60) * 60) + tmpS) * 1000 / 86400)

    'Check out for inettime = 1000 ...
    If itime = 1000 Then
        itime = itime - 1000
    ElseIf itime < 0 Then
        itime = itime + 1000
    End If
    
    'Do I have to say something???
    InternetTime = itime
    
End Function



Private Sub Form_Load()
   
   'Let the form stay on top
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOSIZE + SWP_NOMOVE + SWP_SHOWWINDOW
    
    'Shape the form in the shape of alle shapes
    FormShaper1.ShapeIt
    
    
    Me.BackColor = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "Backcolor")
    lblTime.FontName = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "font")
    lblBeat.FontName = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "font")
    lblTime.ForeColor = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "forecolor")
    lblBeat.ForeColor = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "forecolor")
    lblTime.FontSize = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "fontsize")
    lblBeat.FontSize = GetSettingString(HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "fontsize")
    
        
    'Get Settings
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 1272)
     
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save settings
     If Me.WindowState <> vbMinimized Then
        
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        
    End If
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "Backcolor", Me.BackColor
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "font", lblTime.FontName
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "font", lblBeat.FontName
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "fontsize", lblTime.FontSize
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "fontsize", lblBeat.FontSize
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "forecolor", lblTime.ForeColor
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\TwinWare\TwinBeatSE", "forecolor", lblBeat.ForeColor

    
    End
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
  
      If Me.WindowState <> vbMinimized Then
        
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
     '   SaveSetting App.Title, "Settings", "MainBackcolor", Me.BackColor
        
     End If
     End
     
     
End Sub

Private Sub mnuexit_Click()
    
    'End prog
    Unload Me
    End
    
End Sub

Private Sub mnuhelpabout_Click()
    
    'Call About Form
    Form3.Show 1
    
End Sub

Private Sub mnuhelphelp_Click()
        
    'Call help.html with shell excute
    ShellExecute 0, "Open", App.Path & "\help.html", "", "", vbNormalFocus
       
End Sub

Private Sub mnuMini_Click()
    
    'Minimize -> me
    Me.WindowState = 1
    
End Sub

Private Sub mnuoptionsbackcolor_Click()
    
    'Get backcolor
    CommonDialog1.ShowColor
    Me.BackColor = CommonDialog1.Color
    
End Sub

Private Sub mnuoptionsfont_Click()
        
    On Error Resume Next
    
    'Get font
    CommonDialog1.ShowFont
    lblTime.FontName = CommonDialog1.FontName
    lblTime.FontSize = CommonDialog1.FontSize
    lblBeat.FontName = CommonDialog1.FontName
    lblBeat.FontSize = CommonDialog1.FontSize
    
End Sub

Private Sub mnuoptionstextcolor_Click()
    
    'Get textcolor
    CommonDialog1.ShowColor
    lblTime.ForeColor = CommonDialog1.Color
    lblBeat.ForeColor = CommonDialog1.Color
    
End Sub

'Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
'End Sub

Private Sub mnuX_Click()
    
    'Unload -> me (haha)
    Unload Me
    End
    
End Sub

Private Sub Timer1_Timer()
             
     '.Caption is @ + Internettime but just in 3 Chars (CInt)
     lblBeat.Caption = "@ " & (CInt(InternetTime))
     '.Caption is time
     lblTime.Caption = Time$
     'If minimized then Taskbutton.caption is Internettime
     If Me.WindowState = 1 Then
     Me.Caption = "TwinBeat SE" & " @" & (CInt(InternetTime))
     'Else cation is name
     Else: Me.Caption = "TwinBeat SE"
     End If
     
     
End Sub


