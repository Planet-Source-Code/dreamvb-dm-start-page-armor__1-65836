VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start Page Armor"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7035
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin StartPageArmor.dmHyperLink dmHyperLink2 
      Height          =   180
      Left            =   6285
      TabIndex        =   14
      Top             =   60
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   318
      HoverOut        =   -2147483630
      Caption         =   "Minsize"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
   End
   Begin VB.CheckBox chkAutoRun 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run Start Page Armor when Windows starts"
      Height          =   210
      Left            =   300
      TabIndex        =   13
      Top             =   1350
      Width           =   3675
   End
   Begin VB.CheckBox chkSound 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enable sounds on alerts"
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   1875
      Width           =   5445
   End
   Begin StartPageArmor.Tray Tray1 
      Left            =   5655
      Top             =   15
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   5115
      Top             =   15
   End
   Begin VB.PictureBox pBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00F4F4F4&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   9
      Top             =   3465
      Width           =   7035
      Begin StartPageArmor.dmHyperLink dmHyperLink1 
         Height          =   240
         Left            =   105
         TabIndex        =   15
         Top             =   75
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   423
         HoverOut        =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16053492
      End
   End
   Begin VB.CheckBox chkReadOnly 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lock hosts as ReadOnly"
      Height          =   255
      Left            =   300
      TabIndex        =   8
      Top             =   2115
      Width           =   2310
   End
   Begin VB.CheckBox chkBlockUser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lock user from changing start page"
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   2400
      Width           =   4575
   End
   Begin VB.CheckBox chkAlert 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alert each time an attempt is made to change my Start Page"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   1575
      Width           =   4620
   End
   Begin VB.PictureBox pBar2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   975
      Width           =   1995
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defense Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   105
         TabIndex        =   5
         Top             =   30
         Width           =   1380
      End
   End
   Begin VB.PictureBox PicBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   210
      ScaleHeight     =   330
      ScaleWidth      =   6345
      TabIndex        =   2
      Top             =   465
      Width           =   6345
      Begin VB.Label lblPage 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   105
      End
   End
   Begin VB.PictureBox PicBar1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   195
      Width           =   1995
      Begin VB.Label lblCurPage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current StartPage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   105
         TabIndex        =   1
         Top             =   30
         Width           =   1485
      End
   End
   Begin VB.Shape shpLine 
      BackColor       =   &H0080C0FF&
      BorderColor     =   &H00C0C0C0&
      Height          =   1485
      Left            =   210
      Top             =   1245
      Width           =   6345
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   405
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6105
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3015
      Width           =   390
   End
   Begin VB.Image ImgButt 
      Height          =   450
      Index           =   1
      Left            =   5835
      MousePointer    =   99  'Custom
      Picture         =   "frmmain.frx":0A02
      Stretch         =   -1  'True
      Top             =   2910
      Width           =   990
   End
   Begin VB.Label lblbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4740
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image ImgButt 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   4590
      MousePointer    =   99  'Custom
      Picture         =   "frmmain.frx":2C04
      Stretch         =   -1  'True
      Top             =   2910
      Width           =   1125
   End
   Begin VB.Image ImgEdit 
      Height          =   240
      Left            =   6600
      MouseIcon       =   "frmmain.frx":4E06
      MousePointer    =   99  'Custom
      Picture         =   "frmmain.frx":4F58
      ToolTipText     =   "Edit StartPage"
      Top             =   510
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRes 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnubalnk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_HostsFile As String
Private m_HostLoc(2) As String

Sub CheckColor(p As CheckBox, Optional skip As Boolean = False)
Dim X As Byte
    For X = 0 To frmmain.Controls.Count - 1
        If TypeName(Me.Controls(X)) = "CheckBox" Then
            frmmain.Controls(X).ForeColor = vbBlack
        End If
    Next X
    
    If Not skip Then p.ForeColor = vbBlue
    
End Sub

Private Sub UnloadAll()
    'Clear up variables
    m_HostsFile = ""
    m_CurrentPage = ""
    m_OldPage = ""
    m_CheckPage = ""
    'Stop timer
    Timer1.Enabled = False
    'Unload all forms
    Unload frmabout
    Set frmabout = Nothing
    Unload frmAlert
    Set frmAlert = Nothing
    Unload frmmain
    Set frmmain = Nothing
End Sub

Public Sub chkAlert_Click()
    SaveSetting "StartPageArmor", "cfg", "Alert", Abs(chkAlert.Value)
End Sub

Private Sub chkAlert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckColor chkAlert
End Sub

Private Sub chkAutoRun_Click()
    SaveSetting "StartPageArmor", "cfg", "AutoRun", Abs(chkSound.Value)
End Sub

Private Sub chkAutoRun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckColor chkAutoRun
End Sub

Private Sub chkBlockUser_Click()
Dim iVal As Integer
    If (chkBlockUser) Then iVal = 1 Else iVal = 0
    'Save Reg key
    RegSaveValue HKEY_CURRENT_USER, m_RegPath, "Homepage", REG_DWORD, iVal
End Sub

Private Sub chkBlockUser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckColor chkBlockUser
End Sub

Private Sub chkReadOnly_Click()
On Error Resume Next

    If (chkReadOnly) Then
        SetAttr m_HostsFile, vbReadOnly
    Else
        SetAttr m_HostsFile, vbNormal
    End If
    
End Sub

Private Sub chkReadOnly_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckColor chkReadOnly
End Sub

Private Sub chkSound_Click()
    SaveSetting "StartPageArmor", "cfg", "PlaySound", Abs(chkSound.Value)
End Sub

Private Sub chkSound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckColor chkSound
End Sub

Private Sub Command1_Click()
CheckColor chkAutoRun
End Sub

Private Sub dmHyperLink1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OpenUrl "http://www.eraystudios.co.uk"
    dmHyperLink1.Update
End Sub

Private Sub dmHyperLink2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    frmmain.Visible = False
    Tray1.Visible = True
    dmHyperLink2.Update
End Sub

Private Sub Form_Load()
Dim X As Integer
On Error Resume Next
    
    Set Tray1.Icon = frmmain.Icon
    Tray1.ToolTip = frmmain.Caption
    
    ImgButt(0).MouseIcon = ImgEdit.MouseIcon
    ImgButt(1).MouseIcon = ImgEdit.MouseIcon
    lblBut(0).MouseIcon = ImgButt(0).MouseIcon
    lblBut(1).MouseIcon = ImgButt(0).MouseIcon
    
    dmHyperLink1.Caption = frmmain.Caption & " Copyright © 1990-2006 eRay Studios"
    
    m_HostLoc(0) = "C:\WINDOWS\hosts"
    m_HostLoc(1) = "C:\WINDOWS\SYSTEM32\DRIVERS\ETC\hosts"
    m_HostLoc(2) = "C:\WINNT\SYSTEM32\DRIVERS\ETC\hosts"
    
    For X = 0 To 2
        If LenB(Dir(m_HostLoc(X))) <> 0 Then
            m_HostsFile = m_HostLoc(X)
            X = 2
        End If
    Next X
    X = 0
    
    Erase m_HostLoc
    chkReadOnly.Value = Abs(GetAttr(m_HostsFile) = vbReadOnly)
    chkBlockUser.Value = RegReadString(HKEY_CURRENT_USER, m_RegPath, "Homepage", REG_DWORD)
    chkAlert.Value = Abs(GetSetting("StartPageArmor", "cfg", "Alert", "1"))
    chkSound.Value = Abs(GetSetting("StartPageArmor", "cfg", "PlaySound", 1))
    chkAutoRun.Value = Abs(GetSetting("StartPageArmor", "cfg", "AutoRun", 0))
    'Get the current users start page
    Call DoRegLoad
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckColor chkAutoRun, True
End Sub

Private Sub Form_Paint()
    PicBar.Line (0, 0)-(PicBar.ScaleWidth - 8, PicBar.ScaleHeight - 8), &H8000000F, B
End Sub

Private Sub Form_Resize()
    ImgEdit.Left = (PicBar.ScaleWidth + ImgEdit.Width) + 80
    Line2.X2 = frmmain.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnloadAll
End Sub

Private Sub ImgButt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblbut_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub ImgButt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblbut_MouseUp Index, Button, Shift, X, Y
End Sub

Private Sub ImgEdit_Click()
Dim sUrl As String
    sUrl = Trim(InputBox("Enter the address below for your start page.", "Edit Start Page", lblPage.Caption))
    
    If Len(sUrl) = 0 Then
        Exit Sub
    Else
        'Set New StartPage
        'Save the new start page
        RegSaveValue HKEY_CURRENT_USER, m_HomePagePath, "Start Page", REG_EXPAND_SZ, sUrl
        SaveSetting "StartPageArmor", "cfg", "CurrentPage", sUrl
        Call DoRegLoad
        'Update label with new page
        lblPage.Caption = sUrl
        sUrl = ""
    End If

End Sub

Private Sub lblbut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    lblBut(Index).ForeColor = vbRed
End Sub

Private Sub lblbut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    lblBut(Index).ForeColor = vbBlack
    
    If (Index = 0) Then
        mnuabout_Click
    Else
        Unload frmmain
    End If
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, frmmain
End Sub

Private Sub mnuexit_Click()
    lblbut_MouseUp 1, vbLeftButton, 0, 0, 0
End Sub

Private Sub mnuRes_Click()
    Tray1_MouseDown vbLeftButton
End Sub

Private Sub pBottom_Resize()
    pBottom.Line (0, 0)-(pBottom.ScaleWidth - 1, 0), &HC0C0C0, B
    pBottom.Refresh
End Sub

Private Sub Timer1_Timer()
    m_CheckPage = RegReadString(HKEY_CURRENT_USER, m_HomePagePath, "Start Page", REG_EXPAND_SZ)
    
    If (m_CheckPage) <> (m_OldPage) Then
        If chkAlert Then
            frmAlert.Show vbModal, frmmain
            Exit Sub
        Else
            RegSaveValue HKEY_CURRENT_USER, m_HomePagePath, "Start Page", REG_EXPAND_SZ, m_OldPage
        End If
    End If
    
End Sub

Private Sub Tray1_MouseDown(Button As Integer)

    If Button <> vbLeftButton Then
        PopupMenu mnuFile
    Else
        Tray1.Visible = False
        frmmain.Visible = True
    End If

End Sub

