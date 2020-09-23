VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attention"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAlert 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Do not alert me in the future"
      Height          =   195
      Left            =   285
      TabIndex        =   6
      Top             =   1485
      Width           =   2340
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   105
      Picture         =   "frmAlert.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblBut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6510
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1695
      Width           =   255
   End
   Begin VB.Label lblBut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5250
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1695
      Width           =   330
   End
   Begin VB.Image ImgBut 
      Height          =   390
      Index           =   0
      Left            =   4965
      MousePointer    =   99  'Custom
      Picture         =   "frmAlert.frx":1B42
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   990
   End
   Begin VB.Image ImgBut 
      Height          =   390
      Index           =   1
      Left            =   6165
      MousePointer    =   99  'Custom
      Picture         =   "frmAlert.frx":3D44
      Stretch         =   -1  'True
      Top             =   1605
      Width           =   990
   End
   Begin VB.Label lblask 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to set your home page to the above setting?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1035
      TabIndex        =   3
      Top             =   1110
      Width           =   4260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000F&
      X1              =   1050
      X2              =   6570
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label lblPage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   1065
      TabIndex        =   2
      Top             =   645
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An attempt was made to change your default home page."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1065
      TabIndex        =   1
      Top             =   330
      Width           =   5550
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAlert_Click()
    SaveSetting "StartPageArmor", "cfg", "Alert", Abs(chkAlert.Value)
    frmmain.chkAlert.Value = chkAlert.Value
End Sub

Private Sub chkAlert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkAlert.ForeColor = vbBlue
End Sub

Private Sub Form_Load()
Dim tmpPage As String

    chkAlert.Value = Abs(GetSetting("StartPageArmor", "cfg", "Alert"))

    lblBut(0).MouseIcon = frmmain.ImgButt(0).MouseIcon
    lblBut(1).MouseIcon = frmmain.ImgButt(0).MouseIcon
    ImgBut(0).MouseIcon = lblBut(0).MouseIcon
    ImgBut(1).MouseIcon = lblBut(1).MouseIcon
    
    tmpPage = m_CheckPage
    lblPage.Caption = tmpPage
    'An attemt has been made to change the users home page set back to it's default
    RegSaveValue HKEY_CURRENT_USER, m_HomePagePath, "Start Page", REG_EXPAND_SZ, m_OldPage
    'Play sound
    Call PlayAlert
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    chkAlert.ForeColor = vbBlack
End Sub

Private Sub ImgBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblbut_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub ImgBut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblbut_MouseUp Index, Button, Shift, X, Y
End Sub

Private Sub lblbut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    lblBut(Index).ForeColor = vbRed
End Sub

Private Sub lblbut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    lblBut(Index).ForeColor = vbBlack
    
    If (Index = 1) Then
        'do nothing just unload
        Call DoRegLoad
        Unload frmAlert
    Else
        'Save the new start page
        RegSaveValue HKEY_CURRENT_USER, m_HomePagePath, "Start Page", REG_EXPAND_SZ, lblPage.Caption
        SaveSetting "StartPageArmor", "cfg", "CurrentPage", lblPage.Caption
        Call DoRegLoad
        Unload frmAlert
    End If
    
End Sub
