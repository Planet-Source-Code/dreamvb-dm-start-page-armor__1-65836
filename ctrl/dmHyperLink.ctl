VERSION 5.00
Begin VB.UserControl dmHyperLink 
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   MouseIcon       =   "dmHyperLink.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   210
   ScaleWidth      =   2040
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "dmHyperLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Event MouseOut()
Event MouseIn()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim m_HoverIn As OLE_COLOR
Dim m_HoverOut As OLE_COLOR
Dim m_activeColor As OLE_COLOR

Public Sub Update()
    Call lblLink_MouseMove(1, 0, 0, 0)
End Sub

Sub DoHover(mShow As Boolean)
    If mShow Then
        lblLink.ForeColor = m_HoverIn
    Else
        lblLink.ForeColor = m_HoverOut
    End If
    lblLink.FontUnderline = mShow
End Sub

Sub DoCapture(mCapture As Boolean)
    If mCapture Then
        SetCapture UserControl.hwnd
    Else
        ReleaseCapture
    End If
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Initialize()
    m_HoverIn = vbBlue
    m_HoverOut = ForeColor
    m_activeColor = vbRed
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim mHover As Boolean
    RaiseEvent MouseMove(Button, Shift, x, Y)
    
    If (x < 0 Or Y < 0 Or x > lblLink.Width Or Y > lblLink.Height) Then
        DoCapture False
        mHover = False
        DoHover mHover
        RaiseEvent MouseOut
    ElseIf mHover <> True Then
        DoCapture True
        mHover = True
        DoHover mHover
        RaiseEvent MouseIn
    End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    lblLink.Height = UserControl.Height
    lblLink.Width = UserControl.Width
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    lblLink.ForeColor = m_activeColor
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblLink.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblLink.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Caption() As String
    Caption = lblLink.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblLink.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblLink.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_HoverIn = PropBag.ReadProperty("HoverIn", vbBlue)
    m_HoverOut = PropBag.ReadProperty("HoverOut", vbRed)
    lblLink.Caption = PropBag.ReadProperty("Caption", "Label1")
    Set lblLink.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    lblLink.Enabled = PropBag.ReadProperty("Enabled", True)
    m_activeColor = PropBag.ReadProperty("ActiveColor", vbRed)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", lblLink.ForeColor, &H80000012)
    Call PropBag.WriteProperty("HoverIn", m_HoverIn, vbBlue)
    Call PropBag.WriteProperty("HoverOut", m_HoverOut, vbRed)
    Call PropBag.WriteProperty("Caption", lblLink.Caption, "Label1")
    Call PropBag.WriteProperty("Font", lblLink.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Enabled", lblLink.Enabled, True)
    Call PropBag.WriteProperty("ActiveColor", m_activeColor, vbRed)
End Sub

Public Property Get HoverIn() As OLE_COLOR
    HoverIn = m_HoverIn
End Property

Public Property Let HoverIn(ByVal vNewValue As OLE_COLOR)
    m_HoverIn = vNewValue
End Property

Public Property Get HoverOut() As OLE_COLOR
    HoverOut = m_HoverOut
End Property

Public Property Let HoverOut(ByVal vNewValue As OLE_COLOR)
    m_HoverOut = vNewValue
End Property

Public Property Get Font() As Font
    Set Font = lblLink.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblLink.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    lblLink.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get ActiveColor() As OLE_COLOR
    ActiveColor = m_activeColor
End Property

Public Property Let ActiveColor(ByVal vNewValue As OLE_COLOR)
    m_activeColor = vNewValue
End Property
