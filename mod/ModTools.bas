Attribute VB_Name = "ModTools"
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Reg Paths
Public Const m_RegPath = "Software\Policies\Microsoft\Internet Explorer\Control Panel"
Public Const m_HomePagePath = "Software\Microsoft\Internet Explorer\Main"
Public Const m_RunPath = "Software\Microsoft\Windows\CurrentVersion\Run"
'Sound Flag Consts
Private Const SND_NODEFAULT = &H2
Private Const SND_RESOURCE = &H40004
Private Const SND_SYNC = &H0
'
Private Const GWL_HINSTANCE As Long = -6

Public m_CurrentPage As String
Public m_OldPage As String
Public m_CheckPage As String

Public Sub DoRegLoad()
    m_CurrentPage = RegReadString(HKEY_CURRENT_USER, m_HomePagePath, "Start Page", REG_EXPAND_SZ)
    m_OldPage = GetSetting("StartPageArmor", "cfg", "CurrentPage", "")
    
    If m_OldPage = "" Then
        SaveSetting "StartPageArmor", "cfg", "CurrentPage", m_CurrentPage
        m_OldPage = m_CurrentPage
    End If
    '
    frmmain.lblPage.Caption = m_CurrentPage
End Sub

Public Sub PlayAlert()
On Error Resume Next
    Const sFlags = SND_RESOURCE Or SND_SYNC Or SND_NODEFAULT
    
    'If user wants to play sounds
    If Abs(GetSetting("StartPageArmor", "cfg", "PlaySound", 0)) <> 0 Then
        If (waveOutGetNumDevs >= 1) Then
            'If sound card found play wav
            PlaySound "Attention", ByVal 0&, sFlags
            Exit Sub
        Else
            'Beep the internal speaker
            Beep
        End If
    End If
    
End Sub

Public Sub OpenUrl(sUrl As String)
    Call ShellExecute(frmmain.hwnd, "open", sUrl, vbNullString, vbNullString, 1)
End Sub
