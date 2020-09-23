VERSION 5.00
Begin VB.Form FTray 
   Caption         =   "BackupShield"
   ClientHeight    =   2835
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4875
   Icon            =   "FTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup now"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim Tic As NOTIFYICONDATA

Private Sub Form_Load()
Dim rc As Long
    If App.PrevInstance Then End
    gHW = Me.hwnd
    Hook
    FTray.Hide
    Tic.cbSize = Len(Tic)
    Tic.hwnd = gHW
    Tic.uID = vbNull
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Me.Icon
    Tic.sTip = "Backup-shield" & vbNullChar
    rc = Shell_NotifyIcon(NIM_ADD, Tic)
    LoadSettings
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim MSG As Long
Dim sFilter As String
    
    'If ProgramsStarted <> None_Started Then Exit Sub
    MSG = x / Screen.TwipsPerPixelX
    Select Case MSG
        Case WM_LBUTTONDBLCLK
            If vbYes = MsgBox("Do You want to end Backup-shield ?", vbCritical + vbYesNo, "Backup") Then
                Terminate
            End If
        Case WM_RBUTTONUP, WM_LBUTTONUP
            PopupMenu mnuMain
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Terminate
End Sub

Private Sub Terminate()
Dim rc As Long
    Unhook
    rc = Shell_NotifyIcon(NIM_DELETE, Tic)
    End
End Sub

Private Sub mnuBackup_Click()
    DoBackup
End Sub

Private Sub mnuExit_Click()
    If vbYes = MsgBox("Do You want to end Backup-shield ?", vbCritical + vbYesNo, "Backup") Then
        Terminate
    End If
End Sub

Private Sub mnuSettings_Click()
    Load FSettings
    FSettings.Show
End Sub

Private Sub LoadSettings()
Dim iCount As Integer
Dim strTemp As String
Dim Found As Boolean
If Dir(App.Path & "\Backup.ini") = "Backup.ini" Then
    BackupFolderName = ReadINI(App.Path & "\Backup.ini", "BackupDir", "0")
    Found = False
    While Not Found
        strTemp = ReadINI(App.Path & "\Backup.ini", "SourceDir", CStr(iCount + 1))
        If strTemp = "" Then
            Found = True
        Else
            iCount = iCount + 1
            ReDim Preserve SourceFolder(iCount)
            SourceFolder(iCount) = strTemp
        End If
    Wend
End If
End Sub

