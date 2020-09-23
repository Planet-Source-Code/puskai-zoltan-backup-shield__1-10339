VERSION 5.00
Begin VB.Form FSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup settings"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save list && exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove from backup list"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ListBox lstBackup 
      Height          =   3570
      Left            =   140
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   1320
      Width           =   6135
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "..."
      Height          =   320
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   400
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Drag && drop files and directories in list !"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label lblDestination 
      Caption         =   "Backup directory :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDestination_Click()
Dim strTemp As String

strTemp = fBrowseForFolder(Me.hwnd, "Select backup path")
If strTemp <> "" Then
    txtDestination = strTemp
End If

End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
For i = lstBackup.ListCount - 1 To 0 Step -1
    If lstBackup.Selected(i) Then lstBackup.RemoveItem (i)
Next
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
Dim intFileNr As Integer
    Kill App.Path & "\Backup.ini"
    WriteINI App.Path & "\Backup.ini", "BackupDir", "0", txtDestination.Text
    BackupFolderName = txtDestination.Text
    ReDim SourceFolder(lstBackup.ListCount)
    For i = 0 To lstBackup.ListCount - 1
        WriteINI App.Path & "\Backup.ini", "SourceDir", CStr(i + 1), lstBackup.List(i)
        SourceFolder(i + 1) = lstBackup.List(i)
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    txtDestination.Text = BackupFolderName
    For i = 1 To UBound(SourceFolder)
        lstBackup.AddItem SourceFolder(i)
    Next
End Sub


Private Sub lstBackup_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim numFiles As Integer
    numFiles = Data.Files.Count
    Dim i As Integer
    For i = 1 To numFiles
        'File or directory?
        If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
            lstBackup.AddItem Data.Files(i)
        Else
            lstBackup.AddItem Data.Files(i)
        End If
    Next i

End Sub

Private Sub lstBackup_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy
    Else
        Effect = vbDropEffectNone
    End If

End Sub
