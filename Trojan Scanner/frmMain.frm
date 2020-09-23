VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Trojan Scanner"
   ClientHeight    =   5925
   ClientLeft      =   1455
   ClientTop       =   1230
   ClientWidth     =   10725
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   5925
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Help Me!!!"
      DragIcon        =   "frmMain.frx":0614
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      MouseIcon       =   "frmMain.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   3960
      TabIndex        =   9
      Top             =   4680
      Width           =   1935
   End
   Begin VB.ListBox TrojanList 
      Height          =   3180
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.FileListBox filelist 
      Height          =   675
      Left            =   6120
      TabIndex        =   6
      Top             =   4080
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8520
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.DirListBox Dirlist 
      Height          =   2565
      Left            =   6120
      TabIndex        =   3
      Top             =   1320
      Width           =   4095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000D&
      Caption         =   "Off"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000D&
      Caption         =   "On"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5670
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Height          =   135
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "FireWall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   2400
      Picture         =   "frmMain.frx":0C28
      Top             =   -960
      Width           =   7500
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuh1 
      Caption         =   "Help"
      Begin VB.Menu mnuh 
         Caption         =   "Help me!"
      End
      Begin VB.Menu mnua 
         Caption         =   "About Trojen Scanner..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Dim LastDrive As String
Dim FileNum As Integer
Dim FileBinary As String



Private Sub cmdScan_Click()

    If filelist.ListCount = 0 Then MsgBox "Their is no project files in the folder you specified.", vbQuestion, "Project scanner": Exit Sub

    

    For M = 0 To filelist.ListCount - 1
        filelist.Selected(M) = True
        FileNum = FreeFile
        If Right(Dirlist, 1) = "\" Or Right(Dirlist, 1) = "/" Then
            Open Dirlist.Path & filelist.List(M) For Binary As #FileNum
        Else
            Open Dirlist.Path & "\" & filelist.List(M) For Binary As #FileNum
        End If
            FileBinary = Space(2)
            Get #FileNum, 1, FileBinary
            If LCase(FileBinary) = LCase("MZ") Then TrojanList.AddItem (filelist.List(M))
        Close #FileNum
    Next M
    
    If TrojanList.ListCount = 0 Then
        MsgBox "Files scanning complete. Their was no possible exe trojans detected in source files.", vbOKOnly, "Scan complete"
    Else
        MsgBox "Possible trojens were found in project files. You can view these in the possible trojen list.", vbCritical, "Warning!"
    End If
  
End Sub



Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Scan your File for any Trojan."
End Sub

Private Sub Command1_Click()
FrmHelp.Show 1
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Help for the Trojan Scanner."
End Sub

Private Sub Command2_Click()

Unload Me
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Exit."
End Sub

Private Sub Dirlist_Change()
 filelist.Path = Dirlist.Path
End Sub

Private Sub Drive_Change()
  On Error GoTo FinaliseError
    Dirlist.Path = Drive.Drive
    filelist.Path = Dirlist.Path
    LastDrive = Drive.Drive
    Exit Sub
FinaliseError:
    MsgBox "Error, drive not ready.", vbCritical, "Drive not ready"

End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "Welcome."
End Sub

Private Sub mnua_Click()
Form2.Show 1
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub



Private Sub mnuh_Click()
FrmHelp.Show 1
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "FireWall is Under Construction."
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "FireWall is Under Construction."
End Sub

Private Sub TrojanList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Status.SimpleText = "List Of any Trojan."
End Sub
