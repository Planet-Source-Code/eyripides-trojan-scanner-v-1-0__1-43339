VERSION 5.00
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Trojan Scanner"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin SysTrayCtl.cSysTray Sys 
      Left            =   480
      Top             =   2400
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form1.frx":030A
      TrayTip         =   "Trojan Scanner"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Me!"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "For the &Systray press the minimize button in this Form!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   1680
      Picture         =   "Form1.frx":0624
      Top             =   -840
      Width           =   7500
   End
   Begin VB.Menu mnuFile1 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "Load Me!"
         Index           =   1
      End
      Begin VB.Menu mnu8 
         Caption         =   "Turn Off"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub Form_Initialize()
InitCommonControls

End Sub

Private Sub Command1_Click()
frmMain.Show 1
End
End Sub

Private Sub Command2_Click()
End
End
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Form_Resize()
'If Intray property is true
'icon will be displayed on taskbar
'If false then it will not
If Me.WindowState = vbMinimized Then
Sys.InTray = True
Me.Hide
Else
Sys.InTray = False
End If

End Sub

Private Sub mnuAuthor_Click()
frmAuthor.Show 1
End Sub


Private Sub mnu1_Click(Index As Integer)
frmMain.Show 1
End Sub

Private Sub mnu8_Click()
Unload Me
End Sub

Private Sub Sys_MouseUp(Button As Integer, Id As Long)
'The control supports most of the mouse events
'such as MouseDblClick,MouseUp,etc

'If the icon on the taskbar is clicked
'then show our form
Me.WindowState = vbNormal

If Button = 2 Then
Me.PopupMenu Me.mnuFile1


End If

End Sub

