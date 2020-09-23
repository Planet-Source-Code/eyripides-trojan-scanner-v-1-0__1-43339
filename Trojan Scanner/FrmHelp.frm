VERSION 5.00
Begin VB.Form FrmHelp 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Help me!"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Timer T2 
      Left            =   720
      Top             =   3240
   End
   Begin VB.Timer T1 
      Left            =   120
      Top             =   3240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "FireWall (Under Construction)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Go To &Systray the Load Form."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   240
      Picture         =   "FrmHelp.frx":0442
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "HELP ME!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2775
      Left            =   3600
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Scan your files for Trojan."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

T2.Enabled = True
End Sub

Private Sub Form_Load()
T1.Interval = 5
T2.Interval = 5

T2.Enabled = False
End Sub

Private Sub T1_Timer()
Me.Height = Me.Height + 30
Me.Width = Me.Width + 30

Me.Left = Me.Left + 20
Me.Top = Me.Top + 20

'This command disables T1(Timer1) when Me.Height goes to 6000
If Me.Height = 6000 Then T1.Enabled = False
'Here you can put width too, it doesn't matter we only
'want a value where our T1 commands will stop being repeated

End Sub

Private Sub T2_Timer()
Me.Height = Me.Height - 30
Me.Width = Me.Width - 30

Me.Left = Me.Left - 20
Me.Top = Me.Top - 20

'Now we want the smallest value that our form can reach.
'As we all know the smallest value we can get in non
'negative numbers is 0(zero).
'Well, I found out that the smallest value we can get in
'any form in VB is 120 for Height and 510 for width.
'So, when Me.Height reaches 120 the project unloads!
If Me.Height = 120 Then Unload Me
End Sub
