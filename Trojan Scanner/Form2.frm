VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "About Trojen Scanner"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4335
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   480
      Top             =   3720
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3135
      Left            =   1080
      ScaleHeight     =   3075
      ScaleMode       =   0  'User
      ScaleWidth      =   3379.882
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      Begin VB.Label lblAn 
         BackStyle       =   0  'Transparent
         Caption         =   "                                         Trojen Scanner"
         ForeColor       =   &H8000000B&
         Height          =   225
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   4410
      End
   End
   Begin VB.Label Label1 
      Caption         =   "About Trojan Scanner..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'Load as many lines as you want!

'********************
  LinesNum = 6
'********************


'*************************************
'Initialize and load Labels and Timers
For i = 1 To LinesNum
    Load lblAn(i)
    Load tmr(i)
    lblAn(i).Top = lblAn(i).Top
    lblAn(i).Left = lblAn(i).Left
    lblAn(i).Height = lblAn(i).Height
    lblAn(i).Width = lblAn(i).Width
    lblAn(i).Visible = True
    
    'If its too fast, change here! If its too slow, try changing the speed from the
    'source of the timer.
    tmr(i).Interval = 1
    
    tmr(i).Enabled = False
Next i



'*************************************
'Using the arrays, add your lines!

lblAn(1).Caption = "Trojan Scanner for your PC!"

'If you want more than one lines per scroll, add a vbCrLf like this:
lblAn(2).Caption = "A Nice Trojan Scanner"

lblAn(3).Caption = "Programmed by me!And little Help by Kyriako Nicola"

lblAn(4).Caption = "I hope to vote me!  "

lblAn(5).Caption = "Send e-mail at eyripides_k@hotmail.com!"

lblAn(6).Caption = "" 'Leave the last one blank if you want some seconds before restart the about text.
'*************************************



'*************************************
'Initialize positions

For i = 1 To lblAn.Count - 1
    lblAn(i).FontSize = 9
    lblAn(i).AutoSize = True
    lblAn(i).Alignment = vbCenter
    lblAn(i).Top = Picture1.Height + 10
Next i
'*************************************



'*************************************
'Start the initialization timer! Everything now depends on the timers!
tmr(0).Enabled = True
'*************************************





'*************************************
'Create this beautiful Blue to Black Background

LineBackGround Picture1

'*************************************
End Sub


Private Sub LineBackGround(ctl As PictureBox)
On Error Resume Next

'This line back ground code is taken from the Setup1.VBP (Visual Basic Setup)
'I made some changes...!
    
    Const intBLUESTART% = 255
    Const intBLUEEND% = 0
    Const intBANDHEIGHT% = 50

    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer

    intFormHeight = ctl.ScaleHeight
    intFormWidth = ctl.ScaleWidth
    
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
    sngBlueCur = intBLUESTART

    For intY = 0 To intFormHeight Step intBANDHEIGHT
        ctl.Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
        sngBlueCur = sngBlueCur + sngBlueStep
    Next intY
End Sub


Private Sub tmr_Timer(Index As Integer)
On Error Resume Next
If Index = 0 Then
    
    tmr(0).Enabled = False
    tmr(1).Enabled = True
Else
    
    '*****************************************
    'This is the speed which the labels goes.
    'Default is 5. Change it to make it faster or slower...
    'For Slower, I recommend you to change the timer's Interval.
    lblAn(Index).Top = lblAn(Index).Top - 5
    '*****************************************
    
    FS = (CInt(lblAn(Index).Top) / 100) * 3
    
    
    '************************************************************************
    'Wanna change the color effect? Play with this!
    
    'If you want you can do incredible things with colors.
    'Add onother one variable like FCol (Fcol = foreground color for me)
    'and add it below at the RGB(..) Function and you'll have
    'awesome results!
    
    FCol = CInt((2000 / Picture1.Height) * CInt(lblAn(Index).Top))
    
    ' * FCol2 = CInt((1000 / Picture1.Height) * CInt(lblAn(Index).Top))
    ' * FCol3 = CInt((100 / Picture1.Height) * CInt(lblAn(Index).Top))
    
   
    
    lblAn(Index).ForeColor = RGB(FCol, FCol, FCol) 'White to Black
    
    'This one uses the FCol2 and FCol3... Try it!
    'lblAn(Index).ForeColor = RGB(FCol, FCol2, FCol3)
    '************************************************************************
    
    '************************************************************************
    'Just a check for bugs at the FontSize.
    If FS <= 0 Then FS = 1
    If FS > 9 Then FS = 9
    '************************************************************************
    
    
    '************************************************************************
    'Check for the FontSize and the Fade Away Effect Speed
    
    'Just check if the FontSize is not already gaven (for speed waste)
    If Not lblAn(Index).FontSize = FS Then
        
        lblAn(Index).FontSize = FS
        
        'If you want to change the Fade Away Speed, Change the value below.
        'Default is 2.
        lblAn(Index).Top = lblAn(Index).Top - 2
        
    End If
    '************************************************************************
    
    
    
    '************************************************************************
    'Label is Over the head! It walked fine! Let's go to the next one
    'if there is!
    
    If lblAn(Index).Top <= 0 - lblAn(Index).Height - 10 Then 'Check if its over the PictureBox
        
        tmr(Index).Enabled = False 'Stop current Timer.
        
        tmr(Index + 1).Enabled = True 'Start Next Timer For next Label Scroll
        
        
        'Check if we are at the end of the text(labels).
        If Index + 1 = tmr.Count Then
            
            
            For i = 1 To lblAn.Count
                lblAn(i).FontSize = 5
                lblAn(i).Top = Picture1.Height + 10
            Next i
            
            'Start The First Timer just for time waste.
            tmr(0).Enabled = True
        
        End If
    
    End If

End If



End Sub



