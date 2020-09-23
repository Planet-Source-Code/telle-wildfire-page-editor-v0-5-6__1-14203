VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   1680
   ClientLeft      =   3030
   ClientTop       =   3090
   ClientWidth     =   3210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   624
      Top             =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Can&cel"
      Height          =   312
      Left            =   1176
      TabIndex        =   8
      Top             =   1308
      Width           =   1152
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register "
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   324
      Left            =   2364
      TabIndex        =   7
      Top             =   1296
      Width           =   828
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   660
      TabIndex        =   6
      Top             =   924
      Width           =   2544
   End
   Begin VB.TextBox Text3 
      Height          =   288
      Left            =   660
      TabIndex        =   4
      Top             =   348
      Width           =   2544
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   1272
      MaxLength       =   6
      TabIndex        =   2
      Top             =   24
      Width           =   1932
   End
   Begin VB.TextBox Text1 
      Height          =   252
      Left            =   660
      MaxLength       =   4
      TabIndex        =   1
      Top             =   36
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Organization:"
      Height          =   312
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   264
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   684
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   1260
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Label Label1 
      Caption         =   "Reg #"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   48
      Width           =   1176
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RegNumber, Namey, Organ As Boolean
Dim Trys As Long

Private Sub Command1_Click()
If Text2 = "Nova-Ice" Then
SetReg "Nova Ice Employee", "Ice Works"
Unload Me
Form1.Show
End If
If Trys = 5 Then
MsgBox "Invaild Keys", vbExclamation, "Ending..."
End
Else
If Text1.Text <> "235" Then
MsgBox "Attempt " & Trys & ") Invaild Registration Number!", vbExclamation, "Invaild Reg Number"
Trys = Trys + 1
Else
If Text2.Text <> "682376" Then
MsgBox "Attempt " & Trys & ") Invaild Registration Number!", vbExclamation, "Invaild Reg Number"
Trys = Trys + 1
Else
SetReg Text3, Text4
MsgBox "Regisration Complete!", vbExclamation, "RegComplete"
Unload Me
Form1.Show
End If
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub


Private Sub Form_Load()
Trys = 1
End Sub


Private Sub Text2_Change()
If Text2.Text <> "" Then
RegNumber = True
Else
If Text2.Text = "" Then
RegNumber = False
End If
End If
End Sub


Private Sub Text3_Change()
If Text3.Text <> "" Then
Namey = True
Else
If Text3.Text = "" Then
Namey = False
End If
End If
End Sub


Private Sub Text4_Change()
If Text4.Text <> "" Then
Organ = True
Else
If Text4.Text = "" Then
Organ = False
End If
End If
End Sub


Private Sub Timer1_Timer()
If RegNumber = True Then
If Namey = True Then
If Organ = True Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End If
End If
End Sub


