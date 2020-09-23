VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Frame Wizard"
   ClientHeight    =   6630
   ClientLeft      =   2070
   ClientTop       =   1245
   ClientWidth     =   4605
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   6630
   ScaleWidth      =   4605
   Begin VB.CommandButton SSCommand3 
      Caption         =   "Complete"
      Height          =   390
      Left            =   3120
      TabIndex        =   27
      Top             =   6150
      Width           =   1455
   End
   Begin VB.CommandButton SSCommand1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   15
      TabIndex        =   26
      Top             =   6165
      Width           =   1440
   End
   Begin VB.Frame Frame3 
      Caption         =   "Right Frame"
      Height          =   1692
      Left            =   3120
      TabIndex        =   18
      Top             =   4440
      Width           =   1452
      Begin VB.TextBox RFName 
         Height          =   288
         Left            =   120
         TabIndex        =   20
         Text            =   "RightFrame"
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox RFURL 
         Height          =   288
         Left            =   120
         TabIndex        =   19
         Text            =   "Http://"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label9 
         Caption         =   "URL:"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   852
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Top Frame"
      Height          =   1692
      Left            =   1560
      TabIndex        =   13
      Top             =   4440
      Width           =   1452
      Begin VB.TextBox TFName 
         Height          =   288
         Left            =   120
         TabIndex        =   15
         Text            =   "TopFrame"
         Top             =   480
         Width           =   1212
      End
      Begin VB.TextBox TFURL 
         Height          =   288
         Left            =   120
         TabIndex        =   14
         Text            =   "Http://"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   972
      End
      Begin VB.Label Label7 
         Caption         =   "URL:"
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   852
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Left Frame"
      Height          =   1692
      Left            =   0
      TabIndex        =   8
      Top             =   4440
      Width           =   1452
      Begin VB.TextBox LFURL 
         Height          =   288
         Left            =   120
         TabIndex        =   12
         Text            =   "Http://"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox LFName 
         Height          =   288
         Left            =   120
         TabIndex        =   10
         Text            =   "LeftFrame"
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label6 
         Caption         =   "URL:"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   972
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4212
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4572
      Begin VB.OptionButton Option3 
         Caption         =   "Rows"
         Height          =   192
         Left            =   1800
         TabIndex        =   23
         Top             =   3720
         Width           =   972
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000E&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3435
         ScaleWidth      =   4515
         TabIndex        =   3
         Top             =   0
         Width           =   4575
         Begin VB.Label Row 
            BackColor       =   &H00000000&
            Caption         =   "Label11"
            Height          =   48
            Left            =   0
            TabIndex        =   25
            Top             =   360
            Width           =   4572
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            ForeColor       =   &H000000FF&
            Height          =   3492
            Left            =   480
            TabIndex        =   5
            Top             =   0
            Width           =   50
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000040&
            ForeColor       =   &H000000FF&
            Height          =   50
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   4572
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vertical"
         Height          =   192
         Left            =   120
         TabIndex        =   2
         Top             =   3720
         Width           =   1092
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Horizantal"
         Height          =   192
         Left            =   3360
         TabIndex        =   1
         Top             =   3720
         Width           =   1092
      End
      Begin VB.Label Rows 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         Height          =   252
         Left            =   1800
         TabIndex        =   24
         Top             =   3960
         Width           =   972
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50%"
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   3960
         Width           =   1092
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50%"
         Height          =   252
         Left            =   3360
         TabIndex        =   6
         Top             =   3960
         Width           =   1092
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim V, H, R As Boolean

Private Sub Form_Load()
H = False
V = False
'M = Picture1
Picture2.ScaleHeight = 100
Picture2.ScaleWidth = 100
Label1.Left = 50
Label2.Top = 50
End Sub

Private Sub Label2_Click()
Dim MsG, Button
MsG = MsgBox("Delete This Frame?", vbYesNo + vbQuestion, "Delete Frame")
If MsG = vbYes Then Label4.Caption = "0 %" & Label2.BackStyle = 0
If MsG = vbNo Then Exit Sub
End Sub

Private Sub Option1_Click()
V = True
H = False
End Sub

Private Sub Option2_Click()
H = True
V = False
End Sub

Private Sub Option3_Click()
R = True
H = False
V = False
End Sub

Private Sub Picture2_Click()
If Option1.Value = True Then
 If V = False Then
  V = True
 Else
 V = False
 End If
End If


If Option2.Value = True Then
 If H = False Then
  H = True
 Else
  H = False
 End If
End If

If Option3.Value = True Then
 If R = False Then
  V = True
 Else
 R = False
 End If
End If


End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If V = True Then

Dim Procentage

Label1.Left = X + 1
   Procentage = Int(Label1.Left - 1)
    If Procentage < 0 Then Procentage = 0
    If Procentage > 100 Then Procentage = 100

Label3.Caption = Procentage & "%"
Else
If H = True Then
Label2.Top = Y + 1

Dim Percentage2
Percentage2 = Int(Label2.Top - 1)
If Percentage2 < 0 Then Percentage2 = 0
If Percentage2 > 100 Then Percentage2 = 100
Label4.Caption = Percentage2 & "%"
Else
If R = True Then
Row.Top = Y + 1
Dim Percentage3
Percentage3 = Int(Row.Top - 1)
If Percentage3 < 0 Then Percentage3 = 0
If Percentage3 > 100 Then Percentage3 = 100
Rows.Caption = Percentage3 & "%"
End If
End If
End If
End Sub


Private Sub SSCommand1_Click()
Unload Me
End Sub

Private Sub SSCommand3_Click()
FrameBuild
End Sub


