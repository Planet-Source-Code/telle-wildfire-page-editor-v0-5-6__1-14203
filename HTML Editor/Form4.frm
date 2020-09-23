VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings Wizard"
   ClientHeight    =   4545
   ClientLeft      =   3210
   ClientTop       =   1755
   ClientWidth     =   3960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Templete"
      Height          =   2052
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   3972
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Blank"
         Height          =   765
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form4.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   330
         Width           =   852
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Safty"
         Height          =   765
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form4.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   330
         Width           =   852
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000004&
         Caption         =   "FAQ"
         Height          =   765
         Left            =   2040
         Picture         =   "Form4.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   330
         Width           =   852
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000004&
         Caption         =   "Info/About"
         Height          =   765
         Left            =   3000
         Picture         =   "Form4.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   330
         Width           =   852
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contact"
         Height          =   732
         Left            =   120
         Picture         =   "Form4.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   852
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000004&
         Caption         =   "Gallary"
         Height          =   732
         Left            =   1080
         Picture         =   "Form4.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   852
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H80000004&
         Caption         =   "List"
         Height          =   732
         Left            =   2040
         Picture         =   "Form4.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "Blank"
         Height          =   732
         Left            =   3000
         TabIndex        =   20
         Top             =   1200
         Width           =   852
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Can&cel"
      Height          =   372
      Left            =   2040
      TabIndex        =   11
      Top             =   4080
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O&k"
      Default         =   -1  'True
      Height          =   372
      Left            =   3000
      TabIndex        =   10
      Top             =   4080
      Width           =   972
   End
   Begin VB.Frame Frame2 
      Caption         =   "Form Properties"
      Height          =   732
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   3972
      Begin VB.TextBox Text3 
         Height          =   288
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   2052
      End
      Begin VB.Label Label3 
         Caption         =   "Webpage Title:"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Background Properties"
      Height          =   972
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3972
      Begin VB.CheckBox Check2 
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   252
      End
      Begin VB.CheckBox Check1 
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   252
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1800
         TabIndex        =   4
         Text            =   "Http://"
         Top             =   600
         Width           =   2052
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1800
         TabIndex        =   2
         Text            =   "Http://"
         Top             =   240
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Background Sound:"
         Height          =   252
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Background Image:"
         Height          =   252
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1572
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BgSound, BgImage As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text1.Enabled = True
Else
If Check1.Value = 0 Then
Text1.Enabled = False
End If
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Text2.Enabled = True
Else
If Check2.Value = 0 Then
Text2.Enabled = False
End If
End If
End Sub

Private Sub Command1_Click()
If Label4 = "Blank" Then
BuildBlank Text3, Text1, Text2
Form1.Show
Unload Me
End If
 End Sub

Private Sub Command10_Click()
Label4.Caption = Command10.Caption
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Label4.Caption = Command3.Caption

End Sub

Private Sub Command4_Click()
Label4.Caption = Command4.Caption

End Sub

Private Sub Command5_Click()
Label4.Caption = Command5.Caption

End Sub

Private Sub Command6_Click()
Label4.Caption = Command6.Caption

End Sub

Private Sub Command7_Click()
Label4.Caption = Command7.Caption

End Sub

Private Sub Command9_Click()
Label4.Caption = Command9.Caption

End Sub

