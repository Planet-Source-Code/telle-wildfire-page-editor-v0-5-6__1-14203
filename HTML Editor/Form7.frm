VERSION 5.00
Begin VB.Form Form7 
   ClientHeight    =   2145
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   1560
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   1560
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2895
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   810
      Width           =   750
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = 720
End Sub
