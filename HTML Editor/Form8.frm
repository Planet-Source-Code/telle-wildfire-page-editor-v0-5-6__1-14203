VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Image"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   0
      ScaleHeight     =   3165
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   15
      Width           =   4665
      Begin VB.Line linea 
         Visible         =   0   'False
         X1              =   1080
         X2              =   1950
         Y1              =   1395
         Y2              =   1680
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldXPos As Integer
Dim OldYPos As Integer
Dim wid As Integer
Dim hei As Integer
Dim Color As ColorConstants
Dim centx, centy As Integer
Dim oldxpo, oldypo As Integer
Private Sub Form_Activate()
Mode = 1
End Sub

Private Sub Form_Load()
    OldXPos = -1
    OldYPos = -1
Form_Resize
Mode = 1
End Sub

Private Sub Form_Resize()
Picture1.Height = Me.ScaleHeight
Picture1.Width = Me.ScaleWidth
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Mode = 1 Then
If Button = 1 Then
        Picture1.PSet (X, Y), Form2.Color1.BackColor
        OldXPos = X
        OldYPos = Y
Picture1.PSet (X, Y), Form2.Color1.BackColor
Exit Sub
End If

ElseIf Mode = 2 Then
centx = X
centy = Y


ElseIf Mode = 3 Then
If oldypo = -1 Then
Picture1.PSet (X, Y), Form2.Color1.BackColor
oldxpo = X
oldypo = Y
Exit Sub
End If
Picture1.Line (oldxpo, oldypo)-(X, Y), Form2.Color1.BackColor
oldxpo = X
oldypo = Y

End If

If Mode = 4 Then
' Picture1.hDC, X, Y, Form2.Color1.BackColor, 0
End If

If Mode = 5 Then
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = 0
    If Button = 0 Then Exit Sub 'Do no anything if no mousebutton is pressed
    
    If Mode = 1 Then
    Select Case Button
        Case 1  'Left button
          If OldXPos = -1 Then
          OldXPos = X
          OldYPos = Y
          End If


            
            Me.Picture1.Line (OldXPos, OldYPos)-(X, Y), Form2.Color1.BackColor
            

            
H:
            OldXPos = X
            OldYPos = Y
    End Select
    Exit Sub
    End If
    If Mode = 2 Then
linea.X1 = centx
linea.Y1 = centy
linea.X2 = X
linea.Y2 = Y
    linea.Visible = True
    End If
    
        
        If Mode = 3 Then
    End If
        
        If Mode = 4 Then
        ' Picture1.hDC, X, Y, Form2.Color1.BackColor, 1
    End If
    
        If Mode = 5 Then
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Mode = 1 Then
If Button And 1 Then
        Picture1.PSet (X, Y), Form2.Color1.BackColor
        OldXPos = X
        OldYPos = Y
Picture1.PSet (X, Y), Form2.Color1.BackColor
Exit Sub
End If
End If
If Mode = 2 Then
centx = X
centy = Y
End If

If Mode = 3 Then
If oldypo = -1 Then
Picture1.PSet (X, Y), Form2.Color1.BackColor
oldxpo = X
oldypo = Y
Exit Sub
End If
Picture1.Line (oldxpo, oldypo)-(X, Y), Form2.Color1.BackColor
oldxpo = X
oldypo = Y

End If

If Mode = 4 Then
' Picture1.hDC, X, Y, Form2.Color1.BackColor, 0
End If

If Mode = 5 Then
End If
End Sub
