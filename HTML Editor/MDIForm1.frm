VERSION 5.00
Object = "{30F21E58-B687-11D4-9A46-444553540001}#2.0#0"; "IMAGEFX.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Image Wizard Beta 1 Demo"
   ClientHeight    =   5715
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7545
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   5715
      Left            =   6030
      ScaleHeight     =   5655
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Width           =   1515
      Begin VB.PictureBox Color1 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   105
         ScaleHeight     =   495
         ScaleWidth      =   1005
         TabIndex        =   17
         Top             =   4035
         Width           =   1005
      End
      Begin VB.PictureBox Color2 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   375
         ScaleHeight     =   495
         ScaleWidth      =   1005
         TabIndex        =   18
         Top             =   4200
         Width           =   1005
      End
      Begin VB.OptionButton arrow 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   45
         Picture         =   "MDIForm1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Arrow"
         Top             =   1350
         Width           =   360
      End
      Begin VB.OptionButton optMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   6
         Left            =   405
         Picture         =   "MDIForm1.frx":07FC
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Drop"
         Top             =   1350
         Width           =   360
      End
      Begin VB.OptionButton optMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   11
         Left            =   45
         Picture         =   "MDIForm1.frx":0CEE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Text"
         Top             =   3150
         Width           =   360
      End
      Begin VB.OptionButton Pen 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   405
         Picture         =   "MDIForm1.frx":11E0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Pen"
         Top             =   3150
         Width           =   360
      End
      Begin VB.OptionButton optMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   10
         Left            =   45
         Picture         =   "MDIForm1.frx":16D2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Spray"
         Top             =   2790
         Width           =   360
      End
      Begin VB.OptionButton grandient 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   4
         Left            =   405
         Picture         =   "MDIForm1.frx":1BC4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Grandient"
         Top             =   2790
         Width           =   360
      End
      Begin VB.OptionButton special 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   9
         Left            =   45
         Picture         =   "MDIForm1.frx":20B6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2430
         Width           =   360
      End
      Begin VB.OptionButton brush 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   405
         Picture         =   "MDIForm1.frx":25A8
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Brush"
         Top             =   2415
         Width           =   360
      End
      Begin VB.OptionButton optMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   8
         Left            =   45
         Picture         =   "MDIForm1.frx":2A9A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Shape"
         Top             =   2070
         Width           =   360
      End
      Begin VB.OptionButton optMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   2
         Left            =   405
         Picture         =   "MDIForm1.frx":2F8C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Select Region"
         Top             =   2070
         Width           =   360
      End
      Begin VB.OptionButton fill 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   45
         Picture         =   "MDIForm1.frx":347E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fill"
         Top             =   1710
         Width           =   360
      End
      Begin VB.OptionButton optMenu 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   405
         Picture         =   "MDIForm1.frx":3970
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Magnify"
         Top             =   1710
         Width           =   360
      End
      Begin VB.PictureBox Color 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   735
         ScaleHeight     =   1335
         ScaleWidth      =   720
         TabIndex        =   2
         ToolTipText     =   "Select This Color!"
         Top             =   0
         Width           =   720
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   0
         Picture         =   "MDIForm1.frx":3E62
         ScaleHeight     =   96
         ScaleMode       =   0  'User
         ScaleWidth      =   45
         TabIndex        =   1
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Inten 
         Caption         =   "10"
         Height          =   225
         Left            =   855
         TabIndex        =   16
         Top             =   3600
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Intensity:"
         Height          =   210
         Left            =   60
         TabIndex        =   15
         Top             =   3600
         Width           =   870
      End
   End
   Begin Image FX.ImageFX imgCtl 
      Left            =   7320
      Top             =   3495
      _ExtentX        =   7673
      _ExtentY        =   3678
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "Filters"
      Begin VB.Menu mnuaddN 
         Caption         =   "AddNoise"
      End
      Begin VB.Menu mnuBlur 
         Caption         =   "Blur"
      End
      Begin VB.Menu mnuColorize 
         Caption         =   "Colorize"
      End
      Begin VB.Menu mnusds 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDarken 
         Caption         =   "Darken"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFlipHorizontal 
         Caption         =   "FlipHorizontal"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGrayScale 
         Caption         =   "GrayScale"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "Invert"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLighten 
         Caption         =   "Lighten"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPixelate 
         Caption         =   "Pixelate"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xy, yx As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Private Sub Command1_Click()
imgCtl.Blur Form8.Picture1, Inten

End Sub

Private Sub Command2_Click()
imgCtl.AddNoise Form8.Picture1, Inten
End Sub

Private Sub Command3_Click()
imgCtl.Colorize Form8.Picture1, Color.BackColor


End Sub

Private Sub arrow_Click(Index As Integer)
Mode = 0
End Sub

Private Sub brush_Click()
Mode = 3
oldxpo = -1
oldypo = -1
End Sub

Private Sub Color_Click()
Color1.BackColor = Color.BackColor
End Sub

Private Sub Color2_Click()
c1 = Color1.BackColor
c2 = Color2.BackColor
Color2.BackColor = c1
Color1.BackColor = c2
End Sub

Private Sub fill_Click()
Mode = 4

End Sub

Private Sub grandient_Click(Index As Integer)
Mode = 5
Dim m As New clsGradient
m.Color1 = Color1
m.Color2 = Color2

m.Angle = "95"
m.Draw Form8.Picture1
End Sub

Private Sub Inten_Click()
blah = InputBox("Enter New Intensity:", "Enter Intensity", "10")
Inten.Caption = blah
End Sub

Private Sub MDIForm_Load()
Form8.Show
MsgBox "This Is A Demo!" & vbCrLf & "The Following Functions Work:" & vbCrLf & "AddNoise" & vbCrLf & "Blur" & vbCrLf & "Colorize" & vbCrLf & "Line" & vbCrLf & "Pen" & vbCrLf & "__________________" & vbCrLf & "Note: Grandients Was Not Completed In This Beta 1" & vbCrLf & "Demo, The Other Demos Will Be More Complete", vbInformation, "Information"
End Sub

Private Sub mnuaddN_Click()
imgCtl.AddNoise Form8.Picture1, Inten
End Sub

Private Sub mnuBlur_Click()
imgCtl.Blur Form8.Picture1, Inten
End Sub


Private Sub mnuColorize_Click()
imgCtl.Colorize Form8.Picture1, Color.BackColor
End Sub


Private Sub mnuNew_Click()
Load Form8
End Sub

Private Sub optMenu_Click(Index As Integer)
Mode = 1
oldxpo = -1
oldypo = -1
End Sub


Private Sub optMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Mode = 3
oldxpo = -1
oldypo = -1
End Sub


Private Sub optMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Mode = 3
oldxpo = -1
oldypo = -1
End Sub


Private Sub Pen_Click()
Mode = 1

End Sub

Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
colory = Picture8.Point(X, Y)
Color.BackColor = colory
xy = Picture8.Point(X, Y)
End Sub

Private Sub special_Click(Index As Integer)
MsgBox "Not In Demo Version", vbInformation, "Error"

End Sub


