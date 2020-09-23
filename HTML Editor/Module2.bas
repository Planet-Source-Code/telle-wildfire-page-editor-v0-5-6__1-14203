Attribute VB_Name = "Module2"
Option Explicit
Sub FrameBuild()
With Form1
.Text1.SelText = "<frameset rows=" & Form6.Rows.Caption & ",*>"
.Text1.SelText = "<frame src=" & Form6.TFURL.Text & ">"
.Text1.SelText = "<frameset cols=" & Form6.Label3.Caption & "," & Form6.Label4.Caption & ">"
.Text1.SelText = "<frame src=" & Form6.LFURL.Text & " name=" & Form6.LFName.Text & ">"
.Text1.SelText = "<frame src=" & Form6.RFURL.Text & " name=" & Form6.RFName.Text & ">"
.Text1.SelText = "</frameset>"
Unload Form6
End With
End Sub

Sub BuildBlank(Title As String, Optional BgImage As String, Optional BgSound As String)
Dim Lines As Long
With Form1.Text1
.SelText = "<HTML>" & vbCrLf
.SelText = "    <HEAD>" & vbCrLf
.SelText = "        <TITLE>" & Title & "</TITLE>" & vbCrLf
.SelText = "    </HEAD>" & vbCrLf & vbCrLf
.SelText = "<Body BackGround = """ & BgImage & """ >" & vbCrLf
.SelText = "    <BGSound = """ & BgSound & """>" & vbCrLf
Do Until Lines = 20
.SelText = vbCrLf
Lines = Lines + 1
Loop
.SelText = "</BODY>" & vbCrLf
.SelText = "</HTML>" & vbCrLf
End With
End Sub
