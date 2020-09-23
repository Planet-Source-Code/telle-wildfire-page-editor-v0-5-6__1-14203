Attribute VB_Name = "Module1"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Mode As Integer
Public Saved As Boolean
Public SaveLoc
Public Changed As Boolean
Public Declare Function ExtFloodFill Lib "Gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Sub AddScript(Filter, Title)
'Open
On Error Resume Next
With Form1
.CMD1.Filter = Filter
.CMD1.FilterIndex = 1
.CMD1.Action = 1
.CMD1.DialogTitle = Title
Open .CMD1.FileName For Input As 1
.Text1.Text = Input$(LOF(1), 1)
.Text1.SaveFile "c:\windows\system\test.htm", rtfText

Close 1
End With
End Sub

Sub AddTag(Optional Relative As String, Optional Key As String, Optional RelastionShip As String, Optional Text As String, Optional Image As String, Optional SelImage As String)
Form1.Tags.Nodes.Add Relative, RelastionShip, Key, Text, Image, SelImage
End Sub
Function AddTags()
'-------------------------'
'#    Add basic tags     #'
'-------------------------'
 
 Dim intFileNum As Integer, strFilename As String
 strFilename = App.Path & "\tags.dat"
 intFileNum = FreeFile
 Open strFilename For Input As #intFileNum
 Do While Not EOF(intFileNum)
  Line Input #intFileNum, SValue
  If Not Trim(SValue) = "" Then Form1.Tags.Nodes.Add , , , SValue
 Loop
 Close #intFileNum


'-------------------------'
'# Add property families #'
'-------------------------'
Dim tempNode As Node

'add Anchor
 Set tempNode = Form1.Tags.Nodes.Add(, , "ANCHOR", "<a></a>")
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " href=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " target=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " rel=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " rev=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " type=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " charset=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " hreflang=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " media=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("ANCHOR", tvwChild, , " class=" & Chr(34) & Chr(34))
 

'add Applet
 Set tempNode = Form1.Tags.Nodes.Add(, , "Applet", "<applet></applet>")
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " code=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " codebase=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " archive=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " alt=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Applet", tvwChild, , " width=" & Chr(34) & Chr(34))



'add Area
 Set tempNode = Form1.Tags.Nodes.Add(, , "Area", "<area>")
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " shape=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " coords=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " href=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " target=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Area", tvwChild, , " nohref=" & Chr(34) & Chr(34))



'add Base
 Set tempNode = Form1.Tags.Nodes.Add(, , "Base", "<base>")
 Set tempNode = Form1.Tags.Nodes.Add("Base", tvwChild, , " href=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Base", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Base", tvwChild, , " target=" & Chr(34) & Chr(34))



'add Basefont
 Set tempNode = Form1.Tags.Nodes.Add(, , "Basefont", "<basefont>")
 Set tempNode = Form1.Tags.Nodes.Add("Basefont", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Basefont", tvwChild, , " face=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Basefont", tvwChild, , " size=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Basefont", tvwChild, , " color=" & Chr(34) & Chr(34))



'add bgsound
 Set tempNode = Form1.Tags.Nodes.Add(, , "Bgsound", "<bgsound>")
 Set tempNode = Form1.Tags.Nodes.Add("Bgsound", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Bgsound", tvwChild, , " loop=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Bgsound", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Bgsound", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Bgsound", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Bgsound", tvwChild, , " title=" & Chr(34) & Chr(34))



'add Body
 Set tempNode = Form1.Tags.Nodes.Add(, , "Body", "<body></body>")
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " background=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " bgcolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " text=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " link=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " vlink=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " alink=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " leftmargin=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Body", tvwChild, , " topmargin=" & Chr(34) & Chr(34))



'add Col
 Set tempNode = Form1.Tags.Nodes.Add(, , "Col", "<col>")
 Set tempNode = Form1.Tags.Nodes.Add("Col", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Col", tvwChild, , " span=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Col", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Col", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Col", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Col", tvwChild, , " title=" & Chr(34) & Chr(34))



'add Colgroup
 Set tempNode = Form1.Tags.Nodes.Add(, , "Colgroup", "<colgroup>")
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " valign=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " span=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Colgroup", tvwChild, , " title=" & Chr(34) & Chr(34))


'add Div
 Set tempNode = Form1.Tags.Nodes.Add(, , "div", "<div></div>")
 Set tempNode = Form1.Tags.Nodes.Add("div", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("div", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("div", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("div", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("div", tvwChild, , " align=" & Chr(34) & Chr(34))



'add Embed
 Set tempNode = Form1.Tags.Nodes.Add(, , "Embed", "<embed>")
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " hidden=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Embed", tvwChild, , " class=" & Chr(34) & Chr(34))



'add Font
 Set tempNode = Form1.Tags.Nodes.Add(, , "Font", "<font></font>")
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " face=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " size=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " color=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Font", tvwChild, , " title=" & Chr(34) & Chr(34))



'add Form
 Set tempNode = Form1.Tags.Nodes.Add(, , "Form", "<form></form>")
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " action=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " target=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " method=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " enctype=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Form", tvwChild, , " title=" & Chr(34) & Chr(34))



'add Frame
 Set tempNode = Form1.Tags.Nodes.Add(, , "Frame", "<Frame>")
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " scrolling=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " marginwidth=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " framespacing=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " marginheight=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " noresize")
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " frameborder=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frame", tvwChild, , " bordercolor=" & Chr(34) & Chr(34))



'add Frameset
 Set tempNode = Form1.Tags.Nodes.Add(, , "Frameset", "<frameset></frameset>")
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " rows=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " cols=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " frameborder=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " framespacing=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " border=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " bordercolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Frameset", tvwChild, , " title=" & Chr(34) & Chr(34))


'add H1
 Set tempNode = Form1.Tags.Nodes.Add(, , "h1", "<h1></h1>")
 Set tempNode = Form1.Tags.Nodes.Add("h1", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h1", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h1", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h1", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h1", tvwChild, , " class=" & Chr(34) & Chr(34))



'add H2
 Set tempNode = Form1.Tags.Nodes.Add(, , "h2", "<h2></h2>")
 Set tempNode = Form1.Tags.Nodes.Add("h2", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h2", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h2", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h2", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h2", tvwChild, , " class=" & Chr(34) & Chr(34))



'add H3
 Set tempNode = Form1.Tags.Nodes.Add(, , "h3", "<h3></h3>")
 Set tempNode = Form1.Tags.Nodes.Add("h3", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h3", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h3", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h3", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("h3", tvwChild, , " class=" & Chr(34) & Chr(34))



'add HR
 Set tempNode = Form1.Tags.Nodes.Add(, , "HR", "<hr>")
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " size=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " color=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " noshade")
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("HR", tvwChild, , " title=" & Chr(34) & Chr(34))



'add Iframe
 Set tempNode = Form1.Tags.Nodes.Add(, , "Iframe", "<iframe></iframe>")
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " scrolling=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " marginwidth=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " marginheight=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " frameborder=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Iframe", tvwChild, , " title=" & Chr(34) & Chr(34))



'add IMG
 Set tempNode = Form1.Tags.Nodes.Add(, , "IMG", "<img>")
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " alt=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " border=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " hspace=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " vspace=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " ismap=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " usemap=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " dynsrc=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " start=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " loop=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " controls")
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " loopdelay=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " lowsrc=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("IMG", tvwChild, , " class=" & Chr(34) & Chr(34))




'add INPUT
 Set tempNode = Form1.Tags.Nodes.Add(, , "INPUT", "<input>")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " type=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " value=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " size=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " maxlength=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " tabindex=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " notab")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " checked")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " border")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " width")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " height")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " vspace")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " hspace")
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " accept=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("INPUT", tvwChild, , " title=" & Chr(34) & Chr(34))



'add LINK
 Set tempNode = Form1.Tags.Nodes.Add(, , "LINK", "<link>")
 Set tempNode = Form1.Tags.Nodes.Add("LINK", tvwChild, , " rel=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("LINK", tvwChild, , " href=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("LINK", tvwChild, , " type=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("LINK", tvwChild, , " title=" & Chr(34) & Chr(34))



'add MAP
 Set tempNode = Form1.Tags.Nodes.Add(, , "MAP", "<map></map>")
 Set tempNode = Form1.Tags.Nodes.Add("MAP", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MAP", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MAP", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MAP", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MAP", tvwChild, , " title=" & Chr(34) & Chr(34))



'add MARAQUEE
 Set tempNode = Form1.Tags.Nodes.Add(, , "MARAQUEE", "<marquee></marquee>")
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " behavior=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " direction=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " bgcolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " hspace=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " vspace=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " loop=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " scrollamount=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " scrolldelay=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MARAQUEE", tvwChild, , " title=" & Chr(34) & Chr(34))



'add META
 Set tempNode = Form1.Tags.Nodes.Add(, , "META", "<meta>")
 Set tempNode = Form1.Tags.Nodes.Add("META", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("META", tvwChild, , " http-equiv=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("META", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("META", tvwChild, , " content=" & Chr(34) & Chr(34))


'add MULTICOL
 Set tempNode = Form1.Tags.Nodes.Add(, , "MULTICOL", "<multicol>")
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " cols=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " gutter=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("MULTICOL", tvwChild, , " title=" & Chr(34) & Chr(34))



'add P
 Set tempNode = Form1.Tags.Nodes.Add(, , "PARAGRAPH", "<p>")
 Set tempNode = Form1.Tags.Nodes.Add("PARAGRAPH", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("PARAGRAPH", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("PARAGRAPH", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("PARAGRAPH", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("PARAGRAPH", tvwChild, , " title=" & Chr(34) & Chr(34))




'add Param
 Set tempNode = Form1.Tags.Nodes.Add(, , "Param", "<param></param>")
 Set tempNode = Form1.Tags.Nodes.Add("Param", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Param", tvwChild, , " name=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Param", tvwChild, , " value=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Param", tvwChild, , " type=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("Param", tvwChild, , " valuetype=" & Chr(34) & Chr(34))




'add SCRIPT
 Set tempNode = Form1.Tags.Nodes.Add(, , "SCRIPT", "<script></script>")
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " type=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " language=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " for=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " defer=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " runat=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SCRIPT", tvwChild, , " charset=" & Chr(34) & Chr(34))



'add SOUND
 Set tempNode = Form1.Tags.Nodes.Add(, , "SOUND", "<sound></sound>")
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " src=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " loop=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " delay=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SOUND", tvwChild, , " title=" & Chr(34) & Chr(34))



'add SPAN
 Set tempNode = Form1.Tags.Nodes.Add(, , "SPAN", "<span></span>")
 Set tempNode = Form1.Tags.Nodes.Add("SPAN", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SPAN", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SPAN", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SPAN", tvwChild, , " title=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("SPAN", tvwChild, , " align=" & Chr(34) & Chr(34))



'add STYLE
 Set tempNode = Form1.Tags.Nodes.Add(, , "STYLE", "<style></style>")
 Set tempNode = Form1.Tags.Nodes.Add("STYLE", tvwChild, , " type=" & Chr(34) & Chr(34))



'add TABLE
 Set tempNode = Form1.Tags.Nodes.Add(, , "TABLE", "<table></table>")
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " cellpadding=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " border=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " valign=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " cellspacing=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " nowrap")
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " background=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " bgcolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " bordercolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " cols=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " clear=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " frame=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " rules=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TABLE", tvwChild, , " title=" & Chr(34) & Chr(34))



'add TD
 Set tempNode = Form1.Tags.Nodes.Add(, , "TD", "<td></td>")
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " bgcolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " bordercolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " background=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " rowspan=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " colspan=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " valign=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " nowrap=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TD", tvwChild, , " title=" & Chr(34) & Chr(34))



'add TH
 Set tempNode = Form1.Tags.Nodes.Add(, , "TH", "<th></th>")
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " width=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " height=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " rowspan=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " colspan=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " valign=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " nowrap=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " bgcolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " bordercolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " background=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TH", tvwChild, , " title=" & Chr(34) & Chr(34))



'add TR
 Set tempNode = Form1.Tags.Nodes.Add(, , "TR", "<tr></tr>")
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " bgcolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " bordercolor=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " bordercolorlight=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " bordercolordark=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " align=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " valign=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " nowrap=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " style=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " id=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " class=" & Chr(34) & Chr(34))
 Set tempNode = Form1.Tags.Nodes.Add("TR", tvwChild, , " title=" & Chr(34) & Chr(34))
 'tempNode.Expanded = False

End Function

Sub ColorCode()
Dim OldString, NewString, OldLetter, NewLetter As String
OldString = Form1.Text1.Text
'Form1.Text1.HideSelection = True
OldLetter = "<"
Form1.Text1.SelColor = vbRed
NewLetter = "<"
NewString = Replace(OldString, OldLetter, NewLetter)
Dim OldString1, NewString1, OldLetter1, NewLetter1 As String

Form1.Text1.HideSelection = False
OldString1 = Form1.Text1.Text
Form1.Text1.HideSelection = True
OldLetter1 = ">"
Form1.Text1.SelColor = vbBlue
NewLetter1 = ">"
'Form1.Text1.SelColor = vbBlack
NewString1 = Replace(OldString1, OldLetter1, NewLetter1)
Form1.Text1.HideSelection = False
End Sub

Sub GetReg(User As Label, Company As Label)
User = GetSetting("IW-PE", "Reg", "User")
Company = GetSetting("IW-PE", "Reg", "Company")
End Sub

Sub GetRegStuff()
On Error Resume Next
    FN = FreeFile
    Open "c:\windows\system\WildFire.LT" For Input As FN
    EOF (FN)
    Line Input #FN, nextline$
    Form2.LT = nextline$
    FN = FreeFile
    Open "c:\windows\system\WildFire.regnumber" For Input As FN
    EOF (FN)
    Line Input #FN, nextline$
    Form2.RegNumber = nextline$
End Sub

Sub timeout(interval)
'This pauses a program
'The same as a Pause sub
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Sub OnTop(Form As Form)
SetWinOnTop = SetWindowPos(Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub NewPage()
Dim MsG, Save, Open1
If Changed = True Then
MsG = MsgBox("Would You Like To Save Before Starting A New Page?", vbYesNoCancel + vbQuestion, "Save")
If MsG = vbYes Then
Page "Webpage File|*.html;*.htm|Text File|*.txt|All Files|*.*|", "Save Page", 2
Else
If MsG = vbNo Then
Unload Form1
Form4.Show vbModal, Form1
Form1.Text1 = ""
Else
If MsG = vbCancel Then
Exit Sub
ElseIf Changed = False Then
Unload Form1
Form4.Show vbModal, Form1
Form1.Text1 = ""
End If
End If
End If
End If
End Sub
Sub Page(Filter, name, Action)
Form1.CMD1.Filter = Filter
Form1.CMD1.DialogTitle = name
Form1.CMD1.Action = Action
End Sub
Sub OpenPage()
'Open
On Error Resume Next
With Form1
.CMD1.Filter = "Webpage File|*.html;*.htm|Text File|*.txt|All Files|*.*.*|"
.CMD1.FilterIndex = 1
.CMD1.Action = 1
Open .CMD1.FileName For Input As 1
.Text1.Text = Input$(LOF(1) )
.Text1.SaveFile "c:\windows\system\test.htm", rtfText
.WebBrowser1.Navigate "c:\windows\system\test.htm"

Close 1
End With
End Sub
Sub SavePage()
'Save As
With Form1
On Error Resume Next
If Saved = False Then
.CMD1.Filter = "Webpage File|*.html;*.htm|Text File|*.txt|All Files|*.*.*|"
.CMD1.FilterIndex = 1
.CMD1.Action = 2
Open .CMD1.FileName For Output As #1
Print #1, .Text1.Text
Close #1
SaveLoc = .CMD1.FileName
ElseIf Saved = True Then
Open SaveLoc For Output As #1
Print #1, .Text1.Text
Close #1
End With
End Sub
Sub NewMe()
With Form1
.Text1.HideSelection = True
.Text1.SelColor = vbBlue

.Text1.SelText = "<html>" & Chr(13) & Chr(10)
.Text1.SelText = "<title>"
.Text1.SelColor = vbRed
.Text1.SelText = "Title Here"
.Text1.SelColor = vbBlue
.Text1.SelText = "</title>"
.Text1.SelColor = vbGreen
.Text1.SelText = "<body>"
.Text1.HideSelection = False
End With
End Sub
Sub SetReg(User As String, Company As String)
SaveSetting "IW-PE", "Reg", "User", User
SaveSetting "IW-PE", "Reg", "Company", Company
End Sub
