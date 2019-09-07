<html>
<head>
<title>SoftLine Hit Counter</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<% 'SoftLine counter is written by Bruce Corkhill ©2002
'If you want your own hit counter then goto http://www.coolscripts.cjb.net or http://www.softline.cjb.net %>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center"><b><br>
  </b>
  <%
SET MyFileObject = Server.CreateObject("Scripting.FileSystemObject")

'Edit this loction accordingly
SET MyCouNtFile = MyFileObject.OpenTextFile(Server.MapPath("count.txt"))

IF NOT MyCountFile.AtEndOfStream THEN

'Read the Visitor no.
Visitor = MyCountFile.Readline

Response.Write"This page has been visited "
Response.Write(Visitor)
Response.Write" times"

End IF


'close object
MyCountFile.close

SET MyFileObject = Server.CreateObject("Scripting.FileSystemObject")
SET MyCouNtFile = MyFileObject.CreateTextFile(Server.MapPath("count.txt"))
MyCountFile.WriteLine Visitor+1
MyCountFile.Close


%> </div>
<p align="center">
<font size="1" face="Verdana">Softline Hit Counter</font></p>
</body>
</html>