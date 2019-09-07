<%
   Set MyPageCounter = Server.CreateObject("IISSample.PageCounter")
   HitMe = MyPageCounter.Hits
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%    Set MyPageCounter = Server.CreateObject("IISSample.PageCounter")
%>
This Web page has been viewed <%= MyPageCount.Hits %> times.
<P>
Page Myscript.asp has been viewed <%= MyPageCounter.Hits("/VirtualDir1/Myscript.asp") %> times.
</body>
</html>
