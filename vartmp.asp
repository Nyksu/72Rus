<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var res=0
res=AddShows(0)

%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>
<body bgcolor="#FFFFFF" text="#000000">

<%if (res>0) {%><p>Все робит!!</p><%} else {%><p>Где-то проблемы</p><%}%>

<%
// после блока вставить строку:  AddShows(0)
%>

</body>
</html>
