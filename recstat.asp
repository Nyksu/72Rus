<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var  id=parseInt(Request("rid"))
if (isNaN(id)) {id=0}
var sql=""
var name=""
var coment=""
var rid=0
var ash=0
var ush=0
var ii=0
var ddt=""

if (id!=0) {
	sql="Select * from banblock where state=0 and id="+id
	Records.Source=sql
	Records.Open()
	if (!Records.Eof) {
		name=Records("NAME").Value
		coment=Records("COMENT").Value
	} else {id=0}
	Records.Close()
}
%>

<html>
<head>
<title>Статистика рекламмных мест</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<%
if (id==0) { // список рекламмных мест
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> | 
        <a href="modrec.asp"><font size="2">Управление рекламными местами</font></a><font size="2">  | 
        <a href="addrec.asp">Добавить рекламмное место</a></font> <font size="2">| 
        <a href="banstat.asp">Статистика по банерам</a> |</font></p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#0000FF" size="4">Статистика рекламных блоков</font></b></p>
<p align="center">&nbsp;</p>
<p>Таблица рекламмных блоков и статистика за сегодня</p>
<%
	sql="Select t1.*, t2.ALLSHOW, t2.UNISHOW from banblock t1, blockbanst t2 where t1.state=0 and t1.id=t2.banblock_id and t2.dt='TODAY' order by t1.name"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		name=Records("NAME").Value
		coment=Records("COMENT").Value
		rid=Records("ID").Value
		ash=Records("ALLSHOW").Value
		ush=Records("UNISHOW").Value
		ii=ii+1
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td width="5%"><font size="1"><b>&nbsp;&nbsp;<%=ii%>.</b></font></td>
    <td><font size="1" color="#0033CC"><b><font size="2"><a href="recstat.asp?rid=<%=rid%>"><%=name%></a></font> 
      (</b><font color="#000066"><%=coment%></font><b>)</b></font></td>
    <td width="10%"> 
      <div align="center"><font size="2"><b><font color="#0000CC"><%=ash%></font></b> - хит.</font></div>
    </td>
    <td width="10%"> 
      <div align="center"><font size="2"><b><font color="#0000CC"><%=ush%></font></b> - уник.</font></div>
    </td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
} else { // статистика по рекламмному 
%>
<p><a href="modrec.asp"><font size="2">Управление рекламными местами</font></a><font size="2"> 
  | <a href="recstat.asp">Вернуться к списку рекламмных блоков</a> | <a href="index.asp">Вернуться 
  на главную страницу</a> | <a href="banstat.asp">Статистика по банерам</a></font></p>
<p>Статистика рекламмного места за 30 дней: <font color="#0000CC"><%=name%></font></p>
<p><font size="2"><b><font color="#0066CC">(<%=coment%>)</font></b></font></p>
<%
	sql="Select * from BLOCKBANST where BANBLOCK_ID="+id+" and DT>'TODAY'-30 order by dt desc"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		ash=Records("ALLSHOW").Value
		ush=Records("UNISHOW").Value
		ddt=Records("DT").Value
		ii=ii+1
%>
<table width="40%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td width="5%"><font size="1"><b>&nbsp;&nbsp;<%=ii%>.</b></font></td>
    <td><font size="2" color="#0033CC"><%=ddt%></font></td>
    <td width="30%"> 
      <div align="center"><font size="2"><b><font color="#0000CC"><%=ash%></font></b> - хит.</font></div>
    </td>
    <td width="30%"> 
      <div align="center"><font size="2"><b><font color="#0000CC"><%=ush%></font></b> - уник.</font></div>
    </td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
}
%>
</body>
</html>
