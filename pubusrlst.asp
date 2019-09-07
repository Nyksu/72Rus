<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!
var usok=false
var sql=""

if (String(Session("id_mem"))=="undefined") {
Session("backurl")="pubusrlst.asp"
Response.Redirect("login.asp")
}

if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
	sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
	Records.Source=sql
	Records.Open()
	if (!Records.EOF) {
		usok=true
	}
	Records.Close()
} else {
	usok=true
}

if (!usok) {Response.Redirect("index.asp")}

var sminame=""
Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

var  st=parseInt(Request("st"))
if (isNaN(st)) {st=0}

%>

<html>
<head>
<title>Пользователи СМИ</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="index.asp">Вернуться на главную страницу</a> | <a href="pubarea.asp">Кабинет 
        редактора</a> | <a href="pubusrlst.asp?st=<%=st==0?1:0%>">Пользователи 
        <%if (st==0) {%>
        не 
        <%}%>
        активированые</a> | <a href="addpubusr.asp">Добавить пользователя</a></p>
    </td>
  </tr>
</table>
<h1 align="center">Пользователи СМИ.</h1>
<p align="center"><font color="#0000FF"><b><%=sminame%></b></font></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#FF0000"> 
  <%if (st==0) {%>
  (Активные) 
  <%} else {%>
  (Не активированные) 
  <%}%>
  </font></b></p>
<table width="95%" border="1" align="center" bordercolor="#FFFFFF">
  <tr bgcolor="#CCCCCC" bordercolor="#333333"> 
    <td bordercolor="#333333" width="25%"> 
      <div align="center"> 
        <p><b>ФИО</b></p>
      </div>
    </td>
    <td width="15%"> 
      <div align="center"> 
        <p><b>Псевдоним</b></p>
      </div>
    </td>
    <td width="20%"> 
      <div align="center"> 
        <p><b>Статус</b></p>
      </div>
    </td>
    <td width="15%"> 
      <div align="center"> 
        <p><b>E-mail</b></p>
      </div>
    </td>
    <td> 
      <div align="center"> 
        <p><b>Управление</b></p>
      </div>
    </td>
  </tr>
  <%
 var fio=""
 var nik=""
 var tpnam=""
 var email=""
 var id=0
 sql="Select t1.*, t2.name as tpnam from SMI_USR t1, USR_TYPE t2 where t1.state="+st+" and t1.usr_type_id=t2.id and t1.smi_id="+smi_id
Records.Source=sql
Records.Open()
while (!Records.EOF )
{
fio=String(Records("NAME").Value)
nik=String(Records("NIK_NAME").Value)
tpnam=String(Records("TPNAM").Value)
email=TextFormData(Records("EMAIL").Value,"")
id=Records("ID").Value
 %>
  <tr bordercolor="#333333"> 
    <td bordercolor="#333333" width="25%"> 
      <p align="center"><font color="#0000FF"><%=fio%></font></p>
    </td>
    <td width="15%"> 
      <p align="center"><font color="#0000FF"><%=nik%></font></p>
    </td>
    <td width="20%"> 
      <p align="center"><%=tpnam%></p>
    </td>
    <td width="15%"> 
      <p align="center"> 
        <%if (email!=""){%>
        <a href="mailto:<%=email%>"><%=email%></a> 
        <%} else {%>
        &nbsp; 
        <%}%>
      </p>
    </td>
    <td> 
      <p align="center"><a href="actpubusr.asp?u=<%=id%>&st=<%=st==0?1:0%>"> 
        <%if (st==1) {%>
        Акт. 
        <%} else {%>
        Дизакт. 
        <%}%>
        </a> || Изменить</p>
    </td>
  </tr>
  <%
Records.MoveNext()
} 
Records.Close()
 %>
</table>
<p align="center">&nbsp;</p>
</body>
</html>
