<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\creaters.inc" -->

<%
var fio=""
var usr=parseInt(Request("usr"))
var ShowForm=true
var ErrorMsg=""
var sql=""
var valname=""
var valname1=""

if ((Session("is_adm_mem")!=1)&&(Session("id_mem")!=usr)){
Session("backurl")="pswusr.asp?usr="+usr
Response.Redirect("login.asp")}

Records.Source="Select * from USERS where ID="+usr
Records.Open()
if (Records.EOF){Response.Redirect("area.asp")}
fio=String(Records("FIO").Value)
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){
	valname=TextFormData(Request.Form("valname"),"")
	valname1=TextFormData(Request.Form("valname1"),"")
	if (valname!=valname1) {ErrorMsg+="Пароль не подтвержден или подтвержден с ошибкой.<br>"}
	if (String(valname).length<5) {ErrorMsg+="Пароль не может быть менее 5-и символов.<br>"}
	
	if (ErrorMsg=="") {
		sql=usrpsw
		sql=sql.replace("%ID",usr)
		sql=sql.replace("%PS",valname)
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
		}
		catch(e){
			Connect.RollbackTrans()
			ErrorMsg+=ListAdoErrors()
		}
		if (ErrorMsg==""){
			Connect.CommitTrans()
			ShowForm=false
		}
	}
}
%>

<html>
<head>
<title>Изменение пароля доступа <%=fio%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<p align="center"><b>Изменение пароля доступа.</b> Пользователь: <b><font color="#000099"><%=fio%></font></b></p>
<p>&nbsp;</p>
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 

<p align="center"><a href="area.asp">Кабинет</a> &nbsp;&nbsp;&nbsp;<a href="index.html">Домашняя 
  страница</a></p>
 <%if (ShowForm) {%>
<form name="form1" method="post" action="pswusr.asp">
  <div align="center"> 
    <input type="hidden" name="usr" value="<%=usr%>">
    <table width="70%" border="1" bordercolor="#FFFFFF">
      <tr bordercolor="#0033CC" bgcolor="#99CCCC"> 
        <td width="50%"> 
          <div align="center"><b><font color="#000099">Параметры:</font></b></div>
        </td>
        <td> 
          <div align="center"><b><font color="#000099">Значения:</font></b></div>
        </td>
      </tr>
      <tr bordercolor="#0033CC"> 
        <td width="50%"> 
          <div align="right">Введите новый пароль:&nbsp;&nbsp;</div>
        </td>
        <td>&nbsp;&nbsp;&nbsp; 
          <input type="password" name="valname" maxlength="20" size="20">
        </td>
      </tr>
      <tr bordercolor="#0033CC"> 
        <td width="50%"> 
          <div align="right">Повторите ввод нового пароля:&nbsp;&nbsp;</div>
        </td>
        <td>&nbsp;&nbsp;&nbsp; 
          <input type="password" name="valname1" maxlength="20" size="20">
        </td>
      </tr>
    </table>
    <input type="submit" name="Submit" value="Сохранить">
  </div>
</form>
<%} else {%>
<p>&nbsp; </p>
<center>
  <h2><font color="#3333FF">Пароль изменен!</font></h2>
  <p> 
    <%}%>
  </p>
</center>
</body>
</html>
