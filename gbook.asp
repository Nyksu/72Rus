<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// Код гостевой книги
var guestbook=1
// Изменить для других сайтов
var isadm=0
var sql=""
var tbook=1
var namebook=""
var kvo=0
var name=""
var id=parseInt(Request("id"))
var op=parseInt(Request("op"))
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}
var dat=""
var email=""
var city=""
var txt=""
var st=0
var ii=0
var ErrorMsg=""
var namop=""
var plen=10
var pp=0
var tkvo=0

if (String(Session("gbook"))=="undefined") {Session("gbook")=0}
if ((Session("is_adm_mem")==1) || (Session("is_host")==1)) { isadm=1 }

if (isNaN(id)) {id=0}
if (isNaN(op)) {op=0}
if (Session("gbook")>1) {op=0}
if ((op>1) && (isadm==0)) {op=0}
if ((op>2) && (id==0)) {op=0}
if ((id>0) && (op==0)) {id=0}

switch  (op) {
	case 1 : namop="Добавление сообщения в Гостевую книгу"; break;
	case 2 : namop="Изменение параметров Гостевой книги"; break;
	case 3 : namop="Удаление сообщения"; break;
	case 4 : namop="Блокировка сообщения"; break;
	case 5 : namop="Активация сообщения"; break;
	case 6 : namop="Изменение сообщения"; break;
}

sql="Select * from guestbook where id="+guestbook+" and enterprise_id="+company+" and state=1"
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("index.asp")
}
tbook=Records("STATE_DEF").Value
namebook=Records("NAME").Value
plen=parseInt(Records("PAGELEN").Value)
Records.Close()

if (id>0) {
	sql="Select * from gbmsg where guestbook_id="+guestbook+" and id="+id
	Records.Source=sql
	Records.Open()
	if (Records.EOF){
		Records.Close()
		Response.Redirect("gbook.asp")
	}
	name=Records("AUTORNAME").Value
	id=Records("ID").Value
	dat=Records("POSTDATE").Value
	email=TextFormData(Records("EMAIL").Value,"")
	city=TextFormData(Records("CITY").Value,"")
	txt=TextFormData(Records("COMENT").Value,"")
	st=Records("STATE").Value
	Records.Close()
}

var isFirst=String(Request.Form("Submit"))=="undefined"
if ((!isFirst) && (op>0)){
	sql=""
	if ((op==3) && (Request.Form("kill")==666)) {
		// Удаление 
		sql="Delete from gbmsg where id="+id
		if (fs.FileExists(filename)) {fs.DeleteFile(filename)}
	} else { if (op==3) {Response.Redirect("gbook.asp")}}
	if ((op==4) && (Request.Form("lockit")==33)) {
		// Деактивация 
		sql="Update gbmsg set state=0 where id="+id
	} else { if (op==4) {Response.Redirect("gbook.asp")}}
	if (op==5 && Request.Form("unlockit")==333) {
		// Активация 
		sql="Update gbmsg set state=1 where id="+id
	} else { if (op==5) {Response.Redirect("gbook.asp")}}
	if (op==2) {
		// Изменение атрибутов Гостевой книги
		name=TextFormData(Request.Form("name"),"")
		tbook=Request.Form("tips")==1?1:0
		plen=parseInt(Request.Form("plen"))
		if (isNaN(plen)) {ErrorMsg+="Неверный формат Длины страницы.<br>"}
		if (name.length>100) {ErrorMsg+="Слишком длинное наименование Гостевой книги.<br>"}
		if (name.length<2) {ErrorMsg+="Слишком короткое наименование Гостевой книги.<br>"}
		if (ErrorMsg=="") {
		 sql="Update guestbook set name='%NAME', state_def=%SD, pagelen=%PL where id="+guestbook
		 sql=sql.replace("%NAME",name)
		 sql=sql.replace("%SD",tbook)
		 sql=sql.replace("%PL",plen)
		}
	}
	if (op==1) {
		// Добавление сообщения
		name=TextFormData(Request.Form("name"),"")
		email=TextFormData(Request.Form("email"),"")
		city=TextFormData(Request.Form("city"),"")
		txt=TextFormData(Request.Form("txt"),"")
		while (txt.indexOf("/n")>=0) {txt=txt.rplace("/n"," ")}
		if (name.length>100) {ErrorMsg+="Слишком длинное имя.<br>"}
		if (name.length<2) {ErrorMsg+="Слишком короткое имя.<br>"}
		if (email.length>100) {ErrorMsg+="Слишком длинный E-mail.<br>"}
		if ((email != "") && (email.length<7)) {ErrorMsg+="Недопустимый E-mail.<br>"}
		if (email.length>0) {	 
			if (!/(\w+)@((\w+).)*(\w+)$/.test(email)) {ErrorMsg+="Неверный формат поля 'E-mail'.<br>"}}
		if (city.length>100) {ErrorMsg+="Слишком длинное наименование города (страны).<br>"}
		if (city.length<3) {ErrorMsg+="Слишком короткое наименование города (страны).<br>"}
		if (txt.length>500) {ErrorMsg+="Слишком длинное сообщение.<br>"}
		if (txt.length<3) {ErrorMsg+="Слишком короткое сообщение.<br>"}
		if (ErrorMsg=="") {
			id=NextID()
			sql="Insert into gbmsg (id,guestbook_id,autorname,email,city,postdate,state,coment) "
			sql+="values (%ID, %GB, '%NAME', '%EMAIL', '%CITY', 'NOW', %ST, '%TXT')"
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%ID",id)
			sql=sql.replace("%GB",guestbook)
			sql=sql.replace("%EMAIL",email)
			sql=sql.replace("%CITY",city)
			sql=sql.replace("%ST",tbook)
			sql=sql.replace("%TXT",txt)
			Session("lastadd")=id
		}
	}
	if (op==6) {
		// Добавление сообщения
		name=TextFormData(Request.Form("name"),"")
		email=TextFormData(Request.Form("email"),"")
		city=TextFormData(Request.Form("city"),"")
		txt=TextFormData(Request.Form("txt"),"")
		while (txt.indexOf("/n")>=0) {txt=txt.rplace("/n"," ")}
		if (name.length>100) {ErrorMsg+="Слишком длинное имя.<br>"}
		if (name.length<2) {ErrorMsg+="Слишком короткое имя.<br>"}
		if (email.length>100) {ErrorMsg+="Слишком длинный E-mail.<br>"}
		if ((email != "") && (email.length<7)) {ErrorMsg+="Слишком короткий E-mail.<br>"}
		if (city.length>100) {ErrorMsg+="Слишком длинное наименование города (страны).<br>"}
		if (city.length<3) {ErrorMsg+="Слишком короткое наименование города (страны).<br>"}
		if (txt.length>500) {ErrorMsg+="Слишком длинное сообщение.<br>"}
		if (txt.length<3) {ErrorMsg+="Слишком короткое сообщение.<br>"}
		if (ErrorMsg=="") {
			sql="Update gbmsg set guestbook_id=%GB ,autorname='%NAME' ,email='%EMAIL' ,city='%CITY' ,coment='%TXT'  "
			sql+="where id=%ID"
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%GB",guestbook)
			sql=sql.replace("%ID",id)
			sql=sql.replace("%EMAIL",email)
			sql=sql.replace("%CITY",city)
			sql=sql.replace("%TXT",txt)
		}
	}
	if ((ErrorMsg=="") && (sql!="")) {
			Connect.BeginTrans()
			try{
				Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="Ошибка операции.<br>"
			}
			if (ErrorMsg==""){
				Connect.CommitTrans()
				if (isadm==0 && op==1) {
					Session("gbook")+=1					
				}
		  		Response.Redirect("gbook.asp")
			}
	} 
}

%>

<html>
<head>
<title>Гостевая книга</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>
<body bgcolor="#FFFFFF" text="#000000">

<p><a href="index.asp">Вернуться на главную страницу</a> <%if (isadm==1) {%>| <a href="admarea.asp">Кабинет</a><%}%></p>
<h1 align="center"><b><font size="4" color="#0000CC"><%=namebook%></font></b></h1>
<%
if (isadm==1 && op!=2) {
%>
<p><a href="gbook.asp?op=2&id=0">Изменить параметры Гостевой книги</a></p>
<%
}
%>
<%
if (op==0) {
%>
<p><a href="gbook.asp?op=1&id=0">Добавить новое сообщение в Гостевую книгу</a></p>

<%
sql="Select * from gbmsg where guestbook_id="+guestbook
if (isadm!=1) {
	if (String(Session("lastadd"))=="undefined") {sql+=" and state=1"}
	else {sql+=" and ((state=1 and id<>"+Session("lastadd")+") or (id="+Session("lastadd")+"))"}
}
sql+=" order by id desc"
Records.Source=sql
Records.Open()
if (!Records.EOF) {
kvo=Records.Recordcount
if ((pg+1)*plen > kvo) {tkvo=kvo} else {tkvo=(pg+1)*plen}
%>
<p>Сообщения размещенные в Гостевую книгу:</p>
<p>Размещено сообщений: <font color="#000099"><b><%=kvo%></b></font> </p>
<%
ii=0
while ((!Records.EOF) && (ii<tkvo)) {
	name=Records("AUTORNAME").Value
	id=Records("ID").Value
	dat=Records("POSTDATE").Value
	email=TextFormData(Records("EMAIL").Value,"")
	city=TextFormData(Records("CITY").Value,"")
	txt=TextFormData(Records("COMENT").Value,"")
	while (txt.indexOf("<")>=0) {txt=txt.rplace("<","&lt;")}
	st=Records("STATE").Value
	ii+=1
	if (ii>=(pg*plen+1)) {
%>
<table width="70%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr> 
    <td width="40" bordercolor="#999999"> 
      <div align="center"><b><font size="2" color="#000099"><%=ii%>.</font></b></div>
    </td>
    <td bordercolor="#999999"><b><font size="2" color="#0000FF"><%=dat%></font></b>&nbsp;&nbsp; 
      <font color="#000099"><b><%=name%></b></font> <font size="2">&nbsp;<a href="mailto:<%=email%>">e-mail</a> 
      ( <font color="#0000FF"><%=city%></font> )</font></td>
  </tr>
  <% if (isadm==1) {%>
   <tr> 
    <td width="40">
      <div align="center"><font size="2"><b><font color="#FF0000"><%=st==0?"ост":""%></font></b></font></div>
    </td>
    <td bordercolor="#999999"><font size="2"><a href="gbook.asp?op=6&id=<%=id%>">Изменить</a> 
      | <a href="gbook.asp?op=3&id=<%=id%>">Удалить</a> | <a href="gbook.asp?op=<%=st==1?"4":"5"%>&id=<%=id%>"> 
      <%=st==1?"Остановить":"Активировать"%></a></font></td>
  </tr>
  <%}%>
  <tr> 
    <td width="40">&nbsp;</td>
    <td bordercolor="#999999"><font size="2"><%=txt%></font></td>
  </tr>
  <tr> 
    <td width="40" height="15"></td>
    <td height="15"></td>
  </tr>
</table>
<%
	}
	Records.MoveNext()
}
%>
Страницы: <font size="2">
<%
for ( pp=1; pp<(kvo/plen + 1) ; pp++) {
%>
<% if (pp==(pg+1)) { %><%=pp%> | <%} else {%><a href="gbook.asp?pg=<%=pp-1%>"><%=pp%></a> |  <%}%>
<%
}
%>
</font> 
<%
} else {
%>
<p>Нет сообщений в Гостевой книге</p>
<%
}
Records.Close()
%>
<%
} else {
// выполняем одну из операций
%>
<p><a href="gbook.asp">Вернуться к списку сообщений</a></p>
<p>Операция: <%=namop%></p>
<table width="100%" border="1" cellspacing="1" cellpadding="1">
  <tr>
    <td bgcolor="#F0F0F0">
<% 
if (ErrorMsg!="") {
%>
<div align="center">
  <table width="90%" border="1" cellspacing="1" cellpadding="1" bordercolor="#FF0000" bgcolor="#FFFFCC">
    <tr> 
      <td> 
        <p align="center"><b><font color="#FF0000" size="2">Внимание! Замечены следующие ошибки:</font></b></p>
        <p align="center"><font color="#FF0000" size="2"><%=ErrorMsg%></font></p>
      </td>
    </tr>
  </table>
</div>
<%
}
%>
      <form name="form1" method="post" action="gbook.asp">
          <input type="hidden" name="id" value="<%=id%>">
          <input type="hidden" name="op" value="<%=op%>">
<%
if ((op==1) || (op==6)) {
%>
<table width="90%" border="2" cellspacing="2" cellpadding="1" bordercolor="#FFFFFF" dwcopytype="CopyTableRow">
      <tr> 
        <td bordercolor="#0099FF" bgcolor="#CCCCFF" width="200"> 
          <div align="center"><b>Параметры</b></div>
        </td>
        <td width="15"> 
          <div align="center"><b></b></div>
        </td>
        <td bordercolor="#0099FF" bgcolor="#CCCCFF"> 
          <div align="center"><b>Значения</b></div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
              <div align="right"><font color="#FF0000">Ваше имя или псевдоним:</font> 
              </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" size="30" maxlength="100">
          (до 100 символов)</td>
      </tr>
	  <tr> 
        <td bordercolor="#0099FF" width="200"> 
              <div align="right"><font color="#FF0000">Город (страна):</font> 
              </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="city" value="<%=isFirst?city:Request.Form("city")%>" size="30" maxlength="100">
          (до 100 символов)</td>
      </tr>
	  <tr> 
        <td bordercolor="#0099FF" width="200"> 
              <div align="right">E-mail: </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="email" value="<%=isFirst?email:Request.Form("email")%>" size="30" maxlength="100">
          (до 100 символов)</td>
      </tr>
	  <tr> 
        <td bordercolor="#0099FF" width="200"> 
              <div align="right"><font color="#FF0000">Текст сообщения:</font> 
              </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <textarea name="txt" rows="3" cols="40"><%=txt%></textarea>
              (До 500 символов)</td>
      </tr>
</table>
<%
}
%>
<%
	if (op==3) {
%>
	    <p><b>Удаление сообщения. <font color="#0000CC"><%=name%></font></b></p>
		<p><font size="2" color="#000099"><%=txt%></font>
        <p>
        <p>
          <input type="checkbox" name="kill" value="666">
          Да! Удалить!</p>
<%
	}
%>
<%
	if (op==4) {
%>
	    <p><b>Блокировка сообщения. <font color="#0000CC"><%=name%></font></b></p>
		<p><font size="2" color="#000099"><%=txt%></font>
        <p>
        <p>
          <input type="checkbox" name="lockit" value="33">
          Да! Заблокировать сообщение!</p>
<%
	}
%>
<%
	if (op==5) {
%>
	    <p><b>Активация сообщения. <font color="#0000CC"><%=name%></font></b></p>
		<p><font size="2" color="#000099"><%=txt%></font>
        <p>
        <p>
          <input type="checkbox" name="unlockit" value="333">
          Да! Разблокировать сообщение!</p>
<%
	}
%>
<%
	if (op==2) {
%>
	    <p><b>Параметры гостевой книги:</b></p>
		<p>
          <input type="text" name="name" size="30" maxlength="100" value="<%=isFirst?namebook:Request.Form("name")%>">
          Имя Гостевой книги (до 100 символов)</p>
        <p> 
          <input type="checkbox" name="tips" value="1" <%=tbook==0?"":"checked"%>>
          Разрешить публиковать сообщения сразу без проверки.</p>
		  
        <p>
          <input type="text" name="plen" size="10" maxlength="3" value="<%=isFirst?plen:Request.Form("plen")%>">
          Длина страницы</p>
<%
	}
%>
        <p>
          <input type="submit" name="Submit" value="Продолжить">
        </p>
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<%
}
%>
</body>
</html>
