<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""
var bid=parseInt(Request("bid"))
var Pass=""
var sql=""
var psw=""
var url=""

Session("banlog")="undefined"

if (isNaN(bid)) {bid=0}
isFirst=String(Request.Form("login"))=="undefined"

if (String(Session("backurl"))=="undefined"){Session("backurl")="index.asp"}

if ((!isFirst)&&(bid>0)) {
	Pass=TextFormData(Request.Form("pass"),"")
	Pass=Pass.replace("/*","")
	Pass=Pass.replace("'","")
	
	sql="Select * from baner where state=0 and id="+bid
	Records.Source=sql
	Records.Open()
	if (!Records.Eof) {
		psw=Records("PSW").Value
	} else {bid=0}
	Records.Close()
	
	if (bid>0) {
		if (psw==Pass) {
			Session("banlog")=bid
			url=Session("backurl")+"&bid="+bid
			Response.Redirect(url)
		} else {ErrorMsg="Неверный пароль"}
	} else {ErrorMsg="Неверный код банера"}
}

%>

<html>
<head>
<title>Авторизация просмотра статистики по банеру</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<p>&nbsp;</p>
<p align="center">
  <%if(ErrorMsg!=""){%>
</p>
<center>
  <h2> 
    <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
      <%=ErrorMsg%></p>
  </h2>
</center>
<%}%>
<form name="form1" method="post" action="logban.asp">
  <table width="100%" border="0" cellspacing="4" cellpadding="0">
    <tr valign="middle"> 
      <td width="50%" align="right"> 
        <p><b>Код банера</b></p>
      </td>
      <td width="50%"> 
        <input type="text" name="bid" value="<%=bid==0?"":bid%>" size="20" maxlength="20">
        * </td>
    </tr>
    <tr valign="middle"> 
      <td width="50%" align="right"> 
        <p><b>Пароль</b></p>
      </td>
      <td width="50%"> 
        <input type="password" name="pass" size="20" maxlength="20">
        * </td>
    </tr>
  </table>
  <p align="center"> 
    <input type="submit" name="login" value="Ввод">
  </p>
  <hr size="1">
  <p align="center"><b>* - Обязательные поля</b></p>
  <p align="center">&nbsp;</p>
</form>
</body>
</html>
