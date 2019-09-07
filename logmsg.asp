<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""
var cw=""

if (String(Session("backurl"))=="undefined"){Session("backurl")="index.asp"}

isFirst=String(Request.Form("login"))=="undefined"
if(!isFirst){
	cw=TextFormData(Request.Form("cw"),"")
	Records.Source="Select * from TRADEMSG  where DATE_END>='TODAY' and CODEWORD='"+cw+"'"
	Records.Open()
	if (Records.EOF){ErrorMsg+="Неверный код доступа.<br>"}
	else { if (Records("STATE").Value==0) {
		Session("id_ms")=String(Records("ID").Value)
		} else {
		Session("id_ms")="undefined"
		ErrorMsg+="Нет доступа.<br>"
		}
	}
	Records.Close()
	if (ErrorMsg==""){
		Response.Redirect(Session("backurl"))
	}
}

%>

<html>
<head>
<title>Аутентификация</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>

<body bgcolor="#FFFFFF" text="#000000">

 <p><br>
  </p>
  <p align="center"><%if(ErrorMsg!=""){%> </p>
  <center>
      <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
        <%=ErrorMsg%></p>
  </center>
<%}%>
<form name="form1" method="post" action="logmsg.asp">
  <table width="100%" border="0" cellspacing="4" cellpadding="0">
    <tr valign="middle"> 
      <td width="50%" align="right"> 
        <p><b>Ведите код доступа: </b></p>
      </td>
      <td width="50%"> 
        <input type="text" name="cw" value="<%=isFirst?"":Request.Form("cw")%>" size="30" maxlength="20">
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
