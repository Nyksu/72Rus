<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var cat=parseInt(Request("hid"))
var hid=0
var name=""
var ErrorMsg=""
var ShowForm=true
var tit="" 
var hname=""
var tekhia=0
var hiadr=""


if (isNaN(cat)) {Response.Redirect("catarea.asp?hid=0")}
if (cat==0) {Response.Redirect("catarea.asp?hid=0")}

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog) {
Session("backurl")="edcatarea.asp?hid="+cat
Response.Redirect("login.asp")
}

if (cat>0) {
sql="Select * from catarea where id="+cat+" and catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("catarea.asp?hid=0")
}
name=String(Records("NAME").Value)
hid=Records("HI_ID").Value
Records.Close()
}
tit=name

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){

     name=TextFormData(Request.Form("name"),"")
	 
	 if (name.length<3) {ErrorMsg+="Слишком короткое наименование.<br>"}
	 
	  if (ErrorMsg==""){
	  		sql="Update catarea set name='%NAM' where id=%ID"
			sql=sql.replace("%ID",cat)
			sql=sql.replace("%NAM",name)
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
<title>Изменяем тему каталога URL (<%=tit%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="/72.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="catarea.asp">Каталог 72RUS.RU</a> 
        | <a href="usrarea.asp">Кабинет Пользователя</a> | </p>
    </td>
  </tr>
</table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#666666" bgcolor="#00CCFF">
  <tr>
      
    <td bgcolor="#EAE2DB"> 
      <p>&nbsp; &nbsp; &nbsp; <font color="#000099" size="1"><b> 
        <%
	  tekhia=hid
	  while (tekhia!=0) {
	     sql="Select * from catarea where id="+tekhia
		 Records.Source=sql
		 Records.Open()
		 if (!Records.EOF) {
			hname=String(Records("NAME").Value)
			hiadr="catarea.asp?hid="+tekhia
			tekhia=Records("HI_ID").Value
		%>
        &lt; &nbsp;<a href="<%=hiadr%>"><%=hname%></a> &nbsp; &nbsp; 
        <%	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
        &lt; &nbsp;<a href="catarea.asp?hid=0">Весь каталог</a> &nbsp; 
        <% } else {%>
        &lt; &nbsp;<a href="index.asp">На главную страницу</a> &nbsp; 
        <%} %>
        </b></font> </p>
    </td>
    </tr>
  </table>

<p>&nbsp; </p>
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 
<p>&nbsp; </p>
<%if(ShowForm){%> 
<form name="form1" method="post" action="edcatarea.asp">
  <input type="hidden" name="hid" value="<% =cat %>">
  <p><b><font size="3">Изменяем название раздела <font color="#FF0000"><%=tit%></font> каталога</font></b></p>
  <table width="100%" border="1" bordercolor="#FFFFFF">
    <tr> 
      <td width="380" bgcolor="#00CCCC"> 
        <div align="center"><b>Параметры:</b></div>
      </td>
      <td width="20">&nbsp;</td>
      <td bgcolor="#00CCCC" width="582"> 
        <div align="center"><b>Значения:</b></div>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#0099CC"> 
        <div align="right">Наименование раздела&nbsp;&nbsp;</div>
      </td>
      <td width="20"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#0099CC" width="582"> 
        <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" maxlength="99" size="45">
      </td>
    </tr>
  </table>
 <p>
    <input type="submit" name="Submit" value="Сохранить">
  </p>
</form>
 <%} 
else 
{%> 
<center>
  <h2><font color="#3333FF">Раздел изменен!</font></h2>
  <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную страницу</a></font></p>
  <p><font face="Arial, Helvetica, sans-serif"><a href="catarea.asp">В каталог</a></font></p>
</center>
<%}%>
 
</body>
</html>
