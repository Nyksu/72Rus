<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var  hid=parseInt(Request("hid"))
var name=""
var ErrorMsg=""
var ShowForm=true
var tit="" 
var hname=""
var tekhia=0
var hiadr=""


if (isNaN(hid)) {hid=0}

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && Session("id_mem")!="undefined"){
	sql="Select * from catalog where id="+catalog
	Records.Source=sql
    Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("index.asp")
   }
   if (Session("id_mem")==Records("USERS_ID").Value) {Session("cataloghost")=catalog}
   Records.Close()
}

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog) {
Session("backurl")="addcatarea.asp?hid="+hid
Response.Redirect("login.asp")
}

if (hid>0) {
sql="Select * from catarea where id="+hid+" and catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("catarea.asp?hid=0")
}
hname=String(Records("NAME").Value)
Records.Close()
}
tit=hname
if (tit="") {tit="Корневой раздел"}

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){

     name=TextFormData(Request.Form("name"),"")
	 
	 if (name.length<3) {ErrorMsg+="Слишком короткое наименование.<br>"}
	 
	  if (ErrorMsg==""){
	  		id=NextID("catareaid")
			sql="Insert into catarea(ID,NAME,CATALOG_ID,HI_ID) values (%ID, '%NAM', %CAT, %HID)"
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAM",name)
			sql=sql.replace("%HID",hid)
			sql=sql.replace("%CAT",catalog)
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
<title>Добавляем тему в каталог URL (<%=tit%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="/72.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> |</p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<p align="center"><font size="3"><b><font color="#0000CC">Последние добавленные 
  объявления:</font></b></font></p>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr>
      
    <td bgcolor="#F2F2F2" bordercolor="#CCCCCC"> &nbsp; &nbsp; &nbsp; <font color="#000099" size="1"><b> 
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
      </b></font> </td>
    </tr>
  </table>

<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 
<p>&nbsp; </p>

<%if(ShowForm){%> 
<form name="form1" method="post" action="addcatarea.asp">
  <input type="hidden" name="hid" value="<% =hid %>">
  <p><%=tit%></p>
  <table width="100%" border="1" bordercolor="#FFFFFF">
    <tr> 
      <td width="380" bgcolor="#F2F2F2" bordercolor="#CCCCCC"> 
        <div align="center"> 
          <p><b>Параметры:</b></p>
        </div>
      </td>
      <td width="20">&nbsp;</td>
      <td bgcolor="#F2F2F2" width="582" bordercolor="#CCCCCC"> 
        <div align="center"> 
          <p><b>Значения:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#CCCCCC"> 
        <div align="right"> 
          <p>Наименование раздела&nbsp;&nbsp;</p>
        </div>
      </td>
      <td width="20"> 
        <div align="center">-</div>
      </td>
      <td bordercolor="#CCCCCC" width="582"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="99" size="45">
        </p>
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
  <p><font color="#3333FF"><b>Раздел добавлен!</b></font></p>
  <p>&nbsp;</p>
  <p><a href="admarea.asp">В кабинет</a> | <font face="Arial, Helvetica, sans-serif"><a href="index.asp">
    На главную страницу</a> | <a href="catarea.asp">В каталог</a></font></p>
  </center>
<%}%>
 
</body>
</html>
