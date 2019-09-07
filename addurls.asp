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
var url=""
var about=""
var id=0

var cw=String(Math.random()).substr(3,20)

if(cw.length<20){cw+=String(Math.random()).substr(3,(20-cw.length))}

if (isNaN(hid)) {hid=0}

if (hid==0) {Response.Redirect("catarea.asp?hid=0")}


if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog) {
Session("backurl")="addurls.asp?hid="+hid
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
	 url=TextFormData(Request.Form("url"),"")
	 about=TextFormData(Request.Form("about"),"")
	 cw=TextFormData(Request.Form("cw"),"")
	 
	 while (name.indexOf("'")>-1) {name=name.replace("'","\"")}
	while (name.indexOf("  ")!=-1) {name=name.replace("  "," ")}
	while (name.indexOf(" ")==0) {name=name.replace(" ","")}
	while (about.indexOf("  ")!=-1) {about=about.replace("  "," ")}
	while (about.indexOf("'")>-1) {about=about.replace("'","\"")}
	while (about.indexOf(" ")==0) {about=about.replace(" ","")}
	
	while (url.indexOf(" ")!=-1) {url=url.replace(" ","")}
	 
	 if (name.length<5) {ErrorMsg+="Слишком короткое наименование.<br>"}
	 if (url.length<5) {ErrorMsg+="Слишком короткий URL.<br>"}
	 if (about.length<10) {ErrorMsg+="Слишком короткое описание.<br>"}
	 if (cw.length<5) {ErrorMsg+="Слишком короткое кодовое слово.<br>"}
	 if (url.indexOf("http://")==-1){url="http://"+url}
	sql="Select * from URL where Upper(url)=Upper('"+url+"') and catarea_id="+hid
	Records.Source=sql
    Records.Open()
    if (!Records.EOF){ErrorMsg+="Сайт не может быть размещен дважды в одном разделе каталога!<br>"}
    Records.Close()
	sql="Select * from URL where Upper(name)=Upper('"+name+"') and catarea_id="+hid
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Смените наименование сайта! такое наименование уже встречается в текущем разделе каталога!<br>"}
   Records.Close()
	  if (ErrorMsg==""){
	  		id=NextID("URLID")
			sql="Insert into url (ID,NAME,URL,ABOUT,CATAREA_ID,STATE,REG_DATE,NEW_DATE,KEYWORD,HOST_URL_ID) "
			sql+="values (%ID, '%NAM', '%URL', '%AB', %CAT, 1, 'TODAY', 'TODAY', '%CW', NULL)"
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAM",name)
			sql=sql.replace("%URL",url)
			sql=sql.replace("%AB",about)
			sql=sql.replace("%CW",cw)
			sql=sql.replace("%CAT",hid)
			Connect.BeginTrans()
			try{
			Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="Ошибка вставки.<br>"
			}
			if (ErrorMsg==""){
		   Connect.CommitTrans()
		   ShowForm=false
			}
	  }
	while (name.indexOf("\"")>-1) {name=name.replace("\"","&quot;")}
	while (about.indexOf("\"")>-1) {about=about.replace("\"","&quot;")}
}
if (url.indexOf("http://")!=-1){url=url.replace("http://","")}
%>

<html>
<head>
<title>Добавляем тему в каталог URL (<%=tit%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="/style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> |</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p><table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#00CCFF">
  <tr>
      
    <td bgcolor="#F5F5F5" bordercolor="#CCCCCC"> &nbsp; &nbsp; &nbsp; <font color="#000099" size="1"><b> 
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
<form name="form1" method="post" action="addurls.asp">
  <input type="hidden" name="hid" value="<% =hid %>">
  <p><%=tit%></p>
  <table width="100%" border="1" bordercolor="#FFFFFF">
    <tr> 
      <td width="380" bgcolor="#F0F0F0" bordercolor="#CCCCCC"> 
        <div align="center"> 
          <p><b>Параметры:</b></p>
        </div>
      </td>
      <td width="20"> 
        <p>&nbsp;</p>
      </td>
      <td bgcolor="#F0F0F0" width="582" bordercolor="#CCCCCC"> 
        <div align="center"> 
          <p><b>Значения:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#CCCCCC"> 
        <div align="right"> 
          <p>Наименование сайта&nbsp;&nbsp;</p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#CCCCCC" width="582"> 
        <p> 
          <input type="text" name="name" value="<%=name%>" maxlength="99" size="45">
          (до 100 символов) </p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#CCCCCC"> 
        <div align="right"> 
          <p>URL сайта&nbsp;&nbsp;</p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#CCCCCC" width="582"> 
        <p>http:// 
          <input type="text" name="url" value="<%=url%>" maxlength="250" size="37">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#CCCCCC"> 
        <div align="right"> 
          <p>Краткое описание сайта&nbsp;&nbsp;</p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#CCCCCC" width="582"> 
        <p> 
          <input type="text" name="about" value="<%=about%>" maxlength="250" size="45">
          <font size="2">(до 250 символов) </font></p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#CCCCCC"> 
        <div align="right"> 
          <p>Код владельца сайта&nbsp;&nbsp;</p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#CCCCCC" width="582"> 
        <p> 
          <input type="text" name="cw" value="<%=isFirst?cw:Request.Form("cw")%>" maxlength="20" size="45">
          <font size="2">(до 20 символов, лучше использовать предложенный по умолчанию)</font></p>
      </td>
    </tr>
  </table>
  <p align="center"> 
    <input type="submit" name="Submit" value="Сохранить">
  </p>
</form>
 <%} 
else 
{%> 
<center>
  <h2><font color="#3333FF">Сайт добавлен!</font></h2>
  <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную страницу</a></font></p>
  <p><font face="Arial, Helvetica, sans-serif"><a href="catarea.asp">В каталог</a></font></p>
</center>
<%}%>
 
</body>
</html>
