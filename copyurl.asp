<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<%
var hid=parseInt(Request("hid"))
var sql=""
var hiadr=""
var tit=""
var hname=""
var tekhia=0
var urlcount=0
var endlist=0
var urlname=""
var urlabout=""
var st=parseInt(Request("st"))
var daterenew=""
var urladr=""
var urlid=parseInt(Request("url"))
var dd=0
var cu=0
var isok=true
var cat=0

if (isNaN(hid)) {hid=0}
if (isNaN(urlid)) {urlid=0}
if (urlid==0) {Response.Redirect("admarea.asp")}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="copyurl.asp?url="+urlid+"&hid="+hid
Response.Redirect("login.asp")
}

sql="Select t1.* from url t1, catarea t2 where t1.id="+urlid+" and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
	cat=Records("CATAREA_ID").Value
} else {
	Records.Close()
	Response.Redirect("admarea.asp")
}
Records.Close()

if (isNaN(st)) {st=0}

if (st==1 && hid!=cat) {
	id=NextID("URLID")
	sql="Insert into url (ID,NAME,URL,ABOUT,CATAREA_ID,STATE,REG_DATE,NEW_DATE,KEYWORD,HOST_URL_ID) "
	sql+="Select "+id+", NAME, URL, ABOUT, "+hid+", STATE, REG_DATE, NEW_DATE, KEYWORD, HOST_URL_ID from url where ID="+urlid
	Connect.BeginTrans()
	try{
		Connect.Execute(sql)
	}
		catch(e){
		Connect.RollbackTrans()
		isok=false
		st=0
	}
	if (isok){
		Connect.CommitTrans()
	}
} else {st=0}

%>
<Html>
<Head>
<Title>Копирование ссылки - 72RUS - Тюмень - Региональный портал</Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">


<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr> 
    <td width="150" valign="middle" background="Img/yellow.jpg" height="19"> 
      <div align="center"> <a href="index.asp"><img src="Img/home.gif" width="14" height="14" align="absbottom" border="0" alt="Home"></a> 
        <font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНСКИЙ РЕГИОН</b></font></div>
    </td>
    <td bordercolor="#FFCC00" background="Img/yellow.jpg" height="19"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>WWW.72RUS.RU</b></font></div>
    </td>
    <td width="468" background="Img/yellow.jpg" height="19"> 
      <table width="410" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="Img/tmn.gif" width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНЬ</b></font></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="messages.asp">ОБЪЯВЛЕНИЯ</a></b></font></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="catarea.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>КАТАЛОГ</b></font></a></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="Rail_roads.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>РАСПИСАНИЕ</b></font></a></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="marketBuilder.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>АНАЛИТИКА</b></font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td bgcolor="#ffd34e" width="150" align="left" valign="middle" background="Img/runline.gif"> 
      <div align="center"><a href="index.asp"><img src="Img/%F2%FE%EC%E5%ED%FC%2072rus.gif" alt="www.72rus.ru" border="0"></a></div>
    </td>
    <td bordercolor="#FFCC00" bgcolor="#ffd34e" background="Img/runline.gif"> 
      <div align="center"><object type="text/x-scriptlet" width=171 height="60" data="searchbox.html">
        </object> </div>
    </td>
    <td bgcolor="#ffd34e" align="right" background="Img/runline.gif" width="468"><a href="http://www.rusintel.ru/"><img src="bannes/rol_ba.gif" width="468" height="60" alt="Internet Card - ROL - Tyumen" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0" background="Img/line.gif">
  <tr bgcolor="#996600" background="Img/line.gif"> 
    <td width="150" background="Img/line.gif"> 
      <div align="center"> 
        <p>
      </div>
    </td>
    <td width="490" bgcolor="#996600" background="Img/line.gif"> 
      <p align="left"><b><font color="#FFCC33">- Тюмень -</font></b></p>
    </td>
    <td width="213" background="Img/line.gif"> 
      <p><b><font color="#FFCC33">-- Инфо -- </font></b></p>
    </td>
    <td width="150" bgcolor="#996600" background="Img/line.gif"> 
      <div align="center"> 
        <p align="right"><img src="Img/home.gif" width="14" height="14" align="middle"> 
          <a href="index.asp"><b><font color="#FFCC33">Home</font></b></a></p>
      </div>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" bgcolor="#EBF5ED"> 
      <div align="left"> 
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Тюменский каталог</b></font></td>
          </tr>
          </tbody> 
        </table>
        
      </div>
    </td>
    <td> 
      <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#ffd34e width="100%" height="7"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Копируем URL:</b></font></p>
            </td>
        </tr>
        </tbody> 
      </table>
      <h2 align="left"> <img align=middle height=10 src="img/arrow2.gif" width=11> 
        <b><a href="<%=urladr%>"> 
        <%=urlname%></a></b> - <%=urlabout%></h2>
	  <table width="100%" border="1" bordercolor="#FFFFFF">
          <tr> 
          <td bgcolor="#ffd34e" height="2"> 
            <div align="left"> <font face="Verdana" size="1"><b>Вы здесь</b></font> 
              <font size="2"> 
              <%
	  tekhia=hid
	  while (tekhia!=0) {
	     sql="Select * from catarea where id="+tekhia
		 Records.Source=sql
		 Records.Open()
		 if (!Records.EOF) {
			hname=String(Records("NAME").Value)
			hiadr="copyurl.asp?hid="+tekhia+"&url="+urlid
			tekhia=Records("HI_ID").Value
			if (hiadr != "catarea.asp?hid="+hid) {
		%>
              <img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
              <% if (st!=1)  { %><a href="<%=hiadr%>"><%=hname%></a> <%} else {%><%=hname%><%}%>
              <%
		  	} else {
		%>
              <img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> 
              <%=hname%> 
              <%
		  	}	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
              </b></font> <img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
              <a href="copyurl.asp?hid=0&url=<%=urlid%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Каталог</b></font></a> 
              <% } else {%>
              <img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
              <a href="index.asp"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">На 
              главную страницу</font></b></a> 
              <%} %>
              </font> </div>
			</td>
          </tr>
      </table>
<% if (st!=1)  { %>
<% if (hid>0 && hid!=cat) {%>
	  <p><a href="copyurl.asp?url=<%=urlid%>&hid=<%=hid%>&st=1">Скопировать 
        сайт в текущий каталог</a></p>
		<p>&nbsp;</p>
<% } 
	sql="Select * from catarea where hi_id="+hid+" and catalog_id="+catalog+" order by name"
	Records.Source=sql
	Records.Open()
	while (!Records.EOF) {
		hname=String(Records("NAME").Value)
		tekhia=Records("ID").Value 
		if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {cu=Count_url(5,tekhia)}
		else {cu=Count_url(1,tekhia)}
		Records.MoveNext() %>
        <p align="left"><img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
          <b><a href="copyurl.asp?url=<%=urlid%>&hid=<%=tekhia%>"> 
          <%=hname%></a></b> (<%=cu%>) 
		  <% if (hid!=cat) {%><b>&lt;---</b> <a href="copyurl.asp?url=<%=urlid%>&hid=<%=tekhia%>&st=1">Копировать  сюда!</a></p> <%}%>
        <%	}  Records.Close()
} else {%>
      <p><b><font color="#0000CC">Сайт скопирован в текущую директорию!!</font></b></p>
<% } %>
      </td>
    <td width="150" bgcolor="#EBF5ED"> 
      <div align="left">
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Ссылки</b></font></td>
          </tr>
          </tbody> 
        </table>
      </div>
      <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
        <tr> 
          <td bordercolor="#FFFFFF" valign="top"> 
            <div align="left">
              <h2 align="left">&nbsp;</h2>
              </div>
            </td>
        </tr>
      </table>
      <p align="left">&nbsp;</p>
      <p align="left">&nbsp;</p>
    </td>
  </tr>
</table>
<table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e">
  <tr bgcolor="#EBF5ED" bordercolor="#EBF5ED"> 
    <td bordercolor="#EBF5ED" width="150" valign="middle" align="center">&nbsp;</td>
    <td bordercolor="#EBF5ED"><a href="admarea.asp"><b><font size="2">Кабинет</font></b></a> 
      &nbsp;|&nbsp;<a href="catarea.asp"><b><font size="2">Каталог</font></b></a></td>
    <td width="150" valign="middle" align="center">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
    <td valign="middle" bgcolor="#996600"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
        портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
        Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
        Русинтел</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
    </td>
  </tr>
</table>
</Body>
</Html>
