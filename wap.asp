<% response.ContentType = "text/vnd.wap.wml" %>
<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->
<% response.ContentType = "text/vnd.wap.wml" %>
<%
var hid=parseInt(Request("hid"))
var sql=""
var hiadr=""
var tit=""
var hname=""
var tekhia=0
var urlcount=0
var nextpg=-1
var endlist=0
var urlname=""
var urlid=0
var urlabout=""
var statname=""
var daterenew=""
var urladr=""
var pg=parseInt(Request("pg"))
var dd=0
var cu=0
var kvopub=0
var urlcountall=0

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var nid=0


if (isNaN(hid)) {hid=0}

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

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

sql="Select Count(*) as kvo from url t1, catarea t2 where t1.state=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlcountall=Records("KVO").Value
}
Records.Close()

sql="Select Count(*) as kvo from url t1, catarea t2 where t1.state=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t2.id="+hid
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlcount=Records("KVO").Value
}
Records.Close()

sql="Select Count(*) as kvo from trademsg t1, trade_subj t2 where t1.state=0 and t1.date_end>='TODAY' and t1.trade_subj_id=t2.id and t2.marketplace_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	msgcount=Records("KVO").Value
}
Records.Close()
%>

<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<card id="card1" title="<%=tit%>">

<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  1px 3px 0px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right bgcolor="#CCCCCC"> </td>
  </tr>
  </tbody> 
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td valign="top" align="center"> 
      <table width="100%" border="0" cellspacing="0" bordercolor="#003366">
        <tr bgcolor="#FBF8D7"> 
          <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle"> 
            <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              Вы здесь 
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
			if (hiadr != "catarea.asp?hid="+hid) {
		%>
              / <a href="<%=hiadr%>"><%=hname%></a> 
              <%
		  	} else {
		%>
              / <%=hname%> 
              <%
		  	}	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
              / <a href="catarea.asp?hid=0">Каталог</a> 
              <% } else {%>
              / 
              <%} %>
              / <a href="index.asp">72RUS.RU</a> </p>
          </td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
        <tr> 
          <td bgcolor="#FF9900" align="left"> 
            <h1><b><font color="#FFFFFF">Каталог сайтов: 
              <% if (hid==0) { %>
              WEB ресурсы 
              <%} else {%>
              <%=tit%><%=tit==""?"":":"%> </font></b></h1>
            <%}%>
          </td>
        </tr>
      </table>
      <% if (hid>0) { %>
      <p align="center"> <font color="#FF0000"> + </font> <a href="addurl.asp?hid=<%=hid%>">Добавить 
        сайт в раздел</a> | <a href="searchall.asp">Поиск на сайте</a></p>
      <%}%>
      <%if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {%>
      <table width="98%" border="1" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
        <tr bgcolor="#ffd34e" bordercolor="#996600"> 
          <td colspan="2"> 
            <div align="center"> 
              <p><b><font color="#993300">УПРАВЛЕНИЕ</font><font color="#996600"> 
                :</font></b></p>
            </div>
          </td>
        </tr>
        <tr bgcolor="#FFFFCC" bordercolor="#996600"> 
          <td width="50%"> 
            <div align="center"> 
              <p><font size="1"><a href="addcatarea.asp?hid=<%=hid%>">Добавить 
                тему</a></font></p>
            </div>
          </td>
          <% if (hid>0) { %>
          <td> 
            <div align="center"> 
              <p><font size="1"><a href="addurls.asp?hid=<%=hid%>">Добавить URL</a></font></p>
            </div>
          </td>
          <%}%>
        </tr>
      </table>
      <%}%>
      <%
	sql="Select * from catarea where hi_id="+hid+" and catalog_id="+catalog+" order by name"
	Records.Source=sql
	Records.Open()
	while (!Records.EOF) {
		hname=String(Records("NAME").Value)
		tekhia=Records("ID").Value 
		if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {cu=Count_url(5,tekhia)}
		else {cu=Count_url(1,tekhia)}
		Records.MoveNext() %>
      <table border="1" cellspacing="2" bordercolor="#FFFFFF" width="98%" align="center">
        <tr bgcolor="#FBF8D7"> 
          <td height="18" bgcolor="#4594D8" width="20" align="center" bordercolor="#003366">&nbsp;</td>
          <td colspan="3" height="18" bgcolor="#EBF5ED" bordercolor="#003366"> 
            <p align="left"><b><a href="catarea.asp?hid=<%=tekhia%>"> <%=hname%></a></b> 
              (<%=cu%>) 
              <%if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {%>
              (<a href="edcatarea.asp?hid=<%=tekhia%>">изменить</a>&nbsp;--&nbsp;<a href="delcatarea.asp?hid=<%=tekhia%>">удалить</a>) 
              <%}%>
            </p>
          </td>
        </tr>
      </table>
      <%	}  Records.Close()
%>
      <%
 // Выводим УРЛы
 if (hid>0) {
	 if (isNaN(pg)) {pg=1}
	 if (pg<=0) {pg=1}
	sql="Select * from url where state=1 and catarea_id="+hid+" order by name" // для посетителей
	if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {sql="Select * from url where catarea_id="+hid+" order by name" } // для администратора
	//if (Session("hosturl")==1) {sql="Select * from url where state<4 and catarea_id="+hid} 
	Records.Source=sql.replace("*","count(*) as kvo")
	Records.Open()
	urlcount=Records("KVO").Value
	Records.Close()
	Records.Source=sql
	Records.Open()
	if (!Records.EOF) {
		//urlcount=Records.RecordCount
		if (Session("dd_url")!="undefined" && Session("dd_url")!="") {dd=parseInt(Session("dd_url"))}
		if (isNaN(dd)) {dd=10}
		if (urlcount>((pg-1)*dd)) {
			ii=0
			nextpg=(pg*dd)<=urlcount?(pg+1):-1
			while (ii < (pg-1)*dd) {
			  Records.MoveNext()
			  ii+=1
			}
			endlist=nextpg<0?urlcount:pg*dd
			endlist=endlist-1
			while (ii<=endlist) {
				ii+=1
				urlname=String(Records("NAME").Value)
				urlid=Records("ID").Value
				urlabout=String(Records("ABOUT").Value)
				daterenew=Records("NEW_DATE").Value
				urladr=String(Records("URL").Value)
				switch (Records("STATE").Value) {
					case -2 : statname="В повторной обработке"; break;
					case -1 : statname="В обработке"; break;
					case  0 : statname="Ожидает подтверждения"; break;
					case  1 : statname=""; break;
					case  2 : statname="Остановлен хозяином"; break;
					case  3 : statname="Остановлен администратором"; break;
					case  4 : statname="Удалён"; break;
				}
				Records.MoveNext()
 %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="6">
        <tr> 
          <td></td>
        </tr>
      </table>
      <table width="97%" border="1" cellspacing="0" bordercolor="#C5DCE2">
        <tr> 
          <td bgcolor="#C6DDFF" bordercolor="#E6ECEB"> 
            <p><a href="<%=urladr%>" target="_blank"><b><%=urlname%></b></a> </p>
          </td>
        </tr>
        <tr bordercolor="#FFFFFF"> 
          <td valign="top" bgcolor="#EBF3F5" bordercolor="#EBF3F5"> 
            <table width="100%" border="1" bordercolor="#EBF3F5" cellspacing="2">
              <tr valign="top" bgcolor="#FFFFFF" bordercolor="#C5DCE2"> 
                <td height="23" bordercolor="#C5DCE2" width="90%"> 
                  <p class="digest"><font size="2"><%=urlabout%></font></p>
                </td>
                <td height="23" width="120"> 
                  <p align="left"><font size="1" color="#FF0000"><b><%=statname%></b></font><font size="1">&nbsp;&nbsp;&nbsp; 
                    <%=daterenew%></font></p>
                </td>
              </tr>
            </table>
            <%if ((Session("is_adm_mem")==1)||(Session("cataloghost")==catalog)){%>
            <p>&nbsp; <font size="1"> <a href="delurl.asp?url=<%=urlid%>" target="_blank">удалить</a>&nbsp;&nbsp;&nbsp;<a href="urlresume.asp?url=<%=urlid%>&st=3" target="_blank">остановить</a> 
              &nbsp;&nbsp;&nbsp;<a href="urlresume.asp?url=<%=urlid%>&st=1" target="_blank">разместить</a>&nbsp; 
              &nbsp;<a href="removeurl.asp?url=<%=urlid%>&hid=<%=hid%>" target="_blank">переместить</a>&nbsp;&nbsp; 
              <a href="edurl.asp?url=<%=urlid%>" target="_blank">изменить</a></font></p>
            <%}%>
          </td>
        </tr>
      </table>
      <% } } } Records.Close()
} %>
      <p>&nbsp; </p>
      <p>Страницы:&nbsp;&nbsp; 
        <%
		  for ( ii=1; ii<(urlcount/dd + 1) ; ii++) {
		  if (ii != pg) {
		  %>
        <a href="catarea.asp?hid=<%=hid%>&pg=<%=ii%>"><%=ii%></a>&nbsp; 
        <%} else {%>
        <%=ii%>&nbsp; 
        <%}%>
        <%}%>
      </p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1">
        <tr> 
          <td background="/HeadImg/line_white.jpg"></td>
        </tr>
      </table>
      <div align="left"> 
        <% if (hid>0) { %>
      </div>
      <h2 align="center"><a href="addurl.asp?hid=<%=hid%>">Добавить сайт в раздел 
        <%=tit%><%=tit==""?"":":"%></a> </h2>
      <div align="left"> 
        <%}%>
      </div>
    </td>
  </tr>
</table>

