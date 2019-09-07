<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

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
var urlcountall=0

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""


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
<html>
<head>
<title><%=tit%> / Каталог 72RUS Тюмень, Тюменская Область - <%=tit%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72RUS.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0">
    <tr> 
      <td width="16" bgcolor="#003366">&nbsp;</td>
      <td bgcolor="#003366"> 
        <p><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Тюмень 
          и Тюменская Область ::</b></font></p>
      </td>
      <td bgcolor="#003366" width="23"><a href="admarea.asp"><img src="HeadImg/round_inv.gif" width="23" height="23" border="0"></a></td>
    </tr>
  </table>
  
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="68">
  <tr> 
    <td width="300" align="center" valign="middle"> <a href="/"><img src="HeadImg/logo_72.gif" width="300" height="60" border="0" alt="На главную страницу"></a></td>
    <td align="left" width="468"> 
      <table width="120" border="0" cellspacing="0" height="60">
        <tr> 
          <%
// В переменной bk содержится код блока новостей
var bk=29
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{

	puid=String(Records("ID").Value)



var news_bl=""
filnam=PubFilePath+puid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml2==0) {
	while (!ts.AtEndOfStream){
		news_bl+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
	}
	} else {news_bl=ts.ReadAll()}
	ts.Close()
}

%>
          <td align="CENTER"><%=news_bl%></td>
          <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
        </tr>
      </table>
    </td>
    <td align="left"> 
      <!--begin of Rambler's Top100 code -->
      <a href="http://top100.rambler.ru/top100/"> <img src="http://counter.rambler.ru/top100.cnt?388244" alt="" width=1 height=1 border=0></a> 
      <!--end of Top100 code-->
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="8" bgcolor="#003366">
  <tr> 
    <td></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="8" bgcolor="#003366">
  <tr> 
    <td width="150" align="center"> 
      <p align="center"> <font size="3"><b><font color="#FFFFFF"> Тюмень <a href="/"><font color="#FFFFFF">72rus.ru</font></a></font></b></font></p>
    </td>
    <td bgcolor="#FFFFFF" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="23" valign="top"><img src="HeadImg/round_blu_red.gif" width="23" height="23"></td>
          <td width="285" valign="top"> 
            <table border="0" cellspacing="0" cellpadding="0" width="285" bordercolor="#FFFFFF">
              <tr> 
                <td align="left" width="23" bgcolor="#FF0000">&nbsp;</td>
                <td bgcolor="#FF0000" align="center"> 
                  <p><font size="3"><font color="#FFFFFF"><b>Каталог сайтов</b></font></font> 
                </td>
                <td bgcolor="#FF0000" align="left" width="23"> <img src="HeadImg/round_red_up_r.gif" width="23" height="23"></td>
              </tr>
            </table>
          </td>
          <td align="right">&nbsp;</td>
          <td align="center" width="150"> 
            <%
// В переменной bk содержится код блока новостей
var bk=35
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
            <p><b><font size="3"><%=blokname%></font></b> </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="150" valign="top"> 
      <%
isnews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><b><font  face=Verdana size=1><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <a  class=globalnav href="<%=url%>"><%=hname%></a> </font></b></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <%
} 
Records.Close()
%>
      <%
isnews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><b><font  face=Verdana size=1><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <a  class=globalnav href="<%=url%>"><%=hname%></a></font></b></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <%
} 
Records.Close()
%>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              </font><font 
                  face=Verdana size=1><b><a href="messages.asp">Объявления</a> 
              </b>(<%=msgcount%>) </font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              </font><font 
                  face=Verdana size=1><b><a href="catarea.asp">Каталог сайтов</a></b> 
              (<%=urlcountall%>)</font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#003366 >&nbsp; </td>
          <td bgcolor=#003366 width="23"> <img src="HeadImg/round_inv.gif" width="23" height="23"></td>
        </tr>
        </tbody> 
      </table>
      <%
// В переменной bk содержится код блока новостей
var bk=38
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	puid=String(Records("ID").Value)
	ishtml2=TextFormData(Records("ISHTML").Value,"0")
	filnam=PubFilePath+puid+".pub"
	if (!fs.FileExists(filnam)) { filnam="" }

	if (filnam != "") {
		ts= fs.OpenTextFile(filnam)
		if (ishtml2==0){
			while (!ts.AtEndOfStream){
				news_bl+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
			}
		} else {news_bl=ts.ReadAll()}
		ts.Close()
	}

%>
      <table width="127" border="0" cellspacing="0" height="60">
        <tr> 
          <td align="CENTER"><%=news_bl%></td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
%>
    </td>
    <td valign="top"> 
      <table width="100%" border="0" cellspacing="3" cellpadding="0">
        <tr valign="middle" align="center"> 
          <td> 
            <div align="left"><font face="Verdana" size="1"><b>Вы здесь</b></font> 
              <font size="2"> 
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
              <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <a href="<%=hiadr%>"><%=hname%></a> 
              <%
		  	} else {
		%>
              <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> 
              <%=hname%> 
              <%
		  	}	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
              </b></font> <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <a href="catarea.asp?hid=0"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Каталог</b></font></a> 
              <% } else {%>
              <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <a href="index.asp"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">На 
              главную страницу</font></b></a> 
              <%} %>
              </font></div>
          </td>
        </tr>
      </table>
      <h1><b><font size="3"><%=tit%><%=tit==""?"":":"%> </font></b></h1>
      <% if (hid>0) { %>
      <h2 align="center"><a href="addurl.asp?hid=<%=hid%>">Добавить 
        сайт в раздел</a> </h2>
      <%}%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1">
        <tr> 
          <td background="/HeadImg/line_white.jpg"></td>
        </tr>
      </table>
      <div align="left"> 
        <%if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {%>
        <table width="40%" border="1" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
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
                <p><font size="1"><a href="addurls.asp?hid=<%=hid%>">Добавить 
                  URL</a></font></p>
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
      </div>
      <p align="left"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
        <b><a href="catarea.asp?hid=<%=tekhia%>"> 
        <%=hname%></a></b> (<%=cu%>) 
        <%if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {%>
        (<a href="edcatarea.asp?hid=<%=tekhia%>">изменить</a>&nbsp;--&nbsp;<a href="delcatarea.asp?hid=<%=tekhia%>">удалить</a>) 
        <%}%>
      </p>
      <div align="left"> 
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
        <table width="100%" border="0" bordercolor="#FFFFFF" cellspacing="0">
          <tr> 
            <td width="6%"> 
              <div align="center"> 
                <p><b><%=ii%>. </b></p>
              </div>
            </td>
            <td width="65%" bgcolor="#FFFFCC"> 
              <p><b><a href="<%=urladr%>" target="_blank"><%=urlname%></a></b></p>
            </td>
            <td bgcolor="#FFFFCC"> 
              <div align="center"> 
                <p align="left"><font size="1" color="#FF0000"><b><%=statname%></b></font><font size="1">&nbsp;&nbsp;&nbsp; подтвержден 
                  (<%=daterenew%>)</font></p>
              </div>
            </td>
          </tr>
          <td> 
            <p>&nbsp;</p>
          </td>
          <td> 
            <p><font size="2"><%=urlabout%></font></p>
          </td>
          <td>&nbsp; 
            <%if ((Session("is_adm_mem")==1)||(Session("cataloghost")==catalog)){%>
            <a href="delurl.asp?url=<%=urlid%>">удалить</a>&nbsp;&nbsp;&nbsp;<a href="urlresume.asp?url=<%=urlid%>&st=3">остановить</a> 
            &nbsp;&nbsp;&nbsp;<a href="urlresume.asp?url=<%=urlid%>&st=1">разместить</a> 
            <%}%>
          </td>
          <tr> </tr>
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
      </div>
      <div align="left"> 
        <% if (hid>0) { %>
      </div>
      <h2 align="center"><a href="addurl.asp?hid=<%=hid%>">Добавить 
        сайт в раздел <%=tit%><%=tit==""?"":":"%></a> </h2>
      <div align="left"> 
        <%}%>
      </div>
      <p align="left"></p>
    </td>
    <td valign="top" width="150" align="center"> 
      <h1 align="left"> 
        <%

// В переменной bk содержится код блока новостей
var bk=35
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	puid=String(Records("ID").Value)
	ishtml2=TextFormData(Records("ISHTML").Value,"0")
	filnam=PubFilePath+puid+".pub"
	if (!fs.FileExists(filnam)) { filnam="" }

	if (filnam != "") {
		ts= fs.OpenTextFile(filnam)
		if (ishtml2==0){
			while (!ts.AtEndOfStream){
				news_bl+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
			}
		} else {news_bl=ts.ReadAll()}
		ts.Close()
	}

%>
      </h1>
      <table width="120" border="0" cellspacing="0" height="60">
        <tr> 
          <td align="CENTER"><%=news_bl%></td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
%>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <table width="720" border="0" cellspacing="0" height="60">
    <tr> 
      <%
// В переменной bk содержится код блока новостей
var bk=36
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	puid=String(Records("ID").Value)
	ishtml2=TextFormData(Records("ISHTML").Value,"0")
	filnam=PubFilePath+puid+".pub"
	if (!fs.FileExists(filnam)) { filnam="" }

	if (filnam != "") {
		ts= fs.OpenTextFile(filnam)
		if (ishtml2==0){
			while (!ts.AtEndOfStream){
				news_bl+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
			}
		} else {news_bl=ts.ReadAll()}
		ts.Close()
	}

%>
      <td align="CENTER"><%=news_bl%></td>
      <%
Records.MoveNext()
} 
Records.Close()
%>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#FF0000"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
          портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
          Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
          Русинтел</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
      </td>
    </tr>
  </table>
  <hr size="1">
  <p align="center">| <a href="http://auto.72rus.ru">Авто Тюмень</a> | <a href="http://www.auction.72rus.ru/">Аукцион</a> 
    | <a href="messages.asp">Объявления</a> | <a href="Rail_roads.asp">Расписание</a> 
    | <a href="catarea.asp">Тюменский Каталог</a> | <a href="russia.asp">Глобальный 
    Каталог</a> |<br>
    | <a href="terms.html">Условия использования</a> | <a href="adv.html">Реклама 
    на сервере</a> | 
    <% if (hid>0) { %>
    <a href="addurl.asp?hid=<%=hid%>">Добавить 
    сайт</a> 
    <%}%>
    <br>
    © 2002 <a href="http://www.rusintel.ru">Rusintel Company</a> 
</div>
<p align="center"> 
  <!-- HotLog -->
<script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=46088&im=105&r="+escape(document.referrer)+"&pg="+
escape(window.location.href);
document.cookie="hotlog=1; path=/"; hotlog_r+="&c="+(document.cookie?"Y":"N");
</script><script language="javascript1.1">
hotlog_js="1.1";hotlog_r+="&j="+(navigator.javaEnabled()?"Y":"N")</script>
<script language="javascript1.2">
hotlog_js="1.2";
hotlog_r+="&wh="+screen.width+'x'+screen.height+"&px="+
(((navigator.appName.substring(0,3)=="Mic"))?
screen.colorDepth:screen.pixelDepth)</script>
<script language="javascript1.3">hotlog_js="1.3"</script>
<script language="javascript">hotlog_r+="&js="+hotlog_js;
document.write("<a href='http://click.hotlog.ru/?46088' target='_top'><img "+
" src='http://hit3.hotlog.ru/cgi-bin/hotlog/count?"+
hotlog_r+"&' border=0 width=88 height=31 alt=HotLog></a>")</script>
<noscript><a href=http://click.hotlog.ru/?46088 target=_top><img
src="http://hit3.hotlog.ru/cgi-bin/hotlog/count?s=46088&im=105" border=0 
width="88" height="31" alt="HotLog"></a></noscript>
<!-- /HotLog -->
  <!--Begin of HMAO RATINGS-->
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"> <img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
</body>
</html>
