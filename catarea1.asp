<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

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
<html>
<head>
<title><%=tit%> / Каталог 72RUS Тюмень, Тюменская Область - <%=tit%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  1px 3px 0px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        <%=sminame%> : Каталог интернет сайтов</font></p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"> 
      <a href="/"><img src="images/72rus.gif" width="170" height="87" alt="72RUS.RU Тюменский Регион - информационный портал " border="0"></a> 
    </td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="banshow.asp?rid=4"></script>
    </td>
    <td background="images/fon02.gif" height="87" align="center" width="170">
      <table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="164" bgcolor="#F3F3F3" bordercolor="#E5E5E5"> 
            <p class="digest"> <a href="admarea.asp"><img src="images/e06.gif" width="16" height="9" alt="" border="0"></a> 
              <a href="#" class="digest">Посетителей сейчас:</a> <%=Application("visitors")%><br>
              <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="http://www.72rus.ru/newshow.asp?pid=728" class="digest">Реклама 
              на 72RUS.RU</a><br>
              <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="#" onClick="window.external.AddFavorite(parent.location,document.title)" class="digest">Добавить 
              в избранное</a><br>
			  <img src="images/e06.gif" width="16" height="9" alt="" border="0">
			  <a href="searchall.asp" class="digest">Поиск на сайте</a>
            </p>
</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#FF6600"> 
    <td colspan="4" height="1"></td>
  </tr>
</table>
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right bgcolor="#CCCCCC"> </td>
  </tr>
  </tbody> 
</table>
<table width="173" border="0" cellspacing="0" cellpadding="0" align="left">
  <tr> 
    <td width="172" align="left" valign="top"> 
      <%
// маркек признака новостей
isnews=1
// если необходимо вывести рубрики не новостей то установить в ноль

var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hdd
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hdd+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hdd+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
	}
	recs.Close()
	kvopub=0
	if (name!="") {
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%> 
              [<%=kvopub%>]</a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <%
// маркек признака новостей
isnews=0
// если необходимо вывести рубрики не новостей то установить в ноль
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hdd
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%></a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="messages.asp">Объявления [<%=msgcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="lentamsg.asp">Объявления 
              [новые]</a> </p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" width="170" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
              Каталог [<%=urlcountall%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auto.72rus.ru">Авто 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auction.72rus.ru">Аукцион 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="air_russia.asp">АВИА 
              Расписание</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="Rail_roads.asp">Расписание 
              поездов</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td valign="top"><img src="images/top01.gif" width="170" height="19" alt="" border="0"> 
            <img src="images/fon_menu04.gif" width="172" height="1"></td>
        </tr>
      </table>
      <font color="#000000"> </font> 
      <table width="100%" border="0" cellspacing="0" height="60" align="center" cellpadding="0">
        <tr> 
          <td align="center"> 
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
            <%=news_bl%> 
            <%
Records.MoveNext()
} 
Records.Close()
%>
          </td>
        </tr>
      </table>
 </td>
 </tr>
</table>
<table width="120" border="0" cellspacing="0" cellpadding="0" align="right" height="750">
  <tr> 
    <td valign="top" width="1" rowspan="2"></td>
    <td width="1" valign="top" bgcolor="#003366" rowspan="2"></td>
    <td width="150" valign="top" align="center"> 
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" align="center" bgcolor="#FB9700"> 
            <p class="menu01"> Пользователь</p>
          </td>
        </tr>
      </table>
      <table border=0 cellpadding=0 
      cellspacing=0 width=120 height="80">
        <tbody> 
        <tr> 
          <form name="form" method="post" action="usrarea.asp">
            <td align="center" valign="middle" bgcolor="#6699CC"> 
              <p> 
                <input maxlength=20 name="logname" style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/name.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" type=edit>
                <br>
                <input maxlength=25 name="psw" style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" type=password>
                <input name="Enter" style="BACKGROUND-COLOR: #C6DDFF; COLOR: #003366; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px; HEIGHT: 20px" type=submit value=вход>
                <br>
                <a href="regmemurl.asp"><font color="#FFFFFF">Регистрация</font></a></p>
            </td>
          </form>
        </tr>
        </tbody> 
      </table>
      
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" background="images/top01.gif" align="center" bgcolor="#FF9900"> 
            <p class="menu01"> 
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
              <%=blokname%></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <p> 
        <script language="javascript" src="banshow.asp?rid=5"></script>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" height="950" align="center">
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
          <td bgcolor="#FF9900" align="left" background="images/fon_menu08.gif"> 
            <h1><b><font color="#FFFFFF">Каталог сайтов: 
              <% if (hid==0) { %>
              WEB ресурсы 
              <%} else {%>
              <%=tit%><%=tit==""?"":":"%> </font></b></h1>
            <%}%>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="6">
        <tr> 
          <td></td>
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
          <td height="18" bgcolor="#4594D8" width="20" align="center" bordercolor="#003366"><font 
                  face=Verdana size=1><img src="images/arrow02.gif" width="7" height="7" align="middle"></font></td>
          <td colspan="3" height="18" bgcolor="#EBF5ED" bordercolor="#003366" background="HeadImg/shadow.gif"> 
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
          <td bgcolor="#C6DDFF" bordercolor="#E6ECEB" background="HeadImg/shadow.gif"> 
            <p><img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"><img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"><img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
              <a href="<%=urladr%>" target="_blank"><b><%=urlname%></b></a> </p>
          </td>
        </tr>
        <tr bordercolor="#FFFFFF"> 
          <td valign="top" bgcolor="#EBF3F5" bordercolor="#EBF3F5"> 
            <table width="100%" border="1" bordercolor="#EBF3F5" cellspacing="2">
              <tr valign="top" bgcolor="#FFFFFF" bordercolor="#C5DCE2">
                <td height="23" bordercolor="#C5DCE2" width="90%"> 
                  <p class="digest"><font size="2"><a href="<%=urladr%>" target="_blank"><img src="http://wb1.girafa.com/srv/i?i=979ce4bcc021347a&s=c74f64788bdb9fa4&r=<%=urladr%>" border="0" onLoad="if (this.width>50) this.border=1" alt="<%=urladr%>" align="left"></a><%=urlabout%></font></p>
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
        <table border="0" cellpadding="0" cellspacing="0" width="96%" align="center">
          <tr> 
            <td colspan="3" height="1" bgcolor="#666666"></td>
          </tr>
          <tr> 
            <td colspan="3" height="25" align="center"> 
              <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
                <img src="images/px1.gif" width="1" height="1" alt="" border="0"><a href="http://www.rusintel.ru/newshow.asp?pid=2496">Регистрация 
                доменов RU </a> <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
                <a href="http://www.rusintel.ru/">Разработка интернет сайтов</a> 
                <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
                <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2256">Хостинг</a> 
                <img src="images/e06.gif" width="16" height="9" alt="" border="0"><a href="http://www.72rus.ru/newshow.asp?pid=728"> 
                Размещение рекламы</a></p>
            </td>
          </tr>
        </table>
      </div>
    </td>
  </tr>
  <tr> 
    <td valign="top" align="center" height="15"> 
      <p><font color="#666666" size="1">Для того, чтобы добавить сайт в каталог 
        Вам необходимо пройти <a href="regmemurl.asp">Регистрацию Нового Пользователя</a>. 
        После этого, Вы получаете имя пользователя и пароль и сможете в любое 
        время добавлять, удалять свои сайты, изменять их описание, адрес в случае 
        его изменения.</font> </p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td height="1" bgcolor="#666666"></td>
  </tr>
  <tr> 
    <td height="19" bgcolor="#AFC0D0" align="center">&nbsp; </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td colspan="3"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
  </tr>
  <tr> 
    <td height="70" width="180"> 
      <p align="right">&nbsp;</p>
    </td>
    <td align="center"> 
      <p class="menu02">| 
        <%
// маркек признака новостей
isnews=1
// если необходимо вывести рубрики не новостей то установить в ноль

var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
	}
	recs.Close()
	Records.MoveNext()
%>
        <a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <%
// маркек признака новостей
isnews=0
// если необходимо вывести рубрики не новостей то установить в ноль
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
}
	recs.Close()
	
	Records.MoveNext()
%>
        <a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <br>
        | <a href="http://auto.72rus.ru">Авто.72rus</a> | <a href="http://www.auction.72rus.ru/">Аукцион</a> 
        | <a href="messages.asp">Объявления</a> | <a href="Rail_roads.asp">Расписание</a> 
        | <a href="catarea.asp">Тюменский Каталог</a> | 
    </td>
    <td width="180"> 
      <p>Сделано в <a href="http://www.rusintel.ru">ЗАО Русинтел</a> <br>
        WWW.72RUS.RU © 2002-2004 </p>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=6"></script>
</div>
<hr size="1">
<div align="center"> 
  <%
// В переменной bk содержится код блока новостей
var bk=33
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
  <%=news_bl%> 
  <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
  <!--LiveInternet logo-->
  <a href="http://www.liveinternet.ru/click" target=liveinternet><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="liveinternet.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодня"></a> 
  <!--/LiveInternet-->
  <script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=46088&im=105&r="+escape(document.referrer)+"&pg="+
escape(window.location.href);
document.cookie="hotlog=1; path=/"; hotlog_r+="&c="+(document.cookie?"Y":"N");
</script>
  <script language="javascript1.1">
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
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"><img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
</div>
</body>
</html>
