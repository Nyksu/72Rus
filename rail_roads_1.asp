<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->



<%

// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!

var hid=0
var hname=""
var url=""
var nid=0
var name=""
var ndat=""
var nadr=""
var per=0
var kvopub=0
var pname=""
var pdat=""
var autor=""
var digest=""
var imgLname=""
var imgname=""
var path=""
var hdd=0
var hadr=""
var nm=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var isnews=1
var blokname=""
var tpm=1000
var usok=false
var ishtml=0
var urlname=""
var urlid=0
var urlabout=""
var daterenew=""
var urladr=""
var urlcount=0
var msgcount=0
var sminame=""
var pid=0
var news=""


if (String(Session("id_mem"))=="undefined") {
	if (Session("tip_mem_pub")<3) {usok=true}
	tpm=Session("tip_mem_pub")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			tpm=0
		}
		Records.Close()
	} else {
		usok=true
		tpm=0
	}
}

var getData = Server.CreateObject("SOFTWING.ASPtear")
var dirData=""

sql="Select t1.* from url t1, catarea t2 where t1.recl_id=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
Records.Close()
sql="Select Count(*) as kvo from url t1, catarea t2 where t1.state=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog
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

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()
%>
<html>
<head>
<title>Расписание поездов - Тюмень, Тобольск, Сургут, Пыть-Ях, Лангепас, Мегион, Нижневартовск, Новый Уренгой &lt;&lt; 72rus.ru</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="72RUS.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
<SCRIPT language=JavaScript>


DOM=(document.getElementById)?1:0;
NS4=(document.layers)?1:0;
IE4=(document.all)?1:0;

function CHange_city(nomer){
        if(NS4){document.images['weather'].src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
        if(IE4){document.all['weather'].src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
        if(DOM){getElementsByName('weather').src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
}

</SCRIPT>
<script language="JavaScript">
 function goToUrl(){
  if(NS4)location.href="http://www.gismeteo.ru/weather/towns/"+document.select["City_pogoda"].value+".htm";
  if(IE4)location.href="http://www.gismeteo.ru/weather/towns/"+document.all["City_pogoda"].value+".htm";
  if(DOM)location.href="http://www.gismeteo.ru/weather/towns/"+getElementById("City_pogoda").value+".htm";
 }
</script>
</head>
<script language="JavaScript">
var flag_popup = 1;

var type=navigator.appName
 if (type=="Netscape") 
var lang = navigator.language ;
else 
var lang = navigator.userLanguage;
var lang = lang.substr(0,2) ;

function popup_off() {
	flag_popup = 0;
}

function popup() {
	if (flag_popup == 1 & lang != "ru") {
		window.open('http://www.rusintel.ru/english.asp');
	}
else if (lang == "ru") window.open('http://www.rusintel.ru');	
}

</script>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onUnload=popup();>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr> 
    <td bgcolor="#003366"><font color="#FFFFFF" class="digest"> Информационный 
      портал 72RUS - Тюменский Регион</font></td>
    <td bgcolor="#003366" align="right" width="300"><font size="1" color="#FFFFFF" face="Arial, Helvetica, sans-serif">Сейчас 
      на сайте посетителей: <b><%=Application("visitors")%></b></font><a href="admarea.asp"><img src="HeadImg/round_inv.gif" width="14" height="14" border="0" align="absmiddle"></a></td>
  </tr>
  <tr> </tr>
</table>
<table bgcolor=#000000 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right> </td>
  </tr>
  </tbody> 
</table>
<table border=0 cellpadding=0 cellspacing=0 height=60 width="100%">
  <tbody> 
  <tr> 
    <td align=middle valign=center width=150 bgcolor="#003366"><a href="/"><img height=60 
      src="HeadImg/logo_72_1.gif" width=150 border="0" alt="&gt;&gt; На главную страницу"></a></td>
    <td align=left width=300 valign="middle" bgcolor="#003366" background="HeadImg/logo_72_2.gif">
      <!--RAX counter-->
      <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
      <!--/RAX-->
    </td>
    <td bgcolor="#003366" background="HeadImg/logo_72_3.gif"><script language="javascript" src="banshow.asp?rid=4"></script></td>
  </tr>
  </tbody> 
</table>
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right bgcolor="#CCCCCC"> </td>
  </tr>
  </tbody> 
</table>
<table border=0 cellpadding=0 cellspacing=0 height=30 width="100%">
  <tbody> 
  <tr bgcolor="#003366"> 
    <td align=center valign=middle width=149 bgcolor="#003366" background="HeadImg/pass_bg.gif"> 
      <p><font color="#C6DDFF"><%=sminame%> - Тюмень<br>
        Тюменская область</font></p>
    </td>
    <td align=center width=302 bgcolor="#FF0000" valign="top"> 
      <table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=2 
width="100%">
        <tbody> 
        <tr> 
          <td align=right height="2"> </td>
        </tr>
        </tbody> 
      </table>
      <h1><font color="#FFFFFF">Ж.Д. Расписание</font></h1>
    </td>
    <td bgcolor="#003366" width="8" background="HeadImg/pass_bg.gif">&nbsp; </td>
    <td bgcolor="#003366" valign="middle" background="HeadImg/pass_bg.gif"> 
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=310>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <form name="form2" method="post" action="usrarea.asp">
            <td bgcolor="#003060" align="center" bordercolor="#C6DDFF" valign="middle"> 
              <p><font color="#C6DDFF">Имя:</font> 
                <input maxlength=20 name="logname" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/name.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=edit>
                <font color="#C6DDFF">Пароль:</font> 
                <input maxlength=25 name="psw" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=password>
                <input name="Enter" style="BACKGROUND-COLOR: #C6DDFF; COLOR: #003366; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 50px; HEIGHT: 20px" type=submit value=вход>
              </p>
            </td>
          </form>
        </tr>
        </tbody> 
      </table>
    </td>
  </tr>
  </tbody> 
</table>
<table bgcolor=#000000 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right> </td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="8" bgcolor="#003366">
  <tr> 
    <td></td>
  </tr>
</table>
<table width="151" border="0" cellspacing="0" cellpadding="0" align="left">
  <tr> 
    <th width="150" align="left" valign="top" bgcolor="#EBF3F5"> 
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
	kvopub=0
	if (name!="") {
		recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#EBF3F5"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> <a href="<%=url%>"><%=hname%></a></b></font></p>
          </td>
          <td width="32" bordercolor="#EBF3F5"><font face="Verdana" size="1"><%=kvopub%></font></td>
        </tr>
        </tbody> 
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
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	Records.MoveNext()
%>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#EBF3F5"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> <a href="<%=url%>"><%=hname%></a></b></font></p>
          </td>
          <td width="32" bordercolor="#EBF3F5"><font face="Verdana" size="1"><%=kvopub%></font></td>
        </tr>
        </tbody> 
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font face=Verdana size=1><b> 
              <a 
            href="messages.asp">Объявления</a> </b></font></p>
          </td>
          <td width="32" bordercolor="#EBF3F5" bgcolor="#EBF3F5"><font face="Verdana" size="1"><%=msgcount%></font></td>
        </tr>
        </tbody> 
      </table>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#EBF3F5"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font face=Verdana size=1><b> 
              <a href="lentamsg.asp">Объявления</a></b></font></p>
          </td>
          <td width="40" bordercolor="#EBF3F5" bgcolor="#EBF3F5"><font face="Verdana" size="1" color="#FF0000">Новые</font></td>
        </tr>
        </tbody> 
      </table>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#EBF3F5"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font face=Verdana size=1><b> 
              <a href="catarea.asp">Каталог сайтов</a></b></font></p>
          </td>
          <td width="32" bordercolor="#EBF3F5" bgcolor="#EBF3F5"><font face="Verdana" size="1"><%=urlcount%></font></td>
        </tr>
        </tbody> 
      </table>
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%">
        <tr> 
          <td bgcolor="#003366" align="center"><font color="#EBF3F5" class="digest">:: 
            Новый сайт :: </font> </td>
        </tr>
        <tr> </tr>
      </table>
      <table border="1" bordercolor="#003366" cellspacing="0" width="150">
        <tr> 
          <td bgcolor="#EBF3F5" bordercolor="#EBF3F5"> 
            <p class="digest"><img src="HeadImg/DOT.GIF" width="10" height="10"> 
              <a href="<%=urladr%>"><%=urlname%></a> - <%=urlabout%></p>
          </td>
        </tr>
        <tr> </tr>
      </table>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#C6DDFF" bordercolor="#C6DDFF"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"></font><font face=Verdana size=1><b> 
              <a href="air_russia.asp">АВИА Расписание </a></b></font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#C6DDFF" bordercolor="#C6DDFF"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"></font><font face=Verdana size=1><b> 
              <a 
            href="http://bn.72rus.ru">Баннерообмен</a></b></font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#C6DDFF" bordercolor="#C6DDFF"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"></font><font face=Verdana size=1><b> 
              <a 
            href="http://auction.72rus.ru">Аукцион 72rus.ru</a></b></font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#C6DDFF" bordercolor="#C6DDFF"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"></font><font face=Verdana size=1><b> 
              <a href="http://auto.72rus.ru">Авто.72rus.ru</a></b></font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border="0" cellspacing="0" width="100%" height="1">
        <tr> 
          <td bgcolor="#FFFFFF" bordercolor="#FFFFFF"> </td>
        </tr>
        <tr> </tr>
      </table>
      <font color="#000000"> </font> 
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%">
        <tr> 
          <td bgcolor="#003366" align="center"><font color="#FFFFFF" class="digest"> 
            <font color="#EBF3F5"> 
            <%
// В переменной bk содержится код блока новостей
var bk=32
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
            <%=blokname%> </font></font> </td>
        </tr>
        <tr> </tr>
      </table>
      <font color="#000000"> </font> 
      <table width="150" border="0" cellspacing="0" height="60" align="center" cellpadding="0">
        <tr> 
          <td><font color="#000000"> 
            <%
// В переменной bk содержится код блока новостей
var bk=32
// Не забывать его менять!!

Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)


var news=""
filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}


%>
            <b><a href="newshow.asp?pid=<%=pid%>"> 
            <p class="digest"><%=pname%></p>
            </a></b></font> 
            <%
Records.MoveNext()
} 
Records.Close()
%>
          </td>
          <td align="CENTER" width="1" bgcolor="#003366"></td>
 	</tr>
      </table>
      <font color="#000000"> </font>
<table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" height="1">
  <tr> 
    <td bgcolor="#003366"></td>
  </tr>
  <tr> </tr>
</table>
 </th>
    <td valign="top" width="1"></td>
    <td width="1" valign="top" bgcolor="#003366"></td>
  </tr>
</table>
<table width="151" border="0" cellspacing="0" cellpadding="0" align="right">
  <tr> 
    <td valign="top" width="1" rowspan="2"></td>
    <td width="1" valign="top" bgcolor="#003366" rowspan="2"></td>
    <td width="150" valign="top" bgcolor="#EBF3F5" align="center"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="6">
        <tr>
          <td></td>
        </tr>
      </table>

      <table width="110" border="1" cellspacing="0" cellpadding="0" bordercolor="#003366"><form name="form1" method="post" action="">
          <tr> 
            <td align="center" bgcolor="#003366" bordercolor="#003366"> 
              <p class="digest"><font color="#EBF3F5">Погода в городах</font></p>
            </td>
          </tr>
          <tr> 
            <td align="center" bordercolor="#FFFFFF"> 
              <select name="City_pogoda" onChange="CHange_city(this.options[this.selectedIndex].value)"
			style="BACKGROUND-COLOR: #FFF5CE; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 101px" >
                <option value="27612">Москва</option>
                <option value="28440">Екатеринбург</option>
                <option value="28367" selected>Тюмень</option>
                <option value="23748">Когалым</option>
                <option value="99981">Надым</option>
                <option value="23848">Нефтеюганск</option>
                <option value="23471">Нижневартовск</option>
                <option value="23358">Новый Уренгой</option>
                <option value="99968">Ноябрьск</option>
                <option value="23849">Сургут</option>
                <option value="23933">Ханты-Мансийск</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td align="center" bordercolor="#FFFFFF" height="105" bgcolor="#FFFFFF"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="ФОБОС: подробная погода в городе" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
          </tr>
              </form>
		  </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="6">
        <tr> 
          <td></td>
        </tr>
      </table>
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%">
        <tr> 
          <td bgcolor="#003366" align="center"><font color="#EBF3F5" class="digest">:: 
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
            <%=blokname%> </font> </td>
        </tr>
        <tr> </tr>
      </table>
      <p><table width="120" border="1" cellspacing="0" cellpadding="0" bordercolor="#003366">
        <tr> 
          <td align="CENTER" bordercolor="#EBF5ED"> 
            <script language="javascript" src="banshow.asp?rid=5"></script>
          </td>
        </tr>
      </table></p>
      <p align="CENTER"></p>
      </td>
  </tr>
  <tr>
    <td width="150" valign="top" bgcolor="#003366" align="center" height="1"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" height="850" align="center">
  <tr> 
    <td valign="top" width="1"></td>
    <td valign="top" align="center">
      <table border="1" bordercolor="#FFFFFF" cellspacing="0" width="99%" cellpadding="0">
        <tr bgcolor="#C6DDFF"> 
          <td nowrap bgcolor="#FFFFFF" height="6"></td>
        </tr>
        <tr bgcolor="#C6DDFF"> 
          <td nowrap bgcolor="#C6DDFF" bordercolor="#003366"><font class="digest"> 
            </font> 
            <p align="left"> <font face=Verdana size=1><b><a href="index.asp" title="На главную страницу"><img src="HeadImg/home.gif" width="14" height="14" align="absmiddle" border="0"> 
              <font color="#003366">72rus.ru</font></a></b></font> <font  face=Verdana size=1> 
              <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font> 
              <font  face=Verdana size=1><b> <font color="#003366">Расписание 
              поездов</font></b></font> <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"><font 
                  face=Verdana size=1><b> <a href="air_russia.asp"><font color="#003366">Расписание 
              самолетов</font></a></b></font></p>
          </td>
        </tr>
      </table>

        
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCFF" align="center">
        <tr bordercolor="#CCCCFF" valign="middle" align="center"> 
          <td bordercolor="#CCCCFF" height="11">
            <h1 align="left"><b><font color="#003366">Расписание поездов - Тюменская 
              область </font></b></h1>
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="99%" cellpadding="0">
              <tr bgcolor="#C6DDFF"> 
                <td nowrap bgcolor="#003366" height="6"></td>
              </tr>
              <tr bgcolor="#C6DDFF"> 
                <td nowrap bgcolor="#C6DDFF" background="HeadImg/shadow.gif" align="right"><font class="digest"> 
                  </font> 
                  <p><font 
                  face=Verdana size=1><b> <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
                    </b></font><b><font size="1"><a href="http://www.tours.ru/avia/order.asp?id_partners=159" target="_blank"><font face="Verdana, Arial, Helvetica, sans-serif"> 
                    Бронирование авиабилетов</font></a></font></b> </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class=n3 valign="top"> 
            </td>
        </tr>
      </table>
      <table border="1" cellpadding="0" cellspacing="3" width="98%" align="center" bordercolor="#FFFFFF">
        <tr bgcolor="#EBF3F5" bordercolor="#003366"> 
          <td valign="middle" width="40%"> 
            <p align="right"><b>Расписание Поездов по Станциям</b> </p>
          </td>
          <td valign="middle"> 
            <form action='http://www.poezda.net/st_shed.htm' method=post target="_blank">
              <p><br>
                <select name="st_code" size="1" class=text>
                  <option value="2030100" selected>Тюмень</option>
                  <option value="2030284">Тобольск</option>
                  <option value="2030607">Пыть-Ях</option>
                  <option value="2030600">Сургут</option>
                  <option value="2030153">Лангепас</option>
                  <option value="2030640">Мегион</option>
                  <option value="2030308">Нижневартовск</option>
                  <option value="2030319">Новый Уренгой</option>
                </select>
                &nbsp; 
                <input type=submit value='Поиск' size="5" class=text name="submit2">
                <input type="hidden"  name="pattern" value="Тюмень">
              </p>
            </form>
          </td>
        </tr>
        <tr bgcolor="#EBF3F5" bordercolor="#003366"> 
          <td valign="middle" height="50" width="40%"> 
            <p align="right">Ж.д. расписание по станции:</p>
          </td>
          <td valign="middle" height="50"> 
            <form action="http://www.poezda.net/detail.htm" method="post" target="_blank">
              <p> 
                <input type="text" name="st_name" value size="20" class=text>
                <input type="submit" value="Поиск" size="1" class=text name="submit">
              </p>
            </form>
          </td>
        </tr>
        <tr bgcolor="#EBF3F5" bordercolor="#003366"> 
          <td valign="middle" width="40%" height="53"> 
            <p align="right">Введите номер поезда:</p>
          </td>
          <td valign="middle" height="53"> 
            <form action="http://www.poezda.net/detail.htm" method="post" target="_blank">
              <p> 
                <input type="text" name="tr_num" value size="20" class=text>
                <input type="submit" value="Поиск" size="1" class=text name="submit">
              </p>
            </form>
          </td>
        </tr>
        <tr bgcolor="#C6DDFF" bordercolor="#003366"> 
          <td valign="middle" colspan="2"> 
            <form action="http://www.poezda.net/detail.htm" method="post" target="_blank">
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#EBF3F5"> 
                  <td width="40%"> 
                    <p align="right">Все поезда <b>ОТ</b> станции: </p>
                  </td>
                  <td> 
                    <p> 
                      <input type="text" name="st_from" value="Тюмень" size="20" class=text>
                    </p>
                  </td>
                </tr>
                <tr bgcolor="#EBF3F5"> 
                  <td width="40%"> 
                    <p align="right"><b>ДО</b> станции: </p>
                  </td>
                  <td> 
                    <p> 
                      <input type="text" name="st_to" value="Москва" size="20" class=text>
                      <input type="submit" value="Поиск" size="1" class=text name="submit3">
                    </p>
                  </td>
                </tr>
                <tr bgcolor="#EBF3F5"> 
                  <td width="40%">&nbsp;</td>
                  <td> 
                    <p> 
                      <input type="radio" name="search" value="prefics" checked>
                      Префиксный поиск </p>
                    <p> 
                      <input type="radio" name="search" value="contecst">
                      Контекстный поиск </p>
                  </td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="99%" cellpadding="0">
        <tr bgcolor="#C6DDFF"> 
          <td nowrap bgcolor="#003366" height="6"></td>
        </tr>
        <tr bgcolor="#C6DDFF"> 
          <td nowrap bgcolor="#C6DDFF" background="HeadImg/shadow.gif"><font class="digest"> 
            </font> 
            <p><font 
                  face=Verdana size=1><b> Лента новостей 72RUS.RU &gt;&gt; Тюменский 
              регион</b></font></p>
          </td>
        </tr>
      </table>
      <p align="left">
        <script language="javascript" src="http://72rus.ru/newsjv.asp?hid=125"></script>
      <p align="center"><font size="-2"><a href="http://72rus.ru"> 72RUS.RU - Тюменский регион</a></font></p>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center">
  <script language="javascript" src="banshow.asp?rid=6"></script>
  <hr size="1">
  <table bgcolor=#000000 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
    <tbody> 
    <tr> 
      <td align=right> </td>
    </tr>
    </tbody> 
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#C6DDFF"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b>Информационный 
          портал 72RUS - Тюменская Область </b></font><font size="1"><b>- Программирование 
          и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank">ЗАО 
          Русинтел</a> &copy; 2002 - 2003</font><font size="1"></font></b></p>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#003366"> 
        <p align="center">&nbsp;</p>
      </td>
    </tr>
  </table>
  <hr size="1">
  <p align="center"><font face="Arial, Helvetica, sans-serif"> 
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
    </font><br>
    | <a href="http://auto.72rus.ru">Авто.72rus</a> | <a href="http://www.auction.72rus.ru/">Аукцион</a> 
    | <a href="messages.asp">Объявления</a> | <a href="Rail_roads.asp">Расписание</a> 
    | <a href="catarea.asp">Тюменский Каталог</a> | <br>
    © 2002-2003 <a href="http://www.rusintel.ru">ЗАО Русинтел</a> 
</div>
<p align="center"><a href="http://www.rax.ru/click" target=_blank> 
  <%
// В переменной bk содержится код блока новостей
var bk=33
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"vvr_list.asp.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
var news=""
filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}
}

%>
  <%=news%> 
  <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
  <!-- HotLog -->
  <img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="rax.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодня"> 
  <!-- HotLog -->
  </a> 
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
</body>
</html>
