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
<title>Расписание самолетов - Тарифы на Авиаперелеты - 72RUS Тюмень</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="style1.css" type="text/css">
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
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        <%=sminame%> : Расписание самолетов, Тарифы</font></p>
    </td>
    <td width="170"> 
      <p class="menu01"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="#" onClick="window.external.AddFavorite(parent.location,document.title)"><font color="#000000">Добавить 
        в избранное</font></a></p>
    </td>
    <td align="center" width="200"> 
      <p class="menu01"><a href="admarea.asp"><img src="images/e06.gif" width="16" height="9" alt="" border="0"></a> 
        <font color="#000000">посетителей на сайте: <%=Application("visitors")%></font></p>
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
  </tr>
  <tr bgcolor="#FF6600"> 
    <td colspan="3" height="1"></td>
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
              Каталог [<%=urlcount%>]</a></p>
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
    <td width="150" valign="top" align="center"> 
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" align="center" bgcolor="#FB9700"> 
            <p class="menu01"> Погода</p>
          </td>
        </tr>
      </table>
      <table width="120" border="0" cellspacing="0" cellpadding="0" bordercolor="#003366" align="center">
        <form name="form" method="post" action="">
          <tr> 
            <td align="center" bgcolor="#6699CC" height="20" valign="bottom"> 
              <select name="select" onChange="CHange_city(this.options[this.selectedIndex].value)"
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
            <td align="center" height="106" bgcolor="#6699CC" valign="top"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="ФОБОС: подробная погода в городе" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
          </tr>
        </form>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" align="center" bgcolor="#FB9700"> 
            <p class="menu01">Курс $ ЦБ РФ</p>
          </td>
        </tr>
      </table>
      <table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" height="105" bgcolor="#6699CC"><a href="http://www.informer.ru/cgi-bin/redirect.cgi?id=177_1_1_33_40_1-0&url=http://www.rbc.ru&src_url=usd/eur_cb_forex_000066_88x90.gif"
target="_blank"><img src="http://pics.rbc.ru/img/grinf/usd/eur_cb_forex_000066_88x90.gif " width=88 height="90" border=0></a></td>
        </tr>
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
<table border="0" cellspacing="0" cellpadding="0" height="850" align="center">
  <tr> 
    <td valign="top" width="1"></td>
    <td valign="top" align="center"> 
      <table width="100%" border="0" cellspacing="0" bordercolor="#003366">
        <tr bgcolor="#FBF8D7"> 
          <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle"> 
            <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="index.asp">72RUS.RU</a> / Расписание самолетов /</p>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCFF" align="center">
        <tr bordercolor="#CCCCFF" valign="middle" align="center"> 
          <td bordercolor="#CCCCFF" height="11"> 
            <table width="97%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
              <tr> 
                <td bgcolor="#FF9900" align="left" background="images/fon_menu08.gif"> 
                  <h1><b><font color="#FFFFFF">Расписание самолетов</font></b>&nbsp;</h1>
                </td>
              </tr>
            </table>
            
          </td>
        </tr>
        <tr> 
          <td class=n3 valign="top"> 
            <form action="http://www.polets.ru/cgi-bin/sh.pl" method=get name="f" target="_blank">
              <table border=0 cellpadding=0 cellspacing=0>
                <tbody> 
                <tr valign="top"> 
                  <td height="328"> 
                    <table border=0 cellpadding=5 cellspacing=0 
                  width=99% align="center">
                      <tbody> 
                      <tr> 
                        <td rowspan=10 width=15>&nbsp;</td>
                        <td class=n2b> 
                          <input name="Mode" type=hidden value="Naprav">
                          <input name="Source" type=hidden  value="Brief">
                          <p><b>Город отправления:</b></p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2 bgcolor="#EBF3F5"> 
                          <p> 
                            <select name="CityFrom" size="1" class=text>
                              <option value="0">* Выберите город 
                              <option value="all">* Все направления 
                              <option value="abn">Абакан 
                              <option value="ahl">Айхал 
                              <option value="aau">Актау 
                              <option value="ald">Алдан 
                              <option value="ala">Алматы 
                              <option value="amd">Амдерма 
                              <option value="ade">Амдерма-2 
                              <option value="any">Анадырь 
                              <option value="ana">Анапа 
                              <option value="anv">Андижан 
                              <option value="aph">Апатиты 
                              <option value="arh">Архангельск 
                              <option value="akl">Астана 
                              <option value="asr">Астрахань 
                              <option value="aty'>Атырау 
                                      <option value="a{h">Ашхабад 
                              <option value="aqn">Аян 
                              <option value="bgd">Багдарин 
                              <option value="bkt">Байкит 
                              <option value="lnj">Байконур 
                              <option value="bak">Баку 
                              <option value="ban">Барнаул 
                              <option value="bcn">Барселона 
                              <option  value="btg">Батагай 
                              <option value="bui">Батуми 
                              <option value="bqg">Белая Гора 
                              <option value="bed">Белгород 
                              <option value="blr">Белоярский 
                              <option value="bzk">Березники 
                              <option value="ber">Березово 
                              <option value="bng">Беринговский 
                              <option value="bsk">Бийск 
                              <option value="bi{">Бишкек 
                              <option value="bg}">Благовещенск 
                              <option value="bdb">Бодайбо 
                              <option value="brs">Братск 
                              <option value="bug">Бугульма 
                              <option value="bgs">Бургас 
                              <option value="bur">Буревестник 
                              <option value="bua">Бухара 
                              <option value="prd">Бухта Провидения 
                              <option value="wrq">Ванавара 
                              <option value="war">Варна 
                              <option value="weu">Великий Устюг 
                              <option value="whw">Верхневилюйск 
                              <option value="wik">Вилюйск 
                              <option value="wwo">Владивосток 
                              <option value="wla">Владикавказ 
                              <option value="wgg">Волгоград 
                              <option value="wld">Волгодонск 
                              <option value="wgd">Вологда 
                              <option value="wkt">Воркута 
                              <option value="wrn">Воронеж 
                              <option value="wtg">Вытегра 
                              <option value="haj">Ганновер 
                              <option value="gdv">Геленджик 
                              <option value="g`m">Гюмри 
                              <option value="gnv">Гянджа 
                              <option value="dlc">Далянь 
                              <option value="dep">Депутатский 
                              <option value="dvh">Джебарики-Хая 
                              <option value="dpk">Днепропетровск 
                              <option value="dnc">Донецк 
                              <option value="dnk">Дудинка 
                              <option value="d{b">Душанбе 
                              <option value="DUS">Дюссельдорф 
                              <option value="esk">Ейск 
                              <option value="ekb">Екатеринбург 
                              <option value="eg~">Ербогачен 
                              <option value="ewn">Ереван 
                              <option value="vig">Жиганск 
                              <option value="zla">Залив Лаврентия 
                              <option value="zpv">Запорожье 
                              <option value="zon">Зональное 
                              <option value=znk>Зырянка 
                              <option value="igr">Игарка 
                              <option value=irm>Игрим 
                              <option value="ivw">Ижевск 
                              <option value="int">Инта 
                              <option value="ikt">Иркутск 
                              <option value="i{o">Йошкар-Ола 
                              <option value="kzn">Казань 
                              <option value="kz~">Казачинск 
                              <option value="kld">Калининград 
                              <option value="klg">Калуга 
                              <option value="kgd">Караганда 
                              <option value="kdd">Кедровый 
                              <option  value="krw">Кемерово 
                              <option value="kpm">Кепервеем 
                              <option value="iew">Киев 
                              <option value="krn">Киренск 
                              <option value="kio">Киров 
                              <option value="k~g">Кичменгский Городок 
                              <option value="k{n">Кишинев 
                              <option value="kog">Когалым 
                              <option value="kzi">Кодинск 
                              <option value="ksl">Комсомольск-На-Амуре 
                              <option value="ktn">Костанай 
                              <option value="kts">Котлас 
                              <option value="krr">Краснодар 
                              <option value="kkx">Красноселькуп 
                              <option value="kqa">Красноярск 
                              <option value="kgn">Курган 
                              <option value="kus">Курск 
                              <option value="kis">Кутаиси 
                              <option value="kyy">Кызыл 
                              <option value="lsk">Ленск 
                              <option value="le{">Лешуконское 
                              <option value="lip">Липецк 
                              <option value="lod">Лондон 
                              <option value="lwo">Львов 
                              <option value="mdn">Магадан 
                              <option value="mgn">Маган 
                              <option value="mgs">Магнитогорск 
                              <option value="mkp">Майкоп 
                              <option value="mam">Мама 
                              <option value="mrn">Мариинское 
                              <option value="mko">Марково 
                              <option value="mhl">Махачкала 
                              <option 
                          value="mrw">Минеральные Воды 
                              <option 
                          value="msk">Минск 
                              <option value="mir">Мирный 
                              <option 
                          value="mom">Мома 
                              <option class="hl" value="mow">Москва 
                              <option 
                          value="mtg">Мотыгино 
                              <option value="mun">Мурманск 
                              <option 
                          value="mkm">Мыс Каменный 
                              <option value="m{d">Мыс Шмидта 
                              <option value="m`n">Мюнхен 
                              <option 
                          value="ndm">Надым 
                              <option value="in{">Назрань 
                              <option 
                          value="n~k">Нальчик 
                              <option value="nmg">Наманган 
                              <option 
                          value="nnr">Нарьян-Мар 
                              <option 
                          value="nh~">Нахичевань 
                              <option value="nln">Нелькан 
                              <option 
                          value="nrg">Нерюнгри 
                              <option value="n`g">Нефтеюганск 
                              <option 
                          value="nvg">Нижнеангарск 
                              <option 
                          value="nvw">Нижневартовск 
                              <option 
                          value="nvk">Нижнекамск 
                              <option value="nvq">Нижнеянск 
                              <option 
                          value="nvs">Нижний Новгород 
                              <option 
                          value="nlk">Николаевск-На-Амуре 
                              <option 
                          value="nik">Никольское 
                              <option 
                          value="nwk">Новокузнецк 
                              <option 
                          value="owb">Новосибирск 
                              <option value="nur">Новый Уренгой 
                              <option value="ngl">Ноглики 
                              <option 
                          value="nrs">Норильск 
                              <option value="noq">Ноябрьск 
                              <option 
                          value="nus">Нукус 
                              <option value="n`r">Нюрба 
                              <option 
                          value="nqg">Нягань 
                              <option value="ods">Одесса 
                              <option 
                          value="ozr">Озерная 
                              <option value="omk">Оймякон 
                              <option 
                          value="olk">Олекминск 
                              <option value="oln">Оленек 
                              <option 
                          value="ool">Омолон 
                              <option value="oms">Омск 
                              <option 
                          value="ong">Оренбург 
                              <option value="osk">Орск 
                              <option 
                          value="oso">Оссора 
                              <option value="oha">Оха 
                              <option 
                          value="oht">Охотск 
                              <option value="pwl">Павлодар 
                              <option 
                          value="pan">Палана 
                              <option value="pew">Певек 
                              <option 
                          value="prm">Пермь 
                              <option value="ptz">Петрозаводск 
                              <option 
                          value="prl">Петропавловск-Камчатский 
                              <option 
                          value="p~r">Печора 
                              <option value="pin">Пионерный 
                              <option 
                          value="pts">Подкаменная Тунгуска 
                              <option value="pos">Полины Осипенко 
                              <option value="plq">Полярный 
                              <option 
                          value="rad">Радужный 
                              <option value="rix">Рига 
                              <option 
                          value="row">Ростов-На-Дону 
                              <option 
                          value="sky">Саккырыр 
                              <option value="shd">Салехард 
                              <option 
                          value="sm{">Самара 
                              <option value="skd">Самарканд 
                              <option 
                          value="sag">Сангар 
                              <option class=hl 
                          value="spt">Санкт-Петербург 
                              <option 
                          value="srn">Саранск 
                              <option value="sro">Саратов 
                              <option 
                          value="syh">Саскылах 
                              <option 
                          value="sen">Северо-Енисейск 
                              <option 
                          value="swe">Северо-Эвенск 
                              <option value="ski">Сенаки 
                              <option 
                          value="sel">Сеул 
                              <option value="sip">Симферополь 
                              <option 
                          value="sbo">Соболево 
                              <option value="sog">Советская Гавань 
                              <option value="soj">Советский 
                              <option 
                          value="son">Солнечный 
                              <option value="so~">Сочи 
                              <option 
                          value="srm">Средне-Колымск 
                              <option 
                          value="stw">Ставрополь 
                              <option value="ist">Стамбул 
                              <option 
                          value="sol">Старый Оскол 
                              <option 
                          value="spw">Степанаван 
                              <option value="stv">Стрежевой 
                              <option 
                          value="sun">Сунтар 
                              <option value="sur">Сургут 
                              <option 
                          value="syw">Сыктывкар 
                              <option value="tio">Таксимо 
                              <option 
                          value="tas">Ташкент 
                              <option value="tbs">Тбилиси 
                              <option 
                          value="tlq">Тель-Авив 
                              <option value="tig">Тигиль 
                              <option 
                          value="tsi">Тикси 
                              <option value="til">Тиличики 
                              <option 
                          value="txk">Толька 
                              <option value="tsk">Томск 
                              <option 
                          value="tau">Тура 
                              <option value="trh">Туруханск 
                              <option 
                          value="t`m" selected>Тюмень 
                              <option value="tsn">Тянь-зинь 
                              <option 
                          value="uln">Улан-Батор 
                              <option value="ul|">Улан-Удэ 
                              <option 
                          value="ulk">Ульяновск 
                              <option value="ura">Урай 
                              <option 
                          value="ugn">Ургенч 
                              <option value="usn">Усинск 
                              <option 
                          value="utk">Усть-Илимск 
                              <option 
                          value="ukg">Усть-Каменогорск 
                              <option 
                          value="uk~">Усть-Камчатск 
                              <option 
                          value="uku">Усть-Куйга 
                              <option value="usk">Усть-Кут 
                              <option 
                          value="usm">Усть-Мая 
                              <option value="unr">Усть-Нера 
                              <option 
                          value="uhj">Усть-Хайрюзово 
                              <option 
                          value="utc">Усть-Цильма 
                              <option value="ufa">Уфа 
                              <option 
                          value="uht">Ухта 
                              <option value="fgn">Фергана 
                              <option 
                          value="fra">Франкфурт-На-Майне 
                              <option 
                          value="hbr">Хабаровск 
                              <option value="hdy">Хандыга 
                              <option 
                          value="has">Ханты-Мансийск 
                              <option 
                          value="hrk">Харьков 
                              <option value="hat">Хатанга 
                              <option 
                          value="hel">Хельсинки 
                              <option value="hrp">Херпучи 
                              <option 
                          value="hdt">Худжанд 
                              <option value="~ar">Чара 
                              <option 
                          value="~be">Чебоксары 
                              <option value="~lb">Челябинск 
                              <option 
                          value="~rw">Череповец 
                              <option value="~rs">Черский 
                              <option 
                          value="sht">Чита 
                              <option value="~kd">Чокурдах 
                              <option 
                          value="~mi">Чумикан 
                              <option value="{ah">Шахтерск 
                              <option 
                          value="{nq">Шеньян 
                              <option value="{mt">Шымкент 
                              <option 
                          value="|gt">Эгвекинот 
                              <option value="|li">Элиста 
                              <option 
                          value="`vk">Южно-Курильск 
                              <option 
                          value="`vh">Южно-Сахалинск 
                              <option 
                          value="qkt">Якутск 
                              <option value="qmb">Ямбург 
                              <option value="qrl">Ярославль</option>
                            </select>
                          </p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2b> 
                          <p><b>Город прибытия:</b></p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2 bgcolor="#EBF3F5"> 
                          <p> 
                            <select name="CityTo" size="1" class=text>
                              <option value="0">* Выберите город 
                              <option value="all">* Все направления 
                              <option value="abn">Абакан 
                              <option value="ahl">Айхал 
                              <option value="aau">Актау 
                              <option value="ald">Алдан 
                              <option value="ala">Алматы 
                              <option value="amd">Амдерма 
                              <option value="ade">Амдерма-2 
                              <option value="any">Анадырь 
                              <option value="ana">Анапа 
                              <option value="anv">Андижан 
                              <option value="aph">Апатиты 
                              <option value="arh">Архангельск 
                              <option value="akl">Астана 
                              <option value="asr">Астрахань 
                              <option value="aty'>Атырау 
                                      <option value="a{h">Ашхабад 
                              <option value="aqn">Аян 
                              <option value="bgd">Багдарин 
                              <option value="bkt">Байкит 
                              <option value="lnj">Байконур 
                              <option value="bak">Баку 
                              <option value="ban">Барнаул 
                              <option value="bcn">Барселона 
                              <option  value="btg">Батагай 
                              <option value="bui">Батуми 
                              <option value="bqg">Белая Гора 
                              <option value="bed">Белгород 
                              <option value="blr">Белоярский 
                              <option value="bzk">Березники 
                              <option value="ber">Березово 
                              <option value="bng">Беринговский 
                              <option value="bsk">Бийск 
                              <option value="bi{">Бишкек 
                              <option value="bg}">Благовещенск 
                              <option value="bdb">Бодайбо 
                              <option value="brs">Братск 
                              <option value="bug">Бугульма 
                              <option value="bgs">Бургас 
                              <option value="bur">Буревестник 
                              <option value="bua">Бухара 
                              <option value="prd">Бухта Провидения 
                              <option value="wrq">Ванавара 
                              <option value="war">Варна 
                              <option value="weu">Великий Устюг 
                              <option value="whw">Верхневилюйск 
                              <option value="wik">Вилюйск 
                              <option value="wwo">Владивосток 
                              <option value="wla">Владикавказ 
                              <option value="wgg">Волгоград 
                              <option value="wld">Волгодонск 
                              <option value="wgd">Вологда 
                              <option value="wkt">Воркута 
                              <option value="wrn">Воронеж 
                              <option value="wtg">Вытегра 
                              <option value="haj">Ганновер 
                              <option value="gdv">Геленджик 
                              <option value="g`m">Гюмри 
                              <option value="gnv">Гянджа 
                              <option value="dlc">Далянь 
                              <option value="dep">Депутатский 
                              <option value="dvh">Джебарики-Хая 
                              <option value="dpk">Днепропетровск 
                              <option value="dnc">Донецк 
                              <option value="dnk">Дудинка 
                              <option value="d{b">Душанбе 
                              <option value="DUS">Дюссельдорф 
                              <option value="esk">Ейск 
                              <option value="ekb">Екатеринбург 
                              <option value="eg~">Ербогачен 
                              <option value="ewn">Ереван 
                              <option value="vig">Жиганск 
                              <option value="zla">Залив Лаврентия 
                              <option value="zpv">Запорожье 
                              <option value="zon">Зональное 
                              <option value=znk>Зырянка 
                              <option value="igr">Игарка 
                              <option value=irm>Игрим 
                              <option value="ivw">Ижевск 
                              <option value="int">Инта 
                              <option value="ikt">Иркутск 
                              <option value="i{o">Йошкар-Ола 
                              <option value="kzn">Казань 
                              <option value="kz~">Казачинск 
                              <option value="kld">Калининград 
                              <option value="klg">Калуга 
                              <option value="kgd">Караганда 
                              <option value="kdd">Кедровый 
                              <option  value="krw">Кемерово 
                              <option value="kpm">Кепервеем 
                              <option value="iew">Киев 
                              <option value="krn">Киренск 
                              <option value="kio">Киров 
                              <option value="k~g">Кичменгский Городок 
                              <option value="k{n">Кишинев 
                              <option value="kog">Когалым 
                              <option value="kzi">Кодинск 
                              <option value="ksl">Комсомольск-На-Амуре 
                              <option value="ktn">Костанай 
                              <option value="kts">Котлас 
                              <option value="krr">Краснодар 
                              <option value="kkx">Красноселькуп 
                              <option value="kqa">Красноярск 
                              <option value="kgn">Курган 
                              <option value="kus">Курск 
                              <option value="kis">Кутаиси 
                              <option value="kyy">Кызыл 
                              <option value="lsk">Ленск 
                              <option value="le{">Лешуконское 
                              <option value="lip">Липецк 
                              <option value="lod">Лондон 
                              <option value="lwo">Львов 
                              <option value="mdn">Магадан 
                              <option value="mgn">Маган 
                              <option value="mgs">Магнитогорск 
                              <option value="mkp">Майкоп 
                              <option value="mam">Мама 
                              <option value="mrn">Мариинское 
                              <option value="mko">Марково 
                              <option value="mhl">Махачкала 
                              <option 
                          value="mrw">Минеральные Воды 
                              <option 
                          value="msk">Минск 
                              <option value="mir">Мирный 
                              <option 
                          value="mom">Мома 
                              <option class="hl" value="mow" selected>Москва 
                              <option 
                          value="mtg">Мотыгино 
                              <option value="mun">Мурманск 
                              <option 
                          value="mkm">Мыс Каменный 
                              <option value="m{d">Мыс Шмидта 
                              <option value="m`n">Мюнхен 
                              <option 
                          value="ndm">Надым 
                              <option value="in{">Назрань 
                              <option 
                          value="n~k">Нальчик 
                              <option value="nmg">Наманган 
                              <option 
                          value="nnr">Нарьян-Мар 
                              <option 
                          value="nh~">Нахичевань 
                              <option value="nln">Нелькан 
                              <option 
                          value="nrg">Нерюнгри 
                              <option value="n`g">Нефтеюганск 
                              <option 
                          value="nvg">Нижнеангарск 
                              <option 
                          value="nvw">Нижневартовск 
                              <option 
                          value="nvk">Нижнекамск 
                              <option value="nvq">Нижнеянск 
                              <option 
                          value="nvs">Нижний Новгород 
                              <option 
                          value="nlk">Николаевск-На-Амуре 
                              <option 
                          value="nik">Никольское 
                              <option 
                          value="nwk">Новокузнецк 
                              <option 
                          value="owb">Новосибирск 
                              <option value="nur">Новый Уренгой 
                              <option value="ngl">Ноглики 
                              <option 
                          value="nrs">Норильск 
                              <option value="noq">Ноябрьск 
                              <option 
                          value="nus">Нукус 
                              <option value="n`r">Нюрба 
                              <option 
                          value="nqg">Нягань 
                              <option value="ods">Одесса 
                              <option 
                          value="ozr">Озерная 
                              <option value="omk">Оймякон 
                              <option 
                          value="olk">Олекминск 
                              <option value="oln">Оленек 
                              <option 
                          value="ool">Омолон 
                              <option value="oms">Омск 
                              <option 
                          value="ong">Оренбург 
                              <option value="osk">Орск 
                              <option 
                          value="oso">Оссора 
                              <option value="oha">Оха 
                              <option 
                          value="oht">Охотск 
                              <option value="pwl">Павлодар 
                              <option 
                          value="pan">Палана 
                              <option value="pew">Певек 
                              <option 
                          value="prm">Пермь 
                              <option value="ptz">Петрозаводск 
                              <option 
                          value="prl">Петропавловск-Камчатский 
                              <option 
                          value="p~r">Печора 
                              <option value="pin">Пионерный 
                              <option 
                          value="pts">Подкаменная Тунгуска 
                              <option value="pos">Полины Осипенко 
                              <option value="plq">Полярный 
                              <option 
                          value="rad">Радужный 
                              <option value="rix">Рига 
                              <option 
                          value="row">Ростов-На-Дону 
                              <option 
                          value="sky">Саккырыр 
                              <option value="shd">Салехард 
                              <option 
                          value="sm{">Самара 
                              <option value="skd">Самарканд 
                              <option 
                          value="sag">Сангар 
                              <option class=hl 
                          value="spt">Санкт-Петербург 
                              <option 
                          value="srn">Саранск 
                              <option value="sro">Саратов 
                              <option 
                          value="syh">Саскылах 
                              <option 
                          value="sen">Северо-Енисейск 
                              <option 
                          value="swe">Северо-Эвенск 
                              <option value="ski">Сенаки 
                              <option 
                          value="sel">Сеул 
                              <option value="sip">Симферополь 
                              <option 
                          value="sbo">Соболево 
                              <option value="sog">Советская Гавань 
                              <option value="soj">Советский 
                              <option 
                          value="son">Солнечный 
                              <option value="so~">Сочи 
                              <option 
                          value="srm">Средне-Колымск 
                              <option 
                          value="stw">Ставрополь 
                              <option value="ist">Стамбул 
                              <option 
                          value="sol">Старый Оскол 
                              <option 
                          value="spw">Степанаван 
                              <option value="stv">Стрежевой 
                              <option 
                          value="sun">Сунтар 
                              <option value="sur">Сургут 
                              <option 
                          value="syw">Сыктывкар 
                              <option value="tio">Таксимо 
                              <option 
                          value="tas">Ташкент 
                              <option value="tbs">Тбилиси 
                              <option 
                          value="tlq">Тель-Авив 
                              <option value="tig">Тигиль 
                              <option 
                          value="tsi">Тикси 
                              <option value="til">Тиличики 
                              <option 
                          value="txk">Толька 
                              <option value="tsk">Томск 
                              <option 
                          value="tau">Тура 
                              <option value="trh">Туруханск 
                              <option 
                          value="t`m selected">Тюмень 
                              <option value="tsn">Тянь-зинь 
                              <option 
                          value="uln">Улан-Батор 
                              <option value="ul|">Улан-Удэ 
                              <option 
                          value="ulk">Ульяновск 
                              <option value="ura">Урай 
                              <option 
                          value="ugn">Ургенч 
                              <option value="usn">Усинск 
                              <option 
                          value="utk">Усть-Илимск 
                              <option 
                          value="ukg">Усть-Каменогорск 
                              <option 
                          value="uk~">Усть-Камчатск 
                              <option 
                          value="uku">Усть-Куйга 
                              <option value="usk">Усть-Кут 
                              <option 
                          value="usm">Усть-Мая 
                              <option value="unr">Усть-Нера 
                              <option 
                          value="uhj">Усть-Хайрюзово 
                              <option 
                          value="utc">Усть-Цильма 
                              <option value="ufa">Уфа 
                              <option 
                          value="uht">Ухта 
                              <option value="fgn">Фергана 
                              <option 
                          value="fra">Франкфурт-На-Майне 
                              <option 
                          value="hbr">Хабаровск 
                              <option value="hdy">Хандыга 
                              <option 
                          value="has">Ханты-Мансийск 
                              <option 
                          value="hrk">Харьков 
                              <option value="hat">Хатанга 
                              <option 
                          value="hel">Хельсинки 
                              <option value="hrp">Херпучи 
                              <option 
                          value="hdt">Худжанд 
                              <option value="~ar">Чара 
                              <option 
                          value="~be">Чебоксары 
                              <option value="~lb">Челябинск 
                              <option 
                          value="~rw">Череповец 
                              <option value="~rs">Черский 
                              <option 
                          value="sht">Чита 
                              <option value="~kd">Чокурдах 
                              <option 
                          value="~mi">Чумикан 
                              <option value="{ah">Шахтерск 
                              <option 
                          value="{nq">Шеньян 
                              <option value="{mt">Шымкент 
                              <option 
                          value="|gt">Эгвекинот 
                              <option value="|li">Элиста 
                              <option 
                          value="`vk">Южно-Курильск 
                              <option 
                          value="`vh">Южно-Сахалинск 
                              <option 
                          value="qkt">Якутск 
                              <option value="qmb">Ямбург 
                              <option value="qrl">Ярославль</option>
                            </select>
                            &nbsp; </p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2> 
                          <p>Показывать 
                            <input name="NonDirect"
                        type="checkbox">
                            непрямые рейсы 
                            <input name="Reverse"
                        type="checkbox">
                            обратные рейсы</p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2b> 
                          <p>Для даты (по умолчанию - на сегодня)</p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2> 
                          <p> 
                            <select name="Day" size="1" class=text>
                              <option selected 
                          value="0">-- 
                              <option value="1">1 
                              <option value="2">2 
                              <option value="3">3 
                              <option value="4">4 
                              <option value="5">5 
                              <option value="6">6 
                              <option value="7">7 
                              <option value="8">8 
                              <option value="9">9 
                              <option value="10">10 
                              <option value="11">11 
                              <option value="12">12 
                              <option value="13">13 
                              <option value="14">14 
                              <option value="15">15 
                              <option value="16">16 
                              <option value="17">17 
                              <option value="18">18 
                              <option value="19">19 
                              <option value="20">20 
                              <option value="21">21 
                              <option value="22">22 
                              <option value="23">23 
                              <option value="24">24 
                              <option value="25">25 
                              <option value="26">26 
                              <option value="27">27 
                              <option value="28">28 
                              <option value="29">29 
                              <option value="30">30 
                              <option value="31">31</option>
                            </select>
                            <select 
                        name="Month" size="1" class=text>
                              <option selected value="none">* месяц 
                              <option value="1">Января 
                              <option value="2">Февраля 
                              <option value="3">Марта 
                              <option value="4">Апреля 
                              <option value="5">Мая 
                              <option value="6">Июня 
                              <option value="7">Июля 
                              <option value="8">Августа 
                              <option value="9">Сентября 
                              <option value="10">Октября 
                              <option value="11">Ноября 
                              <option value="12">Декабря</option>
                            </select>
                            &nbsp; вылет после 
                            <select name="Time" size="1" class=text>
                              <option selected 
                          value="00">-- 
                              <option value="01">1 
                              <option value="02">2 
                              <option value="03">3 
                              <option value="04">4 
                              <option value="05">5 
                              <option value="06">6 
                              <option value="07">7 
                              <option value="08">8 
                              <option value="09">9 
                              <option value="10">10 
                              <option value="11">11 
                              <option value="12">12 
                              <option value="13">13 
                              <option value="14">14 
                              <option value="15">15 
                              <option value="16">16 
                              <option value="17">17 
                              <option value="18">18 
                              <option value="19">19 
                              <option value="20">20 
                              <option value="21">21 
                              <option value="22">22 
                              <option value="23">23</option>
                            </select>
                            &nbsp;часов </p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2b> 
                          <p>Сортировать по:</p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=text bgcolor="#EBF3F5"> 
                          <p> 
                            <select name="Sort" size="1" class=text>
                              <option 
                          selected value="w">Времени вылета 
                              <option value="r">Номеру рейса 
                              <option 
                        value="s">Стоимости</option>
                            </select>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                            <input type=submit value="Посмотреть" name="submit" size="1" class=text>
                          </p>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left"> 
                          <hr noshade size="3">
                          <h2 align="left"><a href="http://www.polets.ru/" target="_blank">ЗАО 
                            "Полет-Сирена"</a> представляет самое полное расписание 
                            авиарейсов по России и странам СНГ, фактическое выполнение 
                            рейсов в реальном времени, подбор стыковочных рейсов, 
                            действующие тарифы, последние изменения и отмены в 
                            расписании. </h2>
                        </td>
                      </tr>
                      </tbody> 
                    </table>
                  </td>
                </tr>
                </tbody> 
              </table>
            </form>
            <p align="left"> 
              <script language="javascript" src="http://72rus.ru/newsjv.asp?hid=125"></script>
            <p align="center"><font size="-2"><a href="http://72rus.ru"> 72RUS.RU 
              - Тюменский регион</a></font></p>
            <p align="center">&nbsp; </p>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
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
      <p class="menu02">| <font face="Arial, Helvetica, sans-serif"> 
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
        </font><a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <font face="Arial, Helvetica, sans-serif"> 
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
        </font><a href="<%=url%>"><%=hname%></a> | 
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
