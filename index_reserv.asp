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


var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}


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
<title>72RUS Тюмень</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="style.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 500; margin:  3px 3px 3px 4px}
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
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> 
        <%
// В переменной bk содержится код блока новостей
var bk=30
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
        <font color="#000000"><%=blokname%></font></p>
    </td>
    <td align="center" width="200"> 
      <p class="menu01"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <font color="#000000">посетителей на сайте: <%=Application("visitors")%></font></p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170">
      <h1>72rus.ru</h1>
    </td>
    <td background="images/fon02.gif" height="87"><img src="images/e01.gif" width="2" height="87" alt="" border="0"></td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <table border="0" cellspacing="3" cellpadding="0" width="80%" align="center">
        <tr> 
          <form name="form" method="post" action="search.asp">
            <td height="16" valign="middle" align="center"> 
<img src="images/e06.gif" width="16" height="9" alt="" border="0" align="absmiddle"> 
<input type="text" name="sch" size="40" value="" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 150px" >
<input type="hidden" name="wrds" value="1">
<input type="Image" src="images/b_go.gif" width="19" height="19" alt="" border="0" hspace="10" align="absbottom" name="Findit">
</td>
          </form>
        </tr>
      </table>
      </td>
    <td background="images/fon02.gif" height="87" align="center" width="480">&nbsp; 
    </td>
    <td background="images/fon02.gif" height="87" align="right" width="17"><img src="images/e02.gif" width="17" height="87" alt="" border="0"></td>
  </tr>
  <tr bgcolor="#FF6600"> 
    <td colspan="5" height="1"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="1003">
  <tr> 
    <td valign="top" bgcolor="#0065A8" background="images/fon_menu04.gif" width="172"> 
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
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%> [<%=kvopub%>]</a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="messages.asp">Объявления [<%=msgcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="lentamsg.asp">Объявления 
              [новые]</a> </p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auto.72rus.ru">Авто 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auction.72rus.ru">Аукцион 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="air_russia.asp">АВИА 
              Расписание</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="Rail_roads.asp">Расписание 
              поездов</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td><img src="images/top01.gif" width="156" height="19" alt="" border="0"></td>
        </tr>
      </table>
    </td>
    <td width="300" valign="top"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
        <tr> 
          <td width="300" valign="top" bgcolor="#0365AD"> 
            <%
// В переменной bk содержится код блока новостей
var bk=30
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
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
            <h1><font color="#FFFFFF"><%=pname%></font></h1>
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr bgcolor="#003366"> 
          <td align=left width=300 bgcolor="#666666" valign="top" height="1"></td>
        </tr>
        <tr> 
          <td width="377"> 
            <%
// В переменной bk содержится код блока новостей
var bk=30
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
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
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
            <%if (imgLname != "") {%>
            <a href="<%=url%>"><img src="<%=imgLname%>" border="0" alt="<%=pname%>" ></a> 
            <%}%>
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="61">
        <tr> 
          <td background="images/fon02.gif"> 
            <%
// В переменной bk содержится код блока новостей
var bk=30
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
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
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
            <a href="<%=url%>"> 
            <p class="digest"><%=digest%></p>
            </a> 
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" bgcolor="#0A64AA" background="images/fon_menu05.gif" width="172"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" width="170"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
              Каталог [<%=urlcount%>]</a></p>
          </td>
        </tr>
      </table>
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
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%></a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="141" border="0" cellspacing="0" cellpadding="0" bordercolor="#003366" align="center">
        <form name="form1" method="post" action="">
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
            <td align="center" bordercolor="#FFFFFF" height="105"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="ФОБОС: подробная погода в городе" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
          </tr>
        </form>
      </table>
    </td>
    <td background="images/bg.gif" valign="top" width="342"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" width="170"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
              Каталог [<%=urlcount%>]</a></p>
          </td>
          <td background="images/fon_menu07.gif" height="30" width="170" align="center"> 
            <p class="menu01"> Новые сайты</p>
          </td>
          <td height="30" align="center" background="images/fon_menu07.gif">&nbsp;</td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" background="" width="100%">
        <form action="" method="post">
          <tr valign="middle"> 
            <td width="70"> 
              <p><img src="images/e06.gif" width="16" height="9" alt="" border="0" align="absmiddle">&nbsp;&nbsp;<b>Вход</b></p>
            </td>
            <td> 
              <input maxlength=20 name="logname" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/name.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=edit>
              <input maxlength=25 name="psw" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=password>
              <input type="Image" src="images/b_go.gif" width="19" height="25" alt="" border="0" hspace="10" align="absbottom" name="Image2222">
            </td>
          </tr>
        </form>
      </table>
      <p><img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%></p>
    </td>
    <td height="264" background="images/bg_right.gif" width="17" valign="top">&nbsp;</td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td colspan="4" height="1" bgcolor="#666666"></td>
  </tr>
  <tr bgcolor="#AFC0D0"> 
    <td colspan="4" height="19"> 
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></p>
          </td>
        </tr>
      </table>
      <p class="menu02"><img src="images/e05.gif" alt="" width="16" height="9" border="0"> 
        <img src="images/px1.gif" width="1" height="1" alt="" border="0"><a href="http://www.rusintel.ru/newshow.asp?pid=2496">Регистрация 
        доменов RU COM NET</a> <img src="images/e05.gif" alt="" width="16" height="9" border="0"> 
        <a href="http://www.rusintel.ru/">Разработка интернет сайтов</a> <img src="images/e05.gif" alt="" width="16" height="9" border="0"> 
        <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2256">Хостинг</a> 
        <img src="images/e05.gif" alt="" width="16" height="9" border="0"> <a href="http://www.72rus.ru/newshow.asp?pid=728"> 
        Размещение рекламы</a> <img src="images/e05.gif" alt="" width="16" height="9" border="0"> 
        <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2244"> Интернет 
        карты РОЛ</a> </p>
    </td>
  </tr>
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="640"> 
      <table border="0" cellpadding="0" cellspacing="0" background="">
        <form action="" method="post">
          <tr> 
            <td rowspan="2" width="130"> 
              <p><img src="images/e06.gif" width="16" height="9" alt="" border="0">&nbsp;&nbsp;<b>Вход</b></p>
            </td>
            <td> 
              <input type="Text" name="Input22" value=" USERNAME" size="15">
            </td>
          </tr>
          <tr> 
            <td> 
              <input type="Text" name="Input22" value=" PASSWORD" size="10">
              <input type="Image" src="images/b_go.gif" width="19" height="25" alt="" border="0" hspace="10" align="absbottom" name="Image22">
            </td>
          </tr>
        </form>
      </table>
    </td>
    <td background="images/fon02.gif" height="87" width="22"><img src="images/e01.gif" width="2" height="87" alt="" border="0"></td>
    <td background="images/fon02.gif" height="87" align="center" width="327"> 
      <table border="0" cellpadding="0" cellspacing="0" background="">
        <form action="" method="post">
          <tr> 
            <td width="100"> 
              <p><img src="images/e06.gif" width="16" height="9" alt="" border="0">&nbsp;&nbsp;<b>Поиск</b></p>
            </td>
            <td> 
              <input type="Text" name="Input22" value="" size="12">
            </td>
            <td> 
              <input type="Image" src="images/b_go.gif" width="19" height="25" alt="" border="0" hspace="10" align="absbottom" name="Image22">
            </td>
          </tr>
        </form>
      </table>
    </td>
    <td background="images/fon02.gif" height="87" align="right" width="31"><img src="images/e02.gif" width="21" height="87" alt="" border="0"></td>
  </tr>
  <tr> 
    <td colspan="4" height="21" background="images/fon03.gif"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td valign="top" background="images/fon_menu06.gif" width="172"> 
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td><img src="images/top01.gif" width="156" height="19" alt="" border="0"></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td width="172">&nbsp;</td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td><img src="images/top01.gif" width="156" height="19" alt="" border="0"></td>
        </tr>
      </table>
    </td>
    <td> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr valign="top"> 
          <td width="362"> 
            <p class="t01" style="font-size: 12px;"><img src="images/e05.gif" alt="" width="16" height="9" border="0">&nbsp;&nbsp;<b>Welcome 
              To Our Company</b></p>
            <p class="t01">You must keep a text link on here to <a href="http://www.proeffect.com">Proeffect.com</a>. 
              diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam 
              erat volutpat. Ut wisi enim ad minim veniam, quis nostrud exerci 
              tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo 
              consequat. Lorem ipsum dolor sit amet, consectetuer adipiscing elit, 
              sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna 
              aliquam erat volutpat. Ut wisi enim ad minim veniam, quis nostrud 
              exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea 
              commodo consequat. Lorem ipsum dolor sit amet, consectetuer adipiscing 
              elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore 
              magna aliquam erat volutpat. Ut wisi enim ad minim veniam, quis 
              nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip 
              ex ea commodo consequat. Lorem ipsum dolor sit amet, consectetuer 
              adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet 
              dolore magna aliquam erat volutpat. Ut wisi enim ad minim veniam, 
              quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut 
              aliquip ex ea commodo consequat. Lorem ipsum dolor sit amet, consectetuer 
              adipiscing elit, sed diam nonummy nibh euismod eet dolore magna 
              aliquam erat volutpat. Ut wisi enim ad minim veniam, quis If you 
              need <b>web hosting</b> for your visit <a href="http://www.cheap-webhosting-space.com/" target="_blank">Cheap 
              Web Hosting Space</a>. </p>
          </td>
          <td width="15" bgcolor="#AFC0D0"><img src="images/e03.gif" width="15" height="298" alt="" border="0"></td>
          <td bgcolor="#D1D6DB" background="images/fon01.jpg" width="250" style="background-position: top; background-repeat: repeat-x;"> 
            <p style="font-size: 12px;"><img src="images/e06.gif" width="16" height="9" alt="" border="0">&nbsp;&nbsp;<b>Latest 
              news</b></p>
            <p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp;&nbsp;You 
              must keep a text link on here to <a href="http://www.proeffect.com">Proeffect.com</a>. 
              sed diam nonummy nibh dolor sit elit, sed diam nonummy nibh commodo 
              consequat.</p>
            <p class="left"><a href="">Read more</a></p>
            <p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp;&nbsp;If 
              you need web hosting try out <a href="http://www.cheap-webhosting-space.com/" target="_blank">Cheap 
              Web Hosting Space</a> diam nonummy nibh dolor sit elit, sed diam 
              nonummy nibh commodo consequat.</p>
            <p class="left"><a href="">Read more</a></p>
            <p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp;&nbsp;Lorem 
              ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy 
              nibh dolor sit elit, sed diam Layout by <a href="http://www.proeffect.com">Proeffect.com</a>.</p>
            <p class="left"><a href="">Read more</a></p>
          </td>
          <td width="14" bgcolor="#AFC0D0"><img src="images/e04.gif" width="14" height="298" alt="" border="0"></td>
          <td bgcolor="#AFC0D0"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
        </tr>
      </table>
    </td>
    <td valign="bottom" background="images/bg_right.gif" width="17"><img src="images/bg_right.gif" alt="" width="17" height="16" border="0"></td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td height="1" bgcolor="#666666"></td>
  </tr>
  <tr> 
    <td height="19" background="images/fon01.gif"> 
      <p class="menu02">&nbsp;</p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td colspan="2"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
  </tr>
  <tr bgcolor="#EE7B10"> 
    <td height="19" colspan="2"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
  </tr>
  <tr> 
    <td height="70"> 
      <p>Copyright &copy;2003 CompanyName.com</p>
    </td>
    <td> 
      <p class="menu02"> <a href="">Home</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <a href="">About Us</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="">Support</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <a href="">Services</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="">Contacts</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <a href="">Help</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="">FAQ</a> 
      </p>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr> 
    <td bgcolor="#003366" width="430" background="HeadImg/top_lin_end_bg.gif"><font color="#FFFFFF" class="digest"> 
      <%
// В переменной bk содержится код блока новостей
var bk=30
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
      <%=blokname%></font></td>
    <td background="HeadImg/top_lin_bg.gif" width="43"><img src="HeadImg/top_lin_start.gif" width="43" height="19"></td>
    <td align="right" background="HeadImg/top_lin_bg.gif"><font size="1" face="Arial, Helvetica, sans-serif"></font><a href="admarea.asp"><img src="HeadImg/top_lin_end.gif" width="42" height="19" border="0" align="absmiddle"></a></td>
    <td bgcolor="#003366" align="center" background="HeadImg/top_lin_end_bg.gif" width="199"><font size="1" face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Сейчас 
      на сайте посетителей: <%=Application("visitors")%></font></b></font></td>
  </tr>
  <tr> </tr>
</table>
<table border=0 cellpadding=0 cellspacing=0 height=60 width="100%">
  <tbody> 
  <tr> 
    <td align=middle valign=center width=150><img height=60 
      src="HeadImg/logo_72_1.gif" width=150></td>
    <td align=left width=300 valign="middle" bgcolor="#003366" background="HeadImg/logo_72_2.gif">&nbsp; 
    </td>
    <td bgcolor="#003366" background="HeadImg/logo_72_3.gif"> 
      <script language="javascript" src="banshow.asp?rid=1"></script>
    </td>
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
        Тюменская область</font> 
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
      </p>
    </td>
    <td align=left width=302 bgcolor="#FF0000" valign="top"> 
      <table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=2 
width="100%">
        <tbody> 
        <tr> 
          <td align=right height="2"> </td>
        </tr>
        </tbody> 
      </table>
      <%
// В переменной bk содержится код блока новостей
var bk=30
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
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
      <h1><font color="#FFFFFF"><%=pname%></font></h1>
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
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
                <input name="Enter" style="BACKGROUND-COLOR: #C6DDFF; COLOR: #003366; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 50px; HEIGHT: 20px" type=submit value=Вход>
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
<table border=0 cellpadding=0 cellspacing=0 height=200 width="100%">
  <tbody> 
  <tr> 
    <td align=middle valign=top width=150 bgcolor="#003366"> 
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
        <tr bordercolor="#EBF3F5"> 
          <td bgcolor="#EBF3F5"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> <a href="<%=url%>"><%=hname%></a></b></font></p>
          </td>
          <td width="32" bordercolor="#EBF3F5" bgcolor="#EBF3F5"><font face="Verdana" size="1"><%=kvopub%></font></td>
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
          <td> 
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
          <td> 
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
    </td>
    <td align=left width=300 bgcolor="#000000" valign="middle" background="HeadImg/bg.gif"> 
      <%
// В переменной bk содержится код блока новостей
var bk=30
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
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
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
      <%if (imgLname != "") {%>
      <a href="<%=url%>"><img src="<%=imgLname%>" border="0" alt="<%=pname%>" width="300" ></a> 
      <%}%>
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </td>
    <td valign="middle" align="center" bgcolor="#003366" width="1"></td>
    <td valign="middle"> 
      <table border="0" cellspacing="3" cellpadding="0" width="80%" align="center">
        <tr> 
          <form name="form" method="post" action="search.asp">
            <td height="16" valign="middle" align="center"> 
              <p> 
                <input type="text" name="sch" size="40" value="" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 150px" >
                <input type="hidden" name="wrds" value="1">
                <input type="submit" name="Findit" style="BACKGROUND-IMAGE: url(headimg/lence.gif); COLOR: #003366; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 55px; HEIGHT: 20px" value="     Найти">
              </p>
            </td>
          </form>
        </tr>
      </table>
      <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td bgcolor="#FFFFFF"> 
            <%
// В переменной bk содержится код блока новостей
var bk=31
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
            <%
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=PubImgPath+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=PubImgPath+pid+".jpg"
		if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
	}
	path=""
	hid=String(Records("HEADING_ID").Value)
	hdd=hid
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\"><font color=\"#FFFFFF\">"+nm+"</font></a> ] "+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}
%>
            <hr size=3 noshade width="90%" align="left">
            <p class="digest"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <a href="<%=url%>"><%=pname%></a> <%=digest%></p>
            <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
    </td>
    <td valign="middle" width="120"> 
      <%if (tpm<3) {%>
      <p><a href="addnewsheading.asp?hid=0"><font size="1">Добавить раздел<br>
        </font></a><font size="1"><a href="pubarea.asp">Кабинет редактора</a></font> 
        <%}%>
      </p>
    </td>
  </tr>
  </tbody> 
</table>
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right bgcolor="#CCCCCC" width="450"> </td>
    <td align=right bgcolor="#75B1B7"></td>
  </tr>
  </tbody> 
</table>
<table border=0 cellpadding=0 cellspacing=0 height=60 width="100%">
  <tbody> 
  <tr> 
    <td align=left valign=center width=150 bgcolor="#C6DDFF" height="62"> 
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
              <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></b></font></p>
          </td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td align=left width=300 height="62" valign="middle" bgcolor="#003366"> 
      <%
// В переменной bk содержится код блока новостей
var bk=30
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
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
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
      <a href="<%=url%>"> 
      <p class="digest"> <font color="#FFFFFF"><%=digest%></font></p>
      </a> 
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </td>
    <td bgcolor="#000000" height="62" valign="middle" width="1"></td>
    <td bgcolor="#C6DDFF" height="62" valign="middle"> 
      <p class="digest">В <a href="catarea.asp"><b>Каталоге 72RUS</b></a> новый 
        сайт: <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%></p>
    </td>
  </tr>
  </tbody> 
</table>
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right> </td>
  </tr>
  </tbody> 
</table>
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=8 
width="100%">
  <tbody> 
  <tr> 
    <td align=right width=150 nowrap> <img align=absMiddle border=0 height=23 src="headimg/round_inv.gif" 
      width=23></td>
    <td bgcolor=#FFFFFF valign=middle nowrap> 
      <p class="digest"><img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
        <a href="http://www.rusintel.ru/newshow.asp?pid=2496">Регистрация имени 
        RU COM NET</a> <img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
        <a href="http://www.rusintel.ru/">Разработка интернет сайтов</a> <img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
        <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2256">Хостинг</a> 
        <img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
        <a href="http://www.72rus.ru/newshow.asp?pid=728">Размещение рекламы</a> 
        <img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
        <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2244">Интернет 
        РОЛ</a></p>
    </td>
  </tr>
  </tbody> 
</table>
<table border=0 cellpadding=0 cellspacing=0 height=8 
width="100%">
  <tbody> 
  <tr> 
    <td align=right nowrap width="150" background="HeadImg/top_lin_bg.gif">&nbsp; </td>
    <td valign=top nowrap bgcolor="#003366" background="HeadImg/top_lin_end_bg.gif"><img src="HeadImg/top_lin_end.gif" width="42" height="19" align="absmiddle"><font class="digest"><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"> 
      <a href="russia.asp"><font color="#FFFFFF">Российские и мировые новости</font></a></font></td>
  </tr>
  </tbody> 
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="left"> 
    <td width="150" align="left" bgcolor="#EBF3F5"> <font color="#FFFFFF" class="digest"> 
      <font color="#000000"> </font></font> 
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
        <tr> 
          <td bgcolor="#003366" nowrap align="center"><font color="#C6DDFF"><font class="digest"> 
            <%
// В переменной bk содержится код блока новостей
var bk=37
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
      <table border=1 bordercolor=#003366 cellpadding=0 
      cellspacing=0 width=150>
        <tbody> 
        <tr bordercolor="#EBF3F5" bgcolor="#EBF3F5"> 
          <td bgcolor="#C6DDFF" bordercolor="#C6DDFF"> 
            <p><font><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"></font><font face=Verdana size=1><b> 
              <a 
            href="air_russia.asp">АВИА Расписание</a> </b></font></p>
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
              <a href="Rail_roads.asp">Расписание поездов</a></b></font><font face=Verdana size=1><b> 
              </b></font><font face=Verdana size=1><b> </b></font></p>
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
              <a href="http://www.72rus.ru/pubheading.asp?hid=151">Cотовые телефоны</a></b></font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table width="150" border="0" cellspacing="0" align="center" cellpadding="0">
        <%
// В переменной bk содержится код блока новостей
var bk=37
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)


var news=""
filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+=ts.ReadLine()
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}


%>
        <tr> 
          <td align="CENTER" width="150"><%=news%></td>
          <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
        </tr>
      </table>
      <font color="#000000"> </font> 
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
        <tr> 
          <td bgcolor="#003366" align="center"><font color="#C6DDFF" class="digest"> 
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
            <%=blokname%> </font> </td>
        </tr>
        <tr> </tr>
      </table>
      <font color="#000000"> 
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
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
        <tr> 
          <td bgcolor="#003366" align="center"><font color="#C6DDFF" class="digest"> 
            <%
// В переменной bk содержится код блока новостей
var bk=34
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
      <%
// В переменной bk содержится код блока новостей
var bk=34
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
      <table width="100%" border="0" cellspacing="0" align="center">
        <tr> 
          <td><%=news%></td>
        </tr>
      </table>
      <p> 
        <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </td>
    <td bgcolor="#003366" width="1" ></td>
    <td > 
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
        <tr bgcolor="#C6DDFF"> 
          <td nowrap bgcolor="#C6DDFF"> 
            <p class="digest"> &nbsp;&nbsp;&nbsp; 
              <%
// В переменной bk содержится код блока новостей
var bk=28
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
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bordercolor="#CCCCCC"> 
          <td bordercolor="#CCCCCC"> 
            <dl> 
              <%
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=PubImgPath+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=PubImgPath+pid+".jpg"
		if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
	}
	path=""
	hid=String(Records("HEADING_ID").Value)
	hdd=hid
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\"><font color=\"#FF0000\">"+nm+"</font></a>"+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}
%>
              <dd> 
                <li> 
                  <p> 
                    <%if (imgname != "") {%>
                  <table width="0" border="1" cellspacing="5" cellpadding="0" bordercolor="#FFFFFF" align="left">
                    <tr> 
                      <td width="150" bordercolor="#003366"><a href="<%=url%>?pid=<%=pid%>"><img src="<%=imgname%>" alt="<%=pname%>" border="0" width="150" height="95" ></a></td>
                      <td bordercolor="#003366" width="5" bgcolor="#003366">&nbsp;</td>
                    </tr>
                  </table>
                  <%}else{%>
                  <%}%>
                  <a href="<%=url%>?pid=<%=pid%>"><font size="3"><%=pname%></font></a><br>
                  <font size="2" >[<%=pdat%>] <%=digest%></font><BR clear=all>
                </LI>
                <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
            </dl>
          </td>
        </tr>
      </table>
    </td>
    <td width="1" bgcolor="#003366"></td>
    <td width="150" bgcolor="#EBF3F5" align="center"> 
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
        <tr> 
          <td bgcolor="#C6DDFF" nowrap align="center"> 
            <p class="digest"> 
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
              <%=blokname%> </p>
          </td>
        </tr>
        <tr> </tr>
      </table>
      <p align="CENTER"> 
        <script language="javascript" src="banshow.asp?rid=3"></script>
      </p>
    </td>
  </tr>
</table>
<table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" height="1">
  <tr> 
    <td bgcolor="#003366"></td>
  </tr>
  <tr> </tr>
</table>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=2"></script>
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
    © 2002-2004 <a href="http://www.rusintel.ru">ЗАО Русинтел</a> 
</div>
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
  <!--RAX logo-->
  <a href="http://www.rax.ru/click" target=_blank><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="rax.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодня"></a> 
  <!--/RAX-->
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
