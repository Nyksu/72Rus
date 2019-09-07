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

if (String(Session("urlstat"))=="undefined") {Session("urlstat")="index.asp"}

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
<title>72RUS.RU Тюменский Регион - Тюмень, информационный портал. Объявления, 
Каталог сайтов, Новости, Расписание транспорта</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 500; margin:  3px 3px 3px 4px}
h1 {color: #333333; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
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
<table border="1" cellpadding="0" cellspacing="5" width="1004" height="75" bordercolor="#FFFFFF" align="center">
  <tr> 
    <td width="164" bgcolor="#F3F3F3" bordercolor="#E5E5E5"> 
      <p class="digest"> <a href="admarea.asp"><img src="images/e06.gif" width="16" height="9" alt="" border="0"></a> <a href="#" class="digest">Посетителей 
        сейчас:</a> <%=Application("visitors")%><br>
        <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="http://www.72rus.ru/newshow.asp?pid=728" class="digest">Реклама 
        на 72RUS.RU</a><br>
        <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="#" onClick="window.external.AddFavorite(parent.location,document.title)" class="digest">Добавить 
        в избранное</a></p>
</td>
    <td align="center" bordercolor="#FFFFFF"> 
      <script language="javascript" src="http://www.72rus.ru/banshow.asp?rid=1"></script>
    </td>
    <td width="175" bordercolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center"> 
            <p class="menu01"><a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2244" target="_blank"><img src="HeadImg/rol20.gif" width="120" height="72" align="middle" border="1" alt="Интернет карты РОЛ в Тюмени"></a> 
            </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="1004" align="center">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"><img src="images/72rus.gif" width="170" height="87" alt="72RUS.RU Тюменский Регион - информационный портал "></td>
    <td background="images/fon02.gif" height="87" align="center" width="661"> 
      <table border="0" cellspacing="0" cellpadding="0" align="center" width="468">
        <form name="form" method="get" action="searchall.asp">
          <tr> 
            <td align="center"> 
              <input type="text" name="sch" size="40" value="<%=sch%>" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 275px" >
              <input type="image" src="images/search.gif" width="55" height="20" alt="Найти на сайте" border="0" hspace="10" align="absbottom" name="Findit">
              <select name="wrds" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 100px">
                <option value="1" <%=wrds==1?"selected":""%>>Все слова</option>
                <option value="2" <%=wrds==2?"selected":""%>>Одно из слов</option>
                <option value="3" <%=wrds==3?"selected":""%>>Фраза целиком</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td align="center"> 
              <p class="digest">Везде | <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=2" class="digest">В 
                объявлениях</a> | <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=3" class="digest">В 
                каталоге сайтов</a> | <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=1" class="digest">В 
                публикациях</a> | 
                <input type="checkbox" name="sensation" value="1" <%=sens==1?"checked":""%>>
                Учесть регистр</p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"></td>
          </tr>
        </form>
      </table>
    </td>
    <td background="images/fon02.gif" height="87" width="2" align="center"><img src="images/e01.gif" width="2" height="87" alt="" border="0"></td>
    <td background="images/fon02.gif" height="87" width="170" align="center"><table border="0" cellpadding="0" cellspacing="0" background="" width="100%">
        <form name="form" method="post" action="usrarea.asp">
          <tr valign="middle"> 
            <td width="48" align="center" valign="top"> 
              <p class="digest">Вход</p>
            </td>
            <td> 
              <input maxlength=20 name="logname" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/name.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=edit>
              <input maxlength=25 name="psw" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=password>
              <input type="Image" src="images/b_go.gif" width="19" height="25" alt="" border="0" hspace="10" align="absbottom" name="Enter">
            </td>
          </tr>
        </form>
      </table>
      <p><a href="regmemurl.asp" class="digest"> Регистрация</a></p> </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="1004" align="center">
  <tr> 
  <tr> 
    <td background="images/fon_menu01.gif" height="30" width="166" bgcolor="#3399CC"> 
      <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
        Каталог [<%=urlcount%>]</a></p>
    </td>
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
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	Records.MoveNext()
%>
    <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC" width="172"> 
      <p class="menu01"><img src="images/in_menu01.gif" width="5" height="30" align="absmiddle"><img src="images/arrow.gif" align="absmiddle"> 
        <a href="<%=url%>"><%=hname%> [<%=kvopub%>]</a></p>
    </td>
    <%
} Records.Close()
delete recs
%>  
    <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
      <p class="menu01"><img src="images/in_menu01.gif" width="5" height="30" align="absmiddle"><img src="images/arrow.gif" align="absmiddle"> 
        <a href="messages.asp">Объявления [<%=msgcount%>]</a> <a href="lentamsg.asp">[новые]</a></p>
    </td>
    <td background="images/fon_menu01.gif" align="right" width="172"> 
      <table border="0" cellpadding="0" cellspacing="0" width="172">
        <tr> 
          <td background="images/fon_menu07.gif" height="30" align="center" bgcolor="#FF9900"> 
            <p class="menu01"> <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              Погода</p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="1004" align="center">
  <tr> 
    <td valign="top" width="172" bgcolor="#F3F3F3"> 
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
          <td height="30" background="images/fon_menu01.gif" bgcolor="#3399CC"> 
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
          <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auto.72rus.ru">Авто 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auction.72rus.ru">Аукцион 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="air_russia.asp">АВИА 
              Расписание</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="Rail_roads.asp">Расписание 
              поездов</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td valign="top" width="170"><img src="images/top01.gif" width="170" height="30" alt="" border="0"> 
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" align="center" bgcolor="#F3F3F3" width="344"> 
      <hr size=3 noshade align="center" width="96%">
      <table border="1" cellpadding="0" cellspacing="6" width="100%" align="center" bordercolor="#F3F3F3">
        <tr> 
          <td valign="top" align="center" bgcolor="#FFFFFF" bordercolor="#E5E5E5"> 
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
            <h1><%=pname%></h1>
            <%if (imgLname != "") {%>
            <a href="<%=url%>"><img src="<%=imgLname%>" border="0" alt="<%=pname%>" ></a> 
            <%}%>
            <p class="digest" align="left"><a href="<%=url%>"><%=digest%></a></p>
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
    <td valign="top" bgcolor="#F3F3F3"> 
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
      <p class="digest"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="<%=url%>" target="_blank" class="digest"><%=pname%></a> <%=digest%></p>
      <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </td>
    <td valign="top" width="1" align="center" bgcolor="#F3F3F3" background="images/e01.gif">&nbsp;</td>
    <td valign="top" width="171" align="center" bgcolor="#F3F3F3"> 
      <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="172"> 
            <hr size=3 noshade width="96%" align="center">
            <table width="170" border="0" cellspacing="0" cellpadding="0" bordercolor="#003366" align="center">
              <form name="form1" method="post" action="">
                <tr> 
                  <td align="center" height="30" valign="middle"> 
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
                  <td align="center" height="110"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="ФОБОС: подробная погода в городе" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
                </tr>
              </form>
            </table>
            <hr size=3 noshade width="96%" align="center">
          </td>
        </tr>
      </table>
      <p class="digest><img src="images/e06.gif" width="16" height="9" alt="" border="0" align="left"> 
        <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <img src="images/px1.gif" width="1" height="1" alt="" border="0"><a href="http://www.rusintel.ru/newshow.asp?pid=2496">Регистрация 
        доменов</a><br>
        <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="http://www.rusintel.ru/">Разработка 
        сайтов</a><br>
        <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2256">Хостинг</a></p>
      <hr size=3 noshade width="96%" align="center">
      <a href="http://www.informer.ru/cgi-bin/redirect.cgi?id=177_1_1_33_40_1-0&url=http://www.rbc.ru&src_url=usd/eur_cb_forex_000066_88x90.gif"
target="_blank"><img src="http://pics.rbc.ru/img/grinf/usd/eur_cb_forex_000066_88x90.gif " width=88 height="90" border=0></a> 
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"><script language="JavaScript"> var loc = ''; </script>
<script language="JavaScript1.4">try{ var loc = escape(top.location.href); }catch(e){;}</script>
<script language="JavaScript">
var userid = 48134031; var page = 1;
var rndnum = Math.round(Math.random() * 999111);
document.write('<iframe src="http://ad9.bannerbank.ru/bb.cgi?cmd=ad&hreftarget=_blank&pubid=' + userid + '&pg=' + page + '&vbn=1358&w=120&h=60&num=1&r=ssi&ssi=nofillers&r=ssi&nocache=' + rndnum + '&ref=' + escape(document.referrer) + '&loc=' + loc + '" frameborder=0 vspace=0 hspace=0 width=120 height=60 marginwidth=0 marginheight=0 scrolling=no>');
document.write('<a href="http://ad9.bannerbank.ru/bb.cgi?cmd=go&pubid=' + userid + '&pg=' + page + '&vbn=1358&num=1&w=120&h=60&nocache=' + rndnum + '&loc=' + loc + '&ref=' + escape(document.referrer) + '" target="_blank">');
document.write('<img src="http://ad9.bannerbank.ru/bb.cgi?cmd=ad&pubid=' + userid + '&pg=' + page + '&vbn=1358&num=1&w=120&h=60&nocache=' + rndnum + '&ref=' + escape(document.referrer) + '&loc=' + loc + '" width=120 height=60 Alt="72RUS Banner Network" border=0></a></iframe>');
</script></td>
    <td background="images/fon02.gif" height="87" width="2"><img src="images/e01.gif" width="2" height="87" alt="" border="0"></td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="http://www.72rus.ru/banshow.asp?rid=16"></script>
    </td>
    <td background="images/fon02.gif" height="87" align="center" width="2"><img src="images/e01.gif" width="2" height="87" alt="" border="0"></td>
    <td background="images/fon02.gif" height="87" align="center" width="170"><script language="JavaScript"> var loc = ''; </script>
      <script language="JavaScript1.4">try{ var loc = escape(top.location.href); }catch(e){;}</script>
      <script language="JavaScript">
var userid = 48134031; var page = 1;
var rndnum = Math.round(Math.random() * 999111);
document.write('<iframe src="http://ad9.bannerbank.ru/bb.cgi?cmd=ad&hreftarget=_blank&pubid=' + userid + '&pg=' + page + '&vbn=1358&w=120&h=60&num=2&r=ssi&ssi=nofillers&r=ssi&nocache=' + rndnum + '&ref=' + escape(document.referrer) + '&loc=' + loc + '" frameborder=0 vspace=0 hspace=0 width=120 height=60 marginwidth=0 marginheight=0 scrolling=no>');
document.write('<a href="http://ad9.bannerbank.ru/bb.cgi?cmd=go&pubid=' + userid + '&pg=' + page + '&vbn=1358&num=2&w=120&h=60&nocache=' + rndnum + '&loc=' + loc + '&ref=' + escape(document.referrer) + '" target="_blank">');
document.write('<img src="http://ad9.bannerbank.ru/bb.cgi?cmd=ad&pubid=' + userid + '&pg=' + page + '&vbn=1358&num=2&w=120&h=60&nocache=' + rndnum + '&ref=' + escape(document.referrer) + '&loc=' + loc + '" width=120 height=60 Alt="72RUS Banner Network" border=0></a></iframe>');
</script></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td valign="top" width="168"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
</table>
      <table border="0" cellspacing="0" width="100%" cellpadding="0" height="30">
        <tr> 
          <td bgcolor="#3399CC" align="center" background="images/fon_menu01.gif"> 
            <p class="menu01"> 
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
              <%=blokname%></p>
          </td>
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
      </font> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td> 
            <p class="digest>
              <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="newshow.asp?pid=<%=pid%>">
              <%=pname%> </a></p>
          </td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
%>
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="167" cellpadding="0">
        <tr> 
          <td bgcolor="#006699" align="center" height="30" background="images/fon_menu01.gif"> 
            <p class="menu01"> 
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
              <%=blokname%></p>
          </td>
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
		news+=ts.ReadLine()
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}


%>
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td bgcolor="#FFFFFF"><%=news%></td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="168">
        <tr> 
          <td valign="top"><img src="images/top01.gif" width="167" height="19" alt="" border="0"> 
          </td>
        </tr>
      </table>
    </td>
    <td valign="top"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr valign="top"> 
          <td width="3" bgcolor="#AFC0D0" background="images/e01.gif">&nbsp;</td>
          <td width="468"> 
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0" height="30">
              <tr> 
                <td bgcolor="#3399CC" align="center" background="images/fon_menu01.gif"> 
                  <p class="menu01"><b> 
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
                    <%=blokname%></b></p>
                </td>
              </tr>
              <tr> </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bordercolor="#CCCCCC"> 
                <td bordercolor="#CCCCCC" bgcolor="#FFFFFF"> 
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
                            <td width="100" bordercolor="#003366"><a href="<%=url%>?pid=<%=pid%>"><img src="<%=imgname%>" alt="<%=pname%>" border="0" width="100" height="70" ></a></td>
                            <td bordercolor="#003366" width="5" bgcolor="#FF6600">&nbsp;</td>
                          </tr>
                        </table>
                        <%}else{%>
                        <%}%>
                        <a href="<%=url%>?pid=<%=pid%>"><%=pname%></a><br>
                        [<%=pdat%>] <%=digest%><BR clear=all>
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
          <td width="3" bgcolor="#AFC0D0" background="images/e01.gif">&nbsp;</td>
          <td style="background-position: top; background-repeat: repeat-x;" bgcolor="#FFFFFF"> 
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0" height="30">
              <tr> 
                <td bgcolor="#3399CC" align="center" background="images/fon_menu01.gif"> 
                  <p class="menu01">Новые сайты в каталоге</p>
                </td>
              </tr>
              <tr> </tr>
            </table>
            <%
sql="Select t1.* from url t1, catarea t2 where t1.recl_id>0 and t1.recl_id<5  and t1.catarea_id=t2.id and t2.catalog_id="+catalog+" order by t1.recl_id"

Records.Source=sql
Records.Open()
while (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
%>
            <p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp; 
              <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%> 
              <%
Records.MoveNext()
}
Records.Close()
%>
            </p>
            <p class="left" align="right"> <a href="addurl.asp">Добавить сайт</a></p>
            <hr noshade size="1" width="90%">
            <%
// В переменной bk содержится код блока новостей
var bk=38
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

	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hdd=recs("HI_ID").Value
	recs.Close()
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
}

%>
            <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
              <tbody> 
              <tr> 
                <td align="center" ><%=news%></td>
              </tr>
              </tbody> 
            </table>
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
          <td width="3" bgcolor="#AFC0D0" background="images/e01.gif">&nbsp;</td>
          <td width="120" align="center"> 
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0" height="30">
              <tr> 
                <td bgcolor="#3399CC" align="center" background="images/fon_menu01.gif"> 
                  <p class="menu01"><b> 
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
                    <%=blokname%></b></p>
                </td>
              </tr>
              <tr> </tr>
            </table>
            <script language="javascript" src="banshow.asp?rid=3"></script>
          </td>
        </tr>
      </table>
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
  <script language="javascript" src="banshow.asp?rid=2"></script>
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
  <a href="http://www.liveinternet.ru/stat/72rus.ru/" target=_blank><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="liveinternet.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодня"></a> 
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
  
  <table border="0" cellspacing="0" cellpadding="0" width="1004" align="center">
    <tr> 
      <td valign="top" width="172" bgcolor="#F3F3F3"> 
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
            <td height="30" background="images/fon_menu01.gif" bgcolor="#3399CC"> 
              <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> 
                <a href="<%=url%>"><%=hname%></a></p>
            </td>
          </tr>
        </table>
        <%
} Records.Close()
delete recs
%>
        <table border="0" cellpadding="0" cellspacing="0" width="170">
          <tr> 
            <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
              <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> 
                <a href="http://auto.72rus.ru">Авто 72rus</a></p>
            </td>
          </tr>
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="170">
          <tr> 
            <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
              <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> 
                <a href="http://auction.72rus.ru">Аукцион 72rus</a></p>
            </td>
          </tr>
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="170">
          <tr> 
            <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
              <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> 
                <a href="air_russia.asp">АВИА Расписание</a></p>
            </td>
          </tr>
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="170">
          <tr> 
            <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
              <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> 
                <a href="Rail_roads.asp">Расписание поездов</a></p>
            </td>
          </tr>
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="170">
          <tr> 
            <td background="images/fon_menu01.gif" height="30" bgcolor="#3399CC"> 
              <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> 
                <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></p>
            </td>
          </tr>
        </table>
        <table border="0" cellpadding="0" cellspacing="0" width="156">
          <tr> 
            <td valign="top" width="170"><img src="images/top01.gif" width="170" height="30" alt="" border="0"> 
            </td>
          </tr>
        </table>
      </td>
      <td valign="top" align="center" bgcolor="#F3F3F3" width="344"> 
        <hr size=3 noshade align="center" width="96%">
        <table border="1" cellpadding="0" cellspacing="6" width="100%" align="center" bordercolor="#F3F3F3">
          <tr> 
            <td valign="top" align="center" bgcolor="#FFFFFF" bordercolor="#E5E5E5"> 
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
              <h1><%=pname%></h1>
              <%if (imgLname != "") {%>
              <a href="<%=url%>"><img src="<%=imgLname%>" border="0" alt="<%=pname%>" ></a> 
              <%}%>
              <p class="digest" align="left"><a href="<%=url%>"><%=digest%></a></p>
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
      <td valign="top" bgcolor="#F3F3F3"> 
        <hr size=3 noshade width="98%" align="center">
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=2">Автотранспорт</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=127">Водный воздушный</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=653">Знакомства</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=5">Компьютеры и оргтехника</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=6">Мебель</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=1">Недвижимость</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=372">Нефтегазовый комплекс</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=7">Одежда</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=108">Офисное оборудование</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=43">Продовольствие</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=40">Промоборудование</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=152">Работа</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=3">Стройматериалы</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=86">Услуги</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=4">Электро-бытовая техника</a></p>
        <p class="class"><img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp; 
          <a href="messages.asp?subj=55">Прочее, но полезное</a></p>
      </td>
      <td valign="top" width="1" align="center" bgcolor="#F3F3F3" background="images/e01.gif">&nbsp;</td>
      <td valign="top" width="171" align="center" bgcolor="#F3F3F3"> 
        <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr> 
            <td width="172"> 
              <hr size=3 noshade width="96%" align="center">
              <table width="170" border="0" cellspacing="0" cellpadding="0" bordercolor="#003366" align="center">
                <form name="form1" method="post" action="">
                  <tr> 
                    <td align="center" height="30" valign="middle"> 
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
                    <td align="center" height="110"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="ФОБОС: подробная погода в городе" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
                  </tr>
                </form>
              </table>
              <hr size=3 noshade width="96%" align="center">
            </td>
          </tr>
        </table>
        <p class="digest><img src="images/e06.gif" width="16" height="9" alt="" border="0" align="left"> 
          <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <img src="images/px1.gif" width="1" height="1" alt="" border="0"><a href="http://www.rusintel.ru/newshow.asp?pid=2496">Регистрация 
          доменов</a><br>
          <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="http://www.rusintel.ru/">Разработка 
          сайтов</a><br>
          <img src="images/e06.gif" width="16" height="9" alt="" border="0"> <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2256">Хостинг</a></p>
        <hr size=3 noshade width="96%" align="center">
        <a href="http://www.informer.ru/cgi-bin/redirect.cgi?id=177_1_1_33_40_1-0&url=http://www.rbc.ru&src_url=usd/eur_cb_forex_000066_88x90.gif"
target="_blank"><img src="http://pics.rbc.ru/img/grinf/usd/eur_cb_forex_000066_88x90.gif " width=88 height="90" border=0></a> 
      </td>
    </tr>
  </table>

</div>
</body>
</html>
