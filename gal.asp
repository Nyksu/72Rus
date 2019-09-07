<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\url.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {Response.Redirect("index.asp")}

var hid=0
var sminame=""
var tit=""
var hdd=0
var nm=""
var hiname=""
var period=0
var path=""
var hadr=""
var pname=""
var pdat=""
var autor=""
var news=""
var news_bl=""
var imgname=""
var imgLname=""
var imgRname=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var sos=""
var ishtml=0
var ishtml2=0
var isnews=1
var tpm=1000
var usok=false
var blokname=""
var puid=0

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
var hidimg=""
if (hidimg != "") {
	if (!fs.FileExists(HeadImgPathR+hidimg)) { hidimg="" }
	else {hidimg=HeadImgPath+hidimg}
}

Records.Source="Select t1.* from publication t1, heading t2 where t1.state=1 and t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
hid=Records("heading_id").Value
pname=String(Records("NAME").Value)
pdat=Records("PUBLIC_DATE").Value
autor=TextFormData(Records("AUTOR").Value,"")
ishtml=TextFormData(Records("ISHTML").Value,"0")
Records.Close()

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

tit=sminame

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
		isnews=Records("ISNEWS").Value
		hidimg=TextFormData(Records("PICTURE").Value,"")
	}
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> > "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

var ddt = new Date()
var dt=""
var str=""
var sumdat=Server.CreateObject("datesum.DateSummer")

str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
dt="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
dt=str+dt
dt=sumdat.SummToDate(dt,"-"+period)
if (sumdat.DateComparing(dt,pdat) > 0) {sos="Архив"} else {sos=""}

path="<a href=\"index.asp\">"+sminame+"</a> > "+path
tit+=" | "+hiname

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

imgname=PubImgPath+pid+".gif"
if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
if (imgname=="") {
	imgname=PubImgPath+pid+".jpg"
	if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
}

imgLname=PubImgPath+"l"+pid+".gif"
if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
if (imgLname=="") {
	imgLname=PubImgPath+"l"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
}

imgRname=PubImgPath+"r"+pid+".gif"
if (!fs.FileExists(PubFilePath+"r"+pid+".gif")) { imgRname="" }
if (imgRname=="") {
	imgRname=PubImgPath+"r"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"r"+pid+".jpg")) { imgRname="" }
}

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
%>

<html>
<head>
<title><%=pname%> | <%=tit%> Тюмень, Тюменская область</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72RUS.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
.bar { color: #FFFFFF}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFFFFF; background-color: #FFFFFF}.menu01 {
	color : #FFFFFF;
	margin-top : 0px;
	padding-bottom : 0px;
	margin-bottom : 0px;
	margin-left : 5px;
	margin-right : 0px;
	font-size : 12px;
	font-family : Tahoma, Verdana, Arial;
	font-weight: bold;
}
.menu01 a {
	color : #FFFFFF;
	text-decoration: none;
}
.menu01 a:hover {
	color : #EEEEEE;}.menu01 a:visited { color : #FFFFFF; }
--></style>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        <%=sminame%> : <%=nm%> Тюмень, Тюменская область 
        <%if (tpm<4) {%>
        <a href="headimglst.asp?hid=<%=hid%>"><font color="#000000">Добавить Имидж 
        раздела</font></a> 
        <%}%>
        </font> </p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"> 
      <%if (hidimg != "") {%>
      <a href="/"><img src="<%=hidimg%>" border="0" width="170" height="87" alt="На главную страницу"></a> 
      <%} else {%>
      <a href="/"><img src="images/72rus.gif" width="170" height="87" alt="72RUS.RU Тюменский Регион - информационный портал " border="0"></a> 
      <%}%>
    </td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="banshow.asp?rid=4"></script>
    </td>
    <td background="images/fon02.gif" height="87" align="center" width="170"><table border="0" cellspacing="0" cellpadding="0">
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
				news_bl+=ts.ReadLine()
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
<table border="0" cellspacing="0" cellpadding="0" align="center" height="900">
  <tr valign="top"> 
    <td height="369"> 
      <table width="50%" border="0" cellspacing="0" cellpadding="0" height="6">
        <tr> 
          <td></td>
        </tr>
      </table>
      <table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td valign="middle" height="30"> 
            <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <%=path%> 
              <%if (isnews==1) {%>
              / <a href="archive.asp?hid=<%=hid%>">Архив рубрики</a> 
              <%}%>
            </p>
          </td>
        </tr>
      </table>
      <table width="97%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
        <tr> 
          <td bgcolor="#FF9900" align="left" background="images/fon_menu08.gif"> 
            <h1><b><font color="#FFFFFF"><%=nm%></font></b>&nbsp;</h1>
          </td>
        </tr>
      </table>
      <table width="97%" border="1" cellspacing="0" cellpadding="0" bordercolor="#F4F7FF" align="center">
        <tr bordercolor="#EBF3F5"> 
          <td bordercolor="#EBF3F5" bgcolor="#F7F7F7"> 
            <p class="menu02"> 
              <%
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
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
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
              <img src="images/arrow_next.gif" width="14" height="13" align="absmiddle">&nbsp;&nbsp; 
              <a  href="<%=url%>"><%=hname%></a> <br>
              <%
} 
Records.Close()
delete recs
%>
            </p>
          </td>
        </tr>
      </table>
      <p> 
        <% if (tpm<4) {%>
      </p>
      <p align="center">| <a href="addpubimg.asp?pid=<%=pid%>">Добавить фотографию</a> 
        | 
        <%}%>
        <%if (usok) {%>
        <a href="pubresume.asp?pid=<%=pid%>&st=0">Остановить</a> | <a href="delpub.asp?pid=<%=pid%>">Удалить</a> 
        | <a href="edpub.asp?pid=<%=pid%>"> </a><a href="bloknews.asp?pid=<%=pid%>">Разместить 
        в блок</a> | <a href="edpub.asp?pid=<%=pid%>">Редактировать</a> 
        <%}%>
      </p>
      <table width="100%" border="0" bordercolor="#FFFFFF">
        <tr bordercolor="#CCCCCC"> 
          <td valign="top"> 
            <table border="1" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
              <tr> 
                <td align="left">
                  <h1><b><font color="#000000"><%=pname%></font></b> <font size="2" color="#000000"><%=sos%></font></h1>
                </td>
              </tr>
              <tr> 
                <td bgcolor="#F7F7F7" bordercolor="#EBF3F5"> 
                  <p class="digest"><font size="1"><%=pdat%>&nbsp; <%=autor%>&nbsp; 
                    <%=sos%></font></p>
                </td>
              </tr>
              <tr> 
                <td align="left"> 
                  <table border="0" cellspacing="0" cellpadding="0" align="center" bordercolor="#FFFFFF">
                    <tr> 
                      <td bgcolor="#000000" align="center" height="10" background="HeadImg/lenta_top.gif"></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#000000" align="center"> 
                        <%if (imgRname != "") {%>
                        <img src="<%=imgRname%>" border="1" alt="<%=pname%>" > 
                        <%}else{%>
                        <%}%>
                      </td>
                    </tr>
                    <tr> 
                      <td bgcolor="#000000" background="HeadImg/lenta_bott.gif" height="8"> 
                      </td>
                    </tr>
                  </table>
                  <p class="digest"><%=news%></p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <hr size="3" width="75%" noshade>
      <p align="center"><font color="#666666" size="1">При использовании информационных 
        материалов 72RUS.RU на интернет сайтах, включая тексты, фотографии, изображения, 
        обязательно указание гиперссылки вида: <a href="http://www.72rus.ru">www.72rus.ru</a> 
        Для авторских материалов, также обязательно указывать автора.</font></p>
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
</div>
</body>
</html>
