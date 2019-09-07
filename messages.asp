<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\url.inc" -->


<%
var subj_id=parseInt(Request("subj"))
var tit=""
var path=""
var sbj=0
var nm=""
var namarea=""
var leftit=""
var ritit=""
var canaddmsg=0
var mid=0
var mname=""
var tip=0
var kvo=0
var nik=""
var price=""
var dreg=""
var citynam=""
var rset=CreateRecordSet()

var urlcountall=0
var urlcount=""

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""

if (isNaN(subj_id)) {subj_id=0}
sbj=subj_id

Session("test_subj")=subj_id

while (sbj>0) {
	Records.Source="Select * from trade_subj where id="+sbj+" and marketplace_id="+market
	Records.Open()
	if (Records.EOF){
		Records.Close()
		Response.Redirect("messages.asp")
	}
	nm=String(Records("NAME").Value)
	if (sbj==subj_id) {
		path=nm+path
		namarea=String(Records("NAME").Value)
		leftit=String(Records("IN_NAME").Value)
		ritit=String(Records("OUT_NAME").Value)
		tip=Records("MSG_TYPE").Value
	}
	else {path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> / "+path}
	sbj=Records("HI_ID").Value
	tit=nm+" : "+tit
	Records.Close()
}

if (subj_id>0) {
path="<a href=\"messages.asp\">Объявления 72RUS</a> / "+path
Records.Source="Select * from trade_subj where hi_id="+subj_id+" and marketplace_id="+market
Records.Open()
if (Records.EOF) {canaddmsg=1}
Records.Close()
}
else {path="Объявления 72RUS - Тюменский Регион"}

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
<title><%=tit%> - Доска объявлений 72RUS - Тюмень, Тюменская Область</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  1px 3px 0px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        <%=sminame%> Доска объявлений - Тюмень, Тюменская Область</font> </p>
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
      <font color="#000000"> </font> </td>
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
<table border="0" cellspacing="0" cellpadding="0" height="1000" align="center">
  <tr> 
    <td valign="top" align="center"> 
      <table width="100%" border="0" cellspacing="0" bordercolor="#003366">
        <tr bgcolor="#FBF8D7"> 
          <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle"> 
            <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="index.asp">72RUS.RU</a> / <%=path%></p>
          </td>
        </tr>
      </table>
      <table width="97%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
        <tr> 
          <td bgcolor="#FF9900" align="left" background="images/fon_menu08.gif"> 
            <h1><b><font color="#FFFFFF"> 
              <%if (subj_id>0) {%>
              <%=namarea%> 
              <%} else {%>
              Тюменская доска объявлений 
              <%}%>
              </font></b>&nbsp;</h1>
          </td>
         </tr>
      </table>
<%
if (tip>=10 && tip < 20) 
// Знакомства
{%>      
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="35">
        <tr> 
          <td valign="middle"> 
            <p><img src="/HeadImg/star.gif" width="15" height="16" align="absmiddle"><img src="/HeadImg/star.gif" width="15" height="16" align="absmiddle"><img src="/HeadImg/star.gif" width="15" height="16" align="absmiddle">&nbsp;&nbsp; 
              <a href="banclik.asp?bid=34"> <b>Bride.Ru - знакомства c мужчинами 
              из Западной Европы, США, Канады и Австралии</b>.</a></p>

      </td>
        </tr>
      </table><%}%>
      <%if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && canaddmsg==1) {%>
      <p> <font color="#FF6600">+</font><a href="addms.asp?subj=<%=subj_id%>">Добавить 
        объявление</a></p>
      <%}%>
      <table border="1" cellspacing="5" bordercolor="#FFFFFF" width="99%">
        <%
Records.Source="Select ID,NAME from trade_subj where HI_ID="+subj_id+" and marketplace_id="+market+" order by name"
Records.Open()
while (!Records.EOF)
{
	sbj_id=String(Records("ID").Value)
	subj_name=String(Records("NAME").Value)
	rset.Source="Select COUNT_MSG from GET_COUNT_MSG_ST("+sbj_id+",0)"
	rset.Open()
	kvo=String(rset("COUNT_MSG").Value)
	rset.Close()
	//kol_ads=0
	//Records1.Source="Select '( '||CAST(count_ads as varchar(3))||' )' as kvo from GET_COUNT_ADS_CT("+sbj_id+","+city_id+")"
	//Records1.Open()
	//kol_ads=String(Records1("KVO").Value)
	//Records1.Close()
	Records.MoveNext()
%>
        <tr bgcolor="#FBF8D7"> 
          <td height="18" bgcolor="#4594D8" width="20" align="center" bordercolor="#003366"><font 
                  face=Verdana size=1><img src="images/arrow02.gif" width="7" height="7" align="middle"></font></td>
          <td colspan="3" height="18" bgcolor="#EBF5ED" bordercolor="#003366" background="HeadImg/shadow.gif"> 
            <p><a href="messages.asp?subj=<%=sbj_id%>"><b><%=subj_name%></b></a> 
              (<%=kvo%>) 
              <% if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {%>
              <a href="edsubjmsg.asp?subj=<%=sbj_id%>">изменить</a> 
              <% if (kvo==0) {%>
              | удалить 
              <%}%>
              <%}%>
            </p>
          </td>
        </tr>
        <%} Records.Close()%>
      </table>
      <%
 if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {
 %>
      <table width="97%" border="1" bordercolor="#FFFFFF">
        <tr> 
          <td colspan="2" bgcolor="#FFCC00"> 
            <p align="center"><b>Управление:</b></p>
          </td>
        </tr>
        <tr> 
          <td bordercolor="#FFCC00" width="50%"> 
            <p align="center"><font size="2"><a href="addsubjms.asp?subj=<%=subj_id%>">Добавить 
              раздел</a></font></p>
          </td>
          <%if (subj_id>0) {%>
          <td bordercolor="#FFCC00"> 
            <p align="center"><font size="2"><a href="addms.asp?subj=<%=subj_id%>">Добавить 
              объявление</a></font></p>
          </td>
          <%}%>
        </tr>
      </table>
      <%}%>
      <%
Records.Source="Select * from trademsg where TRADE_SUBJ_ID="+subj_id+"and STATE=0 and DATE_END>='TODAY' order by name"
Records.Open()
if (!Records.EOF) { 
Records.Close()
%>
      <%if (tip>=30 && tip<40) {%>
      <table width="98%" border="1" bordercolor="#FFFFFF">
        <tr> 
          <td bgcolor="#FBF8D7" bordercolor="#003366" height="2" background="HeadImg/shadow.gif"> 
            <div align="center"> 
              <p><b><%=ritit%></b></p>
            </div>
          </td>
        </tr>
      </table>
      <%
Records.Source="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.TRADE_SUBJ_ID="+subj_id+"and t1.IS_FOR_SALE=1 and t1.STATE=0 and t1.DATE_END>='TODAY' order by t1.DATE_CREATE desc"
Records.Open()
while (!Records.EOF) { 
	mid=Records("ID").Value
	mname=Records("NAME").Value
	nik=TextFormData(Records("NIKNAME").Value,"")
	price=TextFormData(Records("PRICE").Value,"")
	dreg=Records("DATE_CREATE").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	while (mname.indexOf("<")>=0) {mname=mname.replace("<","&lt;")}
	while (nik.indexOf("<")>=0) {nik=nik.replace("<","&lt;")}
	while (price.indexOf("<")>=0) {price=price.replace("<","&lt;")}
%>
      <table width="98%" border="2" bordercolor="#FFFFFF">
        <tr> 
          <td bgcolor="#EBF5ED" height="3" bordercolor="#D0E8D5"> 
            <div align="center"> 
              <p align="left"><font size="1" color="#0000FF">(<%=dreg%>)</font> 
                <a href="msg.asp?ms=<%=mid%>"> <%=mname%></a> <font size="2">(<%=nik%>)</font> 
                <font size="1"> <%=citynam%></font></p>
            </div>
          </td>
        </tr>
      </table>
      <%
Records.MoveNext()
}
Records.Close()
%>
      <%} else {%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="8">
        <tr> 
          <td></td>
        </tr>
      </table>
      <table width="97%" border="1" cellspacing="0" bordercolor="#C5DCE2">
        <tr> 
          <td bgcolor="#C6DDFF" bordercolor="#E6ECEB" background="HeadImg/shadow.gif"> 
            <p><b><font color="#003366"><%=leftit%></font></b></p>
          </td>
        </tr>
        <tr bordercolor="#FFFFFF"> 
          <td valign="top" bgcolor="#EBF3F5" bordercolor="#EBF3F5"> 
            <%
Records.Source="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.TRADE_SUBJ_ID="+subj_id+"and t1.IS_FOR_SALE=0 and t1.STATE=0 and t1.DATE_END>='TODAY' order by t1.DATE_CREATE desc"
Records.Open()
while (!Records.EOF) { 
	mid=Records("ID").Value
	mname=Records("NAME").Value
	nik=TextFormData(Records("NIKNAME").Value,"")
	price=TextFormData(Records("PRICE").Value,"")
	dreg=Records("DATE_CREATE").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	while (mname.indexOf("<")>=0) {mname=mname.replace("<","&lt;")}
	while (nik.indexOf("<")>=0) {nik=nik.replace("<","&lt;")}
	while (price.indexOf("<")>=0) {price=price.replace("<","&lt;")}
%>
            <%if (tip<10) {
// Коммерческие объявления
%>
            <table width="100%" border="1" bordercolor="#EBF3F5" cellspacing="2">
              <tr valign="top" bgcolor="#FFFFFF" bordercolor="#C5DCE2"> 
                <td height="23" width="60" align="center"> 
                  <p><font size="1" color="#003366"><%=dreg%></font></p>
                </td>
                <td height="23" bordercolor="#C5DCE2"> 
                  <p><img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
                    <a href="msg.asp?ms=<%=mid%>"> <%=mname%></a> <font size="1" color="#333333">(<%=citynam%>)</font></p>
                </td>
                <td height="23" width="120"> 
                  <p><font color="#006633"><%=price%></font>&nbsp;</p>
                </td>
              </tr>
            </table>
            <%}
if (tip>=10 && tip < 20) {
// Знакомства
%>
            <table width="100%" border="1" bordercolor="#EBF3F5">
              <tr> 
                <td bgcolor="#FCF0ED" bordercolor="#F8E0DA"> 
                  <p><font size="1" color="#0000FF">(<%=dreg%>)</font> <a href="msg.asp?ms=<%=mid%>"><%=nik%></a>: 
                    <font size="2">"<%=mname%>" <font color="#0000FF"><%=price%></font> 
                    </font><font size="1"> (<%=citynam%>)</font></p>
                </td>
              </tr>
            </table>
            <%}
if (tip>=20 && tip < 30) {
// Работа
%>
            <table width="100%" border="1" bordercolor="#EBF3F5">
              <tr> 
                <td bgcolor="#FFFBEA" bordercolor="#FFCC33"> 
                  <p><font size="1" color="#0000FF">(<%=dreg%>)</font> <a href="msg.asp?ms=<%=mid%>"> 
                    <%=mname%></a> <font size="2">(<%=nik%>)</font> <font size="1" color="#FF0000">(<%=citynam%>) 
                    </font><font color="#006600" size="2">з/пл.: <%=price%></font> 
                  </p>
                </td>
              </tr>
            </table>
            <%}%>
            <%
Records.MoveNext()
}
Records.Close()
%>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="8">
        <tr> 
          <td> 
      </td>
        </tr>
      </table>
      <table width="97%" border="1" cellspacing="0" bordercolor="#C5DCE2">
        <tr> 
          <td bgcolor="#C6DDFF" bordercolor="#E6ECEB" background="HeadImg/shadow.gif"> 
            <p><font color="#003366"><b><%=ritit%></b></font></p>
          </td>
        </tr>
        <tr> 
          <td bordercolor="#EBF3F5" valign="top" bgcolor="#EBF3F5"> 
            <%
Records.Source="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.TRADE_SUBJ_ID="+subj_id+"and t1.IS_FOR_SALE=1 and t1.STATE=0 and t1.DATE_END>='TODAY' order by t1.DATE_CREATE desc"
Records.Open()
while (!Records.EOF) { 
	mid=Records("ID").Value
	mname=Records("NAME").Value
	nik=TextFormData(Records("NIKNAME").Value,"")
	price=TextFormData(Records("PRICE").Value,"")
	dreg=Records("DATE_CREATE").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	while (mname.indexOf("<")>=0) {mname=mname.replace("<","&lt;")}
	while (nik.indexOf("<")>=0) {nik=nik.replace("<","&lt;")}
	while (price.indexOf("<")>=0) {price=price.replace("<","&lt;")}
%>
            <%if (tip<10) {
// Коммерческие объявления
%>
            <table width="100%" border="1" bordercolor="#EBF3F5" cellspacing="2">
              <tr valign="top" bordercolor="#C5DCE2" bgcolor="#FFFFFF"> 
                <td height="23" width="60" align="center"> 
                  <p><font size="1" color="#003366"><%=dreg%></font></p>
                </td>
                <td height="23"> 
                  <p><img src="HeadImg/arr_gr.gif" width="5" height="5" align="absmiddle"> 
                    <a href="msg.asp?ms=<%=mid%>"> <%=mname%></a> <font size="1" color="#333333">(<%=citynam%>)</font></p>
                </td>
                <td height="23" width="120"> 
                  <p><font color="#006633" size="1"><%=price%></font><font size="1">&nbsp;</font></p>
                </td>
              </tr>
            </table>
            <%}
if (tip>=10 && tip < 20) {
// Знакомства
%>
            <table width="100%" border="1" bordercolor="#EBF3F5">
              <tr> 
                <td bgcolor="#E0F3F8" height="21" bordercolor="#CFDBE7"> 
                  <p><font size="1" color="#0000FF">(<%=dreg%>)</font> <a href="msg.asp?ms=<%=mid%>"><%=nik%></a>: 
                    <font size="2">&quot;<%=mname%>&quot;&nbsp;<font color="#0000FF"><%=price%></font> 
                    </font><font size="1"> (<%=citynam%>)</font></p>
                </td>
              </tr>
            </table>
            <%}
if (tip>=20 && tip < 30) {
// Работа
%>
            <table width="100%" border="2" bordercolor="#EBF3F5">
              <tr> 
                <td bgcolor="#FFFBEA" bordercolor="#FFCC33"> 
                  <p><font size="1" color="#0000FF">(<%=dreg%>)</font> <a href="msg.asp?ms=<%=mid%>"> 
                    <%=mname%></a> <font size="2">(<%=nik%>)</font><font size="1" color="#FF0000"> 
                    (<%=citynam%>)</font> <font color="#006600" size="2"> з/пл.: 
                    <%=price%> </font> </p>
                </td>
              </tr>
            </table>
            <%}%>
            <%
Records.MoveNext()
}
Records.Close()
%>
          </td>
        </tr>
      </table>
      <%
}
}
else {Records.Close()}
%>
      <p>&nbsp;</p>
      <p class="digest"><font color="#0066CC" size="2">*** Для размещения своего 
        объявления перейдите в соответствующую тематическую рубрику (подрубрику) 
        объявлений и нажмите &quot;Добавить объявление&quot;</font><font color="#C6DDFF" size="2"> 
        </font></p>
      <br>
      <table border="0" cellpadding="0" cellspacing="0" width="98%">
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
