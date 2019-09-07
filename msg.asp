<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\url.inc" -->


<%
var tit=""
var path=""
var sbj=0
var nm=""
var leftit=""
var ritit=""
var canaddmsg=0
var mid=parseInt(Request("ms"))
var namarea=""
var mname=""
var tip=0
var nik=""
var price=""
var dreg=""
var citynam=""
var phone=""
var email=""
var pname=""
var about=""
var gr=""
var sname=""
var filnam=""
var str=""
var ts=""
var imgname=""
var fs= new ActiveXObject("Scripting.FileSystemObject")


var urlcountall=0

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnambl=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""

if (isNaN(mid)) {Response.Redirect("messages.asp")}

Records.Source="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.ID="+mid+"and t1.STATE=0 and t1.DATE_END>='TODAY'"
Records.Open()
if (!Records.EOF) { 
	mname=Records("NAME").Value
	nik=TextFormData(Records("NIKNAME").Value,"")
	price=TextFormData(Records("PRICE").Value,"")
	dreg=Records("DATE_CREATE").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	subj_id=Records("TRADE_SUBJ_ID").Value
	phone=TextFormData(Records("PHONE").Value,"")
	email=TextFormData(Records("E_MAIL").Value,"")
	gr=Records("IS_FOR_SALE").Value
	while (mname.indexOf("<")>=0) {mname=mname.replace("<","&lt;")}
	while (nik.indexOf("<")>=0) {nik=nik.replace("<","&lt;")}
	while (price.indexOf("<")>=0) {price=price.replace("<","&lt;")}
} else {
	Records.Close()
	Response.Redirect("messages.asp")
}
Records.Close()

if (isNaN(subj_id)) {subj_id=0}
sbj=subj_id

while (sbj>0) {
	Records.Source="Select * from trade_subj where id="+sbj+" and marketplace_id="+market
	Records.Open()
	if (Records.EOF){
		Records.Close()
		Response.Redirect("messages.asp")
	}
	nm=String(Records("NAME").Value)
	if (sbj==subj_id) {
		leftit=String(Records("IN_NAME").Value)
		ritit=String(Records("OUT_NAME").Value)
		tip=Records("MSG_TYPE").Value
		sname=String(Records("NAME").Value)
	}
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> / "+path
	sbj=Records("HI_ID").Value
	tit=nm+" : "+tit
	Records.Close()
}

filnam=MsFilePath+mid+".ms"
if (!fs.FileExists(filnam)) { filnam="" }

imgname=MsImgPath+mid+".gif"
if (!fs.FileExists(MsFilePath+mid+".gif")) { imgname="" }
if (imgname=="") {
	imgname=MsImgPath+mid+".jpg"
	if (!fs.FileExists(MsFilePath+mid+".jpg")) { imgname="" }
}

path="<a href=\"messages.asp\">Объявления 72RUS</a> / "+path

if (tip>=10 && tip<20) {pname="Возраст / рост / вес"}
if (tip<10) {pname="Стоимость (цена)"}
if (tip>=20 && tip<30) {pname="Зарплата"}

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
<title><%=mname%> : <%=tit%><%=gr==0?leftit:ritit%>: <%=citynam%> Объявления 72RUS</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
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
          <td align="center" bgcolor="#FFFFFF"> 
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
	filnambl=PubFilePath+puid+".pub"
	if (!fs.FileExists(filnambl)) { filnambl="" }

	if (filnambl != "") {
		ts= fs.OpenTextFile(filnambl)
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
<table border="0" cellspacing="0" cellpadding="0" height="900" align="center">
  <tr> 
    <td valign="top" align="center"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="6">
        <tr> 
          <td></td>
        </tr>
      </table>
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
            <h1><b><font color="#FFFFFF"> Тюменская доска объявлений </font></b>&nbsp;</h1>
          </td>
        </tr>
      </table>
      <table width="97%" border="1" cellspacing="0" bordercolor="#003366">
        <tr bgcolor="#FBF8D7"> 
          <td height="18" bgcolor="#EEEEEE" bordercolor="#EEEEEE"> 
            <table width="99%" border="1" cellspacing="0" cellpadding="0" height="23" bordercolor="#003366" align="center">
              <tr bgcolor="#EBF3F5" bordercolor="#EBF3F5"> 
                <td bordercolor="#FFFFFF" bgcolor="#FFFFFF"> 
                  <h1><b><font color="#003366"> </font><%=mname%><font color="#003366"> 
                    - </font><%=gr==0?leftit:ritit%></b>&nbsp;</h1>
                </td>
                <td bordercolor="#FFFFFF" align="center" bgcolor="#FFFFFF" width="120"> 
                  <p> 
                    <%
 if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {
 %>
                  </p>
                  <p align="center"><font size="1"><a href="delms.asp?ms=<%=mid%>"><font color="#FF0000">Удалить</font></a> 
                    | <a href="edmess.asp?ms=<%=mid%>"><font color="#006633">Изменить</font></a></font></p>
                  <%} else {%>
                  <p><font size="1"><a href="delmess.asp?ms=<%=mid%>"><font color="#FF0000">Удалить</font></a> 
                    | <a href="edmess.asp?ms=<%=mid%>"><font color="#006633">Изменить</font></a></font></p>
                  <%}%>
                </td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" height="3">
              <tr> 
                <td></td>
              </tr>
            </table>
            <table width="99%" border="1" bordercolor="#003366" align="center" cellspacing="0">
              <tr valign="top"> 
                <td colspan="2" height="24" bgcolor="#FFFFFF" bordercolor="#FFFFFF" background="HeadImg/bg_line.jpg"> 
                  <%if (imgname!="") {%>
                  <p align="center"> <img src="<%=imgname%>" border="1" alt="<%=mname%>"></p>
                  <%}%>
                  <p> 
                    <% if (filnam!="") {
				ts= fs.OpenTextFile(filnam)
				while (!ts.AtEndOfStream){
					str=ts.ReadLine()
					while (str.indexOf("<")>=0) {str=str.replace("<","&lt;")}
					while (str.indexOf("\"")>=0) {str=str.replace("\"","&quot;")}
					Response.Write(str+"<br>")
				}
		 } %>
                  </p>
                </td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" height="1">
              <tr> 
                <td></td>
              </tr>
            </table>
            <table width="100%" border="1" bordercolor="#C6DDFF" align="center">
              <tr valign="top" bordercolor="#003366" bgcolor="#FFFFFF"> 
                <td width="200"> 
                  <div align="right"> 
                    <p>Имя / 
                      <%if (tip<10 || tip>=20) {%>
                      организация 
                      <%}%>
                      :&nbsp;&nbsp;</p>
                  </div>
                </td>
                <td> 
                  <p><b><font color="#0000FF">&nbsp;<%=nik%></font></b></p>
                </td>
              </tr>
              <tr valign="top" bordercolor="#003366" bgcolor="#FFFFFF"> 
                <td width="200" height="21"> 
                  <div align="right"> 
                    <p>Дата размещения:&nbsp;&nbsp;</p>
                  </div>
                </td>
                <td height="21"> 
                  <p><b><font color="#0000FF">&nbsp;<%=dreg%></font></b></p>
                </td>
              </tr>
              <tr valign="top" bordercolor="#003366" bgcolor="#FFFFFF"> 
                <td width="200" height="24"> 
                  <div align="right"> 
                    <p><%=pname%>:&nbsp;&nbsp;</p>
                  </div>
                </td>
                <td height="24"> 
                  <p><b><font color="#0000FF">&nbsp;<%=price%></font></b></p>
                </td>
              </tr>
              <tr valign="top" bordercolor="#003366" bgcolor="#FFFFFF"> 
                <td width="200" height="23"> 
                  <div align="right"> 
                    <p>Город:&nbsp;&nbsp;</p>
                  </div>
                </td>
                <td height="23"> 
                  <p><b><font color="#0000FF">&nbsp;<%=citynam%></font></b></p>
                </td>
              </tr>
              <tr valign="top" bordercolor="#003366" bgcolor="#FFFFFF"> 
                <td width="200" height="24"> 
                  <div align="right"> 
                    <p>Телефон:&nbsp;&nbsp;</p>
                  </div>
                </td>
                <td height="24"> 
                  <p><b><font color="#0000FF">&nbsp;<%=phone%></font></b></p>
                </td>
              </tr>
              <tr valign="top" bordercolor="#003366" bgcolor="#FFFFFF"> 
                <td width="200" height="23"> 
                  <div align="right"> 
                    <p>E-mail:&nbsp;&nbsp;</p>
                  </div>
                </td>
                <td height="23"> 
                  <p><b><font color="#0000FF">&nbsp;<%=email%></font></b></p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="3">
        <tr> 
          <td></td>
        </tr>
      </table>
      <hr noshade size="3" width="97%">
      <font color="#0066CC" size="2">*** Для размещения своего объявления перейдите 
      в соответствующую тематическую рубрику (подрубрику) объявлений и нажмите 
      &quot;Добавить объявление&quot;</font><font color="#C6DDFF" size="2"> </font> 
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
