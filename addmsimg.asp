<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var mid=Request("ms")

var divis=0
var hid=0
var name=""
var articul=""
var srname=""
var mname=""
var subj_id=0
var sbj=0
var nm=""
var path=""
var imgname=""
var ext=""
var size=0
var ErrorMsg=""

var smi_id=1
var news_bl=""
var ishtml2=0
var ishtml=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""

if (isNaN(parseInt(mid))) {
	if (Request.ServerVariables("REQUEST_METHOD")=="POST") {
	   updown = Server.CreateObject("ANUPLOAD.OBJ")
		mid=updown.Form("ms")
		if (Session("id_ms")!=mid) {Response.Redirect("messages.asp")}
		if (!isNaN(parseInt(mid))) {
		try {
			updown.Delete(MsFilePath+mid+".gif")
			updown.Delete(MsFilePath+mid+".jpg")
			updown.SavePath = MsFilePath
			size=parseInt(updown.GetSize("file"))
			ext=updown.GetExtension("file").toUpperCase()
			if(ext!="JPG" && ext!="GIF"){throw "Принимаются только JPG или GIF файлы."}
			if (size>20480){throw "Не более 20kB."}
			if (size==0){throw "Нет файла."}
			updown.SaveAs("file",MsFilePath+mid+"."+ext)
			Response.Redirect("msg.asp?ms="+mid)
			}
    	catch(e){ErrorMsg+=String(e.message)=="undefined"?e:e.message}
		
		} else {Response.Redirect("index.html")}
	}
	else {Response.Redirect("index.html")}
}

if (Session("id_ms")!=mid) {Response.Redirect("messages.asp")}

Records.Source="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.ID="+mid+"and t1.STATE=0 and t1.DATE_END>='TODAY'"
Records.Open()
if (!Records.EOF) { 
	mname=Records("NAME").Value
	subj_id=Records("TRADE_SUBJ_ID").Value
	while (mname.indexOf("<")>=0) {mname=mname.replace("<","&lt;")}
} else {
	Records.Close()
	Response.Redirect("messages.asp")
}
Records.Close()

sbj=subj_id

while (sbj>0) {
	Records.Source="Select * from trade_subj where id="+sbj+" and marketplace_id="+market
	Records.Open()
	if (Records.EOF){
		Records.Close()
		Response.Redirect("messages.asp")
	}
	nm=String(Records("NAME").Value)
	if (sbj==subj_id) {srname=nm}
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> / "+path
	sbj=Records("HI_ID").Value
	Records.Close()
}

path="<a href=\"messages.asp\">Все темы</a> /  "+path

%>


<html>
<head>
<title>Загрузка фотографии к объявлению: <%=mname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="/72RUS.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        72RUS.RU Доска объявлений - Тюмень, Тюменская Область</font> </p>
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
<table width="100%" border="0" cellspacing="0" bordercolor="#003366">
  <tr bgcolor="#FBF8D7"> 
    <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle" align="center"> 
      <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="index.asp">72RUS.RU</a> / <%=path%></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
  <tr> 
    <td bgcolor="#FF9900" align="center" background="images/fon_menu08.gif"> 
      <h1><b><font color="#FFFFFF">Вы добавляете фото к объявлению: <%=mname%></font></b></h1>
    </td>
  </tr>
</table>
<div align="center"><b></b> </div>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<form name="form1" method="post" action="addmsimg.asp" enctype="multipart/form-data">
  <input type="hidden" name="ms" value="<% =mid %>">
  <p align="center">&nbsp;</p>
  <div align="center"> 
    <table width="50%" border="2" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#003366" bgcolor="#EBF5ED" background="HeadImg/shadow.gif"> 
          <div align="center"> 
            <p><b><font color="#000099">Выбирите фотографию для загрузки:</font></b></p>
          </div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#003366"> 
          <div align="center"> 
            <input type="file" name="file" size="40">
          </div>
        </td>
      </tr>
    </table>
  </div>
  <p align="center"> 
    <input type="submit" name="Submit" value="Сохранить">
  </p>
</form>
<p align="center">&nbsp;</p>
<p align="center"><a href="%22msg.asp?ms=%22<%=mid%>"><font size="2">Перейти к 
  объявлению </font></a> <font size="2">&nbsp;&nbsp;&nbsp;<a href="index.html">На 
  главную страницу</a><a href="messages.asp?<%=subj_id%>"><br>
  Вернуться в раздел <%=srname%></a></font></p>
<p align="center">&nbsp;</p>
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
