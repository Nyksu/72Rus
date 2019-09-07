<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var id=parseInt(Request("ms"))
var isadm=0
var Login=false
var sql=""
var name=""
var dreg=""
var citynam=""
var subj_id=0
var gr=""
var cww=""
var path=""
var ShowForm=true
var connectOk=false
var filnam=""
var imgname=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var isok=true

if (isNaN(id)) {Response.Redirect("messages.asp")}

if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog ) {
	isadm=1
	connectOk=true
}

sql="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.ID="+id
Records.Source=sql
Records.Open()
if (!Records.EOF) { 
	name=Records("NAME").Value
	dreg=Records("DATE_CREATE").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	subj_id=Records("TRADE_SUBJ_ID").Value
	gr=Records("IS_FOR_SALE").Value
	cww=Records("CODEWORD").Value
	while (name.indexOf("<")>=0) {name=name.replace("<","&lt;")}
	Records.Close()
} else {
	Records.Close()
	Response.Redirect("messages.asp")
}

if (isadm!=1) {
	if (String(Session("msusr"))!=String(id)) {
		var kd=Request("kd")
		if (String(cww)==String(kd)) {
			Session("msusr")=id
			connectOk=true
		}
		else {Login=true}
	} else {connectOk=true}
}

filnam=MsFilePath+id+".ms"
if (!fs.FileExists(filnam)) { filnam="" }

imgname=MsFilePath+id+".gif"
if (!fs.FileExists(imgname)) { imgname="" }
if (imgname=="") {
	imgname=MsFilePath+id+".jpg"
	if (!fs.FileExists(imgname)) { imgname="" }
}

if (String(Request.Form("next1")) != "undefined"){
	if (!connectOk) {Response.Redirect("msg.asp?ms="+id)}
	if (Request.Form("agr")==1) {
		sql="Delete from trademsg where id="+id
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
			if (filnam!="") {fs.DeleteFile(filnam)}
			if (imgname!="") {fs.DeleteFile(imgname)}
		}
		catch(e){
			Connect.RollbackTrans()
			isok=false
		}
		if (isok){
			Connect.CommitTrans()
			var ShowForm=false
		}
	} else {Response.Redirect("msg.asp?ms="+id)}
}
%>

<html>
<head>
<title>Удаление объявления <%=name%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<hr noshade size="1">
<div align="center">
  <!--RAX counter-->
  <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
  <!--/RAX-->
  <script language="javascript" src="banshow.asp?rid=4"></script>
</div>
<hr noshade size="1">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="messages.asp">Объявления 72RUS.RU</a> 
        | </p>
    </td>
  </tr>
</table>
<h1 align="center">&nbsp;</h1>
<h1 align="center"><font size="3">Удаление объявления</font></h1>
<p align="center"><font color="#FF0000"><%=path%> </font></p>
<p align="center"><font color="#0066CC"><b><%=name%></b></font></p>
<%if (Login) {%> 
<form name="frm" method="post" action="delmess.asp">
<input type="hidden" name="ms" value="<% =id %>">
  <p align="center">Введите код доступа к объявлению: 
    <input type="password" name="kd">
    <input type="submit" name="Submit" value="Продолжить">
  </p>
</form>
<%} else  { if (ShowForm) {%>
<p align="center"><b>Вы действительно хотите удалить объявление: <font color="#FF0000"><%=name%></font>? (<a href="msg.asp?ms=<%=id%>">просмотреть</a>)</b></p>
<p>&nbsp;</p>
<p align="center"><b><font size="3" color="#FF0000">&nbsp;</font></b></p>
<p align="left"><b><font size="3" color="#FF0000">&nbsp;Внимание!</font></b> Удаление 
  объявления! Если Вы продолжите то это объявление удлится из системы объявлений 
  72RUS !</p>
<p align="left">Удаление объявления необратимо и его невозможно будет востановить!</p>
<p align="left">Если Вы решили удалить объявление, то поставьте флажок в соответствующем 
  поле и нажмите кнопку &quot;Продолжить&quot;.</p>
      <form name="form1" method="post" action="delmess.asp">
  <input type="hidden" name="ms" value="<% =id %>">
  <p align="center"> 
    <input type="checkbox" name="agr" value="1">
    Да, я хочу удалить это объявление: (<b><font color="#FF0000" size="3"><%=name%></font></b>) из системы объявлений 72RUS ! </p>
  <p align="center">&nbsp;</p>
  <p align="center"> 
    <input type="submit" name="next1" value="Продолжить">
  </p>
</form>
<%} else {%>
<p>&nbsp;</p>
<p align="center"><font color="#FF0000"><b>Объявление удалено!</b></font></p>
<%} }%>
<p align="center"><a href="messages.asp">Вернуться объявлениям</a></p>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=6"></script>
  <hr size="1">
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
    | <a href="catarea.asp">Тюменский Каталог</a> | <br>
    © 2002 <a href="http://www.rusintel.ru">Rusintel Company</a> 
</div>
<p align="center"> 
  <!--RAX logo-->
  <a href="http://www.rax.ru/click" target=_blank><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="rax.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодня"></a> 
  <!--/RAX-->
  <!-- HotLog -->
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
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"> <img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
</body>
</html>
