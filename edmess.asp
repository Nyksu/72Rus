<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\url.inc" -->

<%
Response.Redirect("serv.html")
var puid=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var ishtml2=0
var path=""
var ErrorMsg=""
var name=""
var dreg=""
var citynam=""
var subj_id=0
var gr=""
var cww=""
var nik=""
var price=""
var phone=""
var email=""
var sql=""
var ShowForm=true
var Login=false
var connectOk=false
var id=parseInt(Request("ms"))
var isadm=0
var filnam=""
var imgname=""
var srok=0
var newsrok=""
var sumdat=Server.CreateObject("datesum.DateSummer")
var dats = new Date()
var dend=""
var tip=""
var pname=""
var snam=""
var txt=""
var namarea=""
var sbj=0
var nm=""
var isFirst=true

if (isNaN(id)) {Response.Redirect("messages.asp")}
if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog ) {
	isadm=1
	connectOk=true
	Session("id_ms")=id
}

sql="Select t1.*, t2.name as cityname, (t1.date_end - t1.date_create) as srok from trademsg t1, city t2 where t1.city_id=t2.id and t1.ID="+id
Records.Source=sql
Records.Open()
if (!Records.EOF) { 
	name=Records("NAME").Value
	dreg=Records("DATE_CREATE").Value
	dend=Records("DATE_END").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	subj_id=Records("TRADE_SUBJ_ID").Value
	gr=Records("IS_FOR_SALE").Value
	cww=Records("CODEWORD").Value
	nik=TextFormData(Records("NIKNAME").Value,"")
	price=TextFormData(Records("PRICE").Value,"")
	phone=TextFormData(Records("PHONE").Value,"")
	email=TextFormData(Records("E_MAIL").Value,"")
	srok=Records("SROK").Value
	tip=Records("MSG_TYPE").Value
	snam=Records("IS_FOR_SALE").Value
	newsrok=String(dats.getDate())+"."+String(dats.getMonth()+1)+"."+String(dats.getYear())
	newsrok=sumdat.SummToDate(newsrok,srok+1)
	Records.Close()
} else {
	Records.Close()
	Response.Redirect("messages.asp")
}

if (tip<10) {
	pname="Цена"
	switch (snam) {
		case 0 : snam="Покупка"
		case 1 : snam="Продажа"
	}
}
if (tip>=10 && tip<20) {
	pname="Возраст / рост / вес"
	switch (snam) {
		case 0 : snam="Женщина"
		case 1 : snam="Мужчина"
	}
}
if (tip>=20 && tip<30) {
	pname="Зарплата"
	switch (snam) {
		case 0 : snam="Требуется"
		case 1 : snam="Ищу"
	}
}

if (isadm!=1) {
	if (String(Session("msusr"))!=String(id)) {
		var kd=Request("kd")
		if (String(cww)==String(kd)) {
			Session("msusr")=id
			Session("id_ms")=id
			connectOk=true
		}
		else {Login=true}
	} else {connectOk=true}
}

filnam=MsFilePath+id+".ms"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam!="") {
	ts= fs.OpenTextFile(filnam)
	txt=ts.ReadAll()
	ts.Close()
	while (txt.indexOf("<")>=0) {txt=txt.replace("<","&lt;")}
	while (txt.indexOf("\"")>=0) {txt=txt.replace("\"","&quot;")}
}

imgname=MsImgPath+id+".gif"
if (!fs.FileExists(MsFilePath+id+".gif")) { imgname="" }
if (imgname=="") {
	imgname=MsImgPath+id+".jpg"
	if (!fs.FileExists(MsFilePath+id+".jpg")) { imgname="" }
}

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
		namarea=String(Records("NAME").Value)
	}
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> <img src=\"HeadImg/arrow2.gif\" width=\"11\" height=\"10\" align=\"middle\"> "+path
	sbj=Records("HI_ID").Value
	Records.Close()
}

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst) {
	name=TextFormData(Request.Form("name"),"")
	nik=TextFormData(Request.Form("nik"),"")
	phone=TextFormData(Request.Form("phone"),"")
	email=TextFormData(Request.Form("email"),"")
	price=TextFormData(Request.Form("price"),"")
	txt=TextFormData(Request.Form("txt"),"")
	
	 if (name.length<3) {ErrorMsg+="Слишком короткое наименование.<br>"}
	 if (nik.length==0) {ErrorMsg+="Пустое имя или псевдоним.<br>"}
	 if (txt.length<15) {ErrorMsg+="Слишком короткое объявление.<br>"}
	 if (txt.length>4000) {ErrorMsg+="Слишком длинное объявление.<br>"}
	 if (email.length>0) {	 
	     if (!/(\w+)@((\w+).)*(\w+)$/.test(email)) {ErrorMsg=ErrorMsg+"Неверный формат поля 'E-mail'.<br>"}}
	sql="Update TRADEMSG set NAME='%NAME', NIKNAME='%NIK', E_MAIL='%EML', PHONE='%PH', PRICE='%PR', DATE_END='%DE', DATE_CREATE='TODAY' where ID="+id
	sql=sql.replace("%NAME",name)
	sql=sql.replace("%PH",phone)
	sql=sql.replace("%EML",email)
	sql=sql.replace("%PR",price)
	sql=sql.replace("%NIK",nik)	
	sql=sql.replace("%DE",newsrok)
	Connect.BeginTrans()
	try{
		Connect.Execute(sql)
		ts=fs.OpenTextFile(filnam,2,true)
		ts.Write(txt)
		ts.Close()
	}
	catch(e){
		Connect.RollbackTrans()
		ErrorMsg+=ListAdoErrors()
		ErrorMsg+="Ошибка изменения.<br>"
	}
	if (ErrorMsg==""){
		Connect.CommitTrans()
		ShowForm=false
	}
}

while (name.indexOf("<")>=0) {name=name.replace("<","&lt;")}
while (name.indexOf("\"")>=0) {name=name.replace("\"","&quot;")}
while (nik.indexOf("<")>=0) {nik=nik.replace("<","&lt;")}
while (nik.indexOf("\"")>=0) {nik=nik.replace("\"","&quot;")}
while (price.indexOf("<")>=0) {price=price.replace("<","&lt;")}
%>

<html>
<head>
<title>Редактирование объявления: <%=name%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
<div align="center"><b><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
  <font face=Verdana size=1><b> <a href="index.html">72RUS.RU</a></b></font> <img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"> 
  <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=path%></font></b> 
</div>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b><i>Ошибка!</i></b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<p align="center">Редактирование объявления: <font color="#0000FF"><%=name%></font></p>
<p align="center"><b><font color="#666666">Раздел объявлений: <font color="#3366FF"><%=namarea%></font> 
  ( <font color="#0066CC"><%=snam%></font> )</font></b></p>
<%if (Login) {%>
<form name="frm" method="post" action="edmess.asp">
<input type="hidden" name="ms" value="<% =id %>">
  <p align="center">Введите код доступа к объявлению: 
    <input type="password" name="kd">
    <input type="submit" name="Submit1" value="Продолжить">
  </p>
</form>
<%} else {%>
	<%if (ShowForm){%>
	<form name="form1" method="post" action="edmess.asp">
    <input type="hidden" name="ms" value="<% =id %>">
  <table width="70%" border="1" bordercolor="#FFFFFF" align="center" dwcopytype="CopyTableRow">
    <tr valign="top"> 
      <td width="200" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p>Тема объявления :&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" maxlength="100" size="45">
      </td>
    </tr>
	<tr valign="top"> 
      <td width="200" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p>Имя  
            <%if (tip<10 || tip>=20) {%>
            / организация 
            <%}%>
            :&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <input type="text" name="nik" value="<%=isFirst?nik:Request.Form("nik")%>" maxlength="100" size="45">
      </td>
    </tr>
    <tr valign="top"> 
      <td width="200" height="21" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p>Дата размещения:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td height="21" valign="top"> 
        <p><b><font color="#FF6633">&nbsp;<%=dreg%></font><font color="#0000FF"> 
          ( после редактирования изменится на текущую : <%=String(dats.getDate())+"."+String(dats.getMonth()+1)+"."+String(dats.getYear())%>)</font></b></p>
      </td>
    </tr>
	<tr valign="top"> 
      <td width="200" height="21" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p>Срок размещения до:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td height="21" valign="top"> 
        <p><b><font color="#FF6633">&nbsp;<%=dend%></font><font color="#0000FF"> 
          ( после редактирования изменится на <font color="#003399"><%=newsrok%></font> )</font></b></p>
      </td>
    </tr>
    <tr valign="top"> 
      <td width="200" height="33" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p><%=pname%>:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td height="33" valign="top"> 
        <%if (tip<30) {%><input type="text" name="price" value="<%=isFirst?price:Request.Form("price")%>" maxlength="40"><%} else {%><input type="hidden" name="price" value=""><%}%>
      </td>
    </tr>
    <tr valign="top"> 
      <td width="200" height="24" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p>Телефон:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td height="24" valign="top"> 
        <input type="text" name="phone" value="<%=isFirst?phone:Request.Form("phone")%>" maxlength="40">
      </td>
    </tr>
    <tr valign="top"> 
      <td width="200" height="23" bgcolor="#EBF5ED" valign="middle" bordercolor="#003366"> 
        <div align="right"> 
          <p>E-mail:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td height="23" valign="top"> 
        <input type="text" name="email" value="<%=isFirst?email:Request.Form("email")%>" maxlength="80">
      </td>
    </tr>
	<tr valign="top"> 
      <td width="200" height="23" bgcolor="#EBF5ED" bordercolor="#003366"> 
        <div align="right"> 
          <p>Текст объявления:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td height="23" valign="top"> 
        <textarea name="txt" cols="60" rows="8"><%=isFirst?txt:Request.Form("txt")%></textarea>
      </td>
    </tr>
  </table>
  	
  <div align="center">
    <p>
      <input type="submit" name="Submit" value="Сохранить объявление">
    </p>
    <p>&nbsp; </p>
  </div>
  <div align="center">
    <%if (imgname!="") {%>
    <img src="<%=imgname%>" border="1">
    <%} else {%>
    <p><b><font color="#FF3333">НЕТ ИЛЛЮСТРАЦИИ</font></b></p>
    <%}%>
  <p><a href="addmsimg.asp?ms=<%=id%>">Добавить (изменить) иллюстрацию к объявлению</a></p>
	</div>
	</form>
	<%} else {%>
<h1 align="center"><font color="#3333FF">Изменения сохранены!</font></h1>
<h1 align="center"><font color="#0000FF">Для обеспечения услуг беплатного размещения 
  объявлений на нашем сайте,</font></h1>
<h1 align="center"><font color="#0000FF">просим Вас кликнуть по одной из нижеследующих 
  ссылок наших спонсоров</font>
  <%}%>
  <%}%>
</h1>
<div align="center"> 
  <!-- start Link.Ru -->
  <script language="JavaScript">
// <!--
var LinkRuRND = Math.round(Math.random() * 1000000000);
document.write('<iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF&r='+LinkRuRND+' frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> </iframe>');
// -->
</script>
  <noscript> <iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> 
  </iframe> </noscript> 
  <!-- end Link.Ru -->
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
</div>

</body>
</html>
