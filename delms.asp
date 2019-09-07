<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var id=parseInt(Request("ms"))
var sql=""
var ShowForm=true
var isok=true
var name=""
var dreg=""
var citynam=""
var subj_id=0
var gr=""
var sbj=0
var nm=""
var leftit=""
var sname=""
var ritit=""
var path=""
var filnam=""
var imgname=""
var fs= new ActiveXObject("Scripting.FileSystemObject")


if (isNaN(id)) {Response.Redirect("messages.asp")}
if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog ) {Response.Redirect("msg.asp?ms="+id)}

Records.Source="Select t1.*, t2.name as cityname from trademsg t1, city t2 where t1.city_id=t2.id and t1.ID="+id
Records.Open()
if (!Records.EOF) { 
	name=Records("NAME").Value
	dreg=Records("DATE_CREATE").Value
	citynam=TextFormData(Records("CITYNAME").Value,"")
	subj_id=Records("TRADE_SUBJ_ID").Value
	gr=Records("IS_FOR_SALE").Value
	while (name.indexOf("<")>=0) {name=name.replace("<","&lt;")}
	Records.Close()
} else {
	Records.Close()
	Response.Redirect("messages.asp")
}

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
		sname=String(Records("NAME").Value)
	}
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> <img src=\"Img/arrow2.gif\" width=\"11\" height=\"10\" align=\"middle\"> "+path
	sbj=Records("HI_ID").Value
	Records.Close()
}

filnam=MsFilePath+id+".ms"
if (!fs.FileExists(filnam)) { filnam="" }

imgname=MsFilePath+id+".gif"
if (!fs.FileExists(imgname)) { imgname="" }
if (imgname=="") {
	imgname=MsFilePath+id+".jpg"
	if (!fs.FileExists(imgname)) { imgname="" }
}

if (String(Request.Form("next1")) != "undefined") {
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

<Html>
<Head>
<Title>Удаление объявления <%=name%></Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">

<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr> 
    <td width="150" valign="middle" background="Img/yellow.jpg" height="19"> 
      <div align="center"> <a href="index.asp"><img src="Img/home.gif" width="14" height="14" align="absbottom" border="0" alt="Home"></a> 
        <font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНСКИЙ РЕГИОН</b></font></div>
    </td>
    <td bordercolor="#FFCC00" background="Img/yellow.jpg" height="19"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>WWW.72RUS.RU</b></font></div>
    </td>
    <td width="468" background="Img/yellow.jpg" height="19"> 
      <table width="410" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="Img/tmn.gif" width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНЬ</b></font></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="messages.asp">ОБЪЯВЛЕНИЯ</a></b></font></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="catarea.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>КАТАЛОГ</b></font></a></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="Rail_roads.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>РАСПИСАНИЕ</b></font></a></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="marketBuilder.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>АНАЛИТИКА</b></font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td bgcolor="#ffd34e" width="150" align="left" valign="middle" background="Img/runline.gif"> 
      <div align="center"><a href="index.asp"><img src="Img/log.gif" alt="www.72rus.ru" border="0"></a></div>
    </td>
    <td bordercolor="#FFCC00" bgcolor="#ffd34e" background="Img/runline.gif"> 
      <div align="center"><object type="text/x-scriptlet" width=171 height="60" data="searchbox.html">
        </object> </div>
    </td>
    <td bgcolor="#ffd34e" align="right" background="Img/runline.gif" width="468"><a href="http://www.rusintel.ru/rol.html"><img src="bannes/rol_ba.gif" width="468" height="60" alt="Internet Card - ROL - Tyumen" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0" background="Img/line.gif">
  <tr bgcolor="#996600" background="Img/line.gif"> 
    <td width="150" background="Img/line.gif"> 
      <div align="center"> 
        <p align="left"><b><font color="#FFCC33">- Тюмень -</font></b></p>
      </div>
    </td>
    <td width="490" bgcolor="#996600" background="Img/line.gif"> 
      <p align="left"><b><font color="#FFCC33">- Россия -</font></b> </p>
    </td>
    <td width="213" background="Img/line.gif"> 
      <p><b><font color="#FFCC33">-- Инфо -- </font></b></p>
    </td>
    <td width="150" bgcolor="#996600" background="Img/line.gif"> 
      <div align="center"> 
        <p align="right"><img src="Img/home.gif" width="14" height="14" align="middle"> 
          <a href="index.asp"><b><font color="#FFCC33">Home</font></b></a></p>
      </div>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" bgcolor="#FFFFFF"> 
      <div align="left"> 
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%">
              <div align="right"><img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
                <font face=Verdana size=1><b> <a href="index.asp">72RUS.RU</a></b></font>&nbsp;</div>
            </td>
          </tr>
          </tbody> 
        </table>
        <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
          <tr> 
            <td bordercolor="#FFFFFF" valign="top" height="50" bgcolor="#FFFFFF"> 
              <div align="left"> 
                <table bgcolor=#ffd34e border=3 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#FFFFFF">
                  <tbody> 
                  <tr> 
                    <td bgcolor=#ffd34e width="100%" height="16"> 
                      <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Каталоги</b></font></p>
                    </td>
                  </tr>
                  </tbody> 
                </table>
                <h2><b><a href="catarea.asp">Тюменский Каталог Сайтов</a></b></h2>
                <h2><b><a href="russia.asp">Глобальный Каталог</a></b></h2>
                
              </div>
            </td>
          </tr>
        </table>
        <table bgcolor=#ffd34e border=3 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#FFFFFF">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%" height="16"> 
              <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Новости</b></font></p>
            </td>
          </tr>
          </tbody> 
        </table>
        <h2><a href="polit.asp"><b>Хроника дня</b></a></h2>
        <h2><a href="business.asp">Деловые новости</a></h2>
        <h2><a href="auto_news.html">Автомобильные новости</a></h2>
        <h2><a href="musical_news.html">Музыкальные новости</a></h2>
        <h2><a href="sport.asp">Спортивные новости</a></h2>
        <h2><a href="style.asp">Стиль</a></h2>
        <table bgcolor=#ffd34e border=3 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#FFFFFF">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%" height="16"> 
              <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Расписание</b></font></p>
            </td>
          </tr>
          </tbody> 
        </table>
        <h2 align="left"><a href="Rail_roads.html">Расписание поездов</a></h2>
        <h2 align="left"><a href="air_russia.html">Расписание самолетов</a></h2>
        <h2>&nbsp;</h2>
            </div>
    </td>
    <td> 
      <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#ffd34e width="100%"> <font face=Verdana size=1><b> Удаление 
            объявления из :  <%=path%></b></font> </td>
        </tr>
        </tbody> 
      </table>
       
      <h1 align="left"><font size="3">Удаление объявления</font></h1>
      <p>&nbsp;</p>
	  
<% if (ShowForm) {%>
      <p align="left"><b>Вы действительно хотите удалить объявление: <font color="#FF0000"><%=name%></font>? 
        (<a href="msg.asp?ms=<%=id%>">просмотреть</a>)</b></p>
      <p>&nbsp;</p>
      <h1>Раздел объявлений:<font color="#006600"> <b><%=sname%></b></font></h1>
      <h1>Группа объявлений:<font color="#006600"> <b><%=gr==0?leftit:ritit%></b></font></h1>
      <p>Город:  <%=citynam%></p>
      <p>Дата регистрации: <%=dreg%></p>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;Внимание!</font></b> 
        Удаление объявления! Если Вы продолжите то это объявление удлится из системы 
        объявлений 72RUS !</p>
	<p align="left">Удаление объявления необратимо и его невозможно будет востановить!</p>
      <p align="left">Если Вы решили удалить объявление, то поставьте флажок в 
        соответствующем поле и нажмите кнопку &quot;Продолжить&quot;.</p>
      <form name="form1" method="post" action="delms.asp">
        <input type="hidden" name="ms" value="<% =id %>">
        <p align="left"> 
          <input type="checkbox" name="agr" value="1">
          Да, я хочу удалить это объявление: (<b><font color="#FF0000" size="3"><%=name%></font></b>) 
          из системы объявлений 72RUS !</p>
        <p>&nbsp;</p>
        <p align="left"> 
          <input type="submit" name="next1" value="Продолжить">
        </p>
      </form>
<%} else {%>
      <p>&nbsp;</p>
      <p><font color="#FF0000"><b>Объявление удалено!</b></font></p>
<%}%>	  

	  <p>&nbsp;</p>
	  <p><a href="admarea.asp">Кабинет администратора</a></p>
	  <p><a href="lastmsg.asp?d=2">Последние объявления</a></p>
      <p>&nbsp;</p>
    </td>
    <td width="150" bgcolor="#FFFFFF"> 
      <div align="left">
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Ссылки</b></font></td>
          </tr>
          </tbody> 
        </table>
      </div>
      <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
        <tr> 
          <td bordercolor="#FFFFFF" valign="top" bgcolor="#FFFFFF"> 
            <div align="left"> 
              <h2 align="left"><img src="Img/arrow2.gif" width="11" height="10" align="absmiddle"> 
                <a href="http://www.auction.72rus.ru/">Сибирский Аукцион</a> - 
                проведение индивидуальных торгов.</h2>
              </div>
            </td>
        </tr>
      </table>
      <p align="left">&nbsp;</p>
      <p align="left"><img src="Img/arrow2.gif" width="11" height="10" align="absmiddle"> 
        <a href="http://www.rusintel.ru/internet.html">Твое Имя.RU</a> - <br>
        Регистрация от 20 у.е.<br>
        Хостинг от 5 у.е.</p>
      <p align="left"><img src="Img/arrow2.gif" width="11" height="10" align="absmiddle"> 
        <a href="http://www.rusintel.ru/internet.html">Интернет карты ROL</a>- 
        в Тюмени по новым тарифным планам ROL 20: днем от 0.40 у.е.,<br>
        ночью бесплатно.</p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
    <td valign="middle" bgcolor="#996600"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
        портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
        Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
        Русинтел</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
    </td>
  </tr>
</table>
</Body>
</Html>
