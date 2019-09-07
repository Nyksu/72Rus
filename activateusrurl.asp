<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var sql=""
var tit=""
var logname=""
var pass=""
var ErrorMsg=""
var name=""
var cw=""
var id=parseInt(Request("nik"))
var ShowForm=true
var sta=0


if (isNaN(id)) {Response.Redirect("usrarea.asp")}		

sql="Select * from host_url where id="+id
Records.Source=sql
Records.Open()
if (Records.EOF) {
   Records.Close()
   Response.Redirect("usrarea.asp")
}
name=String(Records("NAME").Value)
email=String(Records("EMAIL").Value)
nikname=Records("NIKNAME").Value
cw=Records("KEYWORD").Value
Records.Close()

if (String(Request.Form("Enter")) != "undefined") {
	if (cw==String(Request.Form("cw"))) {
		sql="Update host_url set state=1 where state=-1 and id="+id
		Connect.BeginTrans()
			try{
			Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="Ошибка вставки.<br>"
			}
			if (ErrorMsg==""){
		   Connect.CommitTrans()
		   ShowForm=false
			}
	}
	else {ErrorMsg="Неверный код активации"}
}

%>

<Html>
<Head>
<Title>Активация зарегистрированного пользователя.</Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72RUS.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr> 
    <td width="150" valign="middle" background="HeadImg/yellow.jpg" height="19"> 
      <div align="center"> <a href="index.asp"><img src="HeadImg/home.gif" width="14" height="14" align="absbottom" border="0" alt="Home"></a> 
        <font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНСКИЙ РЕГИОН</b></font></div>
    </td>
    <td bordercolor="#FFCC00" background="HeadImg/yellow.jpg" height="19"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>WWW.72RUS.RU</b></font></div>
    </td>
    <td width="468" background="HeadImg/yellow.jpg" height="19"> 
      <table width="410" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="HeadImg/tmn.gif" width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНЬ</b></font></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="messages.asp">ОБЪЯВЛЕНИЯ</a></b></font></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><a href="catarea.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>КАТАЛОГ</b></font></a></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><a href="Rail_roads.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>РАСПИСАНИЕ</b></font></a></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><a href="marketBuilder.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>АНАЛИТИКА</b></font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td bgcolor="#ffd34e" width="150" align="left" valign="middle" background="HeadImg/runline.gif"> 
      <div align="center"><a href="index.asp"><img src="Img/%F2%FE%EC%E5%ED%FC%2072rus.gif" alt="www.72rus.ru" border="0"></a></div>
    </td>
    <td bordercolor="#FFCC00" bgcolor="#ffd34e" background="HeadImg/runline.gif"> 
      <div align="center"><object type="text/x-scriptlet" width=171 height="60" data="searchbox.html">
        </object> </div>
    </td>
    <td bgcolor="#ffd34e" align="right" background="HeadImg/runline.gif" width="468"><a href="http://www.rusintel.ru/"><img src="bannes/rol_ba.gif" width="468" height="60" alt="Internet Card - ROL - Tyumen" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0" background="HeadImg/line.gif">
  <tr bgcolor="#996600" background="HeadImg/line.gif"> 
    <td width="150" background="HeadImg/line.gif"> 
      <div align="center"> 
        <p align="left"><b><font color="#FFCC33">- Тюмень -</font></b></p>
      </div>
    </td>
    <td width="490" bgcolor="#996600" background="HeadImg/line.gif"> 
      <p align="left"><b><font color="#FFCC33">- Россия -</font></b> </p>
    </td>
    <td width="213" background="HeadImg/line.gif"> 
      <p><b><font color="#FFCC33">-- Инфо -- </font></b></p>
    </td>
    <td width="150" bgcolor="#996600" background="HeadImg/line.gif"> 
      <div align="center"> 
        <p align="right"><img src="HeadImg/home.gif" width="14" height="14" align="middle"> 
          <a href="index.asp"><b><font color="#FFCC33">Home</font></b></a></p>
      </div>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" bgcolor="#EBF5ED" height="586"> 
      <div align="left"> 
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Тюменский каталог</b></font></td>
          </tr>
          </tbody> 
        </table>
        <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
          <tr> 
            <td bordercolor="#FFFFFF" valign="top"> 
              <div align="left"> 
                <h2 align="left">&nbsp;</h2>
              </div>
            </td>
          </tr>
        </table>
      </div>
    </td>
    <td height="586"> 
      <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#ffd34e width="100%"> <font face="Arial, Helvetica, sans-serif" size="2"><font 
                  face=Verdana size=1><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font> 
            <font 
                  face=Verdana size=1><b> <a href="chat/index.html">72RUS</a></b></font> 
            <font size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font></font><font 
                  face=Verdana size=1><b> <a href="catarea.asp?hid=0">Каталог 
            72RUS</a> </b></font><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
            <b>Активация зарегистрированного пользователя</b></font> </td>
        </tr>
        </tbody> 
      </table>
      <h1 align="left"><font size="3">Активация зарегистрированного пользователя 
        <font color="#FF0000"><%=nikname%></font> </font><font color="#666666">(<%=name%>)</font></h1>
      <%if(ErrorMsg!=""){%>
      <center>
        <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
          <%=ErrorMsg%></p>
      </center>
      <%}%>
      <%
if (ShowForm) {
%>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;!!!</font></b> 
        Вашему аккаунту требуется активация! </p>
      <p align="left">Вам по почте (<%=email%>) был отправлен код активации. Введите 
        его в соответствующее поле и нажмите кнопку АКТИВИРОВАТЬ.</p>
      <form name="form2" method="post" action="activateusrurl.asp">
        <div align="center"> 
          <input type="hidden" name="nik" value="<%=id%>">
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="47%">&nbsp;</td>
              <td width="6%">&nbsp;</td>
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td width="47%"> 
                <div align="right"> 
                  <p><b>Код активации</b></p>
                </div>
              </td>
              <td width="6%"> 
                <p align="center">-</p>
              </td>
              <td valign="top"> 
                <p> 
                  <input type="text" name="cw" maxlength="20" size="25">
                </p>
              </td>
            </tr>
            <tr valign="bottom"> 
              <td width="47%" height="40"> 
                <div align="right"> 
                  <p>Если Вы не получили код активации по почте </p>
                </div>
              </td>
              <td width="6%" height="40"> 
                <p align="center">-</p>
              </td>
              <td height="40"> 
                <p>Вы можете <a href="sendusrkey.asp?nik=<%=id%>">ЗДЕСЬ</a> 
                  повторно отправить его на ваш E-mail адрес, указанный вами при 
                  регистрации.</p>
              </td>
            </tr>
          </table>
        </div>
        <p align="center"> 
          <input type="submit" name="Enter" value="Активировать">
        </p>
      </form>
<%} else {%>
      <p>&nbsp;</p>
      <p align="center"><b><font size="3" color="#FF0000">&nbsp;&nbsp;Активация 
        прошла успешно!!</font></b></p>
      <p align="center">&nbsp;</p>
      <p align="center"><a href="usrarea.asp">Кабинет</a></p>
      <%  }  %>
      <p>&nbsp;</p>
    </td>
    <td width="150" bgcolor="#EBF5ED" height="586"> 
      <div align="left"> 
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> Ссылки</b></font></td>
          </tr>
          </tbody> 
        </table>
      </div>
      <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
        <tr> 
          <td bordercolor="#FFFFFF" valign="top"> 
            <div align="left"> 
              <h2 align="left">&nbsp;</h2>
            </div>
          </td>
        </tr>
      </table>
      <p align="left">&nbsp;</p>
      <p align="left">&nbsp;</p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
    <td valign="middle" bgcolor="#996600" bordercolor="#FFFFFF"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
        портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
        Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
        Русинтел</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
    </td>
  </tr>
</table>
<table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e">
  <tr bgcolor="#EBF5ED" bordercolor="#EBF5ED"> 
    <td bordercolor="#EBF5ED" width="150" valign="middle" align="center"> 
      <!-- HotLog -->
      <script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=46088&im=16&r="+escape(document.referrer)+"&pg="+
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
src="http://hit3.hotlog.ru/cgi-bin/hotlog/count?s=46088&im=16" border=0 
width="88" height="31" alt="HotLog"></a></noscript> 
      <!-- /HotLog -->
    </td>
    <td bordercolor="#EBF5ED"> 
      <div align="center"> 
        <script language="JavaScript" src="http://vbn.tyumen.ru/cgi-bin/hints.cgi?vbn&auto72">
</script>
      </div>
    </td>
    <td width="150" valign="middle" align="center"> 
      <div align="left"> 
        <!--Begin of HMAO RATINGS-->
        <center>
          <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"> <img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
          <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
        </center>
        <!--End of HMAO RATINGS-->
      </div>
    </td>
  </tr>
</table>
</Body>
</Html>
