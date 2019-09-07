<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->



<%
var sql=""
var hiadr=""
var tit=""
var hname=""
var tekhia=0
var urladr=""
var logname=""
var pass=""
var ErrorMsg=""
var name=""
var url=""
var about=""
var id=0
var ShowForm=true
var hid=0
var idh=0
var statname=""
var sta=0
var Recs=CreateRecordSet()

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""	


if (String(Request.Form("Enter")) != "undefined") {
	logname=String(Request.Form("logname"))
	pass=String(Request.Form("psw"))

	logname=logname.replace("/*","")
	logname=logname.replace("'","")
	pass=pass.replace("/*","")
	pass=pass.replace("'","")

	if (logname.length<2) {ErrorMsg+="Слишком короткое имя! <br>"}
	if (pass.length<2) {ErrorMsg+="Слишком короткий пароль! <br>"}
	sql="Select * from host_url where nikname='"+logname+"' and psw='"+pass+"'"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){
		sta=Records("STATE").Value
		if (sta==1) {
		Session("is_rite_connect_url")=1
		Session("id_host_url")=Records("ID").Value
		tit=Records("NAME").Value
		}
		if (sta==0) {
			Session("is_rite_connect_url")=0
			ErrorMsg+="Error JDBC occupation! Virus warning!<br>"
		}
		if (sta==-1) {
			id=Records("ID").Value
			Records.Close()
			Response.Redirect("activateusrurl.asp?nik="+id)
		}
   } else {
		Session("is_rite_connect_url")=0
		ErrorMsg+="Неверный пароль или имя пользователя! <br>"
   }
   Records.Close()
}

if (Session("is_rite_connect_url")==1) {
	sql="Select * from host_url where id="+Session("id_host_url")
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){
		tit=Records("NAME").Value
   } else {
		Session("is_rite_connect_url")=0
		ErrorMsg+="Неверный пароль или имя пользователя! <br>"
   }
   Records.Close()
}

%>

<Html>
<Head>
<Title>Кабинет зарегистрированного пользователя.</Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72RUS.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        72RUS.RU Кабинет пользователя</font></p>
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
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="catarea.asp">Каталог 72RUS.RU</a> 
        | 
        <%
isnews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
        <a  class=globalnav href="<%=url%>"><%=hname%></a> | 
        <%
} 
Records.Close()
%>
        <%
isnews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
        <a  class=globalnav href="<%=url%>"><%=hname%></a> | 
        <%
} 
Records.Close()
%>
        <a href="messages.asp">Объявления</a> | </p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top"> 
    <td width="172" align="center"> 
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
      <table width="100%" border="0" cellspacing="0" height="60">
        <tr> 
          <td align="CENTER"><%=news_bl%></td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
%>
    </td>
    <td align="center" > 
      <p>&nbsp;</p>
      <p><font size="3"><b>Кабинет зарегистрированного пользователя 72RUS.RU</b></font></p>
      <p> 
        <%if(ErrorMsg!=""){%>
      </p>
      <center>
        <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
          <%=ErrorMsg%></p>
      </center>
      <%}%>
      <%
if (Session("is_rite_connect_url")!=1) {
	Session("is_rite_connect_url")=0
%>
      <p align="center">Для входа в кабинет Вам необходимо ввести имя пользователя 
        и пароль.</p>
      <form name="form2" method="post" action="usrarea.asp">
        <div align="center"> 
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="47%">&nbsp;</td>
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td width="47%" height="18"> 
                <div align="right"> 
                  <p><b>Ваш логин</b></p>
                </div>
              </td>
              <td valign="top" height="18"> 
                <p> 
                  <input type="text" name="logname">
                </p>
              </td>
            </tr>
            <tr> 
              <td width="47%"> 
                <div align="right"> 
                  <p><b>Ваш пароль</b></p>
                </div>
              </td>
              <td valign="top"> 
                <p> 
                  <input type="password" name="psw">
                </p>
              </td>
            </tr>
            <tr valign="bottom"> 
              <td width="47%" height="40"> 
                <div align="right"> 
                  <p>Если Вы еще не зарегистрировались</p>
                </div>
              </td>
              <td height="40"> 
                <p>Вы можете зарегистрироваться <a href="regmemurl.asp">ЗДЕСЬ</a></p>
              </td>
            </tr>
          </table>
        </div>
        <p align="center"> 
          <input type="submit" name="Enter" value="Войти">
        </p>
      </form>
      <%} else {%>
      <p align="left">&nbsp;</p>
      <p align="left">Здравствуйте, <%=tit%></p>
      <table width="100%" border="1" cellspacing="2" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
        <tr> 
          <td bgcolor="#0066CC" height="12" bordercolor="#CCCCCC"> 
            <p align="center"><b><font color="#FFFFFF">Управление</font></b></p>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#EEEEEE" bordercolor="#CCCCCC"> 
            <p><a href="catarea.asp?hid=0">Выбрать подкаталог для размещения нового 
              сайта</a></p>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#EEEEEE" height="10" bordercolor="#CCCCCC"> 

              <p align="left"><a href="edmemurl.asp">Изменить реквизиты администратора</a></p>

          </td>
        </tr>
        <tr> 
          <td bgcolor="#EEEEEE" bordercolor="#CCCCCC"> 

              <p align="left"><a href="support.asp">Служба поддержки</a> </p>

          </td>
        </tr>
        <tr>
          <td bgcolor="#EEEEEE" bordercolor="#CCCCCC">
            <p><a href="newshow.asp?pid=728">Разместить рекламу</a></p>
          </td>
        </tr>
      </table>
      <%
	sql="Select * from url where  host_url_id="+Session("id_host_url")
	Records.Source=sql
	Records.Open()
	while (!Records.EOF) {
		name=Records("NAME").Value
		url=Records("URL").Value
		about=Records("ABOUT").Value
		id=Records("ID").Value
		hid=Records("CATAREA_ID").Value
		switch (Records("STATE").Value) {
			case -2 : statname="В повторной обработке"; break;
			case -1 : statname="В обработке"; break;
			case  0 : statname="Недоступен. Ожидает подтверждения"; break;
			case  1 : statname="В эфире!!"; break;
			case  2 : statname="Остановлен Вами"; break;
			case  3 : statname="Остановлен редактором"; break;
			case  4 : statname="Удалён редактором"; break;
		}
		Records.MoveNext()
	%>
      <table width="100%" border="1" cellspacing="2" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
        <tr bgcolor="#0066CC" bordercolor="#CCCCCC" align="center"> 
          <td colspan="2" height="26"> 
            <p><b><font color="#FFFFFF">Зарегистрированные Вами сайты</font></b> 
              <font color="#FFFFFF"> </font></p>
          </td>
        </tr>
        <tr bgcolor="#EEEEEE" bordercolor="#CCCCCC"> 
          <td width="60%" height="26"> 
            <p><b><%=name%></b> <br>
              <a href="<%=url%>"><font color="#0000FF"><%=url%></font></a></p>
          </td>
          <td height="26"> 
            <div align="center"> 
              <p>[<font color="#FF0000"><%=statname%></font>]<br>
                <a href="delurl.asp?url=<%=id%>">удалить</a>&nbsp;&nbsp;&nbsp;<a href="edurl.asp?url=<%=id%>">изменить</a>&nbsp;&nbsp;</p>
            </div>
          </td>
        </tr>
        <tr bordercolor="#CCCCCC"> 
          <td colspan="2" height="14"> 
            <p><font color="#663399"><%=about%></font></p>
          </td>
        </tr>
        <tr bordercolor="#CCCCCC"> 
          <td colspan="2" height="14"> 
            <p><b>Каталог &gt; </b> 
              <%
			  idh=hid
			  while (idh>0) {
			  	sql="Select * from catarea where id="+idh
			  	Recs.Source=sql
			  	Recs.Open()
			  	if (!Recs.EOF) {
			  		idh=Recs("HI_ID").Value
					hname=Recs("NAME").Value %>
              &nbsp;<%=hname%>&nbsp;&gt; 
              <% 
				} else {idh=0}
				Recs.Close()
			  }	%>
            </p>
          </td>
        </tr>
      </table>
      <%
		} 
		Records.Close()
		%>
      <%  } %>
      <hr size="4" noshade>
    </td>
    <td width="150" align="center"> 
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
      <p><b><font size="3" color="#4A824E"><%=blokname%></font></b> </p>
      <script language="javascript" src="banshow.asp?rid=5"></script>
    </td>
  </tr>
</table>
<div align="center"> 
  <hr size="1">
  <div align="center"> 
    <script language="javascript" src="banshow.asp?rid=6"></script>
    <hr size="1">
    <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
      <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
        <td valign="middle" bgcolor="#FF6600" bordercolor="#FFFFFF"> 
          <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
            портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
            Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
            Русинтел</font></a> <font color="#FFFFFF">&copy; 2002 - 2003</font></font></b></p>
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
</Body>
</Html>
