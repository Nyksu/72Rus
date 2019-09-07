<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var hid=parseInt(Request("hid"))
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

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""			

if (isNaN(hid)) {hid=0}

if (hid>0) {
sql="Select * from catarea where id="+hid+" and catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("catarea.asp?hid=0")
}
hname=String(Records("NAME").Value)
Records.Close()
}
tit=hname

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && Session("id_mem")!="undefined"){
	sql="Select * from catalog where id="+catalog
	Records.Source=sql
    Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("index.asp")
   }
   if (Session("id_mem")==Records("USERS_ID").Value) {Session("cataloghost")=catalog}
   Records.Close()
}

if (String(Request.Form("next1")) != "undefined") {
	if (Request.Form("agr")==1) {Session("isagr")=1}
}

if (String(Request.Form("Enter")) != "undefined") {
	logname=String(Request.Form("logname"))
	pass=String(Request.Form("psw"))
	if (logname.length<2) {ErrorMsg+="Слишком короткое имя! <br>"}
	if (pass.length<2) {ErrorMsg+="Слишком короткий пароль! <br>"}
	sql="Select * from host_url where nikname='"+logname+"' and psw='"+pass+"'"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){
		Session("is_rite_connect_url")=1
		Session("id_host_url")=Records("ID").Value
   } else {
		Session("is_rite_connect_url")=0
		ErrorMsg+="Неверный пароль или имя пользователя! <br>"
   }
   Records.Close()
}

isFirst=String(Request.Form("reg")) == "undefined"

if (String(Request.Form("reg")) != "undefined") {
	name=TextFormData(Request.Form("name"),"")
	url=TextFormData(Request.Form("url"),"")
	about=TextFormData(Request.Form("about"),"")
	 
	while (name.indexOf("'")>-1) {name=name.replace("'","\"")}
	while (name.indexOf("  ")!=-1) {name=name.replace("  "," ")}
	while (name.indexOf(" ")==0) {name=name.replace(" ","")}
	while (about.indexOf("  ")!=-1) {about=about.replace("  "," ")}
	while (about.indexOf("'")>-1) {about=about.replace("'","\"")}
	while (about.indexOf(" ")==0) {about=about.replace(" ","")}
	
	while (url.indexOf(" ")!=-1) {url=url.replace(" ","")}
	
	 if (name.length<5) {ErrorMsg+="Слишком короткое наименование.<br>"}
	if (url.length<5) {ErrorMsg+="Слишком короткий URL.<br>"}
	if (about.length<10) {ErrorMsg+="Слишком короткое описание.<br>"}
	if (url.indexOf("?")>-1) {ErrorMsg+="Недопустимый символ в URL (?).<br>"}
	if (url.indexOf("=")>-1) {ErrorMsg+="Недопустимый символ в URL (=).<br>"}
	if (url.indexOf("&")>-1) {ErrorMsg+="Недопустимый символ в URL (&).<br>"}
	if (url.indexOf("@")>-1) {ErrorMsg+="Недопустимый символ в URL (@).<br>"}
	if (url.indexOf("'")>-1) {ErrorMsg+="Недопустимый символ в URL ( ' ).<br>"}
	if (url.indexOf("\"")>-1) {ErrorMsg+="Недопустимый символ в URL ( \" ).<br>"}
	if (url.indexOf("http://")==-1){url="http://"+url}
	sql="Select * from URL where Upper(url)=Upper('"+url+"') and host_url_id<>"+Session("id_host_url")
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Этот сайт уже зарегистрирован другим пользователем!  Если Вы хозяин ресурса , то просим Вас обратиться в нашу службу поддержки!<br>"}
   Records.Close()
   sql="Select * from URL where Upper(url)=Upper('"+url+"') and catarea_id="+hid
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Сайт не может быть размещен дважды в одном разделе каталога!<br>"}
   Records.Close()
   sql="Select * from URL where Upper(name)=Upper('"+name+"') and catarea_id="+hid
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Смените наименование сайта! такое наименование уже встречается в текущем разделе каталога!<br>"}
   Records.Close()
	
	if (ErrorMsg==""){
	  		id=NextID("urlid")
			sql="Insert into url (ID,NAME,URL,ABOUT,CATAREA_ID,STATE,REG_DATE,NEW_DATE,KEYWORD,HOST_URL_ID) "
			sql+="values (%ID, '%NAM', '%URL', '%AB', %CAT, -1, 'TODAY', 'TODAY', '', %HS)"
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAM",name)
			sql=sql.replace("%URL",url)
			sql=sql.replace("%AB",about)
			sql=sql.replace("%CAT",hid)
			sql=sql.replace("%HS",Session("id_host_url"))
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
	  while (name.indexOf("\"")>-1) {name=name.replace("\"","&quot;")}
	  while (about.indexOf("\"")>-1) {about=about.replace("\"","&quot;")}
}
if (url.indexOf("http://")!=-1){url=url.replace("http://","")}
%>

<Html>
<Head>
<Title>Размещение ресурса в каталог 72RUS в подкаталог: <%=tit%></Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<hr noshade size="1">
<div align="center">
  <script language="javascript" src="banshow.asp?rid=4"></script>
</div>
<hr noshade size="1">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="catarea.asp">Каталог 72RUS.RU</a> 
        | <a href="usrarea.asp">Кабинет Пользователя</a> | <a href="regmemurl.asp">Регистрация 
        пользователя</a> | <a href="http://chat.72rus.ru/">Чат</a></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="center"> 
      <!--RAX counter-->
      <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
      <!--/RAX-->
    </td>
    <td> 
      <p align="left">&nbsp;</p>
      <p align="center"><font size="3"><b>Регистрация интернет - ресурса в каталог 
        72RUS</b></font></p>
      <p align="center">&nbsp;</p>
      <%if(ErrorMsg!=""){%>
      <center>
        <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
          <%=ErrorMsg%></p>
      </center>
      <%}%>
      <%
if (Session("isagr")!=1) {
	Session("isagr")=0
%>
      <p align="left"><b>Три простых шага для регистрации Вашего ресурса:</b></p>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;1. </font></b>Для 
        размещения Вашего сайта в каталог 72RUS, ознакомьтесь с ПРАВИЛАМИ размещения:</p>
      <p align="left">&nbsp;</p>
      <p align="left">Администрация 72RUS допускает к размещению сайты соответствующие 
        региональной структуре каталога, имеющие отношение к выбранным темам каталога, 
        имеющие уникальное и полезное содержание.</p>
      <p align="left">&nbsp;</p>
      <p align="left">Для того, чтобы добавить сайт в каталог Вам необходимо пройти 
        <a href="regmemurl.asp">Регистрацию Нового Пользователя</a>. После этого 
        Вы получаете имя пользователя и пароль и сможете в любое время добавлять 
        удалять сайты, изменять их описание, адрес в случае его изменения. </p>
      <p align="left">&nbsp;</p>
      <p align="left">Добавление нового сайта происходит с той страницы каталога, 
        в которую Вы хотите его разместить. </p>
      <p align="left">&nbsp; </p>
      <p align="left">Редакторы каталога оставляют за собой право не размещать 
        ссылки на сайты, удалять из каталога сайты, изменять описание, переносить 
        сайты в другие подкаталоги. Это может происходить по следующим причинам: 
        представленный сайт не соответствует заявленной теме каталога, сайт не 
        имеет полезного содержания, содержание сайта противоречит законодательству 
        РФ, нормам морали, оскорбляет права и достоинство третьих лиц, URL сайта 
        не доступен.</p>
      <p align="left">&nbsp;</p>
      <p align="left">Владельцы сайтов самостоятельно несут ответственность за 
        качество, правдивость, законность содержания сайтов.</p>
      <p align="left">&nbsp;</p>
      <p align="left">Администрация 72RUS не гарантирует получение какой-либо 
        выгоды владельцам сайтов и пользователям каталога, бесперебойность работы 
        каталога. Содержание, правила, структура, дизайн каталога могут быть изменены 
        в любое время без уведомления владельцев сайтов.</p>
      <p align="left">Если Вы не согласны с правилами, пожалуйста не включайте 
        свой ресурс в наш каталог.</p>
      <p align="left">Если Вы согласны с правилами, то поставьте флажок в соответствующем 
        поле и нажмите кнопку &quot;Продолжить&quot;.</p>
      <form name="form1" method="post" action="addurl.asp">
        <input type="hidden" name="hid" value="<% =hid %>">
        <p align="left"> 
          <input type="checkbox" name="agr" value="1">
          Да, я согласен с правилами регистрации сайта в каталог 72RUS и хочу 
          разместить совой ресурс в подраздел <b><font color="#FF0000" size="3"><%=tit%></font></b> данного каталога!</p>
        <p>&nbsp;</p>
        <p align="left"> 
          <input type="submit" name="next1" value="Продолжить">
        </p>
      </form>
      <%} else {%>
      <%
if (Session("is_rite_connect_url")!=1) {
	Session("is_rite_connect_url")=0
%>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;2. </font></b>Следующий 
        шаг - Вам необходимо войти в систему под своим именем и паролем, для этого 
        заполните соответствующие поля и нажмите кнопку &quot;Войти&quot;</p>
      <p align="left">Если у Вас нет имени и пароля в нашей системе (каталог 72RUS) 
        то Вы должны зарегистрироваться. После регистрации сможете продолжить 
        регистрацию Вашего ресурса.</p>
      <form name="form2" method="post" action="addurl.asp">
        <input type="hidden" name="hid" value="<% =hid %>">
        <div align="center"> 
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="47%">&nbsp;</td>
              <td width="6%">&nbsp;</td>
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td width="47%" height="18"> 
                <div align="right"> 
                  <p><b>Ваш логин</b></p>
                </div>
              </td>
              <td width="6%" height="18"> 
                <p align="center">-</p>
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
              <td width="6%"> 
                <p align="center">-</p>
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
              <td width="6%" height="40"> 
                <p align="center">-</p>
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
      <%
if (ShowForm) {
%>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;3. </font></b>Третьим 
        и последним шагом будет заполнение собственно, приведенной ниже, формы 
        регистрации Вашего ресурса:</p>
      <p align="left">Все поля обязательны к заполнению!!</p>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#FFFFFF">
        <tbody> 
        <tr> 
          <td bgcolor=#EEEEEE width="100%" bordercolor="#999999"> 
            <div align="left"> 
              <p><font size="1"><b>Регистрация сайта в раздел</b></font> <b><font size="1"> 
                <%
	  tekhia=hid
	  while (tekhia!=0) {
	     sql="Select * from catarea where id="+tekhia
		 Records.Source=sql
		 Records.Open()
		 if (!Records.EOF) {
			hname=String(Records("NAME").Value)
			hiadr="catarea.asp?hid="+tekhia
			tekhia=Records("HI_ID").Value
		%>
                &lt; <a href="<%=hiadr%>"><%=hname%></a> 
                <%	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
                &lt; <a href="catarea.asp?hid=0">Каталог</a> 
                <% } else {%>
                &lt; <a href="index.asp"> 72RUS.RU</a> 
                <%} %>
                </font></b> </p>
            </div>
          </td>
        </tr>
        </tbody> 
      </table>
      <form name="form3" method="post" action="addurl.asp">
        <input type="hidden" name="hid" value="<% =hid %>">
        <div align="center"> 
          <table width="100%" border="1" cellspacing="2" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
            <tr> 
              <td width="37%" bgcolor="#0066CC" bordercolor="#0000CC"> 
                <div align="center"> 
                  <p><b><font color="#FFFFFF">Параметры</font></b><font color="#FFFFFF">:</font></p>
                </div>
              </td>
              <td bgcolor="#0066CC" bordercolor="#0000CC"> 
                <div align="center"> 
                  <p><b><font color="#FFFFFF">Значения</font></b></p>
                </div>
              </td>
            </tr>
            <tr bordercolor="#999999"> 
              <td width="37%" bgcolor="#EEEEEE"> 
                <div align="right"> 
                  <p>Наименование сайта&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; 
                    <input type="text" name="name" size="35" maxlength="99" value="<%=name%>">
                    <font size="1">(до 100 символов)</font></p>
                </div>
              </td>
            </tr>
            <tr bordercolor="#999999"> 
              <td width="37%" bgcolor="#EEEEEE"> 
                <div align="right"> 
                  <p>URL сайта&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; http:// 
                    <input type="text" name="url" size="27" maxlength="96" value="<%=url%>">
                  </p>
                </div>
              </td>
            </tr>
            <tr bordercolor="#999999"> 
              <td width="37%" bgcolor="#EEEEEE"> 
                <div align="right"> 
                  <p>Краткое описание сайта&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; 
                    <input type="text" name="about" size="35" maxlength="250" value="<%=about%>">
                    <font size="1">(до 250 символов)</font></p>
                </div>
              </td>
            </tr>
            <tr> 
              <td width="37%"> 
                <div align="right">&nbsp;</div>
              </td>
              <td>&nbsp;</td>
            </tr>
          </table>
          <input type="submit" name="reg" value="Зарегистрировать">
          <hr size="4" noshade>
        </div>
      </form>
      <%} else {%>
      <p>&nbsp;</p>
      <p align="center"><font color="#FF0000"> <b><font color="#009966">Ваш ресурс 
        размещен в очередь регистрации ресурса каталога 72RUS.</font></b></font></p>
      <p align="center"><font color="#009966"> <b>После проверки он будет доступен 
        в каталоге.</b></font></p>
      <p align="center">&nbsp;</p>
      <h1><font color="#0000FF">Для обеспечения услуг беплатного размещения Вашего 
        сайта в каталоге,</font></h1>
      <h1><font color="#0000FF">просим Вас кликнуть по одной из нижеследующих 
        ссылок наших спонсоров</font></h1>
      <% } } } %>
      <p>&nbsp; </p>
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
      <p><b><font size="3"><%=blokname%></font></b> </p>
      <table width="120" border="1" cellspacing="0" cellpadding="0" bordercolor="#003366">
        <tr> 
          <td align="CENTER" bordercolor="#EBF5ED"> 
            <script language="javascript" src="banshow.asp?rid=5"></script>
          </td>
        </tr>
      </table>
      <p align="CENTER"></p>
    </td>
  </tr>
</table>
<!-- start Link.Ru -->
<div align="center">
  <script language="JavaScript">
// <!--
var LinkRuRND = Math.round(Math.random() * 1000000000);
document.write('<iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF&r='+LinkRuRND+' frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> </iframe>');
// -->
</script>
  <noscript> <iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> 
  </iframe> </noscript> 
  <!-- end Link.Ru -->
</div>
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
</Body>
</Html>
