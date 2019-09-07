<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->


<%
var hid=0
var id=parseInt(Request("url"))
var sql=""
var hiadr=""
var tit=""
var host=0
var tekhia=0
var urladr=""
var ErrorMsg=""
var name=""
var url=""
var about=""
var ShowForm=true
var urlst=5
var isadm=false
var areaadr=""

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
			

if (isNaN(id)) {Response.Redirect("catarea.asp?hid=0")}
if (id<=0) {Response.Redirect("catarea.asp?hid=0")}


sql="Select t1.* from url t1, catarea t2 where t1.id="+id+" and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("catarea.asp?hid=0")
}
name=String(Records("NAME").Value)
url=String(Records("URL").Value)
about=String(Records("ABOUT").Value)
host=parseInt(Records("HOST_URL_ID").Value)
urlst=Records("STATE").Value
hid=Records("CATAREA_ID").Value
Records.Close()

if (isNaN(host)) {host=0}

tit=name

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog) {
	if (Session("is_rite_connect_url")!=1 || Session("id_host_url")!=host) {Response.Redirect("usrarea.asp")}
} else {isadm=true}

if  (isadm) { if (urlst==1) {areaadr="admarea.asp"} else {areaadr="admarea.asp?st="+urlst}}
else {areaadr="usrarea.asp"}

isFirst=String(Request.Form("reg")) == "undefined"

if (String(Request.Form("reg")) != "undefined") {
	name=TextFormData(Request.Form("name"),"")
	url=String(Request.Form("urls"))
	about=TextFormData(Request.Form("about"),"")
	
	while (name.indexOf("'")>-1) {name=name.replace("'","\"")}
	while (name.indexOf("  ")!=-1) {name=name.replace("  "," ")}
	while (name.indexOf(" ")==0) {name=name.replace(" ","")}
	while (about.indexOf("  ")!=-1) {about=about.replace("  "," ")}
	while (about.indexOf("'")>-1) {about=about.replace("'","\"")}
	while (about.indexOf(" ")==0) {about=about.replace(" ","")}
	
	while (url.indexOf(" ")!=-1) {url=url.replace(" ","")}
	
	
	if (url.length<5) {ErrorMsg+="Слишком короткий URL.<br>"}
	if (name.length<5) {ErrorMsg+="Слишком короткое наименование.<br>"}
	if (about.length<10) {ErrorMsg+="Слишком короткое описание.<br>"}
	if (!isadm) {
		if (url.indexOf("?")>-1) {ErrorMsg+="Недопустимый символ в URL (?).<br>"}
		if (url.indexOf("=")>-1) {ErrorMsg+="Недопустимый символ в URL (=).<br>"}
		if (url.indexOf("&")>-1) {ErrorMsg+="Недопустимый символ в URL (&).<br>"}
	}
	if (url.indexOf("@")>-1) {ErrorMsg+="Недопустимый символ в URL (@).<br>"}
	if (url.indexOf("'")>-1) {ErrorMsg+="Недопустимый символ в URL ( ' ).<br>"}
	if (url.indexOf("\"")>-1) {ErrorMsg+="Недопустимый символ в URL ( \" ).<br>"}
	
	if (url.indexOf("http://")==-1){url="http://"+url}
	
	if (host>0) {
		sql="Select * from URL where Upper(url)=Upper('"+url+"') and host_url_id<>"+host
		Records.Source=sql
    	Records.Open()
   		if (!Records.EOF){ErrorMsg+="Этот сайт уже зарегистрирован другим пользователем!  Если Вы хозяин ресурса , то просим Вас обратиться в нашу службу поддержки!<br>"}
   		Records.Close()
   }
   sql="Select * from URL where Upper(url)=Upper('"+url+"') and catarea_id="+hid+" and id <> "+id
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Сайт не может быть размещен дважды в одном разделе каталога!<br>"}
   Records.Close()
   sql="Select * from URL where Upper(name)=Upper('"+name+"') and catarea_id="+hid+" and id <> "+id
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Смените наименование сайта! такое наименование уже встречается в текущем разделе каталога!<br>"}
   Records.Close()

	if (ErrorMsg==""){
	  		sql="Update URL set name='%NAM', url='%URL', about='%AB', state=%ST, new_date='TODAY' where ID=%ID"
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAM",name)
			sql=sql.replace("%URL",url)
			sql=sql.replace("%AB",about)
			if (!isadm && urlst<3) {sql=sql.replace("%ST",-2)} 
			else {sql=sql.replace("%ST",urlst)}
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
}

while (name.indexOf("\"")>-1) {name=name.replace("\"","&quot;")}
while (about.indexOf("\"")>-1) {about=about.replace("\"","&quot;")}

if (url.indexOf("http://")!=-1){url=url.replace("http://","")}

%>

<Html>
<Head>
<Title>Изменение ресурса: <%=tit%></Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72.css" type="text/css">
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
        | <a href="usrarea.asp">Кабинет Пользователя</a> | </p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" height="618"> 
      <div align="left"> </div>
    </td>
    <td align="left" height="618"> 
      <p align="center">Ваш сайт 
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
        &lt; <a href="index.asp"> На главную страницу</a> 
        <%} %>
      </p>
      <p align="center"><font size="3"><br>
        <b>Изменение реквизитов интернет - ресурса</b></font></p>
      <%if(ErrorMsg!=""){%>
      <center>
        <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
          <%=ErrorMsg%></p>
      </center>
      <%}%>
      <% if (ShowForm) { %>
      <form name="form3" method="post" action="edurl.asp">
        <input type="hidden" name="url" value="<% =id %>">
        <div align="center"> 
          <table width="100%" border="1" cellspacing="2" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
            <tr bordercolor="#0000CC"> 
              <td width="37%" bgcolor="#0066CC"> 
                <div align="center"> 
                  <p><b><font color="#FFFFFF">Параметры</font></b><font color="#FFFFFF">:</font></p>
                </div>
              </td>
              <td bgcolor="#0066CC"> 
                <div align="center"> 
                  <p><b><font color="#FFFFFF">Значения</font></b></p>
                </div>
              </td>
            </tr>
            <tr> 
              <td width="37%" bgcolor="#EEEEEE" bordercolor="#999999"> 
                <div align="right"> 
                  <p align="center">Наименование сайта&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE" bordercolor="#999999"> 
                <div align="left"> 
                  <p align="center">&nbsp;&nbsp;&nbsp; 
                    <input type="text" name="name" size="35" maxlength="99" value="<%=name%>">
                    <br>
                    <font size="1">(до 100 символов)</font></p>
                </div>
              </td>
            </tr>
            <tr> 
              <td width="37%" bgcolor="#EEEEEE" bordercolor="#999999"> 
                <div align="right"> 
                  <p align="center">URL сайта&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE" bordercolor="#999999"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; http:// 
                    <input type="text" name="urls" size="27" maxlength="93" value="<%=url%>">
                  </p>
                </div>
              </td>
            </tr>
            <tr> 
              <td width="37%" bgcolor="#EEEEEE" bordercolor="#999999"> 
                <div align="right"> 
                  <p align="center">Краткое описание сайта&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE" bordercolor="#999999"> 
                <div align="left"> 
                  <p align="center">&nbsp;&nbsp;&nbsp; 
                    <input type="text" name="about" size="35" maxlength="250" value="<%=about%>">
                    <font size="1"><br>
                    (до 250 символов)</font></p>
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
          <input type="submit" name="reg" value="Сохранить">
        </div>
      </form>
      <hr size="4" noshade>
      <%} else {%>
      <p>&nbsp;</p>
      <p align="center"><font color="#FF0000"> <b>Изменения ресурса сохранены!</b></font></p>
      <%if (!isadm && urlst<3) {%>
      <p align="center"><font color="#FF0000"><b>Ресурс поставлен в очередь повторной 
        проверки.</b></font></p>
      <p align="center"><font color="#FF0000"> <b>После проверки он будет доступен 
        через наш каталог.</b></font></p>
      <%}%>
      <p align="center">&nbsp;</p>
      <p align="center"><a href="<%=areaadr%>">В Кабинет пользователя</a></p>
      <% }%>
      <p>&nbsp;</p>
    </td>
    <td width="150" height="618"> 
      <table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="CENTER" bordercolor="#EBF5ED"> 
            <script language="javascript" src="banshow.asp?rid=5"></script>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=6"></script>
  <hr size="1">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#FF6600"> 
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
