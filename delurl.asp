<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var id=parseInt(Request("url"))
var name=""
var host_id=0
var areaadr=""
var admtp=0
var stateurl=0
var hid=0
var retname="Вернутся в кабинет"
var tekhia=0
var sql=""
var hname=""
var hiadr=""
var ShowForm=true
var isok=true

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)&&(Session("is_rite_connect_url")!=1)){
Response.Redirect("catarea.asp")
}

if (isNaN(id)) {Response.Redirect("catarea.asp")}

if ((Session("is_adm_mem")!=1)||(Session("cataloghost")!=catalog)) {
	areaadr="admarea.asp"
	admtp=1
} 
if (Session("is_rite_connect_url")==1) {
	areaadr="usrarea.asp"
	admtp=2
}

sql="Select t1.* from url t1, catarea t2 where t1.id="+id+" and t1. catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect(areaadr)
}
host_id=Records("HOST_URL_ID").Value
name=Records("NAME").Value
stateurl=Records("STATE").Value
hid=Records("CATAREA_ID").Value
Records.Close()

if (admtp==1 && stateurl!=1) {areaadr+="?st="+stateurl}
if (admtp==1 && stateurl==1) {
	areaadr="catarea.asp?hid="+hid
	retname="Обратно в директорию"
}
if (admtp==2 && Session("id_host_url")!=host_id) {Response.Redirect(areaadr)}

if (String(Request.Form("next1")) != "undefined") {
	if (Request.Form("agr")==1) {
		if ((admtp==2) || (admtp==1 && stateurl==4)) {sql="Delete from url where id="+id}
		else {sql="Update url set state=4 where id="+id}
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
		}
		catch(e){
			Connect.RollbackTrans()
			isok=false
		}
		if (isok){
			Connect.CommitTrans()
			var ShowForm=false
		}
	} else {Response.Redirect(areaadr)}
}

%>

<Html>
<Head>
<Title>Удаление ресурса <%=name%></Title>
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
        | <a href="usrarea.asp">Кабинет Пользователя</a> | <a href="regmemurl.asp">Регистрация 
        пользователя</a> | <a href="http://chat.72rus.ru/">Чат</a></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" height="557"> 
      <div align="left"></div>
    </td>
    <td height="557"> 
      <p align="center">&nbsp;</p>
      <p align="center"><font size="3"><b>Удаление ресурса из каталога</b></font></p>
      <p>&nbsp;</p>
      <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#CCCCCC width="100%"> 
            <p><font size=1><b>Удаление ресурса из : </b> <b> 
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
              </b>&lt; <a href="<%=hiadr%>"><%=hname%></a> 
              <%	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
              &lt; <a href="catarea.asp?hid=0">Каталог</a> 
              <% } else {%>
              &lt;<a href="index.asp"> На главную страницу</a> 
              <%} %>
              </font> </p>
          </td>
        </tr>
        </tbody> 
      </table>
      <p>&nbsp;</p>
      <% if (ShowForm) {%>
      <p align="left"><b>Вы действительно хотите удалить ресурс: <font color="#FF0000"><%=name%></font> 
        ?</b></p>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;Внимание!</font></b> 
        Удаление ресурса! Если Вы продолжите то этот ресурс удлится из системы 
        каталога 72RUS !</p>
      <%if (admtp==2 || (admtp==1 && stateurl==4)) {%>
      <p align="left">Удаление ресурса необратимо и ресурс невозможно будет востановить!</p>
      <%}%>
      <p align="left">Если Вы решили удалить ресурс, то поставьте флажок в соответствующем 
        поле и нажмите кнопку &quot;Продолжить&quot;.</p>
      <form name="form1" method="post" action="delurl.asp">
        <input type="hidden" name="url" value="<% =id %>">
        <p align="center"> 
          <input type="checkbox" name="agr" value="1">
          Да, я хочу удалить этот ресурс: (<b><font color="#FF0000" size="3"><%=name%></font></b>) 
          из каталога 72RUS !</p>
        <p align="center">&nbsp;</p>
        <p align="center"> 
          <input type="submit" name="next1" value="Продолжить">
        </p>
      </form>
      <%} else {%>
      <p>&nbsp;</p>
      <p><font color="#FF0000"><b>Ресурс удален!</b></font></p>
      <%}%>
      <p>&nbsp;</p>
      <p><a href="<%=areaadr%>"><%=retname%></a></p>
      <%
	  if (retname!="Вернутся в кабинет") {
	  %>
      <p><a href="admarea.asp">Кабинет администратора</a></p>
      <%
	  }
	  %>
      <p>&nbsp;</p>
    </td>
    <td width="150" height="557"> 
      <table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="CENTER"> 
            <script language="javascript" src="banshow.asp?rid=5"></script>
          </td>
        </tr>
      </table>
      <p align="CENTER"></p>
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
