<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\email.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->


<%
var ErrorMsg=""
var text=""
var sql=""
var email=""
var name=""
var tit=""
var nikname=""
var company=""
var uid=0
var Text=""

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""

if (Session("is_rite_connect_url")==1) {
	sql="Select * from host_url where id="+Session("id_host_url")
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){
		name=Records("NAME").Value
		nikname=Records("NIKNAME").Value
		email=Records("EMAIL").Value
		company=Records("company").Value
		uid=Session("id_host_url")
   } else {
		Records.Close()
		Session("is_rite_connect_url")=0
		Response.Redirect("usrarea.asp")
   }
   Records.Close()
}
else {Response.Redirect("usrarea.asp")}

var eml=Server.CreateObject("JMail.Message")
eml.Logging=false
eml.From=fromaddres
eml.AddRecipient(Recipient)
eml.Charset=characterset
eml.ContentTransferEncoding = "base64"

var isSending=false

isFirst=String(Request.Form("Submit"))=="undefined"
ShowForm=true
if(!isFirst){
	
		//-------------input validation-----------
		tit=TextFormData(Request.Form("Name"),"")
		Text=TextFormData(Request.Form("txt"),"")


		if(Text.length>2000){ErrorMsg+="Сообщение превышает допустимый размер.<br>"}
		if(Text.length<4){ErrorMsg+="Сообщение отсутствует.<br>"}
		if(tit.length<3){ErrorMsg+="Укажите тему сообщения!<br>"}
		

		if (ErrorMsg==""){
			
			try{
				// POP before SMTP ++++++++++++++++++++++
				var  pop3=Server.CreateObject("JMail.POP3")
				pop3.Connect(fromaddres,pswsmtp,servsmtp)
				pop3.Disconnect()
				// POP before SMTP --------------------------------------
				eml.FromName=fromname+nikname
				eml.Subject=tit+" (72RUS)"
				eml.AppendText(" Поддержка 72RUS (каталоги) : "+Subject+"\n")
				eml.AppendText(" Ф.И.О. : "+name+"\n")
				eml.AppendText(" Компания : "+company+"\n")
				eml.AppendText(" Email : "+email+"\n")
				eml.AppendText("\n Сообщение : \n \n")
				eml.AppendText(TextFormData(Request.Form("txt")))
				isSending=eml.Send(servsmtp)
	 			if (isSending) {ShowForm=false}
			}
			catch(e){
				ErrorMsg+="Проблемы с почтой.<br>"
			}
		} else { ShowForm=true }
}
%>
<Html>
<head>
<Title>Служба поддержки пользователей - 72RUS - Тюмень - Тюменская Область</Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72.css" type="text/css">
</head>
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
    <td width="150" align="left"> 
      <div align="left"></div>
    </td>
    <td> 
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td valign="top"> 
            <table width="100%" border="0" cellspacing="6" cellpadding="0" bordercolor="#CCCCFF" align="center">
              <tr bordercolor="#CCCCFF"> 
                <td bordercolor="#CCCCFF" height="11"> 
                  <p><font size="3"><b>Служба поддержки 72RUS.RU</b></font></p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <%if(ErrorMsg!=""){%>
      <center>
        <h2> 
          <p> <font color="#FF3300">Ошибка!</font> <br>
            <%=ErrorMsg%></p>
        </h2>
      </center>
      <%}%>
      <form name="Guest" method="post" action="support.asp">
        <%if(ShowForm){%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr valign="middle"> 
            <td width="120" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Тема 
              вопроса: </font></td>
            <td> 
              <input type="text" name="Name" value="<%=isFirst?"":Request.Form("Name")%>" size="30" maxlength="50">
            </td>
          </tr>
          <tr> 
            <td width="120" align="right" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Ваш 
              вопрос:</font></td>
            <td> 
              <textarea name="txt" cols="30" rows="8"><%=text%></textarea>
            </td>
          </tr>
        </table>
        <p align="center"> 
          <input type="submit" name="Submit" value="Переслать">
          <input type="reset" name="Submit2" value="Очистить">
        </p>
        <%}
else{%>
        <center>
          <h2>Спасибо!<br>
            Ваше сообщение отправлено.</h2>
          <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На 
            главную страницу</a></font></p>
        </center>
        <%}%>
      </form>
      <hr noshade size="4">
    </td>
    <td width="150"> 
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
