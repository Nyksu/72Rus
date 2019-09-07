<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->


<%
Response.Redirect("serv.html")
var sql=""
var sql1=""
var tit=""
var nikname=""
var psw=""
var psw1=""
var ErrorMsg=""
var name=""
var id=0
var ShowForm=true
var email=""
var phone=""
var city=0
var country=0
var company=""
var city_id=0
var re = new RegExp("(\\w+)@((\\w+).)*(\\w+)$")
var cw=String(Math.random()).substr(3,20)

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""	

if(cw.length<20){cw+=String(Math.random()).substr(3,(20-cw.length))}			


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

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){
	name=String(Request.Form("name"))
	email=String(Request.Form("email"))
	phone=Request.Form("pone")
	nikname=String(Request.Form("nikname"))
	psw=String(Request.Form("psw"))
	psw1=String(Request.Form("psw1"))
	city=Request.Form("city")
	area=Request.Form("region")
	country=Request.Form("country")
	company=Request.Form("company")
	
	if (!re.test(email)) {ErrorMsg+="Некорректный E-mail.<br>"}
	if (name.length<10) {ErrorMsg+="Необходимо указать полнностью ФИО!<br>"}
	if (phone != "" && phone.length<6) {ErrorMsg+="Слишком короткий телефон.<br>"}
	if (nikname.length < 3) {ErrorMsg+="Слишком короткое имя пользователя.<br>"}
	if (psw.length < 6) {ErrorMsg+="Слишком короткий пароль! Необходимо не менее 6-и символов!<br>"}
	if (psw != psw1) {ErrorMsg+="Пароли не совпадают!  Повторите ввод паролей!<br>"}
	if (String(city).length < 2) {ErrorMsg+="Слишком короткое наименование города.<br>"}
	while (String(company).indexOf("'")>-1) {company=String(company).replace("'","\"")}
	
	
   sql="Select * from HOST_URL where Upcase(nikname)=Upcase('"+nikname+"')"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Такое имя пользователя занято! Попробуте зарегестрировать другое!<br>"}
   Records.Close()
   sql="Select * from HOST_URL where Upper(email)=Upper('"+email+"')"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Такой адрес электронной почты уже зарегистрирован!<br>"}
   Records.Close()
   if (String(company).length>0) {
   sql="Select * from HOST_URL where Upcase(company)=Upcase('"+company+"')"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Такое юридическое лицо уже зарегистрировано!<br>"}
   Records.Close()
   }
   sql1=""
   sql="Select * from CITY where Upcase(name)=Upcase('"+city+"')"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF) {city_id=Records("ID").Value} else {city_id=-1}
   Records.Close()

   
   if (ErrorMsg=="") {
		//ShowForm=false
		sql=insurlhost
		id=NextID("universal")
		if (city_id<0) {
		city_id=NextID("cityid")
		sql1="Insert into CITY (ID,NAME) values ( %ID, '%NAME' )"
		sql1=sql1.replace("%ID",city_id)
		sql1=sql1.replace("%NAME",city)
		}
		sql=sql.replace("%ID",id)
		sql=sql.replace("%EMAIL",email)
		sql=sql.replace("%NAME",name)
		sql=sql.replace("%PHONE",phone)
		sql=sql.replace("%NIK",nikname)
		sql=sql.replace("%PS",psw)
		sql=sql.replace("%CW",cw)
		sql=sql.replace("%CITY",city_id)
		sql=sql.replace("%AREA",area)
		sql=sql.replace("%LAND",country)
		sql=sql.replace("%COMP",company)
		
		Connect.BeginTrans()
		try{
			if (sql1 != "") {Connect.Execute(sql1)}
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
		   Session("is_rite_connect_url")=1
		   Session("id_host_url")=id
		}
   }
}

%>

<Html>
<Head>
<Title><%=tit%>Регистрация Участника 72RUS - Тюменская Область</Title>
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
        <a href="messages.asp">Объявления</a> | <a href="http://chat.72rus.ru/">Чат</a></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top"> 
    <td width="150" align="center"> 
      <!--RAX counter-->
      <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
      <!--/RAX-->
    </td>
    <td align="center"> 
      <p align="center">&nbsp;</p>
      <p align="center"><font size="3"><b>Регистрация пользователя 72RUS.RU</b></font></p>
	  
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%>

        <%if(ShowForm){%>

<form name="form1" method="post" action="regmemurl.asp">
        <p>Параметры, помеченные знаком <b><font color="#FF0000" size="3">*</font></b> 
          обязательны к заполнению!</p>
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
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Ф.И.О. *&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <div align="left"> 
                <p>&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="name" size="35" maxlength="99" value="<%=isFirst?"":Request.Form("name")%>">
                  *</p>
              </div>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>E - mail *&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <div align="left"> 
                <p>&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="email" size="35" maxlength="40" value="<%=isFirst?"":Request.Form("email")%>">
                  *</p>
              </div>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Телефон&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <div align="left"> 
                <p>&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="phone" size="35" maxlength="40" value="<%=isFirst?"":Request.Form("phone")%>">
                </p>
              </div>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Имя пользователя*&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <p>&nbsp;&nbsp;&nbsp; 
                <input type="text" name="nikname" size="35" maxlength="99" value="<%=isFirst?"":Request.Form("nikname")%>">
                *</p>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Пароль*&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <p>&nbsp;&nbsp;&nbsp; 
                <input type="password" name="psw" size="15" maxlength="20" value="">
                * <font size="1">повторите пароль</font>: 
                <input type="password" name="psw1" size="15" maxlength="20">
                *</p>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Страна*&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <p>&nbsp;&nbsp;&nbsp; 
                <select name="country">
                  <%
				sql="Select * from  country where regionality=2 and rus_name is not null order by rus_name"
				Records.Source=sql
				Records.Open()
				while(!Records.EOF){
				%>
                  <option value="<%=Records("ID").Value%>"
				<%=isFirst&&(Records("ID").Value==190)?"selected":""%>
				<%=!isFirst&&(Records("ID").Value==Request.Form("country"))?"selected":""%>> 
                  <%=Records("RUS_NAME").Value%> </option>
                  <%
				Records.MoveNext()
				}
				Records.Close()
				%>
                </select>
                *</p>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Регион*&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <p>&nbsp;&nbsp;&nbsp; 
                <select name="region">
                  <%
				sql="Select * from  area where state=0 order by id"
				Records.Source=sql
				Records.Open()
				while(!Records.EOF){
				%>
                  <option value="<%=Records("ID").Value%>"
				<%=isFirst&&(Records("ID").Value==54)?"selected":""%>
				<%=!isFirst&&(Records("ID").Value==Request.Form("region"))?"selected":""%>> 
                  <%=Records("NAME").Value%> </option>
                  <%
				Records.MoveNext()
				}
				Records.Close()
				%>
                </select>
                *</p>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Город* &nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <p>&nbsp;&nbsp;&nbsp; 
                <input type="text" name="city" size="35" maxlength="99" value="<%=isFirst?"":Request.Form("city")%>">
                *</p>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%"> 
              <div align="right"> 
                <p>Компания&nbsp;&nbsp;</p>
              </div>
            </td>
            <td> 
              <div align="left"> 
                <p>&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="company" size="35" maxlength="100" value="<%=isFirst?"":Request.Form("company")%>">
                </p>
              </div>
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <p>
          <input type="submit" name="Submit" value="Зарегистрировать">
        </p>
      </form>
<%} else {%>
	  <p><font color="#FF3300" size="3">Пользователь зарегистрирован!</font></p>
      <center>
        <p><a href="usrarea.asp">Вход в кабинет пользователя</a></p>
        <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную 
          страницу</a></font></p>
	<p><font face="Arial, Helvetica, sans-serif"><a href="catarea.asp">В каталог</a></font></p>
	</center>
      <p><%}%>
      </p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <hr noshade size="4">
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
