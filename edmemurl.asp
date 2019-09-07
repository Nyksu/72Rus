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
var sql=""
var sql1=""
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


var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""				


if (Session("is_rite_connect_url")!=1) {Response.Redirect("usrarea.asp")}

id=Session("id_host_url")
sql="Select * from host_url where id="+id
Records.Source=sql
Records.Open()
if (Records.EOF) {
   Records.Close()
   Response.Redirect("usrarea.asp")
}
name=String(Records("NAME").Value)
email=String(Records("EMAIL").Value)
phone=Records("PHONE").Value
if (String(phone)=="undefined") {phone=""}
nikname=Records("NIKNAME").Value
city=Records("CITY_ID").Value
area=Records("AREA_ID").Value
country=Records("COUNTRY_ID").Value
company=Records("COMPANY").Value
Records.Close()

sql="Select * from city where id="+city
Records.Source=sql
Records.Open()
city=Records("NAME").Value
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){
	name=String(Request.Form("name"))
	email=String(Request.Form("email"))
	phone=Request.Form("phone")
	city=Request.Form("city")
	area=Request.Form("region")
	country=Request.Form("country")
	company=Request.Form("company")
	
	if (!re.test(email)) {ErrorMsg+="Некорректный E-mail.<br>"}
	if (name.length<10) {ErrorMsg+="Необходимо указать полнностью ФИО!<br>"}
	if (phone != "" && phone.length<6) {ErrorMsg+="Слишком короткий телефон.<br>"}
	if (String(city).length < 2) {ErrorMsg+="Слишком короткое наименование города.<br>"}
	if (company.length>0) {
		while (String(company).indexOf("'")>-1) {company=String(company).replace("'","\"")}
		while (company.indexOf("  ")!=-1) {company=company.replace("  "," ")}
		while (company.indexOf(" ")==0) {company=company.replace(" ","")}
	}
	while (name.indexOf("  ")!=-1) {name=name.replace("  "," ")}
	while (name.indexOf(" ")==0) {name=name.replace(" ","")}
	
	
   sql="Select * from HOST_URL where Upper(email)=Upper('"+email+"') and id <> "+id
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="Такой адрес электронной почты уже зарегистрирован!<br>"}
   Records.Close()
   if (String(company).length>0) {
   sql="Select * from HOST_URL where Upcase(company)=Upcase('"+company+"') and id <> "+id
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
		sql=edurlhost
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
			ErrorMsg+="Ошибка обновления.<br>"
		}
		if (ErrorMsg==""){
		   Connect.CommitTrans()
		   ShowForm=false
		}
   }
}

if (company.length>0) {while (company.indexOf("\"")>-1) {company=company.replace("\"","&quot;")}}

%>

<Html>
<Head>
<Title>Изменение реквизитов пользователя сервера 72RUS - Тюменская Область</Title>
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
        | <a href="usrarea.asp">Кабинет Пользователя</a> | <a href="newshow.asp?pid=728">Разместить 
        рекламу</a></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left"> 
      <div align="left"> 
        <!--RAX counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/RAX-->
      </div>
    </td>
    <td> 
      <p align="center">&nbsp;</p>
      <p align="center"><font size="3"><b>Реквизиты пользователя: <%=nikname%></b></font></p>
      <%if(ErrorMsg!=""){%>
      <center>
        <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
          <%=ErrorMsg%></p>
      </center>
      <%}%>
      <p>&nbsp;</p>
      <%if(ShowForm){%>
      <form name="form1" method="post" action="edmemurl.asp">
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
                  <input type="text" name="name" size="35" maxlength="99" value="<%=name%>">
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
            <td bgcolor="#EEEEEE"> 
              <div align="left"> 
                <p>&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="email" size="35" maxlength="40" value="<%=email%>">
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
                  <input type="text" name="phone" size="35" maxlength="40" value="<%=phone%>">
                </p>
              </div>
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
				<%=Records("ID").Value==country?"selected":""%>> <%=Records("RUS_NAME").Value%> 
                  </option>
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
				<%=Records("ID").Value==area?"selected":""%>> <%=Records("NAME").Value%> 
                  </option>
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
                <input type="text" name="city" size="35" maxlength="99" value="<%=city%>">
                *</p>
            </td>
          </tr>
          <tr bgcolor="#EEEEEE" bordercolor="#999999"> 
            <td width="37%" height="49"> 
              <div align="right"> 
                <p>Компания&nbsp;&nbsp;</p>
              </div>
            </td>
            <td height="49"> 
              <div align="left"> 
                <p>&nbsp;&nbsp;&nbsp; 
                  <input type="text" name="company" size="35" maxlength="100" value="<%=company%>">
                </p>
              </div>
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <p> 
          <input type="submit" name="Submit" value="Сохранить">
        </p>
      </form>
      <%} else {%>
      <p>&nbsp; </p>
      <center>
        <h1><font color="#FF3300" size="3">Реквизиты сохранены!</font></h1>
        <p>&nbsp;</p>
        <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную 
          страницу</a></font></p>
        <p><font face="Arial, Helvetica, sans-serif"><a href="usrarea.asp">В кабинет</a></font></p>
      </center>
      <%}%>
      <hr size="4" noshade width="80%">
    </td>
    <td width="150"> 
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
