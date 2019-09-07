<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\Creaters.inc" -->

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
var id=parseInt(Request("nik"))
var re = new RegExp("(\\w+)@((\\w+).)*(\\w+)$")			


if (isNaN(id)) {Response.Redirect("admarea.asp")}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="edusrurl.asp?nik="+id
Response.Redirect("login.asp")
}

sql="Select * from host_url where id="+id
Records.Source=sql
Records.Open()
if (Records.EOF) {
   Records.Close()
   Response.Redirect("admarea.asp")
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


<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" bgcolor="#EBF5ED"> 
      <div align="left"></div>
    </td>
    <td> 
      <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="2"><font 
                  face=Verdana size=1><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font> 
            <font 
                  face=Verdana size=1><b> <a href="index.html">72rus.ru </a></b></font><font 
                  face=Verdana size=1><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font> 
            <font 
                  face=Verdana size=1><b> <a href="catarea.asp?hid=0">Каталог</a></b></font> 
            <font size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font></font><font 
                  face=Verdana size=1><b> Регистрация</b></font></td>
        </tr>
        </tbody> 
      </table>
      <h1 align="left"><font size="3">Реквизиты администратора сайта: <%=nikname%></font></h1>
	  
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%>
<p>&nbsp;</p>
<%if(ShowForm){%>
      <form name="form1" method="post" action="edmemurl.asp">
        <p> 
          <input type="hidden" name="nik" value="<%=id%>">
          Параметры, помеченные знаком <b><font color="#FF0000" size="3">*</font></b> 
          обязательны к заполнению!</p>
        <table width="90%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
          <tr> 
            <td width="37%" bgcolor="#000066"> 
              <div align="center"><font color="#FFFF00"><b>Параметры</b></font>:</div>
            </td>
            <td width="6%">&nbsp; </td>
            <td bgcolor="#000066"> 
              <div align="center"><font color="#FFFF00"><b>Значения</b></font></div>
            </td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#CCCCFF"> 
              <div align="right">Ф.И.О. *&nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#CCCCFF"> 
              <div align="left">&nbsp;&nbsp;&nbsp; 
                <input type="text" name="name" size="35" maxlength="99" value="<%=name%>">
                *</div>
            </td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#99CCFF"> 
              <div align="right">E - mail *&nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#99CCFF"> 
              <div align="left">&nbsp;&nbsp;&nbsp; 
                <input type="text" name="email" size="35" maxlength="40" value="<%=email%>">
                *</div>
            </td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#CCCCFF"> 
              <div align="right">Телефон&nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#CCCCFF"> 
              <div align="left">&nbsp;&nbsp;&nbsp; 
                <input type="text" name="phone" size="35" maxlength="40" value="<%=phone%>">
              </div>
            </td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#99CCFF"> 
              <div align="right">Страна*&nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#99CCFF">&nbsp;&nbsp;&nbsp; 
              <select name="country">
                <%
				sql="Select * from  country where regionality=2 and rus_name is not null order by rus_name"
				Records.Source=sql
				Records.Open()
				while(!Records.EOF){
				%>
                <option value="<%=Records("ID").Value%>"
				<%=Records("ID").Value==country?"selected":""%>> 
                <%=Records("RUS_NAME").Value%> </option>
                <%
				Records.MoveNext()
				}
				Records.Close()
				%>
              </select>
              *</td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#CCCCFF"> 
              <div align="right">Регион*&nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#CCCCFF">&nbsp;&nbsp;&nbsp; 
              <select name="region">
                <%
				sql="Select * from  area where state=0 order by id"
				Records.Source=sql
				Records.Open()
				while(!Records.EOF){
				%>
                <option value="<%=Records("ID").Value%>"
				<%=Records("ID").Value==area?"selected":""%>> 
                <%=Records("NAME").Value%> </option>
                <%
				Records.MoveNext()
				}
				Records.Close()
				%>
              </select>
              *</td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#99CCFF"> 
              <div align="right">Город* &nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#99CCFF">&nbsp;&nbsp;&nbsp; 
              <input type="text" name="city" size="35" maxlength="99" value="<%=city%>">
              *</td>
          </tr>
          <tr> 
            <td width="37%" bgcolor="#CCCCFF"> 
              <div align="right">Компания&nbsp;&nbsp;</div>
            </td>
            <td width="6%">&nbsp;</td>
            <td bgcolor="#CCCCFF"> 
              <div align="left">&nbsp;&nbsp;&nbsp; 
                <input type="text" name="company" size="35" maxlength="100" value="<%=company%>">
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
	    <p><font face="Arial, Helvetica, sans-serif"><a href="admarea.asp?usr=1">В 
          кабинет</a></font></p>
	</center>
<%}%>
    </td>
    <td width="150" bgcolor="#EBF5ED"> 
      <div align="left"> </div>
      <p align="left">&nbsp;</p>
      <p align="left">&nbsp;</p>
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
