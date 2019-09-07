<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var uid=parseInt(Request("nik"))
var sql=""
var name=""
var email=""
var phone=""
var nikname=""
var city=0
var area=""
var country=0
var company=""
var psw=""
var cw=""
var url=""
var about=""
var id=0
var ShowForm=true
var hid=0
var statname=""
var idh=0
var Recs=CreateRecordSet()
var hname=""


if (isNaN(uid)) {Response.Redirect("admarea.asp")}

sql="Select t1.*, t2.name as area, t3.name as city, t4.rus_name as land from host_url t1, area t2, city t3, country t4 where t1.area_id=t2.id and t1.city_id=t3.id and t1.country_id=t4.id and t1.id="+uid
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
city=Records("CITY").Value
area=Records("AREA").Value
country=Records("LAND").Value
company=Records("COMPANY").Value
psw=Records("PSW").Value
cw=Records("KEYWORD").Value
Records.Close()

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="shownik.asp?nik="+uid
Response.Redirect("login.asp")
}

if (company.length>0) {while (company.indexOf("\"")>-1) {company=company.replace("\"","&quot;")}}

%>

<Html>
<Head>
<Title>Реквизитов пользователя сервера 72RUS - Тюменская Область</Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">


<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> |</p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#0000FF" size="4">Пользователь WEB каталога</font></b></p>
<p align="center">&nbsp;</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left"> 
      <div align="left"> </div>
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
                  face=Verdana size=1><b> Пользователь</b></font></td>
        </tr>
        </tbody> 
      </table>
      <h1 align="left"><font size="3">Реквизиты администратора сайта: <%=nikname%></font></h1>
	  <p align="left">&nbsp;</p>
      <p align="left">ФИО: <b><%=name%></b></p>
      <p align="left">E-Mail: <b><%=email%></b></p>
      <p align="left">Телефон: <b><%=phone%></b></p>
      <p align="left">Компания: <b><%=company%></b></p>
      <p align="left">Страна: <b><%=country%></b></p>
	  <p align="left">Регион: <b><%=area%></b></p>
	  <p align="left">Город: <b><%=city%></b></p>
	  <p align="left">PSW: <b><%=psw%></b></p>
	  <p align="left">CW: <b><%=cw%></b></p>
	  <p align="left">&nbsp;</p>
      <div align="center">
<% 
	sql="Select * from url where  host_url_id="+uid
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
        <table width="90%" border="1" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
          <tr> 
            <td width="60%" height="26" bgcolor="#ffd34e"> 
              <p><%=name%><br>
                <font color="#0000FF"><%=url%></font></p>
            </td>
            <td height="26" bgcolor="#ffd34e"> 
              <div align="center"> 
                <p>(<font color="#FF0000"><%=statname%></font>)<br>
                   <font size="2"><a href="urlresume.asp?url=<%=id%>&st=1">разместить</a>&nbsp;&nbsp; 
                  <a href="delurl.asp?url=<%=id%>">удалить</a>&nbsp;&nbsp; <a href="edurl.asp?url=<%=id%>">изменить</a>&nbsp;&nbsp; 
                  <a href="urlresume.asp?url=<%=id%>&st=3">остановить</a>&nbsp;&nbsp; 
                  <a href="removeurl.asp?url=<%=id%>&hid=<%=hid%>">переместить</a>&nbsp;&nbsp; 
                  <a href="copyurl.asp?url=<%=id%>&hid=<%=hid%>">скопировать</a></font>
                </p>
              </div>
            </td>
          </tr>
          <tr> 
            <td colspan="2" height="14"> 
              <p><font color="#663399" size="1"><%=about%></font></p>
            </td>
          </tr>
          <tr> 
            <td colspan="2" bordercolor="#ffd34e" height="14"> 
              <p><b>Путь размещения:</b> <font color="#996600"> 
                <%
			  idh=hid
			  while (idh>0) {
			  sql="Select * from catarea where id="+idh
			  Recs.Source=sql
			  Recs.Open()
			  if (!Recs.EOF) {
			  	idh=Recs("HI_ID").Value
				hname=Recs("NAME").Value
			  %>
                &nbsp;<%=hname%>&nbsp;&gt; 
                <% }
				else {idh=0}
				Recs.Close()
				} 
				%>
                </font></p>
            </td>
          </tr>
        </table>
        <p><font color="#996600">--------------------------------------------------------------------</font> 
          <%} Records.Close()%>
        </p>
        <p>&nbsp;</p>
     </div>
      <p align="left">&nbsp;</p>
      <p align="left">&nbsp;</p>
	  <p><font face="Arial, Helvetica, sans-serif"><a href="admarea.asp">Вернуться 
        в кабинет</a></font></p>
      <p><font face="Arial, Helvetica, sans-serif"><a href="admarea.asp?usr=1">Вернуться 
        к списку пользователей</a></font></p>
      <p>&nbsp; </p>
    </td>
    <td width="150"> 
      <div align="left">
        <p>&nbsp;</p>
      </div>
      </td>
  </tr>
</table>
</Body>
</Html>
