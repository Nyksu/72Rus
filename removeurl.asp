<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\count_url.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<%
var hid=parseInt(Request("hid"))
var sql=""
var hiadr=""
var tit=""
var hname=""
var tekhia=0
var urlcount=0
var endlist=0
var urlname=""
var urlabout=""
var st=parseInt(Request("st"))
var daterenew=""
var urladr=""
var urlid=parseInt(Request("url"))
var dd=0
var cu=0
var isok=true

if (isNaN(hid)) {hid=0}
if (isNaN(urlid)) {urlid=0}
if (urlid==0) {Response.Redirect("admarea.asp")}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="removeurl.asp?url="+urlid+"&hid="+hid
Response.Redirect("login.asp")
}

sql="Select t1.* from url t1, catarea t2 where t1.id="+urlid+" and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
} else {
	Records.Close()
	Response.Redirect("admarea.asp")
}
Records.Close()

if (isNaN(st)) {st=0}

if (st==1) {
	sql="Update url set catarea_id="+hid+" where id="+urlid
	Connect.BeginTrans()
	try{
		Connect.Execute(sql)
	}
		catch(e){
		Connect.RollbackTrans()
		isok=false
		st=0
	}
	if (isok){
		Connect.CommitTrans()
	}
}

%>
<Html>
<Head>
<Title>Перемещение ссылки - 72RUS - Тюмень - Региональный портал</Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">


<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr bgcolor="#6699CC"> 
    <td width="150" valign="middle" height="19"> 
      <div align="center"> <font face="Arial, Helvetica, sans-serif" size="-2"><b><a href="/"><font size="2" color="#FFFFFF">НА 
        ГЛАВНУЮ</font></a></b></font></div>
    </td>
    <td bordercolor="#FFCC00" height="19" width="150"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><a href="catarea.asp" target="_blank"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><font size="2" color="#FFFFFF">КАТАЛОГ</font></b></font></a></font></div>
    </td>
    <td height="19">&nbsp; </td>
  </tr>
</table>
<br>
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
          <td bgcolor=#CCCCCC width="100%" height="7"> 
            <p><font 
                  face=Verdana size=1><b> Перемещаем URL:</b></font></p>
            </td>
        </tr>
        </tbody> 
      </table>
      <h2 align="left"> <b><a href="<%=urladr%>"> <%=urlname%></a></b> - <%=urlabout%></h2>
	  <table width="100%">
        <tr> 
          <td bgcolor="#F7F7F7" height="2"> 
            <div align="left"> <font face="Verdana" size="1"><b>Вы здесь</b></font> 
              <font size="2"> 
              <%
	  tekhia=hid
	  while (tekhia!=0) {
	     sql="Select * from catarea where id="+tekhia
		 Records.Source=sql
		 Records.Open()
		 if (!Records.EOF) {
			hname=String(Records("NAME").Value)
			hiadr="removeurl.asp?hid="+tekhia+"&url="+urlid
			tekhia=Records("HI_ID").Value
			if (hiadr != "catarea.asp?hid="+hid) {
		%>
              &lt; <a href="<%=hiadr%>"><%=hname%></a> 
              <%
		  	} else {
		%>
              &lt;<font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b> 
              <%=hname%> 
              <%
		  	}	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
              </b></font> &lt; <a href="removeurl.asp?hid=0&url=<%=urlid%>"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Каталог</b></font></a> 
              <% } else {%>
              &lt;<a href="index.asp"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">На 
              главную страницу</font></b></a> 
              <%} %>
              </font> </div>
			</td>
          </tr>
      </table>
      <% if (st!=1)  { %>
      <% if (hid>0) {%>
      <p><a href="removeurl.asp?url=<%=urlid%>&hid=<%=hid%>&st=1"><b><br>
        Разместить сайт в текущий каталог</b></a></p>
		<p>&nbsp;</p>
<% } 
	sql="Select * from catarea where hi_id="+hid+" and catalog_id="+catalog+" order by name"
	Records.Source=sql
	Records.Open()
	while (!Records.EOF) {
		hname=String(Records("NAME").Value)
		tekhia=Records("ID").Value 
		if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog) {cu=Count_url(5,tekhia)}
		else {cu=Count_url(1,tekhia)}
		Records.MoveNext() %>
        
      <p align="left">&gt;<b><a href="removeurl.asp?url=<%=urlid%>&hid=<%=tekhia%>"> 
        <%=hname%></a></b> (<%=cu%>) <b>&lt;---</b> <a href="removeurl.asp?url=<%=urlid%>&hid=<%=tekhia%>&st=1">Перенести 
        сюда!</a></p>
        <%	}  Records.Close()
} else {%>
      <p></p>
      <p><b><font color="#0000CC">Сайт перемещен в текущую директорию!! </font></b></p>
      <b><font color="#0000CC">
      <FORM>
        <div align="center"> 
          <INPUT TYPE="button" VALUE="Закрыть окно" onClick=opener.location.reload();;window.parent.close()>
          <br>
        </div>
      </FORM>
      </font></b>
      <% } %>
    </td>
    <td width="150"> 
      <div align="left"> </div>
      <p align="left">&nbsp;</p>
      <p align="left">&nbsp;</p>
    </td>
  </tr>
</table>
<br>
<br>
<table width="100%" cellspacing="0" cellpadding="0" border="0">
  <tr> 
    <td width="150" valign="middle" align="center">&nbsp;</td>
    <td bgcolor="#F7F7F7"> 
      <div align="center"><a href="admarea.asp"><b><font size="2">Кабинет</font></b></a> 
        &nbsp;| &nbsp;<a href="catarea.asp"><b><font size="2">Каталог</font></b></a></div>
    </td>
    <td width="150" valign="middle" align="center">&nbsp;</td>
  </tr>
</table>
</Body>
</Html>
