<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var  id=parseInt(Request("bid"))
var dcl=parseInt(Request("dd"))
var ddd=parseInt(Request("ddd"))
if (isNaN(id)) {id=0}
if (isNaN(dcl)) {dcl=0}
if (isNaN(ddd)) {ddd=30}
var sql=""
var name=""
var coment=""
var bid=0
var ash=0
var ush=0
var ii=0
var ddt=""
var psw=""
var pathh=""
var btype=""
var wi=0
var he=0
var fc=""
var bnam=""

if ((Session("is_adm_mem")!=1)&&(Session("banlog")!=id)) {
	Session("backurl")="banstat.asp?dd="+dcl+"&ddd="+ddd
	Response.Redirect("logban.asp?bid="+id)
}


var dtt=""

if (id!=0) {
	sql="Select * from baner where state=0 and id="+id
	Records.Source=sql
	Records.Open()
	if (!Records.Eof) {
		name=Records("NAME").Value
		coment=Records("COMENT").Value
		psw=Records("PSW").Value
		pathh=Records("PATHNAME").Value
		btype=Records("SOURSTYPE").Value
		wi=Records("WIDTH").Value
		he=Records("HEIGHT").Value
		fc=Records("FLASHCOD").Value
	} else {id=0}
	Records.Close()
}
%>

<html>
<head>
<title>Статистика по банерам</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<style type="text/css">
<!--
body {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #333333;
	background-color: #FFFFFF;
}
td {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #333333;
}
.breadcrumb {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #333333;
}
a {
	color: #666666;
	text-decoration:underline;
}
a:hover {
	color: #000000;
	text-decoration:none;
}
.CatTop {
	font-size:13px;
}
A.CatTop {
	font-size:13px;
}
A:hover.CatTop {
	font-size:13px;
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<%
if (Session("is_adm_mem")==1) {
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> | 
        <a href="modrec.asp"><font size="2">Управление рекламными местами</font></a><font size="2"> 
        | <a href="recstat.asp">Статистика по рекламным местам</a> | <a href="addrec.asp">Добавить 
        рекламмное место</a></font> <font size="2">| </font></p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#0000FF" size="4">Баннеры</font></b></p>
<p>| + <font size="2"><a href="addbaner.asp">Добавить банер</a></font> | 
  <%if (id>0) {%>
  <a href="banstat.asp">Вернуться к списку банеров</a> 
  <%}%>
</p>
<%
}
%>
<%
if (dcl>0) {
%>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#0000FF" size="4">Статистика переходов</font></b></p>
<p align="center"><a href="/">На главную</a> | <a href="banstat.asp?bid=<%=id%>">Статистика 
  по показам банера</a></p>
<%
} else { 
	if (id>0) {
%>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#0000FF" size="4">Cтатистика показов</font></b></p>
<p align="center"><a href="/">На главную</a> | <a href="banstat.asp?bid=<%=id%>"></a>Статистика 
  переходов по банеру за: <a href="banstat.asp?bid=<%=id%>&dd=1">1</a> | <a href="banstat.asp?bid=<%=id%>&dd=2">2</a> 
  | <a href="banstat.asp?bid=<%=id%>&dd=5">5</a> | <a href="banstat.asp?bid=<%=id%>&dd=7">7</a> 
  | <a href="banstat.asp?bid=<%=id%>&dd=10">10</a> | <a href="banstat.asp?bid=<%=id%>&dd=30">30</a> 
  | последних дней</p>
<%
	}
}
%>
<%
if (id==0) {
	// Список рекламмных банеров
%>
	
<p><b>Таблица активных рекламмных банеров</b></p>
<%
	sql="Select * from baner where state=0"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		name=Records("NAME").Value
		coment=Records("COMENT").Value
		bid=Records("ID").Value
		ii=ii+1
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td width="0"> 
      <p><font size="1"><b>&nbsp;&nbsp;<%=ii%>.</b></font></p>
    </td>
    <td> 
      <p><font size="1" color="#0033CC"><b><font size="2"><a href="banstat.asp?bid=<%=bid%>"><%=name%></a></font><br>
        <%if (coment!="") {%>
        ( </b><font color="#666666"><%=coment%></font><b> )</b> 
        <%}%>
        </font></p>
    </td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
} else {
	// показы банера
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#FFFFFF"> 
    <td width="0"> 
      <p>&nbsp;</p>
    </td>
    <td align="center"> 
      <p><font size="3" color="#000099"><b><%=name%></b></font><br>
        <%if (coment!="") {%>
        <font size="2">( </font><font color="#666666" size="2"><%=coment%></font><font size="2"> 
        )</font> 
        <%}%>
      </p>
    </td>
  </tr>
</table>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td rowspan="2" width="0">&nbsp; 
      <%if (wi<=he){%>
      <%if (btype==1) { // Это имидж!! %>
      <img src="<%=pathh%>" width="<%=wi%>" height="<%=he%>" border="0"> 
      <%}%>
      <%if (btype==2) { // Это Флэшь!! %>
      <object classid="clsid:<%=fc%>" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="<%=wi%>" height="<%=he%>">
			<param name=movie value="<%=pathh%>"><param name=quality value=high>
			<embed src="<%=pathh%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="<%=wi%>" height="<%=he%>">
			</embed>
      </object> 
      <%}%>
      <% 
		if (btype==3) { // Это Джава-скрипт
			var ts=""
			var fs= new ActiveXObject("Scripting.FileSystemObject")
			var filnam=""
			var str=""
			filnam=JSPath+id+".jsf"
			if (!fs.FileExists(filnam)) { filnam="" }
			if (filnam!="") {
				ts= fs.OpenTextFile(filnam)
				str=ts.ReadAll()
		%>
      <%=str%> 
      <%
				
			}
		}
		%>
      <%}%>
    </td>
    <td bordercolor="#CCCCCC" align="center"> 
      <%if (wi>=he){%>
      <%if (btype==1) { // Это имидж!! %>
      <img src="<%=pathh%>" width="<%=wi%>" height="<%=he%>" border="0"> 
      <%}%>
      <%if (btype==2) { // Это Флэшь!! %>
      <object classid="clsid:<%=fc%>" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="<%=wi%>" height="<%=he%>">
			<param name=movie value="<%=pathh%>"><param name=quality value=high>
			<embed src="<%=pathh%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="<%=wi%>" height="<%=he%>">
			</embed>
      </object> 
      <%}%>
      <% 
		if (btype==3) { // Это Джава-скрипт
			var ts=""
			var fs= new ActiveXObject("Scripting.FileSystemObject")
			var filnam=""
			var str=""
			filnam=JSPath+id+".jsf"
			if (!fs.FileExists(filnam)) { filnam="" }
			if (filnam!="") {
				ts= fs.OpenTextFile(filnam)
				str=ts.ReadAll()
		%>
      <%=str%> 
      <%
				
			}
		}
		%>
      <%}%>
    </td>
  </tr>
  <tr> 
    <td bordercolor="#CCCCCC" valign="top"> 
      <div align="center">
<%
	if (dcl==0) {
%>
        <p><b>Статистика банера по рекламным местам за 30 дней</b></p>
		<table width="99%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
          <tr bordercolor="#CCCCCC" bgcolor="#E5E5E5"> 
            <td width="35"> 
              <div align="center"> 
                <p><b><font color="#0000CC">П.п.</font></b></p>
              </div>
            </td>
            <td width="60"> 
              <div align="center"> 
                <p><b><font color="#0000CC">Дата</font></b></p>
              </div>
            </td>
            <td> 
              <div align="center"> 
                <p><b><font color="#0000CC">Рекламные места </font></b></p>
              </div>
            </td>
            <td width="80"> 
              <div align="center"> 
                <p><b><font color="#0000CC">Хиты</font></b></p>
              </div>
            </td>
            <td width="80"> 
              <div align="center"> 
                <p><b><font color="#0000CC">Посетители</font></b></p>
              </div>
            </td>
          </tr>
        </table>
<%
	sql="Select t1.id, t1.dt, t2.name, t1.allshow, t1.unishow from banblockst t1, banblock t2 where t1.baner_id="+id+" and t1.banblock_id=t2.id and t1.dt>'TODAY'-"+ddd+" order by t1.dt desc, t2.name"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		ash=Records("ALLSHOW").Value
		ush=Records("UNISHOW").Value
		ddt=Records("DT").Value
		bnam=Records("NAME").Value
		ii=ii+1
%>
	    <table width="99%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
          <tr bordercolor="#CCCCCC"> 
            <td width="35"> 
              <p><font size="1"><b>&nbsp;&nbsp;<%=ii%>.</b></font></p>
            </td>
            <td width="60"> 
              <p><font size="2" color="#0033CC"><%=ddt%></font></p>
            </td>
            <td> 
              <p><font size="2" color="#0033CC"><%=bnam%></font></p>
            </td>
            <td width="80"> 
              <div align="center"> 
                <p><font size="2" color="#0000CC"><%=ash%></font></p>
              </div>
            </td>
            <td width="80"> 
              <div align="center"> 
                <p><font size="2" color="#0000CC"><%=ush%></font></p>
              </div>
            </td>
          </tr>
        </table>
<%
		Records.MoveNext()
	}
	Records.Close()
%>
<%
	} else {
	// статистика по кликам банера
%>
        <p><b>Статистика переходов по банеру</b></p>
        <table width="100%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
          <tr bgcolor="#E1F2EC" bordercolor="#D8D8D8"> 
            <td width="200"> 
              <div align="center"> 
                <p><b>Время перехода</b></p>
              </div>
            </td>
            <td width="150" bgcolor="#E1F2EC"> 
              <div align="center"> 
                <p><b>IP адрес</b></p>
              </div>
            </td>
            <td bgcolor="#E1F2EC"> 
              <div align="center"> 
                <p><b>Со страницы</b></p>
              </div>
            </td>
          </tr>
          <%
  var olddt=""
  var ipa=""
  var ura=""
  var dtr=""
  Records.Source="Select * from BANCLIKS where baner_id="+id+" and dtr>'TODAY'-"+dcl+" order by dt desc"
  Records.Open()
  while (!Records.EOF) {
  	dtt=Records("DT").Value
	ipa=Records("IP_ADR").Value
	ura=Records("URL").Value
	dtr=Records("DTR").Value
  %>
          <%if (String(dtr)!=String(olddt)) {%>
          <tr bgcolor="#F3F3F3" bordercolor="#D8D8D8"> 
            <td colspan="3"> 
              <div align="center"> 
                <p><b><font size="3" color="#0000FF">На дату: <font color="#000099"><%=dtr%></font></font></b></p>
              </div>
            </td>
          </tr>
          <%}%>
          <tr bordercolor="#D8D8D8"> 
            <td width="200"> 
              <p><font size="2" color="#0033FF"><%=dtt%></font></p>
            </td>
            <td width="150"> 
              <p><font size="2" color="#0033FF"><%=ipa%></font></p>
            </td>
            <td> 
              <p><font size="2" color="#0033FF"><%=ura%></font></p>
            </td>
          </tr>
          <%
  	olddt=dtr
  	Records.MoveNext()
  }
  Records.Close()
  %>
        </table>
<%
	}
%>
	</div>
    </td>
  </tr>
</table>
<%
}
%>
</body>
</html>
