<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var opr=parseInt(Request("opr"))
var id=parseInt(Request("id"))
var isadm=0
if ((Session("is_adm_mem")==1) || (Session("is_host")==1)) { isadm=1 }
//if (isadm==0) {Response.Redirect("index.asp")}
if (isadm==0) {Response.Redirect("index.html")}
var ErrorMsg=""
var name=""
var coment=""
var bid=0
var ii=0
var sql=""
var tp=""
var tpp=""
var tpnam=""
var ww=""
var hh=""
var jscr=""
var path=""
var flash=""
var psw1=""
var psw2=""
var url=""
if (isNaN(id)) {id=0}
var filename=JSPath+id+".jsf"
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""


if (isNaN(opr)) {opr=0}
if (opr>0 && id==0) {opr=0}
if (opr>0 && id>0) {
	sql="Select * from baner where id="+id
	Records.Source=sql
	Records.Open()
	if (Records.Eof) {
		Records.Close()
		Response.Redirect("modbaner.asp")
	}
	name=TextFormData(Records("NAME").Value,"")
	coment=TextFormData(Records("COMENT").Value,"")
	url=TextFormData(Records("URL").Value,"")
	path=TextFormData(Records("PATHNAME").Value,"")
	flash=TextFormData(Records("flashcod").Value,"")
	ww=parseInt(Records("width").Value)
	hh=parseInt(Records("height").Value)
	tpp=parseInt(Records("sourstype").Value)
	Records.Close()
	if (tpp==3) {
		if (fs.FileExists(filename)) { 
			ts= fs.OpenTextFile(filename)
			jscr=ts.ReadAll()
			ts.Close()
		}
	}
}

var isFirst=String(Request.Form("Submit"))=="undefined"

if ((!isFirst) && (opr>0 && id>0)){
	sql=""
	if ((opr==1) && (Request.Form("kill")==666)) {
		// Удаление банера
		sql="Delete from baner where id="+id
		if (fs.FileExists(filename)) {fs.DeleteFile(filename)}
	} else { if (opr==1) {Response.Redirect("modbaner.asp")}}
	if ((opr==2) && (Request.Form("lockit")==33)) {
		// Деактивация банера
		sql="Update baner set state=1 where id="+id
	} else { if (opr==2) {Response.Redirect("modbaner.asp")}}
	if (opr==4 && Request.Form("unlockit")==333) {
		// Активация банера
		sql="Update baner set state=0 where id="+id
	} else { if (opr==4) {Response.Redirect("modbaner.asp")}}
	if (opr==3) {
		// Изменение параметров банера
		name=TextFormData(Request.Form("name"),"")
		url=TextFormData(Request.Form("url"),"")
		coment=TextFormData(Request.Form("coment"),"")
		path=TextFormData(Request.Form("path"),"")
		flash=TextFormData(Request.Form("flash"),"")
		psw1=TextFormData(Request.Form("psw1"),"")
		psw2=TextFormData(Request.Form("psw2"),"")
		ww=parseInt(Request.Form("ww"))
		hh=parseInt(Request.Form("hh"))
		tpp=parseInt(Request.Form("tpp"))
		jscr=TextFormData(Request.Form("jscr"),"")
		if (name.length>100) {ErrorMsg+="Слишком длинное наименование.<br>"}
		if (name.length<2) {ErrorMsg+="Слишком короткое наименование.<br>"}
		if (psw1.length<2) {ErrorMsg+="Слишком короткий пароль.<br>"}
		if (psw1!=psw2) {ErrorMsg+="Неверно повторен пароль.<br>"}
		if (isNaN(ww)) {ww=0}
		if (isNaN(hh)) {hh=0}
		if ((ww==0)||(hh==0)) {ErrorMsg+="Некорректные размеры банера.<br>"}
		if (isNaN(tpp)) {tpp=0}
		if ((tpp==0)||(tpp>3)) {ErrorMsg+="Некорректный тип банера.<br>"}
		if (tpp<4) {ttt=tpp}
		if (tpp==1) {
			if (path.length<5) {ErrorMsg+="Некорректный путь.<br>"}
			if (url.length<5) {ErrorMsg+="Некорректный адрес перехода.<br>"}
		}
		if (tpp==2) {
			if (path.length<5) {ErrorMsg+="Некорректный путь.<br>"}
			if (flash.length<25) {ErrorMsg+="Некорректный код Flash.<br>"}
		}
		if (tpp==3) {
			if (jscr.length<10) {ErrorMsg+="Слишком короткий скрипт.<br>"}
		}
		sql="Update baner set name='%NAME',coment='%COM',pathname='%PATH',flashcod='%FLASH',psw='%PSW',width=%WW,height=%HH,sourstype=%TT,url='%URL' "
		sql+="where id="+id
		sql=sql.replace("%COM",coment)
		sql=sql.replace("%NAME",name)
		sql=sql.replace("%PATH",path)
		sql=sql.replace("%FLASH",flash)
		sql=sql.replace("%PSW",psw1)
		sql=sql.replace("%WW",ww)
		sql=sql.replace("%HH",hh)
		sql=sql.replace("%TT",tpp)
		sql=sql.replace("%URL",url)
		if (fs.FileExists(filename)) {fs.DeleteFile(filename)}
	}
	if ((ErrorMsg=="") && (sql!="")) {
			Connect.BeginTrans()
			try{
				Connect.Execute(sql)
				if (tpp==3) {
					ts=fs.OpenTextFile(filename,2,true)
					ts.Write(jscr)
					ts.Close()
				}
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="Ошибка изменения.<br>"
			}
			if (ErrorMsg==""){
				Connect.CommitTrans()
		  		Response.Redirect("modbaner.asp")
			}
	} 
}

%>

<html>
<head>
<title>Изменить реквизиты банера.</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><a href="admarea.asp">Вернуться в кабинет</a> | <a href="index.asp">Перейти 
  на главную страницу</a></p>
<p align="center"><b><font color="#0000FF" size="4">Управление банерами.</font></b></p>
<%
if (opr==0) {
%>
<div align="center"><a href="addbaner.asp">Добавить банер</a> | <a href="banstat.asp">Статистика 
  по показам банеров</a> | <a href="modrec.asp">Управление рекламными местами</a></div>
<p>Активные банеры:</p>
<%
	sql="Select * from baner where state=0 order by name"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		name=TextFormData(Records("NAME").Value,"")
		coment=TextFormData(Records("COMENT").Value,"")
		bid=Records("ID").Value
		tp=parseInt(Records("SOURSTYPE").Value)
		if (isNaN(tp)) {tp=0}
		tpnam=""
		switch (tp) {
			case 1 : tpnam="Имидж"; break;
			case 2 : tpnam="FLASH"; break;
			case 3 : tpnam="JScript"; break;
		}
		ww=parseInt(Records("width").Value)
		hh=parseInt(Records("height").Value)
		if (isNaN(ww)) {ww=0}
		if (isNaN(hh)) {hh=0}
		ii=ii+1
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#9999FF"> 
    <td width="60"> 
      <div align="center"> <b><font size="2"><%=ii%>.</font></b></div>
    </td>
    <td><font size="1" color="#FF0000"><%=bid%> ( <%=tpnam%> )</font> <b><font size="2"><a href="banstat.asp?bid=<%=bid%>"><%=name%></a></font></b> <font size="2" color="#3300CC">( <%=coment%> )</font></td>
  </tr>
  <tr bordercolor="#9999FF"> 
    <td width="60"> 
      <div align="center"> <b><font size="1"><%=ww%>x<%=hh%></font></b></div>
    </td>
    <td><font size="1"><b></b></font><font size="1" color="#FF0000"><b>Управление:</b></font><font size="1" color="#0033CC"> 
      &nbsp;&nbsp;&nbsp;<a href="modbaner.asp?id=<%=bid%>&opr=1">удалить</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modbaner.asp?id=<%=bid%>&opr=2">заблокировать</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modbaner.asp?id=<%=bid%>&opr=3">изменить 
      реквизиты</a></font></td>
  </tr>
  <tr bordercolor="#9999FF"> 
    <td colspan="2" height="8"> </td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
%>
<%=ii==0?"<p><font size=\"2\"><b><font color=\"#FF6633\">Нет активных банеров.</font></b></font></p>":""%> 
<p>не активные банеры:</p>
<%
	sql="Select * from baner where state=1 order by name"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		name=TextFormData(Records("NAME").Value,"")
		coment=TextFormData(Records("COMENT").Value,"")
		bid=Records("ID").Value
		tp=parseInt(Records("SOURSTYPE").Value)
		if (isNaN(tp)) {tp=0}
		tpnam=""
		switch (tp) {
			case 1 : tpnam="Имидж"; break;
			case 2 : tpnam="FLASH"; break;
			case 3 : tpnam="JScript"; break;
		}
		ww=parseInt(Records("width").Value)
		hh=parseInt(Records("height").Value)
		if (isNaN(ww)) {ww=0}
		if (isNaN(hh)) {hh=0}
		ii=ii+1
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td width="60"> 
      <div align="center"> <b><font size="2"><%=ii%>.</font></b></div>
    </td>
    <td><font size="1" color="#FF0000"><%=bid%> ( <%=tpnam%> )</font> <b><font size="2"><a href="banstat.asp?bid=<%=bid%>"><%=name%></a></font></b> <font size="2" color="#3300CC">( <%=coment%> )</font></td>
  </tr>
  <tr bordercolor="#CCCCCC"> 
    <td width="60"> 
      <div align="center"> <b><font size="1"><%=ww%>x<%=hh%></font></b></div>
    </td>
    <td><font size="1"><b></b></font><font size="1" color="#FF0000"><b>Управление:</b></font><font size="1" color="#0033CC"> 
      &nbsp;&nbsp;&nbsp;<a href="modbaner.asp?id=<%=bid%>&opr=1">удалить</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modbaner.asp?id=<%=bid%>&opr=4">активировать</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modbaner.asp?id=<%=bid%>&opr=3">изменить 
      реквизиты</a></font></td>
  </tr>
  <tr bordercolor="#9999FF"> 
    <td colspan="2" height="8"> </td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
%>
<%=ii==0?"<p><font size=\"2\"><b><font color=\"#FF6633\">Нет не активных банеров.</font></b></font></p>":""%> 
<%
} else {
%>
<table width="100%" border="1" cellspacing="2" cellpadding="0">
  <tr>
    <td bgcolor="#CCCCCC">
      <p><a href="modbaner.asp">Вернуться к списку банеров</a></p>
<% 
if (ErrorMsg!="") {
%>
<div align="center">
  <table width="90%" border="1" cellspacing="1" cellpadding="1" bordercolor="#FF0000" bgcolor="#FFFFCC">
    <tr> 
      <td> 
        <p align="center"><b><font color="#FF0000" size="2">Внимание! Замечены следующие ошибки:</font></b></p>
        <p align="center"><font color="#FF0000" size="2"><%=ErrorMsg%></font></p>
      </td>
    </tr>
  </table>
</div>
<%
}
%>
	<form name="form1" method="post" action="modbaner.asp">
        <input type="hidden" name="id" value="<%=id%>">
        <input type="hidden" name="opr" value="<%=opr%>">
<%
	if (opr==1) {
%>
	    <p><b>Удаление банера. <font color="#0000CC"><%=name%></font></b></p>
        <p>
          <input type="checkbox" name="kill" value="666">
          Да! Убить!</p>
<%
	}
%>
<%
	if (opr==2) {
%>
	    <p><b>Деактивировать банер. <font color="#0000CC"><%=name%></font></b></p>
        <p>
          <input type="checkbox" name="lockit" value="33">
          Да! Стопорнуть!</p>
<%
	}
%>
<%
	if (opr==4) {
%>
	    <p><b>Активировать банер. <font color="#0000CC"><%=name%></font></b></p>
        <p>
          <input type="checkbox" name="unlockit" value="333">
          Да! Активировать!</p>
<%
	}
%>
<%
	if (opr==3) {
%>
    <table width="90%" border="2" cellspacing="2" cellpadding="1" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#0099FF" bgcolor="#CCCCFF" width="200"> 
          <div align="center"><b>Параметры</b></div>
        </td>
        <td width="15"> 
          <div align="center"><b></b></div>
        </td>
        <td bordercolor="#0099FF" bgcolor="#CCCCFF"> 
          <div align="center"><b>Значения</b></div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><font color="#FF0000">Наименование:</font> </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" size="30" maxlength="100">
          (до 100 символов)</td>
      </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right">Коментарий: </div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="text" name="coment" value="<%=isFirst?coment:Request.Form("coment")%>" size="30" maxlength="499">
              (до 500 символов)</td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right">Путь до файла: </div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="text" name="path" value="<%=isFirst?path:Request.Form("path")%>" size="30" maxlength="250">
              (только для IMG &amp; FLASH)</td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right">FLASH code: </div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="text" name="flash" value="<%=isFirst?flash:Request.Form("flash")%>" size="30" maxlength="90">
              (только для FLASH)</td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right"><font color="#FF0000">Пароль пользователя: </font></div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="password" name="psw1" size="20" maxlength="20">
              повторите пароль: 
              <input type="password" name="psw2" size="20" maxlength="20">
            </td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right"><font color="#FF0000">Ширина: </font></div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="text" name="ww" value="<%=isFirst?ww:Request.Form("ww")%>" size="6" maxlength="3">
            </td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right"><font color="#FF0000">Высота: </font></div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="text" name="hh" value="<%=isFirst?hh:Request.Form("hh")%>" size="6" maxlength="3">
            </td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" height="20" width="200"> 
              <div align="right"><font color="#FF0000">Тип банера: </font></div>
            </td>
            <td width="15" height="20">&nbsp;</td>
            <td bordercolor="#0099FF" height="20"> 
              <input type="radio" name="tpp" value="1" <%=tpp==1?"checked":""%> >
              - Имидж 
              <input type="radio" name="tpp" value="2" <%=tpp==2?"checked":""%>>
              - Flash 
              <input type="radio" name="tpp" value="3" <%=tpp==3?"checked":""%>>
              - JScript</td>
          </tr>
          <tr> 
            <td bordercolor="#0099FF" width="200"> 
              <div align="right">URL перехода: </div>
            </td>
            <td width="15">&nbsp;</td>
            <td bordercolor="#0099FF"> 
              <input type="text" name="url" value="<%=isFirst?url:Request.Form("url")%>" size="30" maxlength="200" >
            </td>
          </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right">JScript: </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <textarea name="jscr" rows="3" cols="40"><%=jscr%></textarea>
          (только для банеров JScript)</td>
      </tr>
    </table>
 <%
	}
%>   
		<p>Нажмите &quot;Применить&quot; подверждения операции.</p>
	    <input type="submit" name="Submit" value="Применить">
	</form>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<%
}
%>
<p>&nbsp;</p>
</body>
</html>
