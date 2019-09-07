<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var isadm=0
if ((Session("is_adm_mem")==1) || (Session("is_host")==1)) { isadm=1 }
if (isadm==0) {Response.Redirect("index.asp")}
//if (isadm==0) {Response.Redirect("index.html")}
var ErrorMsg=""
var ShowForm=true
var isFirst=String(Request.Form("Submit"))=="undefined"
var name=""
var url=""
var coment=""
var path=""
var flash=""
var psw1=""
var psw2=""
var ww=""
var hh=""
var tpp=0
var jscr=""
var id=0
var sql=""
var filename=""
var fs=""
var ts=""
var ttt=0


if(!isFirst){
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
	// проверка значений параметров
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
	// сохраняем!
	if (ErrorMsg=="") {
		sql="Insert into baner (id,name,coment,state,pathname,flashcod,psw,width,height,sourstype,url) "
		sql+="values (%ID,'%NAME','%COM',0,'%PATH','%FLASH','%PSW',%WW,%HH,%TT,'%URL')"
		id=NextID("bann")
		filename=JSPath+id+".jsf"
		var fs= new ActiveXObject("Scripting.FileSystemObject")
		sql=sql.replace("%ID",id)
		sql=sql.replace("%COM",coment)
		sql=sql.replace("%NAME",name)
		sql=sql.replace("%PATH",path)
		sql=sql.replace("%FLASH",flash)
		sql=sql.replace("%PSW",psw1)
		sql=sql.replace("%WW",ww)
		sql=sql.replace("%HH",hh)
		sql=sql.replace("%TT",tpp)
		sql=sql.replace("%URL",url)
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
			ErrorMsg+="Ошибка вставки.<br>"
		}
		if (ErrorMsg==""){
		  Connect.CommitTrans()
		  ShowForm=false
		}
	}
}
%>


<html>
<head>
<title>Добавление банера в систему</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><a href="admarea.asp">Вернуться в кабинет</a> | <a href="index.asp">Перейти 
  на главную страницу</a></p>
<p align="center"><b><font color="#000099" size="4">Добавление банера в систему.</font></b></p>
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
<%
if (ShowForm) {
%>
<form name="form1" method="post" action="addbaner.asp">
  <div align="center">
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
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" size="30" maxlength="100">
          (до 100 символов)</td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right">Коментарий: </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="coment" value="<%=isFirst?"":Request.Form("coment")%>" size="30" maxlength="499">
          (до 500 символов)</td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right">Путь до файла: </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="path" value="<%=isFirst?"":Request.Form("path")%>" size="30" maxlength="250">
          (только для IMG &amp; FLASH)</td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right">FLASH code: </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="flash" value="<%=isFirst?"":Request.Form("flash")%>" size="30" maxlength="90">
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
          <input type="text" name="ww" value="<%=isFirst?"0":Request.Form("ww")%>" size="6" maxlength="3">
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><font color="#FF0000">Высота: </font></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="hh" value="<%=isFirst?"0":Request.Form("hh")%>" size="6" maxlength="3">
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" height="20" width="200"> 
          <div align="right"><font color="#FF0000">Тип банера: </font></div>
        </td>
        <td width="15" height="20">&nbsp;</td>
        <td bordercolor="#0099FF" height="20"> 
          <input type="radio" name="tpp" value="1" <%=isFirst?"checked":""%><%=Request.Form("tpp")==1?"checked":""%> >
          - Имидж 
          <input type="radio" name="tpp" value="2" <%=Request.Form("tpp")==2?"checked":""%>>
          - Flash 
          <input type="radio" name="tpp" value="3" <%=Request.Form("tpp")==3?"checked":""%>>
          - JScript</td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right">URL перехода: </div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="url" value="<%=isFirst?"":Request.Form("url")%>" size="30" maxlength="200" >
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
    <input type="submit" name="Submit" value="Сохранить">
  </div>
</form>
<%
} else {
%>
<div align="center">
  <table width="90%" border="1" cellspacing="1" cellpadding="1" bordercolor="#FFFFFF">
    <tr> 
      <td bordercolor="#0066FF"> 
        <p align="center">&nbsp;</p>
        <p align="center"><b><font color="#0000FF" size="4">Банер добавлен в систему!</font></b></p>
        <p align="center"><a href="javascript:" onClick="win1=window.open ('upldfile.asp?tp=<%=ttt%>','w1','width=450,height=370,toolbar=no,location=no, directories=no,status=no,menubar=no,scrollbars=no');return true;" >ЗДЕСЬ</a> можно загрузить файл банера.</p>
        <p align="center">&nbsp;</p>
      </td>
    </tr>
  </table>
</div>
<%
}
%>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</body>
</html>
