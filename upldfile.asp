<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var isadm=0
if ((Session("is_adm_mem")==1) || (Session("is_host")==1)) { isadm=1 }
//if (isadm==0) {Response.Redirect("index.asp")}
if (isadm==0) {Response.Redirect("index.html")}

var formname="любой файл *.*"
var fnam=""
var ext=""
var filename=""
var size=0
var ErrorMsg=""
var isEnd=false

var tp=parseInt(Request("tp"))

if (Request.ServerVariables("REQUEST_METHOD")=="POST") {
	updown = Server.CreateObject("ANUPLOAD.OBJ")
	tp=parseInt(updown.Form("tp"))
	if (isNaN(tp)) {tp=0}
	switch (tp) {
		case 1 : formname=" графика *.GIF, *.JPG"; break;
		case 2 : formname=" Flash графика *.SWF"; break;
	}
	fnam=updown.GetFileName("file")
	ext=updown.GetExtension("file").toUpperCase()
	size=parseInt(updown.GetSize("file"))
	fileneme=BanDefPath+fnam
	if (size==0){throw "Нет файла."}
	switch (tp) {
		case 1 : if(ext!="JPG" && ext!="GIF"){throw "Принимаются только JPG или GIF файлы."}; break;
		case 2 : if(ext!="SWF"){throw "Принимаются только SWF файлы."}; break;
	}
	try {
		updown.Delete(fileneme)
		updown.SaveAs("file",fileneme)
		isEnd=true
	}
	catch(e){ErrorMsg+=String(e.message)=="undefined"?e:e.message}
} else {
	if (isNaN(tp)) {tp=0}
	switch (tp) {
		case 1 : formname=" графика *.GIF, *.JPG"; break;
		case 2 : formname=" Flash графика *.SWF"; break;
	}
}

%>

<html>
<head>
<title>Загрузка файла</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<p align="center">Загрузка файла:<br>
  (формат файла -  <%=formname%>)</p>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%><%
if (! isEnd) {
%>
<form name="form1" method="post" action="upldfile.asp" enctype="multipart/form-data">
  <div align="center"> 
    <table width="300" border="2" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#333333" bgcolor="#CCCCCC"> 
          <div align="center"> 
            <p><b><font size="2">Выбирите файл для загрузки на сервер:</font></b></p>
          </div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#333333"> 
          <div align="left"> 
            <p align="center"> 
              <input type="file" name="file" size="30">
            </p>
          </div>
        </td>
      </tr>
    </table>
  </div>
  <p align="center"> 
    <input type="submit" name="Submit" value="Сохранить">
  </p>
</form>
<%} else {%>
<p align="center"><font color="#0000FF"><b><i>Загрузка (<%=fileneme%>) завершена!</i></b></font></p>
<%}%>
<FORM>
  <div align="center">
    <INPUT TYPE="button" VALUE="Закрыть окно" onClick=window.parent.close()>
  </div>
</FORM>
</body>
</html>
