<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var pid=Request("pid")

// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!

var tpm=1000
var sql=""
if (String(Session("id_mem"))=="undefined") {
	if (String(Session("tip_mem_pub"))=="undefined") {Response.Redirect("index.asp")}
	tpm=Session("tip_mem_pub")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			tpm=0
		}
		Records.Close()
	} else {
		tpm=0
	}
}
if (tpm>3) {Response.Redirect("index.asp")}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

var ext=""
var size=0
var ErrorMsg=""
var fnam=""
if (Request.ServerVariables("REQUEST_METHOD")=="POST") {
	updown = Server.CreateObject("ANUPLOAD.OBJ")
	fnam=updown.GetFileName("file")
	ext=updown.GetExtension("file").toUpperCase()
	size=parseInt(updown.GetSize("file"))
	if(ext!="JPG" && ext!="GIF"){throw "Принимаются только JPG или GIF файлы."}
	if (size>20480){throw "Не более 20kB."}
	if (size==0){throw "Нет файла."}
	try {
		updown.Delete(HeadImgPathR+fnam)
		updown.SaveAs("file",HeadImgPathR+fnam)
		Response.Redirect("headimglst.asp")
	}
	catch(e){ErrorMsg+=String(e.message)=="undefined"?e:e.message}
}

%>

<html>
<head>
<title>Загрузка фотографии к рубрикам и разделам сайта</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#FFCC00"> 
    <td colspan="3" height="17" bgcolor="#CCCCCC" bordercolor="#CCCCCC"> 
      <p align="left"><font face="Arial, Helvetica, sans-serif"><a href="/">На 
        главную страницу</a> | <a href="pubarea.asp">Кабинет редактора</a> | <a href="headimglst.asp">Список 
        загруженных имиджей</a></font></p>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<form name="form1" method="post" action="addhidimg.asp" enctype="multipart/form-data">
  <h1 align="center"> <b>Загрузка имиджей к рубрикам и разделам сайта</b></h1>
  <p>
  <div align="center"> 
    <p>графические файлы GIF или JPG до 20 килобайт</p>
  </div>
  <div align="center"> 
    <table width="60%" border="2" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#333333" bgcolor="#CCCCCC"> 
          <div align="center"> 
            <p><b>Выбирите имидж на своем компьютере для загрузки на сервер:</b></p>
          </div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#333333"> 
          <div align="left"> 
            <p align="center"> 
              <input type="file" name="file" size="40">
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
<p align="center">&nbsp;</p>
<hr width="500">
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</body>
</html>
