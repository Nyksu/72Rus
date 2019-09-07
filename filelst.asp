<%@LANGUAGE="JScript"%>
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
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

var hid=parseInt(Request("hid"))
if (isNaN(hid)) {hid=0}

var hname=""
var tekimg=""
if (hid != 0) {
	Records.Source="Select * from heading where id="+hid+" and smi_id="+smi_id
	Records.Open()
	if (Records.EOF) {
		Records.Close()
		Response.Redirect("index.asp")
	}
	hname=Records("NAME").Value
	tekimg=TextFormData(Records("PICTURE").Value,"")
Records.Close()
}

%>

<html>
<head>
<title>Имиджи для рубрик и разделов сайта</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<LINK REL="stylesheet" HREF="style.css" TYPE="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333" height="2"> 
      <p><font face="Arial, Helvetica, sans-serif"><a href="/">На главную страницу</a> 
        | <a href="pubarea.asp">Кабинет редактора</a> | <a href="addhidimg.asp">Загрузить 
        имидж</a></font></p>
    </td>
  </tr>
</table>
<h2 align="center">Имиджи для рубрик и разделов сайта</h2>
<%if (hid !=0 ) {%>
<p align="center">Добавление имиджа в раздел: <a href="pubheading.asp?hid=<%=hid%>"><%=hname%></a></p>
<%}%>
<table width="100%" border="1" align="center" bordercolor="#FFFFFF">
  <tr bordercolor="#666666"> 
    <td width="40" bgcolor="#CCCCCC"> 
      <p align="center"><b>№</b></p>
    </td>
    <td width="468" bgcolor="#CCCCCC"> 
      <p align="center"><b>Имидж</b></p>
    </td>
    <td bgcolor="#CCCCCC"> 
      <p align="center"><b>Управление</b></p>
    </td>
  </tr>
  <%
var fso, f, f1, fc, s, ii, fl, nam, ext;
   fso = new ActiveXObject("Scripting.FileSystemObject");
   f = fso.GetFolder(HeadImgPathR);
   fc = new Enumerator(f.files);
   s = "";
   ii=0;
   for (; !fc.atEnd(); fc.moveNext())
   {
      s  = fc.item();
	  fl = fso.GetFile(s);
	 nam=fl.Name;
	 ext="";
	 if (String(nam).indexOf(".",nam.length-4)!=-1) {
	 ext=String(nam).substring(String(nam).indexOf(".",nam.length-4)+1,nam.length)
	 ext=ext.toUpperCase()
	 }
	if (ext != "" && (ext=="JPG" || ext=="GIF") ) {
	ii += 1;
%>
  <tr bordercolor="#666666"> 
    <td width="40"> 
      <div align="right"><%=ii%>.</div>
    </td>
    <td width="468"> 
      <div align="center"><%=nam%></div>
    </td>
    <td bordercolor="#666666"> 
      <div align="center"> 
        <%if (tekimg != nam) {%>
        <%if (hid==0) {%>
        <a href="delheadimg.asp?img=<%=nam%>">удалить</a> 
        <%} else {%>
        <a href="setheadimg.asp?hid=<%=hid%>&img=<%=nam%>">Выбрать 
        этот</a> 
        <%}
		} 
		%>
      </div>
    </td>
  </tr>
  <%
   }}
 %>
</table>
<p>&nbsp;</p>
</body>
</html>
