<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var pid=Request("pid")

// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=1
// +++  smi_id - ��� ��� � ������� SMI !!

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
	if(ext!="JPG" && ext!="GIF"){throw "����������� ������ JPG ��� GIF �����."}
	if (size>20480){throw "�� ����� 20kB."}
	if (size==0){throw "��� �����."}
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
<title>�������� ���������� � �������� � �������� �����</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#FFCC00"> 
    <td colspan="3" height="17" bgcolor="#CCCCCC" bordercolor="#CCCCCC"> 
      <p align="left"><font face="Arial, Helvetica, sans-serif"><a href="/">�� 
        ������� ��������</a> | <a href="pubarea.asp">������� ���������</a> | <a href="headimglst.asp">������ 
        ����������� �������</a></font></p>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<form name="form1" method="post" action="addhidimg.asp" enctype="multipart/form-data">
  <h1 align="center"> <b>�������� ������� � �������� � �������� �����</b></h1>
  <p>
  <div align="center"> 
    <p>����������� ����� GIF ��� JPG �� 20 ��������</p>
  </div>
  <div align="center"> 
    <table width="60%" border="2" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#333333" bgcolor="#CCCCCC"> 
          <div align="center"> 
            <p><b>�������� ����� �� ����� ���������� ��� �������� �� ������:</b></p>
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
    <input type="submit" name="Submit" value="���������">
  </p>
</form>
<p align="center">&nbsp;</p>
<hr width="500">
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</body>
</html>
