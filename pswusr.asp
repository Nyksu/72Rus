<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\sql.inc" -->
<!-- #include file="inc\creaters.inc" -->

<%
var fio=""
var usr=parseInt(Request("usr"))
var ShowForm=true
var ErrorMsg=""
var sql=""
var valname=""
var valname1=""

if ((Session("is_adm_mem")!=1)&&(Session("id_mem")!=usr)){
Session("backurl")="pswusr.asp?usr="+usr
Response.Redirect("login.asp")}

Records.Source="Select * from USERS where ID="+usr
Records.Open()
if (Records.EOF){Response.Redirect("area.asp")}
fio=String(Records("FIO").Value)
Records.Close()

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){
	valname=TextFormData(Request.Form("valname"),"")
	valname1=TextFormData(Request.Form("valname1"),"")
	if (valname!=valname1) {ErrorMsg+="������ �� ����������� ��� ����������� � �������.<br>"}
	if (String(valname).length<5) {ErrorMsg+="������ �� ����� ���� ����� 5-� ��������.<br>"}
	
	if (ErrorMsg=="") {
		sql=usrpsw
		sql=sql.replace("%ID",usr)
		sql=sql.replace("%PS",valname)
		Connect.BeginTrans()
		try{
			Connect.Execute(sql)
		}
		catch(e){
			Connect.RollbackTrans()
			ErrorMsg+=ListAdoErrors()
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
<title>��������� ������ ������� <%=fio%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<p align="center"><b>��������� ������ �������.</b> ������������: <b><font color="#000099"><%=fio%></font></b></p>
<p>&nbsp;</p>
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>������!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 

<p align="center"><a href="area.asp">�������</a> &nbsp;&nbsp;&nbsp;<a href="index.html">�������� 
  ��������</a></p>
 <%if (ShowForm) {%>
<form name="form1" method="post" action="pswusr.asp">
  <div align="center"> 
    <input type="hidden" name="usr" value="<%=usr%>">
    <table width="70%" border="1" bordercolor="#FFFFFF">
      <tr bordercolor="#0033CC" bgcolor="#99CCCC"> 
        <td width="50%"> 
          <div align="center"><b><font color="#000099">���������:</font></b></div>
        </td>
        <td> 
          <div align="center"><b><font color="#000099">��������:</font></b></div>
        </td>
      </tr>
      <tr bordercolor="#0033CC"> 
        <td width="50%"> 
          <div align="right">������� ����� ������:&nbsp;&nbsp;</div>
        </td>
        <td>&nbsp;&nbsp;&nbsp; 
          <input type="password" name="valname" maxlength="20" size="20">
        </td>
      </tr>
      <tr bordercolor="#0033CC"> 
        <td width="50%"> 
          <div align="right">��������� ���� ������ ������:&nbsp;&nbsp;</div>
        </td>
        <td>&nbsp;&nbsp;&nbsp; 
          <input type="password" name="valname1" maxlength="20" size="20">
        </td>
      </tr>
    </table>
    <input type="submit" name="Submit" value="���������">
  </div>
</form>
<%} else {%>
<p>&nbsp; </p>
<center>
  <h2><font color="#3333FF">������ �������!</font></h2>
  <p> 
    <%}%>
  </p>
</center>
</body>
</html>
