<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""
// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=1
// +++  smi_id - ��� ��� � ������� SMI !!
var Pass=""
var Nik=""


if (String(Session("backurl"))=="undefined"){Session("backurl")="index.html"}

isFirst=String(Request.Form("login"))=="undefined"
if(!isFirst){	Pass=TextFormData(Request.Form("pass"),"")
	Nik=TextFormData(Request.Form("nik"),"")

	Nik=Nik.replace("/*","")
	Nik=Nik.replace("'","")
	Pass=Pass.replace("/*","")
	Pass=Pass.replace("'","")
	Records.Source="Select * from SMI_USR where PSW='"+Pass+"' and NIK_NAME='"+Nik+"' and SMI_ID="+smi_id
	Records.Open()
	if (Records.EOF){ErrorMsg+="�������� '������ ��� ���������'.<br>"}
	else { if (Records("STATE").Value==0){
		Session("id_mem_pub")=String(Records("ID").Value)
		Session("name_mem_pub")=String(Records("NAME").Value)
		Session("nik_mem_pub")=String(Records("NIK_NAME").Value)
		Session("tip_mem_pub")=String(Records("USR_TYPE_ID").Value)
		Session("email_mem_pub")=TextFormData(Records("EMAIL").Value,"")
		} else {
		Session("id_mem_pub")="undefined"
		Session("name_mem_pub")="undefined"
		Session("nik_mem_pub")="undefined"
		Session("tip_mem_pub")="undefined"
		ErrorMsg+="�������� '������ ��� ���������'.<br>"
		}
	}
	Records.Close()
	if (ErrorMsg==""){
		Response.Redirect(Session("backurl"))
	}
}

%>

<html>
<head>
<title>��������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>

<body bgcolor="#FFFFFF" text="#000000">

 <p><br>
  </p>
  <p align="center"><%if(ErrorMsg!=""){%> </p>
  <center>
    <h2> 
      <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
        <%=ErrorMsg%></p>
    </h2>
  </center>
  <%}%>
  
<p align="center">���� ������������.</p> 
  
<form name="form1" method="post" action="logpub.asp">
  <table width="100%" border="0" cellspacing="4" cellpadding="0"> 
      <tr valign="middle"> 
        <td width="50%" align="right"> 
          <p><b>��������� (�����)</b></p>
        </td>
        <td width="50%"> 
          <input type="text" name="nik" value="<%=isFirst?"":Request.Form("nik")%>" size="20" maxlength="20">
          * </td>
      </tr>
      <tr valign="middle"> 
        <td width="50%" align="right"> 
          <p><b>������</b></p>
        </td>
        <td width="50%"> 
          <input type="password" name="pass" size="20" maxlength="20">
          * </td>
      </tr> 
    </table>
    <p align="center"> 
      <input type="submit" name="login" value="����">
    </p>
    <hr size="1">
    <p align="center"><b>* - ������������ ����</b></p>
    
  <p align="center"><font color="#FF0000">����������� ��� serg � ������ 123</font></p>
  </form>
  

<p><a href="index.html">��������� �� ������� ��������</a></p>
<p><a href="index.asp">��������� � �����������</a></p>
</body>
</html>
