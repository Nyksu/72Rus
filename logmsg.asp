<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var ErrorMsg=""
var cw=""

if (String(Session("backurl"))=="undefined"){Session("backurl")="index.asp"}

isFirst=String(Request.Form("login"))=="undefined"
if(!isFirst){
	cw=TextFormData(Request.Form("cw"),"")
	Records.Source="Select * from TRADEMSG  where DATE_END>='TODAY' and CODEWORD='"+cw+"'"
	Records.Open()
	if (Records.EOF){ErrorMsg+="�������� ��� �������.<br>"}
	else { if (Records("STATE").Value==0) {
		Session("id_ms")=String(Records("ID").Value)
		} else {
		Session("id_ms")="undefined"
		ErrorMsg+="��� �������.<br>"
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
      <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
        <%=ErrorMsg%></p>
  </center>
<%}%>
<form name="form1" method="post" action="logmsg.asp">
  <table width="100%" border="0" cellspacing="4" cellpadding="0">
    <tr valign="middle"> 
      <td width="50%" align="right"> 
        <p><b>������ ��� �������: </b></p>
      </td>
      <td width="50%"> 
        <input type="text" name="cw" value="<%=isFirst?"":Request.Form("cw")%>" size="30" maxlength="20">
        * </td>
    </tr>
  </table>
  <p align="center"> 
    <input type="submit" name="login" value="����">
  </p>
  <hr size="1">
  <p align="center"><b>* - ������������ ����</b></p>
  <p align="center">&nbsp;</p>
</form>
  

</body>
</html>
