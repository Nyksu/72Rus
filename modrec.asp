<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var opr=parseInt(Request("opr"))
var id=parseInt(Request("id"))
var sql=""
var name=""
var coment=""
var rid=0
var ii=0
var isadm=0
var bnm=""
var bid=0
var maxs=0
var dtstart=""
var dtend=""
var rban=0
var dban=0
var ErrorMsg=""
var dbnam=""
var rbnam=""

if (isNaN(opr)) {opr=0}
if (isNaN(id)) {id=0}
if ((Session("is_adm_mem")==1) || (Session("is_host")==1)) { isadm=1 }
if (isadm==0) {Response.Redirect("index.asp")}
//if (isadm==0) {Response.Redirect("index.html")}
if (opr>0 && id==0) {Response.Redirect("modrec.asp")}
if (opr>0 && id>0) {
	sql="Select * from banblock where id="+id
	Records.Source=sql
	Records.Open()
	if (Records.Eof) {
		Records.Close()
		Response.Redirect("modrec.asp")
	}
	name=TextFormData(Records("NAME").Value,"")
	coment=TextFormData(Records("COMENT").Value,"")
	rban=Records("baner_id").Value
	dban=Records("baner_id_def").Value
	dtstart=Records("startdate").Value
	dtend=Records("stopdate").Value
	maxs=Records("maxshows").Value
	Records.Close()
}
var isFirst=String(Request.Form("Submit"))=="undefined"

if ((!isFirst) && (opr>0 && id>0)){
	sql=""
	if ((opr==1) && (Request.Form("kill")==666)) {
		// �������� ���������� �����
		sql="Delete from banblock where id="+id
	} else { if (opr==1) {Response.Redirect("modrec.asp")}}
	if ((opr==2) && (Request.Form("lockit")==33)) {
		// ����������� ���������� �����
		sql="Update banblock set state=1 where id="+id
	} else { if (opr==2) {Response.Redirect("modrec.asp")}}
	if (opr==5 && Request.Form("unlockit")==333) {
		// ��������� ���������� �����
		sql="Update banblock set state=0 where id="+id
	} else { if (opr==5) {Response.Redirect("modrec.asp")}}
	if (opr==3) {
		// ��������� ���������� �����
		name=TextFormData(Request.Form("name"),"")
		coment=TextFormData(Request.Form("coment"),"")
		dban=parseInt(Request.Form("dban"))
		if (isNaN(dban)) {ErrorMsg+="��������� ����� �� ������.<br>"}
		if (name.length<2) {ErrorMsg+="������� �������� ��� ������.<br>"}
		if (ErrorMsg=="") {sql="Update banblock set name='"+name+"', coment='"+coment+"', baner_id_def="+dban+" where id="+id}
	}
	if (opr==4) {
		// ��������� ��������� �������� �����
		rban=parseInt(Request.Form("rban"))
		maxs=parseInt(Request.Form("maxs"))
		dtstart=TextFormData(Request.Form("dtstart"),"")
		dtend=TextFormData(Request.Form("dtend"),"")
		if (isNaN(maxs)) {maxs=0}
		if (maxs==0){ErrorMsg+="�� ������ ������������ ���-�� �������.<br>"}
		if (!/(\d{1,2}).(\d{1,2}).(\d{4})$/.test(dtend)){ErrorMsg+="'���� ���������' ������������.<br>"}
		if (!/(\d{1,2}).(\d{1,2}).(\d{4})$/.test(dtstart)){ErrorMsg+="'���� ������' ������������.<br>"}
		if (isNaN(rban)) {ErrorMsg+="���������� ����� �� ������.<br>"}
		if (ErrorMsg=="") {sql="Update banblock set baner_id="+rban+", maxshows="+maxs+", stopdate='"+dtend+"', startdate='"+dtstart+"' where id="+id}
	}
	if ((ErrorMsg=="") && (sql!="")) {
			Connect.BeginTrans()
			try{
				Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="������ �������.<br>"
			}
			if (ErrorMsg==""){
				Connect.CommitTrans()
		  		Response.Redirect("modrec.asp")
			}
	} 
}

%>

<html>
<head>
<title>���������� ���������� �������</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<p><a href="admarea.asp">��������� � �������</a> | <a href="index.asp">������� 
  �� ������� ��������</a></p>
<p align="center">&nbsp;</p>
<p align="center"><b><font color="#0000FF" size="4">���������� ���������� �������.</font></b></p>
<%
if (opr==0) {
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">�������</a> | 
        <font size="2"><a href="banstat.asp">���������� �� �������</a> | <a href="addbaner.asp">�������� 
        �����</a> | <a href="modbaner.asp">���������� ��������</a></font></p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<p align="center"><a href="addrec.asp">+ �������� ���������� �����</a> | <a href="recstat.asp">���������� 
  ���������� ����</a> <font size="2"> |<a href="addrec.asp"></a></font></p>
<p><b>�������� ��������� �����:</b></p>
<%
	sql="Select t1.*, t2.name as rbnam, t3.name as dbnam from banblock t1, baner t2, baner t3 where t1.state=0 and t1.baner_id=t2.id and t1.baner_id_def=t3.id order by t1.name"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		name=TextFormData(Records("NAME").Value,"")
		coment=TextFormData(Records("COMENT").Value,"")
		dbnam=TextFormData(Records("dbnam").Value,"")
		rbnam=TextFormData(Records("rbnam").Value,"")
		dtstart=Records("startdate").Value
		dtend=Records("stopdate").Value
		maxs=Records("maxshows").Value
		rid=Records("ID").Value
		ii=ii+1
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#666666"> 
    <td width="5%"> 
      <p><font size="1"><b>&nbsp;&nbsp;<%=rid%>.</b></font></p>
    </td>
    <td> 
      <p><font color="#0033CC"><b><a href="recstat.asp?rid=<%=rid%>"><%=name%></a> 
        </b><font color="#000066">(<%=coment%>)</font></font></p>
    </td>
  </tr>
  <tr bordercolor="#999999"> 
    <td colspan="2"> 
      <p><font size="2">��������� �����:<font color="#009933"><%=rbnam%></font><font size="2"> 
        | ��������� �����:</font> <font color="#FF0000"><%=dbnam%></font></font></p>
    </td>
  </tr>
  <tr bordercolor="#999999"> 
    <td colspan="2" bgcolor="#E1F2EC"> 
      <p><font size="2">���� ������ ��������� �������� : <font color="#009933"><%=dtstart%></font> 
        | ���� ���������: <font color="#009933"><%=dtend%></font> | ������������ 
        ���-�� �������: <font color="#009933"><%=maxs%></font></font></p>
    </td>
  </tr>
  <tr bordercolor="#999999"> 
    <td colspan="2" bgcolor="#F4F4F4"> 
      <p><font size="1"><b>&nbsp;</b></font><font size="1" color="#FF0000"><b>����������:</b></font><font size="1" color="#0033CC"> 
        &nbsp;&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=1">�������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=2">�������������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=3">�������� 
        ���������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=4">�������� 
        ��������� �������</a></font></p>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="10"></td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
%>
<%=ii==0?"<p><font size=\"2\"><b><font color=\"#FF6633\">��� �������� ���������� ����.</font></b></font></p>":""%> 
<p><b>�� �������� ��������� �����:</b></p>
<%
	sql="Select * from banblock where state=1 order by name"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.Eof) {
		name=TextFormData(Records("NAME").Value,"")
		coment=TextFormData(Records("COMENT").Value,"")
		rid=Records("ID").Value
		ii=ii+1
%>
<table width="90%" border="1" cellspacing="2" cellpadding="2" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td width="5%"><font size="1"><b>&nbsp;&nbsp;<%=rid%>.</b></font></td>
    <td><font size="1" color="#0033CC"><b><font size="2"><a href="recstat.asp?rid=<%=rid%>"><%=name%></a></font> 
      (</b><font color="#000066"><%=coment%></font><b>)</b></font></td>
  </tr>
  <tr bordercolor="#CCCCCC" bgcolor="#F4F4F4"> 
    <td colspan="2"> 
      <p><font size="1"><b>&nbsp;</b></font><font size="1" color="#FF0000"><b>����������:</b></font><font size="1" color="#0033CC"> 
        &nbsp;&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=1">�������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=5">��������������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=3">�������� 
        ���������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="modrec.asp?id=<%=rid%>&opr=4">�������� 
        ��������� �������</a></font></p>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="10"></td>
  </tr>
</table>
<%
		Records.MoveNext()
	}
	Records.Close()
%>
<%=ii==0?"<p><font size=\"2\"><b><font color=\"#FF6633\">��� �� �������� ���������� ����.</font></b></font></p>":""%>
<%
} else {
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" bordercolor="#FFFFFF">
  <tr>
    <td bgcolor="#F4F4F4"> 
      <p><font size="2"><a href="modrec.asp">��������� � ������ ��������� ������</a></font></p>
<% 
if (ErrorMsg!="") {
%>
<div align="center">
  <table width="90%" border="1" cellspacing="1" cellpadding="1" bordercolor="#FF0000" bgcolor="#FFFFCC">
    <tr> 
      <td> 
        <p align="center"><b><font color="#FF0000" size="2">��������! �������� ��������� ������:</font></b></p>
        <p align="center"><font color="#FF0000" size="2"><%=ErrorMsg%></font></p>
      </td>
    </tr>
  </table>
</div>
<%
}
%>
	<form name="form1" method="post" action="modrec.asp">
        <input type="hidden" name="id" value="<%=id%>">
        <input type="hidden" name="opr" value="<%=opr%>">
<%
	if (opr==1) {
%>
	    <p><b>�������� ����������� �����. <font color="#0000CC"><%=name%></font></b></p>
        <p>
          <input type="checkbox" name="kill" value="666">
          ��! �����!</p>
<%
	}
%>
<%
	if (opr==2) {
%>
	    <p><b>�������������� ���������� �����. <font color="#0000CC"><%=name%></font></b></p>
        <p>
          <input type="checkbox" name="lockit" value="33">
          ��! ����������!</p>
<%
	}
%>
<%
	if (opr==3) {
%>
	    <p><b>�������� ��������� ���������� �����. <font color="#0000CC"><%=name%></font></b></p>
        <p> <b><font color="#FF0000" size="2">������������:</font></b> 
          <input type="text" name="name" value="<%=isFirst?name:Request.Form("name")%>" size="30" maxlength="100">
          <font size="2"> (�� 100 ��������)</font></p>
        <p><b><font size="2">����������: 
          <input type="text" name="coment" value="<%=isFirst?coment:Request.Form("coment")%>" size="30" maxlength="250">
          </font></b><font size="2"><font size="2">(�� 250 ��������)</font></font></p>
        <p><b><font color="#FF0000" size="2">��������� �����: </font></b>
		<select name="dban">
    <%
	sql="Select * from baner where state=0"
	Records.Source=sql
	Records.Open()
	while (!Records.Eof) {
		bnm=Records("NAME").Value
		bid=Records("ID").Value
	%>
            <option value="<%=bid%>" <%=dban==bid?"selected":""%> ><%=bnm%></option>
            <%
		Records.MoveNext()
	}
	Records.Close()
	%>
          </select>
		</p>
<%
	}
%>
<%
	if (opr==4) {
%>
	    <p><b>��������� ��������� ���������� �����. <font color="#0000CC"><%=name%></font></b></p>
        <p><b><font color="#FF0000" size="2">���������� �����: </font></b>
		<select name="rban">
            <%
	sql="Select * from baner where state=0"
	Records.Source=sql
	Records.Open()
	while (!Records.Eof) {
		bnm=Records("NAME").Value
		bid=Records("ID").Value
	%>
            <option value="<%=bid%>" <%=rban==bid?"selected":""%>><%=bnm%></option>
    <%
		Records.MoveNext()
	}
	Records.Close()
	%>
          </select>
		</p>
        <p><b><font color="#FF0000" size="2">MAX �-�� �������: </font></b><input type="text" name="maxs" value="<%=isFirst?maxs:Request.Form("maxs")%>" size="20" maxlength="9"></p>
        <p><b><font color="#FF0000" size="2">���� ������ ������: </font></b><input type="text" name="dtstart" value="<%=isFirst?dtstart:Request.Form("dtstart")%>" size="30" maxlength="10"></p>
        <p><b><font color="#FF0000" size="2">���� ��������� ������: </font></b><input type="text" name="dtend" value="<%=isFirst?dtend:Request.Form("dtend")%>" size="30" maxlength="10"> 
        </p>
<%
	}
%>
<%
	if (opr==5) {
%>
	    <p><b>������������ ���������� �����. <font color="#0000CC"><%=name%></font></b></p>
        <p>
          <input type="checkbox" name="unlockit" value="333">
          ��! ������������!</p>
<%
	}
%>
	    <p>������� &quot;���������&quot; ������������ ��������.</p>
	    <input type="submit" name="Submit" value="���������">
      </form>
	</td>
  </tr>
</table>
<%
}
%>
<p>&nbsp;</p>
</body>
</html>
