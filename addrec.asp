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
var coment=""
var id=0
var sql=""
var bid=0
var bnm=""
var rban=0
var dban=0
var maxs=0
var dtstart=""
var dtend=""

if(!isFirst){
	name=TextFormData(Request.Form("name"),"")
	coment=TextFormData(Request.Form("coment"),"")
	rban=parseInt(Request.Form("rban"))
	dban=parseInt(Request.Form("dban"))
	maxs=parseInt(Request.Form("maxs"))
	dtstart=TextFormData(Request.Form("dtstart"),"")
	dtend=TextFormData(Request.Form("dtstart"),"")
	
	if (!/(\d{1,2}).(\d{1,2}).(\d{4})$/.test(dtend)){ErrorMsg+="'���� ���������' ������������.<br>"}
	if (!/(\d{1,2}).(\d{1,2}).(\d{4})$/.test(dtstart)){ErrorMsg+="'���� ������' ������������.<br>"}
	if (isNaN(rban)) {ErrorMsg+="���������� ����� �� ������.<br>"}
	if (isNaN(dban)) {ErrorMsg+="��������� ����� �� ������.<br>"}
	if (isNaN(maxs)) {maxs=0} 
	if (maxs==0){ErrorMsg+="�� ������ ������������ ���-�� �������.<br>"}
	if (name.length<2) {ErrorMsg+="������� �������� ��� ������.<br>"}
	else {
		sql="Select * from baner where name='"+name+"'"
		Records.Source=sql
		Records.Open()
		if (!Records.Eof) {ErrorMsg+="��� ������ ������!<br>"}
		Records.Close()
	}
	if (ErrorMsg=="") {
		id=NextID("RECBLOCK")
		sql="Insert into banblock (id,name,baner_id,stopdate,maxshows,startdate,baner_id_def,coment,state,tipblock) "
		sql+="values (%ID, '%NAM', %BID, '%ED', %MS, '%SD', %BIDD, '%COM', 0,1)"
		sql=sql.replace("%ID",id)
		sql=sql.replace("%COM",coment)
		sql=sql.replace("%NAM",name)
		sql=sql.replace("%BID",rban)
		sql=sql.replace("%BIDD",dban)
		sql=sql.replace("%MS",maxs)
		sql=sql.replace("%SD",dtstart)
		sql=sql.replace("%ED",dtend)
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
		  ShowForm=false
		}
	}
}
%>

<html>
<head>
<title>���������� ����������� �����</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p align="center"><a href="admarea.asp">��������� � �������</a> | <a href="index.asp">������� 
  �� ������� ��������</a></p>
<p align="center"><b><font color="#000099" size="4">���������� ����������� ����� 
  � �������.</font></b></p>
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
<%
if (ShowForm) {
%>
<form name="form1" method="post" action="addrec.asp">
  <div align="center">
    <table width="90%" border="2" cellspacing="2" cellpadding="1" bordercolor="#FFFFFF">
      <tr> 
        <td bordercolor="#0099FF" bgcolor="#CCCCFF" width="200"> 
          <div align="center"><b>���������</b></div>
        </td>
        <td width="15"> 
          <div align="center"><b></b></div>
        </td>
        <td bordercolor="#0099FF" bgcolor="#CCCCFF"> 
          <div align="center"><b>��������</b></div>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font color="#FF0000" size="2">������������:</font> 
            </b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" size="30" maxlength="100">
          <font size="2"> (�� 100 ��������)</font></td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font size="2">����������: </font></b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="coment" value="<%=isFirst?"":Request.Form("coment")%>" size="30" maxlength="250">
          <font size="2">(�� 250 ��������)</font></td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font color="#FF0000" size="2">���������� �����:</font> 
            </b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <select name="rban">
            <%
	sql="Select * from baner where state=0"
	Records.Source=sql
	Records.Open()
	while (!Records.Eof) {
		bnm=Records("NAME").Value
		bid=Records("ID").Value
	%>
            <option value="<%=bid%>" <%=Request.Form("rban")==bid?"selected":""%>><%=bnm%></option>
    <%
		Records.MoveNext()
	}
	Records.Close()
	%>
          </select>
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font color="#FF0000" size="2">MAX �-�� �������:</font> 
            </b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="maxs" value="<%=isFirst?"":Request.Form("maxs")%>" size="20" maxlength="9">
        </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font color="#FF0000" size="2">���� ������ ������: 
            </font></b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="dtstart" value="<%=isFirst?"":Request.Form("dtstart")%>" size="30" maxlength="10">
          <font size="2">(������: &quot;��.��.����&quot;)</font> </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font color="#FF0000" size="2">���� ��������� 
            ������: </font></b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <input type="text" name="dtend" value="<%=isFirst?"":Request.Form("dtend")%>" size="30" maxlength="10">
          <font size="2">(������: &quot;��.��.����&quot;)</font> </td>
      </tr>
      <tr> 
        <td bordercolor="#0099FF" width="200"> 
          <div align="right"><b><font color="#FF0000" size="2">��������� �����: 
            </font></b></div>
        </td>
        <td width="15">&nbsp;</td>
        <td bordercolor="#0099FF"> 
          <select name="dban">
    <%
	sql="Select * from baner where state=0"
	Records.Source=sql
	Records.Open()
	while (!Records.Eof) {
		bnm=Records("NAME").Value
		bid=Records("ID").Value
	%>
            <option value="<%=bid%>" <%=Request.Form("bban")==bid?"selected":""%> ><%=bnm%></option>
            <%
		Records.MoveNext()
	}
	Records.Close()
	%>
          </select>
        </td>
      </tr>
    </table>
    <input type="submit" name="Submit" value="���������">
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
        <p align="center"><b><font color="#0000FF" size="4">���������� ����� (<%=id%>) ��������� � �������!</font></b></p>
        <p align="center"><font size="2"><b>��� HTML: <font color="#0033CC">&lt;script 
          language=&quot;javascript&quot; src=&quot;</font><font color="#0033CC">banshow.asp?rid=<%=id%>&amp;u=http://72rus.ru/<font color="#FF0000">��� ��������</font>.html&quot;&gt;</font><font size="2"><b><font size="2"><b><font color="#0033CC">&lt;/script</font></b></font><font color="#0033CC">&gt;</font></b></font></b></font></p>
        <p align="center"><font size="2"><b>��� ASP: <font color="#0033CC">&lt;script 
          language=&quot;javascript&quot; src=&quot;banshow.asp?rid=<%=id%>&quot;&gt;</font><font size="2"><b><font color="#0033CC">&lt;/script</font></b></font><font color="#0033CC">&gt;</font></b></font></p>
        <p align="center">&nbsp;</p>
      </td>
    </tr>
  </table>
</div>
<%
}
%>

</body>
</html>
