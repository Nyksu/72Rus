<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=1
// +++  smi_id - ��� ��� � ������� SMI !!

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {pid=0}

var usok=false
var id_usr=0
var sql=""
var bid=0
var name=""
var sname=""
var sminame=""
var isfull=""

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="bloknews.asp"
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<5) {usok=true}
	id_usr=TextFormData(Session("id_mem_pub"),"0")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
		}
		Records.Close()
	} else {
		usok=true
	}
}

if (!usok) {Response.Redirect("index.asp")}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

%>

<html>
<head>
<title>�������������� ����� �������� � ����������</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><font face="Arial, Helvetica, sans-serif"><a href="/">�� ������� ��������</a> 
        | <a href="pubarea.asp">������� ���������</a> | </font></p>
    </td>
  </tr>
</table>
<h1 align="center">���: <%=sminame%></h1>
<h1 align="center">����� ��������:</h1>
<h1 align="center">(* - ���������� �������� ������� �������)</h1>
<p align="center"><a href="addbloknews.asp">�������� ����� ����</a></p>
<%
var recs=CreateRecordSet()
Records.Source="Select * from block_news where smi_id="+smi_id+" order by id"
Records.Open()
while (!Records.EOF )
{
	bid=String(Records("ID").Value)
	name=String(Records("NAME").Value)
	sname=TextFormData(Records("SUBJ").Value,"")
%>
<table width="780" border="1" bordercolor="#FFFFFF" align="center">
  <tr bgcolor="#33CCFF" bordercolor="#333333"> 
    <td width="150" bgcolor="#CCCCCC"> 
      <p><b>���� �</b> <b><font color="#000099"><%=bid%></font></b></p>
    </td>
    <td bgcolor="#FFFFFF"> 
      <p>|| <a href="block.asp?bk=<%=bid%>">������� 
        �����</a> || <a href="edblock.asp?bk=<%=bid%>">�������� 
        ���������</a> || <a href="delblock.asp?bk=<%=bid%>">�������</a> 
        || </p>
    </td>
  </tr>
  <tr> 
    <td colspan="2" height="8"> 
      <p><b>������� �������� ����:</b> <font color="#0000FF"><%=sname%><br>
        </font>������� ��������: <b><%=name%> </b></p>
    </td>
  </tr>
</table>
<%if (pid>0) {%>
<p align="center">���������� ���������� � �������: || 
  <%
recs.Source="Select * from news_pos where block_news_id="+bid+" order by posit"
recs.Open()
while (!recs.EOF )
{
	pos=String(recs("POSIT").Value)
	isfull="*"
	if (String(recs("PUBLICATION_ID").Value)=="null") {isfull=""}
%>
  <a href="setposblok.asp?bk=<%=bid%>&pos=<%=pos%>&pid=<%=pid%>"><%=pos%><%=isfull%></a>|| 
  <%
	recs.MoveNext()
} 
recs.Close()%>
  <%}%>
<hr width="780">
<%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
<p>&nbsp; </p>
</body>
</html>
