<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!

var hid=parseInt(Request("hid"))
if (isNaN(hid)) {hid=0}

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {pid=0}

if (pid==0) {Response.Redirect("public.asp")}

var st=parseInt(Request("st"))
if (isNaN(st)) {st=0}

var tpm=1000
var isok=false
var sql=""
var sminame=""
var name=""
var nm=""
var hadr=""
var path=""
var hiname=""

if (String(Session("id_mem"))=="undefined") {
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

if (tpm>2) {Response.Redirect("public.asp")}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

var tit=sminame

if (hid>0) {
	Records.Source="Select * from heading where id="+hid+" and smi_id="+smi_id
	Records.Open()
	if (Records.EOF) {
		Records.Close()
		Response.Redirect("public.asp")
	}
	Records.Close()
}

Records.Source="Select  t1.*  from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.id="+pid
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("public.asp")
}
name=String(Records("NAME").Value)
Records.Close()


var hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	if (hdd==hid) {
		path=nm+" | "+path
		hiname=String(Records("NAME").Value)
	}
	else {
		path="<a href=\"removepub.asp?hid="+hdd+"&pid="+pid+"\">"+nm+"</a> | "+path
	}
	hdd=Records("HI_ID").Value
	Records.Close()
}

path="<a href=\"removepub.asp?hid=0&pid="+pid+"\">"+sminame+"</a> | "+path

tit+=" | "+hiname

if (st==1) {
// перемещаем!
	sql="Update publication set heading_id="+hid+" where id="+pid
	Connect.BeginTrans()
	try{
			Connect.Execute(sql)
			Connect.CommitTrans()
			isok=true
	}
	catch(e){
			Connect.RollbackTrans()
	}
}

%>

<html>
<head>
<title>Публикации новостей: <%=tit%></title> <meta http-equiv="Content-Type" content="text/html; charset=windows-1251"> 
<link rel="stylesheet" href="style.css" type="text/css"> 
</head>

<body bgcolor="#FFFFFF" text="#000000" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0">
<TABLE WIDTH="100%" BORDER="1" CELLSPACING="0" CELLPADDING="0" BORDERCOLOR="#FFFFFF"> 
<TR> <TD BGCOLOR="#CCCCCC" BORDERCOLOR="#333333"> <P><FONT FACE="Arial, Helvetica, sans-serif"><A HREF="/">На 
главную страницу</A> | <A HREF="pubarea.asp">Кабинет редактора</A> | </FONT></P></TD></TR> 
</TABLE><table width="90%" border="1" BORDERCOLOR="#808080"> <tr> <td>&nbsp;<%=path%></td></tr> </table><p>ПЕРЕМЕЩЕНИЕ публикации.</p><p>Рубрика: <font color="#003399"><b><font size="3"><%=tit%></font></b></font></p>
<p><%if ( ! isok) {%><a href="removepub.asp?st=1&hid=<%=hid%>&pid=<%=pid%>">Разместить в текущую 
  рубрику</a><%} else {%><font color="#FF0000"><b>Публикация размещена в текущую рубрику!!</b></font><%}%></p>
<p>
  <%
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	Records.MoveNext()
%>
</p>
<table width="90%" border="1">
  <tr> 
    <td height="7"> 
      <p><b><a href="removepub.asp?hid=<%=hdd%>&pid=<%=pid%>"><%=hname%></a></b></p>
    </td>
    <td width="25%"> 
      <p align="center"><b><font color="#FF0000">&lt;-------&lt;</font></b> <a href="removepub.asp?st=1&hid=<%=hdd%>&pid=<%=pid%>">разместить 
        в рубрику</a></p>
    </td>
  </tr>
</table>
<%
} 
Records.Close()
%> <p>&nbsp;</p>
</body>
</html>
