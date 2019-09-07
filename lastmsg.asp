<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%
var dd=parseInt(Request("d"))
if (isNaN(dd)) {dd=1}
if (dd==0) {dd=1}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="lastmsg.asp?d="+dd
Response.Redirect("login.asp")
}

var path=""
var mid=0
var sql=""
var name=""
var subj_id=0
var subj=0
var citynam=""
var inn=""
var onn=""
var iss=0
var tpnam=""
var nompp="1"
var dc=""
var nm=""


%>

<html>
<head>
<title>Последние объявления</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> |</p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<p align="center"><font size="3"><b><font color="#0000CC">Последние добавленные 
  объявления:</font></b></font></p>
<p>За последние (<b><font color="#FF0000"><%=dd%></font></b>) дня</p>
<%
var recs=CreateRecordSet()
nompp=1
sql="Select t1.*, t2.name as citynam, t3.in_name, t3.out_name from trademsg t1, city t2, trade_subj t3 where t1.city_id=t2.id and t1.trade_subj_id=t3.id and t3.marketplace_id="+market+" and t1.date_create>'TODAY'-"+dd+" order by t1.id desc"
Records.Source=sql
Records.Open()
while (!Records.EOF) { 
	mid=Records("ID").Value
	name=Records("NAME").Value
	subj_id=Records("trade_subj_id").Value
	citynam=Records("CITYNAM").Value
	inn=Records("IN_NAME").Value
	onn=Records("OUT_NAME").Value
	iss=Records("IS_FOR_SALE").Value
	dc=Records("DATE_CREATE").Value
	if (iss==0) {tpnam=inn} else {tpnam=onn}
	sbj=subj_id
	path=""
	while (sbj>0) {
		recs.Source="Select * from trade_subj where id="+sbj+" and marketplace_id="+market
		recs.Open()
		nm=String(recs("NAME").Value)
		path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> <img src=\"Img/arrow2.gif\" width=\"11\" height=\"10\" align=\"middle\"> "+path
		sbj=recs("HI_ID").Value
		recs.Close()
	}
%>
<table width="95%" border="1">
  <tr> 
    <td width="100"> 
      <div align="center"> 
        <p><b><a href="msg.asp?ms=<%=mid%>" target="_blank"><%=nompp%>.</a></b></p>
      </div>
    </td>
    <td width="150"> 
      <div align="center">
        <p><b><%=citynam%></b></p>
      </div>
    </td>
    <td>
      <p>(<%=tpnam%>) <b><font color="#0000CC"><%=name%></font></b></p>
    </td>
  </tr>
  <tr> 
    <td width="100"> 
      <div align="center"> 
        <p><b><%=dc%></b></p>
      </div>
    </td>
    <td width="150"> 
      <div align="center"><font size="2"><a href="delms.asp?ms=<%=mid%>">Удалить</a> 
        Переместить</font></div>
    </td>
    <td> 
      <p><b><font color="#FF0000">Путь:</font> <%=path%></b></p>
    </td>
  </tr>
</table>
<hr>
<%
	nompp=nompp+1
	Records.MoveNext()
}
Records.Close()
delete recs
%>
<p>&nbsp;</p>

</body>
</html>
