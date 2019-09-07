<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\bans.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var  id=parseInt(Request("rid"))
var uur=Request("u")
var rrr=""
if (String(uur)=="undefined") {uur=""}
if (String(Session("urlstat"))=="undefined") {Session("urlstat")=""}
if (uur=="") {uur=Session("urlstat")}
Session("urlstat")=String(uur)

if (isNaN(id)) {id=0}

if (id!=0) {
	var  sql=""
	var rst=""
	var bid=0
	var dbid=0
	var isOk=true
	var pathh=""
	var btype=""
	var wi=0
	var he=0
	var fc=""
	
	sql="Select * from banblock where state=0 and id="+id
	Records.Source=sql
	Records.Open()
	if (Records.Eof) { isOk=false } 
	else {
		bid=Records("BANER_ID").Value
		dbid=Records("BANER_ID_DEF").Value
	}
	Records.Close()
	if (isOk) {
		sql="Select t1.* from banblock t1, baner t3 where t1.state=0 and t1.stopdate > 'TODAY' and t1.startdate <= 'TODAY' and ((t1.maxshows > (Select sum(t2.allshow) from banblockst t2 where t1.id=t2.banblock_id and t1.baner_id=t2.baner_id)) or (not exists (Select * from banblockst t2 where t1.id=t2.banblock_id and t1.baner_id=t2.baner_id) )) and t1.baner_id=t3.id and t3.state=0 and t1.id="+id
		Records.Source=sql
		Records.Open()
		if (Records.Eof) {bid=dbid} // Показываем банер по-умолчанию
		Records.Close()
		if (bid!=null) {
			sql="Select * from baner where id="+bid
			Records.Source=sql
			Records.Open()
			pathh=Records("PATHNAME").Value
			btype=Records("SOURSTYPE").Value
			wi=Records("WIDTH").Value
			he=Records("HEIGHT").Value
			fc=Records("FLASHCOD").Value
			Records.Close()
%>

<%if (btype==1) { // Это имидж!! %>
document.write('<a href="http://72rus.ru/banclik.asp?bid=<%=bid%>" target="_blank"><img src="<%=pathh%>" width="<%=wi%>" height="<%=he%>" border="0"></a>');
<%}%>

<%if (btype==2) { // Это Флэшь!! %>
document.write('<object classid="clsid:<%=fc%>" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="<%=wi%>" height="<%=he%>">');
document.write('  <param name=movie value="<%=pathh%>">');
document.write('  <param name=quality value=high>');
document.write('  <embed src="<%=pathh%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="<%=wi%>" height="<%=he%>">');
document.write('  </embed> ');
document.write('</object>');
<%}%>

<% 
if (btype==3) { // Это Джава-скрипт
var ts=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var filnam=""
var str=""
filnam=JSPath+bid+".jsf"
if (!fs.FileExists(filnam)) { filnam="" }
if (filnam!="") {
	ts= fs.OpenTextFile(filnam)
	while (!ts.AtEndOfStream){
		str=ts.ReadLine()
%>
		document.write('<%=str%>'); 
<%
	}
}
}
%>

<%
		}
		rrr=AddShows(id)
	}
 }
 %>