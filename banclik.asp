<%@LANGUAGE="JAVASCRIPT"%>
<%
if (String(Session("urlstat"))=="undefined") {Session("urlstat")="Страница не определена"}
var urrl=Session("urlstat")
%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\bans.inc" -->

<%
var  id=parseInt(Request("bid"))
if (isNaN(id)) {id=0}
var sql=""
var aurl=""

if (id!=0) {
	sql="Select * from baner where id="+id   
	Records.Source=sql
	Records.Open()
	if (!Records.Eof) {
		aurl=Records("URL").Value
		Records.Close()
		AddClik(id,urrl)
		Response.Redirect(aurl)
	}
	Records.Close()
} 

//Response.Redirect("index.html")
Response.Redirect("index.asp")

%>
