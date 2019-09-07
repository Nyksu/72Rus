<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var id=parseInt(Request("url"))
var setst=parseInt(Request("st"))

if (isNaN(setst)) {Response.Redirect("admarea.asp")}
if (isNaN(id)) {id=0}
if (id==0) {Response.Redirect("admarea.asp")}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="urlresume.asp?url="+id+"&st="+setst
Response.Redirect("login.asp")
}

var hid=0
var isok=true

sql="Select t1.* from url t1, catarea t2 where t1.id="+id+" and t1. catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("admarea.asp")
}
Records.Close()

sql="Update url set state="+setst+" where id="+id
Connect.BeginTrans()
try{
	Connect.Execute(sql)
}
	catch(e){
	Connect.RollbackTrans()
	isok=false
}
if (isok){
	Connect.CommitTrans()
}
Response.Redirect("admarea.asp")
%>

