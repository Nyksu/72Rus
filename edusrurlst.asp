<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var id=parseInt(Request("nik"))
var setst=parseInt(Request("st"))

if (isNaN(setst)) {Response.Redirect("admarea.asp?usr=1")}
if (isNaN(id)) {id=0}
if (id==0) {Response.Redirect("admarea.asp?usr=1")}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="edusrurlst.asp?nik="+id+"&st="+setst
Response.Redirect("login.asp")
}

var hid=0
var isok=true
var stat=0
var newstat=5

sql="Select * from host_url where id="+id
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("admarea.asp?usr=1")
}
stat=Records("STATE").Value
Records.Close()

newstat=setst

if (stat==1 && setst==1) {newstat=0}
if (stat==2 && setst==1) {newstat=0}

sql="Update host_url set state="+newstat+" where id="+id

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
Response.Redirect("admarea.asp?usr=1")
%>

