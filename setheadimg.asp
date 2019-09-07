<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
// ********************
var smi_id=1
// +++  smi_id -  SMI !!

var tpm=1000
var sql=""
if (String(Session("id_mem"))=="undefined") {
	if (String(Session("tip_mem_pub"))=="undefined") {Response.Redirect("index.asp")}
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
if (tpm>3) {Response.Redirect("index.asp")}

var hid=parseInt(Request("hid"))
if (isNaN(hid)){Response.Redirect("index.asp")}
var img=String(Request("img"))
if (img=="undefined"){Response.Redirect("headimglst.asp?hid="+hid)}

var isok=true
sql="Update HEADING set PICTURE='"+img+"' where id="+hid+" and smi_id="+smi_id
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
	Response.Redirect("pubheading.asp?hid="+hid)
}
else {Response.Redirect("headimglst.asp?hid="+hid)}

%>
