<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%

var smi_id=1


var hid=parseInt(Request("hid"))
if (isNaN(hid)) {Response.Redirect("index.asp")}

var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}

if (hid==0) {Response.Redirect("index.asp")}

var hname=""
var hiname=""
var url=""
var pid=0
var pname=""
var pdat=""
var autor=""
var digest=""
var sminame=""
var period=0
var pos=0
var lpag=0
var pp=0
var hdd=0
var path=""
var nm=""
var hadr=""
var imgname=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var nid=0
var name=""
var ndat=""
var nadr=""
var per=0
var kvopub=0
var isnews=1
var tpm=1000
var usok=false
var blokname=""
var ishtml=0
var urlcount=0
var msgcount=0

if (String(Session("id_mem"))=="undefined") {
	if (Session("tip_mem_pub")<3) {usok=true}
	tpm=Session("tip_mem_pub")
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			tpm=0
		}
		Records.Close()
	} else {
		usok=true
		tpm=0
	}
}

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

tit=sminame

Records.Source="Select * from heading where id="+hid+" and smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
isnews=Records("ISNEWS").Value
Records.Close()

var hidimg=""
hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		path=nm+" "+path
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
		lpag=Records("PAGE_LENGTH").Value
		hidimg=TextFormData(Records("PICTURE").Value,"")
	}
	else {
		path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> > "+path
	}
	hdd=Records("HI_ID").Value
	Records.Close()
}

path="<a href=\"index.asp\">"+sminame+"</a> > "+path
tit+=" | "+hiname

if (pg>0) {pp=pg*lpag}

if (hidimg != "") {
	if (!fs.FileExists(HeadImgPathR+hidimg)) { hidimg="" }
	else {hidimg=HeadImgPath+hidimg}
}
sql="Select Count(*) as kvo from url t1, catarea t2 where t1.state=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlcount=Records("KVO").Value
}
Records.Close()

sql="Select Count(*) as kvo from trademsg t1, trade_subj t2 where t1.state=0 and t1.date_end>='TODAY' and t1.trade_subj_id=t2.id and t2.marketplace_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	msgcount=Records("KVO").Value
}
Records.Close()
%>
<%
if (isnews==1){
Records.Source="Select * from publication where heading_id="+hid+"and state=1 and public_date>='TODAY'-"+period+" and public_date<='TODAY' order by public_date desc, id desc"
} else {
Records.Source="Select * from publication where heading_id="+hid+"and state=1 and public_date<='TODAY' order by public_date desc, id desc"
}
Records.Open()
while (!Records.EOF && pos<=lpag*(1+pg))
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	if  (pos>=pp) {
		imgname=PubImgPath+pid+".gif"
    	if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
		if (imgname=="") {
			imgname=PubImgPath+pid+".jpg"
			if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
		}
%>document.write('[<%=pdat%>] <a href="http://72rus.ru/<%=url%>" target="_blank"><%=pname%></a><br>');
<%
	}
	Records.MoveNext()
	pos+=1
} 
Records.Close()
%>
