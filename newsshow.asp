<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=1
// +++  smi_id - ��� ��� � ������� SMI !!

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {Response.Redirect("index.asp")}

var hid=0
var sminame=""
var tit=""
var hdd=0
var nm=""
var hiname=""
var period=0
var path=""
var hadr=""
var pname=""
var pdat=""
var autor=""
var news=""
var imgname=""
var imgLname=""
var imgRname=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var sos=""
var sql=""
var ishtml=0
var usok=false
var id_usr=0

if (String(Session("id_mem"))=="undefined") {
	if (String(Session("id_mem_pub"))=="undefined") {
		Session("backurl")="newsshow.asp?pid="+pid
		Response.Redirect("logpub.asp")
	}
	if (Session("tip_mem_pub")<7) {usok=true}
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

if (!usok) {Response.Redirect("newshow.asp?pid="+pid)}

Records.Source="Select t1.* from publication t1, heading t2 where t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("pubarea.asp")
}
hid=Records("heading_id").Value
pname=String(Records("NAME").Value)
pdat=Records("PUBLIC_DATE").Value
autor=TextFormData(Records("AUTOR").Value,"")
ishtml=TextFormData(Records("ISHTML").Value,"0")
Records.Close()

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

tit=sminame

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
	}
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> | "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

var ddt = new Date()
var dt=""
var str=""
var sumdat=Server.CreateObject("datesum.DateSummer")

str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
dt="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
dt=str+dt
dt=sumdat.SummToDate(dt,"-"+period)
if (sumdat.DateComparing(dt,pdat) > 0) {sos="� ������"} else {sos="� ����������"}

path="<a href=\"index.asp\">"+sminame+"</a> | "+path
tit+=" | "+hiname

filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}

imgname=PubImgPath+pid+".gif"
if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
if (imgname=="") {
	imgname=PubImgPath+pid+".jpg"
	if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
}

imgLname=PubImgPath+"l"+pid+".gif"
if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
if (imgLname=="") {
	imgLname=PubImgPath+"l"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
}

imgRname=PubImgPath+"r"+pid+".gif"
if (!fs.FileExists(PubFilePath+"r"+pid+".gif")) { imgRname="" }
if (imgRname=="") {
	imgRname=PubImgPath+"r"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"r"+pid+".jpg")) { imgRname="" }
}

%>

<html>
<head>
<title><%=tit%> | <%=pname%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="90%" border="1">
  <tr> 
    <td>&nbsp;<%=path%></td>
  </tr>
</table>
<p>�������: <font color="#003399"><b><font size="3"><%=tit%></font></b></font></p>
<table width="90%" border="1" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#000099" colspan="2"> 
      <p><b><font color="#FFFF00"> || <%=pname%> || <font color="#FFFFFF"><%=sos%></font></font></b></p>
    </td>
  </tr>
  <tr> 
    <td bordercolor="#CCCCCC" width="81%" height="65"> 
      <p>���� ����������: <%=pdat%><br>
        �����: <%=autor%></p>
      <%=news%> </td>
    <td bordercolor="#CCCCCC" width="20%" height="65" valign="top"> 
      <div align="center"> 
        <%if (imgname != "") {%>
        <img src="<%=imgname%>" > 
        <%}else{%>
        . 
        <%}%>
      </div>
    </td>
  </tr>
</table>
<table width="90%" border="1" bordercolor="#FFFFFF">
  <tr bordercolor="#CCCCCC"> 
    <td width="50%" valign="top"> 
      <div align="center"> 
        <%if (imgLname != "") {%>
        <img src="<%=imgLname%>" > 
        <%}else{%>
        . 
        <%}%>
      </div>
    </td>
    <td width="50%" valign="top"> 
      <div align="center"> 
        <%if (imgRname != "") {%>
        <img src="<%=imgRname%>" > 
        <%}else{%>
        . 
        <%}%>
      </div>
    </td>
  </tr>
</table>
<p><a href="addpubimg.asp?pid=<%=pid%>">�������� 
  ����������</a></p>
<p><a href="index.html">��������� �� ������� ��������</a></p>
<p><a href="pubarea.asp">��������� � �������</a></p>
<p><a href="index.asp">��������� � �����������</a></p>
<p>&nbsp;</p>
</body>
</html>
