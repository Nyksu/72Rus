<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->


<%

// ��� ������� ��� ���... �� ������ �������� ��� � ������ ������!!
var smi_id=1
// +++  smi_id - ��� ��� � ������� SMI !!

var hid=0
var hname=""
var url=""
var nid=0
var name=""
var ndat=""
var nadr=""
var per=0
var kvopub=0
var pname=""
var pdat=""
var autor=""
var digest=""
var imgLname=""
var imgname=""
var path=""
var hdd=0
var hadr=""
var nm=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var isnews=1
var blokname=""
var tpm=1000
var usok=false
var ishtml=0
var urlname=""
var urlid=0
var urlabout=""
var daterenew=""
var urladr=""
var urlcount=0
var msgcount=0
var sminame=""
var pid=0
var news=""

if (String(Session("urlstat"))=="undefined") {Session("urlstat")="index.asp"}

var tps=parseInt(Request("tps"))
var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
if (isNaN(tps)) {tps=1}
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=parseInt(Request("pg"))
if (isNaN(pg)) {pg=0}


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

var getData = Server.CreateObject("SOFTWING.ASPtear")
var dirData=""

sql="Select t1.* from url t1, catarea t2 where t1.recl_id=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog

Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
Records.Close()
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

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()
%>
<html>
<head>
<title>72RUS.RU ��������� ������ - ������, �������������� ������. ����������, ������� ������, �������, ���������� ����������</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 500; margin:  3px 3px 3px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
<SCRIPT language=JavaScript>


DOM=(document.getElementById)?1:0;
NS4=(document.layers)?1:0;
IE4=(document.all)?1:0;

function CHange_city(nomer){
        if(NS4){document.images['weather'].src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
        if(IE4){document.all['weather'].src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
        if(DOM){getElementsByName('weather').src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
}

</SCRIPT>
<script language="JavaScript">
 function goToUrl(){
  if(NS4)location.href="http://www.gismeteo.ru/weather/towns/"+document.select["City_pogoda"].value+".htm";
  if(IE4)location.href="http://www.gismeteo.ru/weather/towns/"+document.all["City_pogoda"].value+".htm";
  if(DOM)location.href="http://www.gismeteo.ru/weather/towns/"+getElementById("City_pogoda").value+".htm";
 }
</script>
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> 
        <%
// � ���������� bk ���������� ��� ����� ��������
var bk=30
// �� �������� ��� ������!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
        <font color="#000000"><%=blokname%> </font> </p>
    </td>
    <td width="170"> 
      <p class="menu01"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <A href="#" onclick="window.external.AddFavorite(parent.location,document.title)"><font color="#000000">�������� 
        � ���������</font></A></p>
      </td>
    <td align="center" width="200"> 
      <p class="menu01"><a href="admarea.asp"><img src="images/e06.gif" width="16" height="9" alt="" border="0"></a> 
        <font color="#000000">����������� �� �����: <%=Application("visitors")%></font></p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"> 
      <img src="images/72rus.gif" width="170" height="87" alt="72RUS.RU ��������� ������ - �������������� ������ "> 
    </td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="banshow.asp?rid=1"></script>
    </td>
  </tr>
  <tr bgcolor="#FF6600"> 
    <td colspan="3" height="1"></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td valign="top" bgcolor="#0065A8" background="images/fon_menu04.gif" width="172"> 
      <%
// ������ �������� ��������
isnews=1
// ���� ���������� ������� ������� �� �������� �� ���������� � ����

var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
	}
	recs.Close()
	kvopub=0
	if (name!="") {
		recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%> 
              [<%=kvopub%>]</a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="messages.asp">���������� [<%=msgcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="lentamsg.asp">���������� 
              [�����]</a> </p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auto.72rus.ru">���� 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auction.72rus.ru">������� 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="air_russia.asp">���� 
              ����������</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="Rail_roads.asp">���������� 
              �������</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="http://bn.72rus.ru">������������ 72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td valign="top"><img src="images/top01.gif" width="170" height="19" alt="" border="0"> 
            <img src="images/fon_menu04.gif" width="172" height="1"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center"> 
            <p class="menu01">&nbsp; </p>
          </td>
        </tr>
      </table>
    </td>
    <td width="300" valign="top" bgcolor="CECECE"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
        <tr> 
          <td width="300" valign="top" bgcolor="#0365AD"> 
            <%
// � ���������� bk ���������� ��� ����� ��������
var bk=30
// �� �������� ��� ������!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
}

%>
            <h1><font color="#FFFFFF"><%=pname%></font></h1>
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr bgcolor="#003366"> 
          <td align=left width=300 bgcolor="#666666" valign="top" height="1"></td>
        </tr>
        <tr> 
          <td width="377"> 
            <%
// � ���������� bk ���������� ��� ����� ��������
var bk=30
// �� �������� ��� ������!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
}

%>
            <%if (imgLname != "") {%>
            <a href="<%=url%>"><img src="<%=imgLname%>" border="0" alt="<%=pname%>" ></a> 
            <%}%>
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" height="87">
        <tr> 
          <td background="images/fon02.gif"> 
            <%
// � ���������� bk ���������� ��� ����� ��������
var bk=30
// �� �������� ��� ������!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
    news=""
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
}

%>
            <a href="<%=url%>"> 
            <p class="digest"><%=digest%></p>
            </a> 
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" width="172" background="images/fon_menu04.gif" bgcolor="#0065A8"> 
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" width="170"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
              ������� [<%=urlcount%>]</a></p>
          </td>
        </tr>
      </table>
      <%
// ������ �������� ��������
isnews=0
// ���� ���������� ������� ������� �� �������� �� ���������� � ����
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%></a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <table width="141" border="0" cellspacing="0" cellpadding="0" bordercolor="#003366" align="center">
        <form name="form1" method="post" action="">
          <tr> 
            <td align="center" bordercolor="#FFFFFF"> 
              <select name="City_pogoda" onChange="CHange_city(this.options[this.selectedIndex].value)"
			style="BACKGROUND-COLOR: #FFF5CE; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 101px" >
                <option value="27612">������</option>
                <option value="28440">������������</option>
                <option value="28367" selected>������</option>
                <option value="23748">�������</option>
                <option value="99981">�����</option>
                <option value="23848">�����������</option>
                <option value="23471">�������������</option>
                <option value="23358">����� �������</option>
                <option value="99968">��������</option>
                <option value="23849">������</option>
                <option value="23933">�����-��������</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td align="center" bordercolor="#FFFFFF" height="105"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="�����: ��������� ������ � ������" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
          </tr>
        </form>
      </table>
      <br>
      <img src="images/fon_menu04.gif" width="172" height="1"></td>
    <td background="images/bg.gif" valign="top"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td background="images/fon_menu07.gif" height="30" width="170" align="center"> 
            <p class="menu01"> <a href="regmemurl.asp"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              �����������</a></p>
          </td>
          <td height="30" align="center" background="images/fon_menu08.gif"><!--LiveInternet counter--><script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script><!--/LiveInternet--></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" background="" width="100%">
        <form name="form" method="post" action="usrarea.asp">
          <tr valign="middle"> 
            <td width="48" align="center"> 
              <p>&nbsp;<b>����</b></p>
            </td>
            <td> 
              <input maxlength=20 name="logname" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/name.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=edit>
              <input maxlength=25 name="psw" 
            style="BACKGROUND-COLOR: #cfdbe7; BACKGROUND-IMAGE: url(headimg/pass.gif); BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 72px" 
            type=password>
              <input type="Image" src="images/b_go.gif" width="19" height="25" alt="" border="0" hspace="10" align="absbottom" name="Enter">
            </td>
          </tr>
        </form>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <%
// � ���������� bk ���������� ��� ����� ��������
var bk=31
// �� �������� ��� ������!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
            <%
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=PubImgPath+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=PubImgPath+pid+".jpg"
		if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
	}
	path=""
	hid=String(Records("HEADING_ID").Value)
	hdd=hid
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\"><font color=\"#FFFFFF\">"+nm+"</font></a> ] "+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}
%>
            <hr size=3 noshade width="90%" align="left">
            <p><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="<%=url%>" target="_blank"><%=pname%></a> <%=digest%></p>
            <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td colspan="3" height="1" bgcolor="#666666"></td>
  </tr>
  <tr> 
    <td colspan="3" height="19"> 
      <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <img src="images/px1.gif" width="1" height="1" alt="" border="0"><a href="http://www.rusintel.ru/newshow.asp?pid=2496">����������� 
        ������� RU COM NET</a> <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="http://www.rusintel.ru/">���������� �������� ������</a> <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2256">�������</a> 
        <img src="images/e06.gif" width="16" height="9" alt="" border="0"><a href="http://www.72rus.ru/newshow.asp?pid=728"> 
        ���������� �������</a></p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center"> 
            <p class="menu01"><a href="http://www.rusintel.ru/goodslst.asp?divis=4&hid=2244" target="_blank"><img src="HeadImg/rol20.gif" width="120" height="72" align="middle" border="0" alt="�������� ����� ��� � ������"></a> 
            </p>
          </td>
        </tr>
      </table>
    </td>
    <td background="images/fon02.gif" height="87" width="2"><img src="images/e01.gif" width="2" height="87" alt="" border="0"></td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="banshow.asp?rid=16"></script>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr> 
    <td valign="top" width="168"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
 <tr> 
          <td align="CENTER">
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
              <tr> 
                <td bgcolor="#006699" align="center"> 
                  <p class="menu01"> �����</p>
                </td>
              </tr>
              <tr> </tr>
            </table>
            <table border="0" cellspacing="0" cellpadding="0" width="170" align="center">
              <tr> 
                <form name="form" method="post" action="search.asp">
                  <td valign="middle" align="center"> 
                    <p> &nbsp; 
                      <input type="text" name="sch" size="40" value="" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 150px" >
                      <input type="hidden" name="wrds" value="1">
                      <input type="submit" name="Findit" style="BACKGROUND-IMAGE: url(headimg/lence.gif); COLOR: #003366; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 55px; HEIGHT: 20px" value="     �����">
                    </p>
                  </td>
                </form>
              </tr>
            </table>
 </td>
</tr>
      </table>
      <table border="0" cellspacing="0" width="100%" cellpadding="0">
        <tr> 
          <td bgcolor="#006699" align="center"> 
            <p class="menu01"> 
              <%
// � ���������� bk ���������� ��� ����� ��������
var bk=32
// �� �������� ��� ������!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
              <%=blokname%></p>
          </td>
        </tr>
        <tr> </tr>
      </table>
      <font color="#000000"> 
      <%
// � ���������� bk ���������� ��� ����� ��������
var bk=32
// �� �������� ��� ������!!

Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)


var news=""
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


%>
      </font> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr> 
          <td bgcolor="#FFFFFF"> 
            <p><a href="newshow.asp?pid=<%=pid%>"> <%=pname%> </a></p>
          </td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
%>
      <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="167" cellpadding="0">
        <tr> 
          <td bgcolor="#006699" align="center"> 
            <p class="menu01"> 
              <%
// � ���������� bk ���������� ��� ����� ��������
var bk=34
// �� �������� ��� ������!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
              <%=blokname%></p>
          </td>
        </tr>
        <tr> </tr>
      </table>
      <%
// � ���������� bk ���������� ��� ����� ��������
var bk=34
// �� �������� ��� ������!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}

var news=""
filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+=ts.ReadLine()
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}


%>
      <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td bgcolor="#FFFFFF"><%=news%></td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="168">
        <tr> 
          <td valign="top"><img src="images/top01.gif" width="167" height="19" alt="" border="0"> 
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" background="images/bg.gif"> 
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr valign="top"> 
          <td width="3" bgcolor="#AFC0D0">&nbsp;</td>
          <td width="468"> 
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
              <tr> 
                <td bgcolor="#006699" align="center"> 
                  <p class="menu01"><b> 
                    <%
// � ���������� bk ���������� ��� ����� ��������
var bk=28
// �� �������� ��� ������!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
                    <%=blokname%></b></p>
                </td>
              </tr>
              <tr> </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bordercolor="#CCCCCC"> 
                <td bordercolor="#CCCCCC" bgcolor="#FFFFFF"> 
                  <dl> 
                    <%
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2, block_news t3 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+"and t2.block_news_id=t3.id and t3.smi_id="+smi_id+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	imgname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgname=PubImgPath+pid+".gif"
    if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
	if (imgname=="") {
		imgname=PubImgPath+pid+".jpg"
		if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
	}
	path=""
	hid=String(Records("HEADING_ID").Value)
	hdd=hid
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"pubheading.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\"><font color=\"#FF0000\">"+nm+"</font></a>"+path
	hdd=recs("HI_ID").Value
	recs.Close()
	}
%>
                    <dd> 
                      <li> 
                        <p> 
                          <%if (imgname != "") {%>
                        <table width="0" border="1" cellspacing="5" cellpadding="0" bordercolor="#FFFFFF" align="left">
                          <tr> 
                            <td width="100" bordercolor="#003366"><a href="<%=url%>?pid=<%=pid%>"><img src="<%=imgname%>" alt="<%=pname%>" border="0" width="100" height="70" ></a></td>
                            <td bordercolor="#003366" width="5" bgcolor="#FF6600">&nbsp;</td>
                          </tr>
                        </table>
                        <%}else{%>
                        <%}%>
                        <a href="<%=url%>?pid=<%=pid%>"><%=pname%></a><br>
                        [<%=pdat%>] <%=digest%><BR clear=all>
                      </LI>
                      <%
	Records.MoveNext()
} 
Records.Close()
delete recs
%>
                  </dl>
                </td>
              </tr>
            </table>
          </td>
          <td width="3" bgcolor="#AFC0D0">&nbsp;</td>
          <td style="background-position: top; background-repeat: repeat-x;" bgcolor="#FFFFFF"> 
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
              <tr> 
                <td bgcolor="#006699" align="center"> 
                  <p class="menu01">����� ����� � ��������</p>
                </td>
              </tr>
              <tr> </tr>
            </table>
<%
sql="Select t1.* from url t1, catarea t2 where t1.recl_id=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog

Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
%>
			<p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp; 
              <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%> 
                                   <%
Records.MoveNext()
Records.Close()
delete recs
%>
<%
sql="Select t1.* from url t1, catarea t2 where t1.recl_id=2 and t1.catarea_id=t2.id and t2.catalog_id="+catalog

Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
%>
			<p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp; 
              <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%> 
                                   <%
Records.MoveNext()
Records.Close()
delete recs
%>
<%
sql="Select t1.* from url t1, catarea t2 where t1.recl_id=3 and t1.catarea_id=t2.id and t2.catalog_id="+catalog

Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
%>
			<p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp; 
              <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%> 
                                   <%
Records.MoveNext()
Records.Close()
delete recs
%>
<%
sql="Select t1.* from url t1, catarea t2 where t1.recl_id=4 and t1.catarea_id=t2.id and t2.catalog_id="+catalog

Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
%>
			<p class="left"><img src="images/dot_g.gif" width="5" height="5" alt="" border="0" align="middle">&nbsp; 
              <a href="<%=urladr%>" target="_blank"><%=urlname%></a> - <%=urlabout%> 
                                   <%
Records.MoveNext()
Records.Close()
delete recs
%>
</p>
            <p class="left" align="right"> <a href="addurl.asp">�������� ����</a></p>
<hr noshade size="1" width="90%">

              <%
// � ���������� bk ���������� ��� ����� ��������
var bk=38
// �� �������� ��� ������!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}

	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hdd=recs("HI_ID").Value
	recs.Close()
var news=""
filnam=PubFilePath+pid+".pub"
if (!fs.FileExists(filnam)) { filnam="" }

if (filnam != "") {
	ts= fs.OpenTextFile(filnam)
	if (ishtml==0) {
	while (!ts.AtEndOfStream){
		news+=ts.ReadLine()
	}
	} else {news=ts.ReadAll()}
	ts.Close()
}
}

%>

            <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
              <tbody> 
              <tr> 
                <td align="center" ><%=news%></td>
              </tr>
              </tbody> 
            </table>
            <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
          </td>
          <td width="3" bgcolor="#AFC0D0">&nbsp;</td>
          <td width="120" align="center"> 
            <table border="0" bordercolor="#FFFFFF" cellspacing="0" width="100%" cellpadding="0">
              <tr> 
                <td bgcolor="#006699" align="center"> 
                  <p class="menu01"><b> 
                    <%
// � ���������� bk ���������� ��� ����� ��������
var bk=35
// �� �������� ��� ������!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
                    <%=blokname%></b></p>
                </td>
              </tr>
              <tr> </tr>
            </table>
            <script language="javascript" src="banshow.asp?rid=3"></script>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td height="1" bgcolor="#666666"></td>
  </tr>
  <tr> 
    <td height="19" bgcolor="#AFC0D0" align="center">&nbsp; </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td colspan="3"><img src="images/px1.gif" width="1" height="1" alt="" border="0"></td>
  </tr>
  <tr> 
    <td height="70" width="180"> 
      <p align="right">&nbsp;</p>
    </td>
    <td align="center"> 
      <p class="menu02">| 
        <%
// ������ �������� ��������
isnews=1
// ���� ���������� ������� ������� �� �������� �� ���������� � ����

var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
	}
	recs.Close()
	Records.MoveNext()
%>
        <a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <%
// ������ �������� ��������
isnews=0
// ���� ���������� ������� ������� �� �������� �� ���������� � ����
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hid
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hid+" and public_date<='TODAY' order by public_date desc, id desc"
	}
	recs.Open()
	if (!recs.EOF) {
		nid=String(recs("ID").Value)
		name=String(recs("NAME").Value)
		nadr=TextFormData(recs("URL").Value,"newshow.asp")
		nadr+="?pid="+nid
		ndat=recs("PUBLIC_DATE").Value
	} else {
		nid=0
		name=""
		nadr=""
		ndat=""
}
	recs.Close()
	
	Records.MoveNext()
%>
        <a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <br>
        | <a href="http://auto.72rus.ru">����.72rus</a> | <a href="http://www.auction.72rus.ru/">�������</a> 
        | <a href="messages.asp">����������</a> | <a href="Rail_roads.asp">����������</a> 
        | <a href="catarea.asp">��������� �������</a> | 
    </td>
    <td width="180">
<p>������� � <a href="http://www.rusintel.ru">��� ��������</a> <br>
        WWW.72RUS.RU � 2002-2004 </p>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=2"></script>
</div>
<hr size="1">
<div align="center"> 
  <%
// � ���������� bk ���������� ��� ����� ��������
var bk=33
// �� �������� ��� ������!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
	pid=String(Records("ID").Value)
	pname=String(Records("NAME").Value)
	url=TextFormData(Records("URL").Value,"newsshow.asp")
	url+="?pid="+pid
	pdat=Records("PUBLIC_DATE").Value
	autor=TextFormData(Records("AUTOR").Value,"")
	digest=TextFormData(Records("DIGEST").Value,"")
	imgLname=PubImgPath+"l"+pid+".gif"
    if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
	if (imgLname=="") {
		imgLname=PubImgPath+"l"+pid+".jpg"
		if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
	}
	path=""
	//hid=String(Records("HEADING_ID").Value)
	//hdd=hid
	hdd=String(Records("HEADING_ID").Value)
	while (hdd>0) {
	recs.Source="Select * from heading where id="+hdd
	recs.Open()
	nm=String(recs("NAME").Value)
	hadr=TextFormData(recs("URL").Value,"vvr_list.asp.asp")
	path="<a href=\""+hadr+"?hid="+hdd+"\">"+nm+"</a> &gt; "+path
	hdd=recs("HI_ID").Value
	recs.Close()
var news=""
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
}

%>
  <%=news%> 
  <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
  <!--LiveInternet logo--><a href="http://www.liveinternet.ru/click" target=liveinternet><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="liveinternet.ru: �������� ����� ����� �� 24 ����, ����������� �� 24 ���� � �� �������"></a><!--/LiveInternet-->
 
  <script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=46088&im=105&r="+escape(document.referrer)+"&pg="+
escape(window.location.href);
document.cookie="hotlog=1; path=/"; hotlog_r+="&c="+(document.cookie?"Y":"N");
</script>
  <script language="javascript1.1">
hotlog_js="1.1";hotlog_r+="&j="+(navigator.javaEnabled()?"Y":"N")</script>
  <script language="javascript1.2">
hotlog_js="1.2";
hotlog_r+="&wh="+screen.width+'x'+screen.height+"&px="+
(((navigator.appName.substring(0,3)=="Mic"))?
screen.colorDepth:screen.pixelDepth)</script>
  <script language="javascript1.3">hotlog_js="1.3"</script>
  <script language="javascript">hotlog_r+="&js="+hotlog_js;
document.write("<a href='http://click.hotlog.ru/?46088' target='_top'><img "+
" src='http://hit3.hotlog.ru/cgi-bin/hotlog/count?"+
hotlog_r+"&' border=0 width=88 height=31 alt=HotLog></a>")</script>
  <noscript><a href=http://click.hotlog.ru/?46088 target=_top><img
src="http://hit3.hotlog.ru/cgi-bin/hotlog/count?s=46088&im=105" border=0 
width="88" height="31" alt="HotLog"></a></noscript> 
  <!-- /HotLog -->
  <!--Begin of HMAO RATINGS-->
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"><img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
</div>
</body>
</html>
