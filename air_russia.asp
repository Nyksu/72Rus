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
<title>���������� ��������� - ������ �� ������������ - 72RUS ������</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
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
<script language="JavaScript">
var flag_popup = 1;

var type=navigator.appName
 if (type=="Netscape") 
var lang = navigator.language ;
else 
var lang = navigator.userLanguage;
var lang = lang.substr(0,2) ;

function popup_off() {
	flag_popup = 0;
}

function popup() {
	if (flag_popup == 1 & lang != "ru") {
		window.open('http://www.rusintel.ru/english.asp');
	}
else if (lang == "ru") window.open('http://www.rusintel.ru');	
}

</script>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onUnload=popup();>
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        <%=sminame%> : ���������� ���������, ������</font></p>
    </td>
    <td width="170"> 
      <p class="menu01"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="#" onClick="window.external.AddFavorite(parent.location,document.title)"><font color="#000000">�������� 
        � ���������</font></a></p>
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
      <a href="/"><img src="images/72rus.gif" width="170" height="87" alt="72RUS.RU ��������� ������ - �������������� ������ " border="0"></a> 
    </td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="banshow.asp?rid=4"></script>
    </td>
  </tr>
  <tr bgcolor="#FF6600"> 
    <td colspan="3" height="1"></td>
  </tr>
</table>
<table bgcolor=#003366 border=0 cellpadding=0 cellspacing=0 height=1 
width="100%">
  <tbody> 
  <tr> 
    <td align=right bgcolor="#CCCCCC"> </td>
  </tr>
  </tbody> 
</table>
<table width="173" border="0" cellspacing="0" cellpadding="0" align="left">
  <tr> 
    <td width="172" align="left" valign="top"> 
      <%
// ������ �������� ��������
isnews=1
// ���� ���������� ������� ������� �� �������� �� ���������� � ����

var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hdd
	if (isnews==1) {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hdd+" and public_date>='TODAY'-"+per+" and public_date<='TODAY' order by public_date desc, id desc"
	} else {
	recs.Source="Select * from PUBLICATION where state=1 and heading_id="+hdd+" and public_date<='TODAY' order by public_date desc, id desc"
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
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%> 
              [<%=kvopub%>]</a></p>
          </td>
        </tr>
      </table>
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
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="pubheading.asp"}
	url+="?hid="+hdd
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%></a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="messages.asp">���������� [<%=msgcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="lentamsg.asp">���������� 
              [�����]</a> </p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" width="170" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
              ������� [<%=urlcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auto.72rus.ru">���� 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auction.72rus.ru">������� 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="air_russia.asp">���� 
              ����������</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="Rail_roads.asp">���������� 
              �������</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
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
      <font color="#000000"> </font> 
      <table width="100%" border="0" cellspacing="0" height="60" align="center" cellpadding="0">
        <tr> 
          <td align="center"> 
            <%
// � ���������� bk ���������� ��� ����� ��������
var bk=38
// �� �������� ��� ������!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	puid=String(Records("ID").Value)
	ishtml2=TextFormData(Records("ISHTML").Value,"0")
	filnam=PubFilePath+puid+".pub"
	if (!fs.FileExists(filnam)) { filnam="" }

	if (filnam != "") {
		ts= fs.OpenTextFile(filnam)
		if (ishtml2==0){
			while (!ts.AtEndOfStream){
				news_bl+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
			}
		} else {news_bl=ts.ReadAll()}
		ts.Close()
	}

%>
            <%=news_bl%> 
            <%
Records.MoveNext()
} 
Records.Close()
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="120" border="0" cellspacing="0" cellpadding="0" align="right" height="750">
  <tr> 
    <td width="150" valign="top" align="center"> 
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" align="center" bgcolor="#FB9700"> 
            <p class="menu01"> ������</p>
          </td>
        </tr>
      </table>
      <table width="120" border="0" cellspacing="0" cellpadding="0" bordercolor="#003366" align="center">
        <form name="form" method="post" action="">
          <tr> 
            <td align="center" bgcolor="#6699CC" height="20" valign="bottom"> 
              <select name="select" onChange="CHange_city(this.options[this.selectedIndex].value)"
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
            <td align="center" height="106" bgcolor="#6699CC" valign="top"><a href="javascript:goToUrl()"><img width=100 height=100 border=0 alt="�����: ��������� ������ � ������" name="weather" src="http://img.gismeteo.ru/informer/28367-6.GIF"></a></td>
          </tr>
        </form>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" align="center" bgcolor="#FB9700"> 
            <p class="menu01">���� $ �� ��</p>
          </td>
        </tr>
      </table>
      <table width="120" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" height="105" bgcolor="#6699CC"><a href="http://www.informer.ru/cgi-bin/redirect.cgi?id=177_1_1_33_40_1-0&url=http://www.rbc.ru&src_url=usd/eur_cb_forex_000066_88x90.gif"
target="_blank"><img src="http://pics.rbc.ru/img/grinf/usd/eur_cb_forex_000066_88x90.gif " width=88 height="90" border=0></a></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" background="images/top01.gif" align="center" bgcolor="#FF9900"> 
            <p class="menu01"> 
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
              <%=blokname%></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <p> 
        <script language="javascript" src="banshow.asp?rid=5"></script>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" height="850" align="center">
  <tr> 
    <td valign="top" width="1"></td>
    <td valign="top" align="center"> 
      <table width="100%" border="0" cellspacing="0" bordercolor="#003366">
        <tr bgcolor="#FBF8D7"> 
          <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle"> 
            <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="index.asp">72RUS.RU</a> / ���������� ��������� /</p>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCFF" align="center">
        <tr bordercolor="#CCCCFF" valign="middle" align="center"> 
          <td bordercolor="#CCCCFF" height="11"> 
            <table width="97%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
              <tr> 
                <td bgcolor="#FF9900" align="left" background="images/fon_menu08.gif"> 
                  <h1><b><font color="#FFFFFF">���������� ���������</font></b>&nbsp;</h1>
                </td>
              </tr>
            </table>
            
          </td>
        </tr>
        <tr> 
          <td class=n3 valign="top"> 
            <form action="http://www.polets.ru/cgi-bin/sh.pl" method=get name="f" target="_blank">
              <table border=0 cellpadding=0 cellspacing=0>
                <tbody> 
                <tr valign="top"> 
                  <td height="328"> 
                    <table border=0 cellpadding=5 cellspacing=0 
                  width=99% align="center">
                      <tbody> 
                      <tr> 
                        <td rowspan=10 width=15>&nbsp;</td>
                        <td class=n2b> 
                          <input name="Mode" type=hidden value="Naprav">
                          <input name="Source" type=hidden  value="Brief">
                          <p><b>����� �����������:</b></p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2 bgcolor="#EBF3F5"> 
                          <p> 
                            <select name="CityFrom" size="1" class=text>
                              <option value="0">* �������� ����� 
                              <option value="all">* ��� ����������� 
                              <option value="abn">������ 
                              <option value="ahl">����� 
                              <option value="aau">����� 
                              <option value="ald">����� 
                              <option value="ala">������ 
                              <option value="amd">������� 
                              <option value="ade">�������-2 
                              <option value="any">������� 
                              <option value="ana">����� 
                              <option value="anv">������� 
                              <option value="aph">������� 
                              <option value="arh">����������� 
                              <option value="akl">������ 
                              <option value="asr">��������� 
                              <option value="aty'>������ 
                                      <option value="a{h">������� 
                              <option value="aqn">��� 
                              <option value="bgd">�������� 
                              <option value="bkt">������ 
                              <option value="lnj">�������� 
                              <option value="bak">���� 
                              <option value="ban">������� 
                              <option value="bcn">��������� 
                              <option  value="btg">������� 
                              <option value="bui">������ 
                              <option value="bqg">����� ���� 
                              <option value="bed">�������� 
                              <option value="blr">���������� 
                              <option value="bzk">��������� 
                              <option value="ber">�������� 
                              <option value="bng">������������ 
                              <option value="bsk">����� 
                              <option value="bi{">������ 
                              <option value="bg}">������������ 
                              <option value="bdb">������� 
                              <option value="brs">������ 
                              <option value="bug">�������� 
                              <option value="bgs">������ 
                              <option value="bur">����������� 
                              <option value="bua">������ 
                              <option value="prd">����� ���������� 
                              <option value="wrq">�������� 
                              <option value="war">����� 
                              <option value="weu">������� ����� 
                              <option value="whw">������������� 
                              <option value="wik">������� 
                              <option value="wwo">����������� 
                              <option value="wla">����������� 
                              <option value="wgg">��������� 
                              <option value="wld">���������� 
                              <option value="wgd">������� 
                              <option value="wkt">������� 
                              <option value="wrn">������� 
                              <option value="wtg">������� 
                              <option value="haj">�������� 
                              <option value="gdv">��������� 
                              <option value="g`m">����� 
                              <option value="gnv">������ 
                              <option value="dlc">������ 
                              <option value="dep">����������� 
                              <option value="dvh">���������-��� 
                              <option value="dpk">�������������� 
                              <option value="dnc">������ 
                              <option value="dnk">������� 
                              <option value="d{b">������� 
                              <option value="DUS">����������� 
                              <option value="esk">���� 
                              <option value="ekb">������������ 
                              <option value="eg~">��������� 
                              <option value="ewn">������ 
                              <option value="vig">������� 
                              <option value="zla">����� ��������� 
                              <option value="zpv">��������� 
                              <option value="zon">��������� 
                              <option value=znk>������� 
                              <option value="igr">������ 
                              <option value=irm>����� 
                              <option value="ivw">������ 
                              <option value="int">���� 
                              <option value="ikt">������� 
                              <option value="i{o">������-��� 
                              <option value="kzn">������ 
                              <option value="kz~">��������� 
                              <option value="kld">����������� 
                              <option value="klg">������ 
                              <option value="kgd">��������� 
                              <option value="kdd">�������� 
                              <option  value="krw">�������� 
                              <option value="kpm">��������� 
                              <option value="iew">���� 
                              <option value="krn">������� 
                              <option value="kio">����� 
                              <option value="k~g">����������� ������� 
                              <option value="k{n">������� 
                              <option value="kog">������� 
                              <option value="kzi">������� 
                              <option value="ksl">�����������-��-����� 
                              <option value="ktn">�������� 
                              <option value="kts">������ 
                              <option value="krr">��������� 
                              <option value="kkx">������������� 
                              <option value="kqa">���������� 
                              <option value="kgn">������ 
                              <option value="kus">����� 
                              <option value="kis">������� 
                              <option value="kyy">����� 
                              <option value="lsk">����� 
                              <option value="le{">����������� 
                              <option value="lip">������ 
                              <option value="lod">������ 
                              <option value="lwo">����� 
                              <option value="mdn">������� 
                              <option value="mgn">����� 
                              <option value="mgs">������������ 
                              <option value="mkp">������ 
                              <option value="mam">���� 
                              <option value="mrn">���������� 
                              <option value="mko">������� 
                              <option value="mhl">��������� 
                              <option 
                          value="mrw">����������� ���� 
                              <option 
                          value="msk">����� 
                              <option value="mir">������ 
                              <option 
                          value="mom">���� 
                              <option class="hl" value="mow">������ 
                              <option 
                          value="mtg">�������� 
                              <option value="mun">�������� 
                              <option 
                          value="mkm">��� �������� 
                              <option value="m{d">��� ������ 
                              <option value="m`n">������ 
                              <option 
                          value="ndm">����� 
                              <option value="in{">������� 
                              <option 
                          value="n~k">������� 
                              <option value="nmg">�������� 
                              <option 
                          value="nnr">������-��� 
                              <option 
                          value="nh~">���������� 
                              <option value="nln">������� 
                              <option 
                          value="nrg">�������� 
                              <option value="n`g">����������� 
                              <option 
                          value="nvg">������������ 
                              <option 
                          value="nvw">������������� 
                              <option 
                          value="nvk">���������� 
                              <option value="nvq">��������� 
                              <option 
                          value="nvs">������ �������� 
                              <option 
                          value="nlk">����������-��-����� 
                              <option 
                          value="nik">���������� 
                              <option 
                          value="nwk">����������� 
                              <option 
                          value="owb">����������� 
                              <option value="nur">����� ������� 
                              <option value="ngl">������� 
                              <option 
                          value="nrs">�������� 
                              <option value="noq">�������� 
                              <option 
                          value="nus">����� 
                              <option value="n`r">����� 
                              <option 
                          value="nqg">������ 
                              <option value="ods">������ 
                              <option 
                          value="ozr">������� 
                              <option value="omk">������� 
                              <option 
                          value="olk">��������� 
                              <option value="oln">������ 
                              <option 
                          value="ool">������ 
                              <option value="oms">���� 
                              <option 
                          value="ong">�������� 
                              <option value="osk">���� 
                              <option 
                          value="oso">������ 
                              <option value="oha">��� 
                              <option 
                          value="oht">������ 
                              <option value="pwl">�������� 
                              <option 
                          value="pan">������ 
                              <option value="pew">����� 
                              <option 
                          value="prm">����� 
                              <option value="ptz">������������ 
                              <option 
                          value="prl">�������������-���������� 
                              <option 
                          value="p~r">������ 
                              <option value="pin">��������� 
                              <option 
                          value="pts">����������� �������� 
                              <option value="pos">������ �������� 
                              <option value="plq">�������� 
                              <option 
                          value="rad">�������� 
                              <option value="rix">���� 
                              <option 
                          value="row">������-��-���� 
                              <option 
                          value="sky">�������� 
                              <option value="shd">�������� 
                              <option 
                          value="sm{">������ 
                              <option value="skd">��������� 
                              <option 
                          value="sag">������ 
                              <option class=hl 
                          value="spt">�����-��������� 
                              <option 
                          value="srn">������� 
                              <option value="sro">������� 
                              <option 
                          value="syh">�������� 
                              <option 
                          value="sen">������-�������� 
                              <option 
                          value="swe">������-������ 
                              <option value="ski">������ 
                              <option 
                          value="sel">���� 
                              <option value="sip">����������� 
                              <option 
                          value="sbo">�������� 
                              <option value="sog">��������� ������ 
                              <option value="soj">��������� 
                              <option 
                          value="son">��������� 
                              <option value="so~">���� 
                              <option 
                          value="srm">������-������� 
                              <option 
                          value="stw">���������� 
                              <option value="ist">������� 
                              <option 
                          value="sol">������ ����� 
                              <option 
                          value="spw">���������� 
                              <option value="stv">��������� 
                              <option 
                          value="sun">������ 
                              <option value="sur">������ 
                              <option 
                          value="syw">��������� 
                              <option value="tio">������� 
                              <option 
                          value="tas">������� 
                              <option value="tbs">������� 
                              <option 
                          value="tlq">����-���� 
                              <option value="tig">������ 
                              <option 
                          value="tsi">����� 
                              <option value="til">�������� 
                              <option 
                          value="txk">������ 
                              <option value="tsk">����� 
                              <option 
                          value="tau">���� 
                              <option value="trh">��������� 
                              <option 
                          value="t`m" selected>������ 
                              <option value="tsn">����-���� 
                              <option 
                          value="uln">����-����� 
                              <option value="ul|">����-��� 
                              <option 
                          value="ulk">��������� 
                              <option value="ura">���� 
                              <option 
                          value="ugn">������ 
                              <option value="usn">������ 
                              <option 
                          value="utk">����-������ 
                              <option 
                          value="ukg">����-����������� 
                              <option 
                          value="uk~">����-�������� 
                              <option 
                          value="uku">����-����� 
                              <option value="usk">����-��� 
                              <option 
                          value="usm">����-��� 
                              <option value="unr">����-���� 
                              <option 
                          value="uhj">����-��������� 
                              <option 
                          value="utc">����-������ 
                              <option value="ufa">��� 
                              <option 
                          value="uht">���� 
                              <option value="fgn">������� 
                              <option 
                          value="fra">���������-��-����� 
                              <option 
                          value="hbr">��������� 
                              <option value="hdy">������� 
                              <option 
                          value="has">�����-�������� 
                              <option 
                          value="hrk">������� 
                              <option value="hat">������� 
                              <option 
                          value="hel">��������� 
                              <option value="hrp">������� 
                              <option 
                          value="hdt">������� 
                              <option value="~ar">���� 
                              <option 
                          value="~be">��������� 
                              <option value="~lb">��������� 
                              <option 
                          value="~rw">��������� 
                              <option value="~rs">������� 
                              <option 
                          value="sht">���� 
                              <option value="~kd">�������� 
                              <option 
                          value="~mi">������� 
                              <option value="{ah">�������� 
                              <option 
                          value="{nq">������ 
                              <option value="{mt">������� 
                              <option 
                          value="|gt">��������� 
                              <option value="|li">������ 
                              <option 
                          value="`vk">����-�������� 
                              <option 
                          value="`vh">����-��������� 
                              <option 
                          value="qkt">������ 
                              <option value="qmb">������ 
                              <option value="qrl">���������</option>
                            </select>
                          </p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2b> 
                          <p><b>����� ��������:</b></p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2 bgcolor="#EBF3F5"> 
                          <p> 
                            <select name="CityTo" size="1" class=text>
                              <option value="0">* �������� ����� 
                              <option value="all">* ��� ����������� 
                              <option value="abn">������ 
                              <option value="ahl">����� 
                              <option value="aau">����� 
                              <option value="ald">����� 
                              <option value="ala">������ 
                              <option value="amd">������� 
                              <option value="ade">�������-2 
                              <option value="any">������� 
                              <option value="ana">����� 
                              <option value="anv">������� 
                              <option value="aph">������� 
                              <option value="arh">����������� 
                              <option value="akl">������ 
                              <option value="asr">��������� 
                              <option value="aty'>������ 
                                      <option value="a{h">������� 
                              <option value="aqn">��� 
                              <option value="bgd">�������� 
                              <option value="bkt">������ 
                              <option value="lnj">�������� 
                              <option value="bak">���� 
                              <option value="ban">������� 
                              <option value="bcn">��������� 
                              <option  value="btg">������� 
                              <option value="bui">������ 
                              <option value="bqg">����� ���� 
                              <option value="bed">�������� 
                              <option value="blr">���������� 
                              <option value="bzk">��������� 
                              <option value="ber">�������� 
                              <option value="bng">������������ 
                              <option value="bsk">����� 
                              <option value="bi{">������ 
                              <option value="bg}">������������ 
                              <option value="bdb">������� 
                              <option value="brs">������ 
                              <option value="bug">�������� 
                              <option value="bgs">������ 
                              <option value="bur">����������� 
                              <option value="bua">������ 
                              <option value="prd">����� ���������� 
                              <option value="wrq">�������� 
                              <option value="war">����� 
                              <option value="weu">������� ����� 
                              <option value="whw">������������� 
                              <option value="wik">������� 
                              <option value="wwo">����������� 
                              <option value="wla">����������� 
                              <option value="wgg">��������� 
                              <option value="wld">���������� 
                              <option value="wgd">������� 
                              <option value="wkt">������� 
                              <option value="wrn">������� 
                              <option value="wtg">������� 
                              <option value="haj">�������� 
                              <option value="gdv">��������� 
                              <option value="g`m">����� 
                              <option value="gnv">������ 
                              <option value="dlc">������ 
                              <option value="dep">����������� 
                              <option value="dvh">���������-��� 
                              <option value="dpk">�������������� 
                              <option value="dnc">������ 
                              <option value="dnk">������� 
                              <option value="d{b">������� 
                              <option value="DUS">����������� 
                              <option value="esk">���� 
                              <option value="ekb">������������ 
                              <option value="eg~">��������� 
                              <option value="ewn">������ 
                              <option value="vig">������� 
                              <option value="zla">����� ��������� 
                              <option value="zpv">��������� 
                              <option value="zon">��������� 
                              <option value=znk>������� 
                              <option value="igr">������ 
                              <option value=irm>����� 
                              <option value="ivw">������ 
                              <option value="int">���� 
                              <option value="ikt">������� 
                              <option value="i{o">������-��� 
                              <option value="kzn">������ 
                              <option value="kz~">��������� 
                              <option value="kld">����������� 
                              <option value="klg">������ 
                              <option value="kgd">��������� 
                              <option value="kdd">�������� 
                              <option  value="krw">�������� 
                              <option value="kpm">��������� 
                              <option value="iew">���� 
                              <option value="krn">������� 
                              <option value="kio">����� 
                              <option value="k~g">����������� ������� 
                              <option value="k{n">������� 
                              <option value="kog">������� 
                              <option value="kzi">������� 
                              <option value="ksl">�����������-��-����� 
                              <option value="ktn">�������� 
                              <option value="kts">������ 
                              <option value="krr">��������� 
                              <option value="kkx">������������� 
                              <option value="kqa">���������� 
                              <option value="kgn">������ 
                              <option value="kus">����� 
                              <option value="kis">������� 
                              <option value="kyy">����� 
                              <option value="lsk">����� 
                              <option value="le{">����������� 
                              <option value="lip">������ 
                              <option value="lod">������ 
                              <option value="lwo">����� 
                              <option value="mdn">������� 
                              <option value="mgn">����� 
                              <option value="mgs">������������ 
                              <option value="mkp">������ 
                              <option value="mam">���� 
                              <option value="mrn">���������� 
                              <option value="mko">������� 
                              <option value="mhl">��������� 
                              <option 
                          value="mrw">����������� ���� 
                              <option 
                          value="msk">����� 
                              <option value="mir">������ 
                              <option 
                          value="mom">���� 
                              <option class="hl" value="mow" selected>������ 
                              <option 
                          value="mtg">�������� 
                              <option value="mun">�������� 
                              <option 
                          value="mkm">��� �������� 
                              <option value="m{d">��� ������ 
                              <option value="m`n">������ 
                              <option 
                          value="ndm">����� 
                              <option value="in{">������� 
                              <option 
                          value="n~k">������� 
                              <option value="nmg">�������� 
                              <option 
                          value="nnr">������-��� 
                              <option 
                          value="nh~">���������� 
                              <option value="nln">������� 
                              <option 
                          value="nrg">�������� 
                              <option value="n`g">����������� 
                              <option 
                          value="nvg">������������ 
                              <option 
                          value="nvw">������������� 
                              <option 
                          value="nvk">���������� 
                              <option value="nvq">��������� 
                              <option 
                          value="nvs">������ �������� 
                              <option 
                          value="nlk">����������-��-����� 
                              <option 
                          value="nik">���������� 
                              <option 
                          value="nwk">����������� 
                              <option 
                          value="owb">����������� 
                              <option value="nur">����� ������� 
                              <option value="ngl">������� 
                              <option 
                          value="nrs">�������� 
                              <option value="noq">�������� 
                              <option 
                          value="nus">����� 
                              <option value="n`r">����� 
                              <option 
                          value="nqg">������ 
                              <option value="ods">������ 
                              <option 
                          value="ozr">������� 
                              <option value="omk">������� 
                              <option 
                          value="olk">��������� 
                              <option value="oln">������ 
                              <option 
                          value="ool">������ 
                              <option value="oms">���� 
                              <option 
                          value="ong">�������� 
                              <option value="osk">���� 
                              <option 
                          value="oso">������ 
                              <option value="oha">��� 
                              <option 
                          value="oht">������ 
                              <option value="pwl">�������� 
                              <option 
                          value="pan">������ 
                              <option value="pew">����� 
                              <option 
                          value="prm">����� 
                              <option value="ptz">������������ 
                              <option 
                          value="prl">�������������-���������� 
                              <option 
                          value="p~r">������ 
                              <option value="pin">��������� 
                              <option 
                          value="pts">����������� �������� 
                              <option value="pos">������ �������� 
                              <option value="plq">�������� 
                              <option 
                          value="rad">�������� 
                              <option value="rix">���� 
                              <option 
                          value="row">������-��-���� 
                              <option 
                          value="sky">�������� 
                              <option value="shd">�������� 
                              <option 
                          value="sm{">������ 
                              <option value="skd">��������� 
                              <option 
                          value="sag">������ 
                              <option class=hl 
                          value="spt">�����-��������� 
                              <option 
                          value="srn">������� 
                              <option value="sro">������� 
                              <option 
                          value="syh">�������� 
                              <option 
                          value="sen">������-�������� 
                              <option 
                          value="swe">������-������ 
                              <option value="ski">������ 
                              <option 
                          value="sel">���� 
                              <option value="sip">����������� 
                              <option 
                          value="sbo">�������� 
                              <option value="sog">��������� ������ 
                              <option value="soj">��������� 
                              <option 
                          value="son">��������� 
                              <option value="so~">���� 
                              <option 
                          value="srm">������-������� 
                              <option 
                          value="stw">���������� 
                              <option value="ist">������� 
                              <option 
                          value="sol">������ ����� 
                              <option 
                          value="spw">���������� 
                              <option value="stv">��������� 
                              <option 
                          value="sun">������ 
                              <option value="sur">������ 
                              <option 
                          value="syw">��������� 
                              <option value="tio">������� 
                              <option 
                          value="tas">������� 
                              <option value="tbs">������� 
                              <option 
                          value="tlq">����-���� 
                              <option value="tig">������ 
                              <option 
                          value="tsi">����� 
                              <option value="til">�������� 
                              <option 
                          value="txk">������ 
                              <option value="tsk">����� 
                              <option 
                          value="tau">���� 
                              <option value="trh">��������� 
                              <option 
                          value="t`m selected">������ 
                              <option value="tsn">����-���� 
                              <option 
                          value="uln">����-����� 
                              <option value="ul|">����-��� 
                              <option 
                          value="ulk">��������� 
                              <option value="ura">���� 
                              <option 
                          value="ugn">������ 
                              <option value="usn">������ 
                              <option 
                          value="utk">����-������ 
                              <option 
                          value="ukg">����-����������� 
                              <option 
                          value="uk~">����-�������� 
                              <option 
                          value="uku">����-����� 
                              <option value="usk">����-��� 
                              <option 
                          value="usm">����-��� 
                              <option value="unr">����-���� 
                              <option 
                          value="uhj">����-��������� 
                              <option 
                          value="utc">����-������ 
                              <option value="ufa">��� 
                              <option 
                          value="uht">���� 
                              <option value="fgn">������� 
                              <option 
                          value="fra">���������-��-����� 
                              <option 
                          value="hbr">��������� 
                              <option value="hdy">������� 
                              <option 
                          value="has">�����-�������� 
                              <option 
                          value="hrk">������� 
                              <option value="hat">������� 
                              <option 
                          value="hel">��������� 
                              <option value="hrp">������� 
                              <option 
                          value="hdt">������� 
                              <option value="~ar">���� 
                              <option 
                          value="~be">��������� 
                              <option value="~lb">��������� 
                              <option 
                          value="~rw">��������� 
                              <option value="~rs">������� 
                              <option 
                          value="sht">���� 
                              <option value="~kd">�������� 
                              <option 
                          value="~mi">������� 
                              <option value="{ah">�������� 
                              <option 
                          value="{nq">������ 
                              <option value="{mt">������� 
                              <option 
                          value="|gt">��������� 
                              <option value="|li">������ 
                              <option 
                          value="`vk">����-�������� 
                              <option 
                          value="`vh">����-��������� 
                              <option 
                          value="qkt">������ 
                              <option value="qmb">������ 
                              <option value="qrl">���������</option>
                            </select>
                            &nbsp; </p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2> 
                          <p>���������� 
                            <input name="NonDirect"
                        type="checkbox">
                            �������� ����� 
                            <input name="Reverse"
                        type="checkbox">
                            �������� �����</p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2b> 
                          <p>��� ���� (�� ��������� - �� �������)</p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2> 
                          <p> 
                            <select name="Day" size="1" class=text>
                              <option selected 
                          value="0">-- 
                              <option value="1">1 
                              <option value="2">2 
                              <option value="3">3 
                              <option value="4">4 
                              <option value="5">5 
                              <option value="6">6 
                              <option value="7">7 
                              <option value="8">8 
                              <option value="9">9 
                              <option value="10">10 
                              <option value="11">11 
                              <option value="12">12 
                              <option value="13">13 
                              <option value="14">14 
                              <option value="15">15 
                              <option value="16">16 
                              <option value="17">17 
                              <option value="18">18 
                              <option value="19">19 
                              <option value="20">20 
                              <option value="21">21 
                              <option value="22">22 
                              <option value="23">23 
                              <option value="24">24 
                              <option value="25">25 
                              <option value="26">26 
                              <option value="27">27 
                              <option value="28">28 
                              <option value="29">29 
                              <option value="30">30 
                              <option value="31">31</option>
                            </select>
                            <select 
                        name="Month" size="1" class=text>
                              <option selected value="none">* ����� 
                              <option value="1">������ 
                              <option value="2">������� 
                              <option value="3">����� 
                              <option value="4">������ 
                              <option value="5">��� 
                              <option value="6">���� 
                              <option value="7">���� 
                              <option value="8">������� 
                              <option value="9">�������� 
                              <option value="10">������� 
                              <option value="11">������ 
                              <option value="12">�������</option>
                            </select>
                            &nbsp; ����� ����� 
                            <select name="Time" size="1" class=text>
                              <option selected 
                          value="00">-- 
                              <option value="01">1 
                              <option value="02">2 
                              <option value="03">3 
                              <option value="04">4 
                              <option value="05">5 
                              <option value="06">6 
                              <option value="07">7 
                              <option value="08">8 
                              <option value="09">9 
                              <option value="10">10 
                              <option value="11">11 
                              <option value="12">12 
                              <option value="13">13 
                              <option value="14">14 
                              <option value="15">15 
                              <option value="16">16 
                              <option value="17">17 
                              <option value="18">18 
                              <option value="19">19 
                              <option value="20">20 
                              <option value="21">21 
                              <option value="22">22 
                              <option value="23">23</option>
                            </select>
                            &nbsp;����� </p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=n2b> 
                          <p>����������� ��:</p>
                        </td>
                      </tr>
                      <tr> 
                        <td class=text bgcolor="#EBF3F5"> 
                          <p> 
                            <select name="Sort" size="1" class=text>
                              <option 
                          selected value="w">������� ������ 
                              <option value="r">������ ����� 
                              <option 
                        value="s">���������</option>
                            </select>
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                            <input type=submit value="����������" name="submit" size="1" class=text>
                          </p>
                        </td>
                      </tr>
                      <tr> 
                        <td align="left"> 
                          <hr noshade size="3">
                          <h2 align="left"><a href="http://www.polets.ru/" target="_blank">��� 
                            "�����-������"</a> ������������ ����� ������ ���������� 
                            ���������� �� ������ � ������� ���, ����������� ���������� 
                            ������ � �������� �������, ������ ����������� ������, 
                            ����������� ������, ��������� ��������� � ������ � 
                            ����������. </h2>
                        </td>
                      </tr>
                      </tbody> 
                    </table>
                  </td>
                </tr>
                </tbody> 
              </table>
            </form>
            <p align="left"> 
              <script language="javascript" src="http://72rus.ru/newsjv.asp?hid=125"></script>
            <p align="center"><font size="-2"><a href="http://72rus.ru"> 72RUS.RU 
              - ��������� ������</a></font></p>
            <p align="center">&nbsp; </p>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
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
      <p class="menu02">| <font face="Arial, Helvetica, sans-serif"> 
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
        </font><a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <font face="Arial, Helvetica, sans-serif"> 
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
        </font><a href="<%=url%>"><%=hname%></a> | 
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
  <script language="javascript" src="banshow.asp?rid=6"></script>
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
  <!--LiveInternet logo-->
  <a href="http://www.liveinternet.ru/click" target=liveinternet><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="liveinternet.ru: �������� ����� ����� �� 24 ����, ����������� �� 24 ���� � �� �������"></a> 
  <!--/LiveInternet-->
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
