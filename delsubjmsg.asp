<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->

<%
var id=parseInt(Request("url"))
var name=""
var host_id=0
var areaadr=""
var admtp=0
var stateurl=0
var hid=0
var retname="�������� � �������"
var tekhia=0
var sql=""
var hname=""
var hiadr=""
var ShowForm=true
var isok=true

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)&&(Session("is_rite_connect_url")!=1)){
Response.Redirect("catarea.asp")
}

if (isNaN(id)) {Response.Redirect("catarea.asp")}

if ((Session("is_adm_mem")!=1)||(Session("cataloghost")!=catalog)) {
	areaadr="admarea.asp"
	admtp=1
} 
if (Session("is_rite_connect_url")==1) {
	areaadr="usrarea.asp"
	admtp=2
}

sql="Select t1.* from url t1, catarea t2 where t1.id="+id+" and t1. catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect(areaadr)
}
host_id=Records("HOST_URL_ID").Value
name=Records("NAME").Value
stateurl=Records("STATE").Value
hid=Records("CATAREA_ID").Value
Records.Close()

if (admtp==1 && stateurl!=1) {areaadr+="?st="+stateurl}
if (admtp==1 && stateurl==1) {
	areaadr="catarea.asp?hid="+hid
	retname="������� � ����������"
}
if (admtp==2 && Session("id_host_url")!=host_id) {Response.Redirect(areaadr)}

if (String(Request.Form("next1")) != "undefined") {
	if (Request.Form("agr")==1) {
		if ((admtp==2) || (admtp==1 && stateurl==4)) {sql="Delete from url where id="+id}
		else {sql="Update url set state=4 where id="+id}
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
			var ShowForm=false
		}
	} else {Response.Redirect(areaadr)}
}

%>

<Html>
<Head>
<Title>��������������� <%=name%></Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">

<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr> 
    <td width="150" valign="middle" background="Img/yellow.jpg" height="19"> 
      <div align="center"> <a href="index.asp"><img src="Img/home.gif" width="14" height="14" align="absbottom" border="0" alt="Home"></a> 
        <font face="Arial, Helvetica, sans-serif" size="-2"><b>��������� ������</b></font></div>
    </td>
    <td bordercolor="#FFCC00" background="Img/yellow.jpg" height="19"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>WWW.72RUS.RU</b></font></div>
    </td>
    <td width="468" background="Img/yellow.jpg" height="19"> 
      <table width="410" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="Img/tmn.gif" width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>������</b></font></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="messages.asp">����������</a></b></font></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="catarea.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>�������</b></font></a></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="Rail_roads.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>����������</b></font></a></div>
          </td>
          <td width="82" height="20" background="Img/tmn.gif"> 
            <div align="center"><a href="marketBuilder.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>���������</b></font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td bgcolor="#ffd34e" width="150" align="left" valign="middle" background="Img/runline.gif"> 
      <div align="center"><a href="index.asp"><img src="Img/%F2%FE%EC%E5%ED%FC%2072rus.gif" alt="www.72rus.ru" border="0"></a></div>
    </td>
    <td bordercolor="#FFCC00" bgcolor="#ffd34e" background="Img/runline.gif"> 
      <div align="center"><object type="text/x-scriptlet" width=171 height="60" data="searchbox.html">
        </object> </div>
    </td>
    <td bgcolor="#ffd34e" align="right" background="Img/runline.gif" width="468"><a href="http://www.rusintel.ru/"><img src="bannes/rol_ba.gif" width="468" height="60" alt="Internet Card - ROL - Tyumen" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0" background="Img/line.gif">
  <tr bgcolor="#996600" background="Img/line.gif"> 
    <td width="150" background="Img/line.gif"> 
      <div align="center"> 
        <p align="left"><b><font color="#FFCC33">- ������ -</font></b></p>
      </div>
    </td>
    <td width="490" bgcolor="#996600" background="Img/line.gif"> 
      <p align="left"><b><font color="#FFCC33">- ������ -</font></b> </p>
    </td>
    <td width="213" background="Img/line.gif"> 
      <p><b><font color="#FFCC33">-- ���� -- </font></b></p>
    </td>
    <td width="150" bgcolor="#996600" background="Img/line.gif"> 
      <div align="center"> 
        <p align="right"><img src="Img/home.gif" width="14" height="14" align="middle"> 
          <a href="index.asp"><b><font color="#FFCC33">Home</font></b></a></p>
      </div>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="left" bgcolor="#EBF5ED"> 
      <div align="left"> 
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%">
              <div align="right"><img src="Img/arrow2.gif" width="11" height="10" align="middle"> 
                <font face=Verdana size=1><b> <a href="index.html">72RUS.RU </a></b></font></div>
            </td>
          </tr>
          </tbody> 
        </table>
        <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
          <tr> 
            <td bordercolor="#FFFFFF" valign="top"> 
              <div align="left"> 
                <h2 align="left">&nbsp;</h2>
              </div>
            </td>
          </tr>
        </table>
      </div>
    </td>
    <td> 
      <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#ffd34e width="100%"> <font face=Verdana size=1><b>. <img src="Img/arrow2.gif" width="11" height="10" align="absmiddle"> 
            <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=path%></font> 
            </b></font></td>
        </tr>
        </tbody> 
      </table>
       
      <h1 align="left"><font size="3">�������� ������� ����������</font></h1>
      <p>&nbsp;</p>
	  
      <% if (ShowForm) {%>
      <p align="left"><b>�� ������������� ������ ������� ������: <font size="3" color="#990000"><%=namarea%></font> 
        ?</b></p>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;��������!</font></b> 
        �������� �������! ���� �� ���������� �� ���� ������ ������� �� ������� 
        �������� 72RUS !</p>
	  <p align="left">�������� ������� ���������� � ������ ���������� ����� �����������!</p>
      <p align="left">������ ����� ������ ������ � ��� ������, ���� � ��� ��� 
        ����������. </p>
      <p align="left">���� �� ������ ������� ������, �� ��������� ������ � ��������������� 
        ���� � ������� ������ &quot;����������&quot;.</p>
      <form name="form1" method="post" action="delurl.asp">
        <input type="hidden" name="url" value="<% =id %>">
        <p align="left"> 
          <input type="checkbox" name="agr" value="1">
          ��, � ���� ������� ���� ������: (<b><font size="3" color="#990000"><%=namarea%></font></b>) 
          �� �������� 72RUS !</p>
        <p>&nbsp;</p>
        <p align="left"> 
          <input type="submit" name="next1" value="����������">
        </p>
      </form>
<%} else {%>
      <p>&nbsp;</p>
      <p><font color="#FF0000"><b>������ ������!</b></font></p>
<%}%>	  

	  <p>&nbsp;</p>
	<p><a href="admarea.asp">������� ��������������</a></p>
	  <p>&nbsp;</p>
    </td>
    <td width="150" bgcolor="#EBF5ED"> 
      <div align="left">
        <table bgcolor=#ffd34e border=0 cellpadding=0 cellspacing=0 
            width="100%">
          <tbody> 
          <tr> 
            <td bgcolor=#ffd34e width="100%"><font face="Arial, Helvetica, sans-serif" size="1"><img src="Img/arrow2.gif" width="11" height="10" align="middle"></font><font 
                  face=Verdana size=1><b> ������</b></font></td>
          </tr>
          </tbody> 
        </table>
      </div>
      <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e" height="100%">
        <tr> 
          <td bordercolor="#FFFFFF" valign="top"> 
            <div align="left">
              <h2 align="left">&nbsp;</h2>
              </div>
            </td>
        </tr>
      </table>
      <p align="left">&nbsp;</p>
      <p align="left">&nbsp;</p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
  <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
    <td valign="middle" bgcolor="#996600"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>�������������� 
        ������ 72RUS - ��������� ������� </b></font><font color="#FFFFFF" size="1"><b>- 
        ���������������� � ������</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">��� 
        ��������</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
    </td>
  </tr>
</table>
</Body>
</Html>
