<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="admarea.asp"
Response.Redirect("login.asp")
}

var url=""
var about=""
var id=0
var hid=0
var hname=""
var sql=""
var name=""
var statname=""
var idh=0
var ii=0
var dreg=""
var zgl=""
var nikname=""
var nikid=0
var cw=""
var statid=0
var stt=parseInt(Request("st"))
var usr=parseInt(Request("usr"))
if (isNaN(stt)) {stt=-1}
if (isNaN(usr)) {usr=0}

var Recs=CreateRecordSet()

switch (stt) {
			case -3 : zgl="����� ������������"; break;
			case -2 : zgl="����� � ��������� ���������"; break;
			case -1 : zgl="����� � ���������"; break;
			case  0 : zgl="����������� �����. ������� �������������"; break;
			case  1 : zgl="����� � �����"; break;
			case  2 : zgl="�����, ������������ ��������"; break;
			case  3 : zgl="�����, ������������ ���������������"; break;
			case  4 : zgl="�������� �����"; break;
}

%>


<html>
<head>
<title>����������������� �����.</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr valign="middle" align="center" bgcolor="#F8F8F8"> 
    <td width="150" height="19"> 
      <div align="center"> <a href="index.asp"><img src="Img/home.gif" width="14" height="14" align="absbottom" border="0" alt="Home"></a> 
        <font face="Arial, Helvetica, sans-serif" size="-2"><b>��������� ������</b></font></div>
    </td>
    <td bordercolor="#FFCC00" height="19"> 
      <p>������ �����������: <b><%=Application("visitors")%></b></p>
    </td>
    <td width="400" height="19" align="right"> 
      <table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="index.asp">Index.asp</a></b></font></div>
          </td>
          <td width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b><a href="pubheading.asp?hid=125">�������</a></b></font></div>
          </td>
          <td width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="catarea.asp">�������</a></b></font></div>
          </td>
          <td height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="pubarea.asp">������� 
              ��������� </a></b></font></div>
            </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0">
  <tr bgcolor="#0066FF"> 
    <td height="9"> 
      <div align="center"> 
        <p><b><font color="#FFFF00" size="4">������� ��������������</font></b></p>
      </div>
    </td>
  </tr>
</table>
<table width="100%" border="0">
  <tr>
    <td bgcolor="#F2F4F4"> 
      <div align="center">
        <p><font color="#0000FF"><b>������������, <%=Session("name_mem")%> !</b></font></p>
      </div>
    </td>
  </tr>
</table>
<table width="100%" border="1">
  <tr bordercolor="#0066FF" bgcolor="#0066FF"> 
    <td bordercolor="#0066FF" width="50%"> 
      <div align="center"> 
        <p><b><font color="#FFFFFF">��������</font></b></p>
      </div>
    </td>
    <td width="50%"> 
      <div align="center"> 
        <p><b><font color="#FFFFFF">�����������</font></b></p>
      </div>
    </td>
  </tr>
  <tr bordercolor="#0066FF" valign="top"> 
    <td width="50%" height="137"> 
      <ul>
        <li> 
          <p><a href="pswusr.asp?usr=<%=Session("id_mem")%>">�������� 
            ������</a></p>
        </li>
        <li> 
          <p><a href="pubheading.asp?hid=125">������ &quot;�������&quot;</a></p>
        </li>
        <li> 
          <p><a href="catarea.asp">������������� WEB �������</a></p>
        </li>
        <li> 
          <p><a href="messages.asp">������������� ����������</a></p>
        </li>
        <li> 
          <p><a href="addpubusr.asp">�������� ������������ ���</a></p>
        </li>
        <li> 
          <p><a href="addbaner.asp">�������� �����</a></p>
        </li>
        <li> 
          <p><a href="modrec.asp">���������� ���������� ������</a></p>
        </li>
        <li>
          <p><a href="modbaner.asp">��������� ��������</a></p>
        </li>
      </ul>
    </td>
    <td width="50%" height="137"> 
      <ul>
        <li> 
          <p><a href="admarea.asp?st=3">������������� �����</a></p>
        </li>
        <li> 
          <p><a href="admarea.asp?st=4">��������� �����</a></p>
        </li>
        <li> 
          <p><a href="admarea.asp?st=-2">����� � ��������� ���������</a></p>
        </li>
        <li> 
          <p><a href="admarea.asp?st=0">����������� �����</a></p>
        </li>
        <li> 
          <p><a href="admarea.asp?st=2">����� ������������� ��������������</a></p>
        </li>
        <li> 
          <p><a href="admarea.asp?st=-1">����� � ���������</a></p>
        </li>
        <li>
          <p>������������ �����</p>
        </li>
        <li>
          <p>������� ��������������</p>
        </li>
        <li> 
          <p><a href="admarea.asp?usr=1">������������</a></p>
        </li>
        <li> 
          <p><a href="lastmsg.asp?d=2">��������� ����������</a> (<font size="2">�� 
            ��������� 2 ���</font>)</p>
        </li>
      </ul>
    </td>
  </tr>
</table>
<%if (usr==0) {%>
<p><font color="#000066"><b><%=zgl%>:</b></font></p>
<%
	sql="Select t1.*, t3.nikname from url t1, catarea t2, host_url t3 where t1.host_url_id=t3.id and t1.state="+stt+" and t1.catarea_id=t2.id and t2.catalog_id="+catalog
	sql+=" UNION Select t1.*, cast('' as varchar(20)) as nikname from url t1, catarea t2 where t1.host_url_id is null and t1.state="+stt+" and t1.catarea_id=t2.id and t2.catalog_id="+catalog+" order by 9,1"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.EOF) {
		ii+=1
		name=Records("NAME").Value
		url=Records("URL").Value
		about=Records("ABOUT").Value
		nikname=Records("NIKNAME").Value
		nikid=Records("HOST_URL_ID").Value
		id=Records("ID").Value
		hid=Records("CATAREA_ID").Value
		cw=Records("KEYWORD").Value
		switch (Records("STATE").Value) {
			case -2 : statname="� ��������� ���������"; break;
			case -1 : statname="� ���������"; break;
			case  0 : statname="����������. ������� �������������"; break;
			case  1 : statname=""; break;
			case  2 : statname="���������� ��������"; break;
			case  3 : statname="���������� ���������������"; break;
			case  4 : statname="�����"; break;
		}
		dreg=Records("REG_DATE").Value
		Records.MoveNext()
	%>
        
<table width="90%" border="1" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" bordercolor="#FFFFFF" align="center">
  <tr bgcolor="#ffd34e" bordercolor="#ffd34e"> 
    <td width="45%"> 
      <p><b><%=ii%>.</b>&nbsp;&nbsp;<%=name%>&nbsp;&nbsp;<a href="<%=url%>" target="blank"><font color="#0000FF"><%=url%></font></a></p>
              </td>
            
    <td> 
      <p><font size="2"><a href="urlresume.asp?url=<%=id%>&st=1">����������</a>&nbsp;&nbsp; 
        <a href="delurl.asp?url=<%=id%>" target="_blank">�������</a>&nbsp;&nbsp; 
        <a href="edurl.asp?url=<%=id%>" target="_blank">��������</a>&nbsp;&nbsp; 
        <a href="urlresume.asp?url=<%=id%>&st=3" target="_blank">����������</a>&nbsp;&nbsp; 
        <a href="removeurl.asp?url=<%=id%>&hid=<%=hid%>" target="_blank">�����������</a>&nbsp;&nbsp; 
        <a href="copyurl.asp?url=<%=id%>&hid=<%=hid%>" target="_blank">�����������</a></font></p>
              
            </td>
          </tr>
		  <tr> 
            
    <td colspan="2"> 
      <p>���� �����������: <%=dreg%>&nbsp;&nbsp;|&nbsp;<font color="#663399" size="1"><%=about%></font></p>
              </td>
          </tr>
          <tr> 
    <td colspan="2" bordercolor="#ffd34e"> 
      <p><b>���� ����������:</b> <font color="#996600">
			  <%
			  idh=hid
			  while (idh>0) {
			  sql="Select * from catarea where id="+idh
			  Recs.Source=sql
			  Recs.Open()
			  if (!Recs.EOF) {
			  	idh=Recs("HI_ID").Value
				hname=Recs("NAME").Value
			  %>
                &nbsp;<%=hname%>&nbsp;&gt; 
                <% }
				else {idh=0}
				Recs.Close()
				} 
				%>
			  </font></p>
              </td>
          </tr>
		  <tr>
		  	<td>
      <p><font size="2"><b><%if (nikname!="") {%>���:&nbsp;&nbsp;<a href="shownik.asp?nik=<%=nikid%>"><%=nikname%></a>&nbsp;&nbsp;|&nbsp;&nbsp;<%}%>
        ���������� � ��������� �����:</b> <a href="reclurl.asp?url=<%=id%>&st=1">1</a>&nbsp;&nbsp;<a href="reclurl.asp?url=<%=id%>&st=2">2</a>&nbsp;&nbsp;<a href="reclurl.asp?url=<%=id%>&st=3">3</a>&nbsp;&nbsp;<a href="reclurl.asp?url=<%=id%>&st=4">4</a>&nbsp;&nbsp;<a href="reclurl.asp?url=<%=id%>&st=5">5</a>&nbsp;&nbsp;<a href="reclurl.asp?url=<%=id%>&st=6">6</a>&nbsp;&nbsp;7&nbsp;&nbsp;8 
        </font></p>			
			</td>
			<td>
			<p><font size="2"><b>CW: <%=cw%></b></font></p>
			</td>
		  </tr>
        </table>
<%} Records.Close()%>
<%}%>
<% if (usr==1) {%>
<p align="center"><b>������ ������������������ ������������� ��������:</b></p>
<%
	sql="Select * from HOST_URL Order by date_reg desc, nikname"
	Records.Source=sql
	Records.Open()
	ii=0
	while (!Records.EOF) {
		ii+=1
		name=Records("NAME").Value
		nikname=Records("NIKNAME").Value
		id=Records("ID").Value
		url=Records("EMAIL").Value
		cw=Records("KEYWORD").Value
		statid=Records("STATE").Value
		switch (statid) {
			case  0 : statname="����������"; break;
			case  1 : statname="��������"; break;
			case  -1 : statname="������� ���������"; break;
			case 2 : statname="������"; break;
		}
		dreg=Records("DATE_REG").Value
		Records.MoveNext()%>
<table width="90%" border="1" bordercolor="#FFFFFF" align="center">
  <tr> 
    <td bgcolor="#FFEFCE" bordercolor="#FFCC00"> 
      <p><b><%=ii%>.&nbsp;&nbsp;<a href="shownik.asp?nik=<%=id%>"><%=nikname%></a></b>&nbsp;&nbsp; 
        <%=name%>&nbsp;&nbsp;<%=dreg%>&nbsp;&nbsp;<a href="mailto:<%=url%>"><%=url%></a>&nbsp;&nbsp; 
        <b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">| <%=statname%></font></b></p>
	</td>
    <td width="20%" bordercolor="#FFCC66"> 
      <div align="center"> 
        <p><font size="2"><a href="edusrurl.asp?nik=<%=id%>"><font size="1">���.</font></a> 
          <font size="1"><a href="edusrurlst.asp?nik=<%=id%>&st=1"> 
          <%if (statid==0){%>
          �������.
<%} else {%>
          ����.
<%}%>
          </a> del test</font></font></p>
      </div>
    </td>
  </tr>
</table>
<%} Records.Close()%>
<%}%>
<p>&nbsp;</p>


</body>
</html>
