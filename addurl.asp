<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var hid=parseInt(Request("hid"))
var sql=""
var hiadr=""
var tit=""
var hname=""
var tekhia=0
var urladr=""
var logname=""
var pass=""
var ErrorMsg=""
var name=""
var url=""
var about=""
var id=0
var ShowForm=true	

var smi_id=1
var news_bl=""
var ishtml2=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""			

if (isNaN(hid)) {hid=0}

if (hid>0) {
sql="Select * from catarea where id="+hid+" and catalog_id="+catalog
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("catarea.asp?hid=0")
}
hname=String(Records("NAME").Value)
Records.Close()
}
tit=hname

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && Session("id_mem")!="undefined"){
	sql="Select * from catalog where id="+catalog
	Records.Source=sql
    Records.Open()
   if (Records.EOF){
	Records.Close()
	Response.Redirect("index.asp")
   }
   if (Session("id_mem")==Records("USERS_ID").Value) {Session("cataloghost")=catalog}
   Records.Close()
}

if (String(Request.Form("next1")) != "undefined") {
	if (Request.Form("agr")==1) {Session("isagr")=1}
}

if (String(Request.Form("Enter")) != "undefined") {
	logname=String(Request.Form("logname"))
	pass=String(Request.Form("psw"))
	if (logname.length<2) {ErrorMsg+="������� �������� ���! <br>"}
	if (pass.length<2) {ErrorMsg+="������� �������� ������! <br>"}
	sql="Select * from host_url where nikname='"+logname+"' and psw='"+pass+"'"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){
		Session("is_rite_connect_url")=1
		Session("id_host_url")=Records("ID").Value
   } else {
		Session("is_rite_connect_url")=0
		ErrorMsg+="�������� ������ ��� ��� ������������! <br>"
   }
   Records.Close()
}

isFirst=String(Request.Form("reg")) == "undefined"

if (String(Request.Form("reg")) != "undefined") {
	name=TextFormData(Request.Form("name"),"")
	url=TextFormData(Request.Form("url"),"")
	about=TextFormData(Request.Form("about"),"")
	 
	while (name.indexOf("'")>-1) {name=name.replace("'","\"")}
	while (name.indexOf("  ")!=-1) {name=name.replace("  "," ")}
	while (name.indexOf(" ")==0) {name=name.replace(" ","")}
	while (about.indexOf("  ")!=-1) {about=about.replace("  "," ")}
	while (about.indexOf("'")>-1) {about=about.replace("'","\"")}
	while (about.indexOf(" ")==0) {about=about.replace(" ","")}
	
	while (url.indexOf(" ")!=-1) {url=url.replace(" ","")}
	
	 if (name.length<5) {ErrorMsg+="������� �������� ������������.<br>"}
	if (url.length<5) {ErrorMsg+="������� �������� URL.<br>"}
	if (about.length<10) {ErrorMsg+="������� �������� ��������.<br>"}
	if (url.indexOf("?")>-1) {ErrorMsg+="������������ ������ � URL (?).<br>"}
	if (url.indexOf("=")>-1) {ErrorMsg+="������������ ������ � URL (=).<br>"}
	if (url.indexOf("&")>-1) {ErrorMsg+="������������ ������ � URL (&).<br>"}
	if (url.indexOf("@")>-1) {ErrorMsg+="������������ ������ � URL (@).<br>"}
	if (url.indexOf("'")>-1) {ErrorMsg+="������������ ������ � URL ( ' ).<br>"}
	if (url.indexOf("\"")>-1) {ErrorMsg+="������������ ������ � URL ( \" ).<br>"}
	if (url.indexOf("http://")==-1){url="http://"+url}
	sql="Select * from URL where Upper(url)=Upper('"+url+"') and host_url_id<>"+Session("id_host_url")
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="���� ���� ��� ��������������� ������ �������������!  ���� �� ������ ������� , �� ������ ��� ���������� � ���� ������ ���������!<br>"}
   Records.Close()
   sql="Select * from URL where Upper(url)=Upper('"+url+"') and catarea_id="+hid
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="���� �� ����� ���� �������� ������ � ����� ������� ��������!<br>"}
   Records.Close()
   sql="Select * from URL where Upper(name)=Upper('"+name+"') and catarea_id="+hid
	Records.Source=sql
    Records.Open()
   if (!Records.EOF){ErrorMsg+="������� ������������ �����! ����� ������������ ��� ����������� � ������� ������� ��������!<br>"}
   Records.Close()
	
	if (ErrorMsg==""){
	  		id=NextID("urlid")
			sql="Insert into url (ID,NAME,URL,ABOUT,CATAREA_ID,STATE,REG_DATE,NEW_DATE,KEYWORD,HOST_URL_ID) "
			sql+="values (%ID, '%NAM', '%URL', '%AB', %CAT, -1, 'TODAY', 'TODAY', '', %HS)"
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAM",name)
			sql=sql.replace("%URL",url)
			sql=sql.replace("%AB",about)
			sql=sql.replace("%CAT",hid)
			sql=sql.replace("%HS",Session("id_host_url"))
			Connect.BeginTrans()
			try{
			Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="������ �������.<br>"
			}
			if (ErrorMsg==""){
		   Connect.CommitTrans()
		   ShowForm=false
			}
	  }
	  while (name.indexOf("\"")>-1) {name=name.replace("\"","&quot;")}
	  while (about.indexOf("\"")>-1) {about=about.replace("\"","&quot;")}
}
if (url.indexOf("http://")!=-1){url=url.replace("http://","")}
%>

<Html>
<Head>
<Title>���������� ������� � ������� 72RUS � ����������: <%=tit%></Title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72.css" type="text/css">
</Head>
<BODY bgColor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0">
<hr noshade size="1">
<div align="center">
  <script language="javascript" src="banshow.asp?rid=4"></script>
</div>
<hr noshade size="1">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="catarea.asp">������� 72RUS.RU</a> 
        | <a href="usrarea.asp">������� ������������</a> | <a href="regmemurl.asp">����������� 
        ������������</a> | <a href="http://chat.72rus.ru/">���</a></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="top" align="center"> 
    <td width="150" align="center"> 
      <!--RAX counter-->
      <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
      <!--/RAX-->
    </td>
    <td> 
      <p align="left">&nbsp;</p>
      <p align="center"><font size="3"><b>����������� �������� - ������� � ������� 
        72RUS</b></font></p>
      <p align="center">&nbsp;</p>
      <%if(ErrorMsg!=""){%>
      <center>
        <p> <font color="#FF3300" size="2"><b>������!</b></font> <br>
          <%=ErrorMsg%></p>
      </center>
      <%}%>
      <%
if (Session("isagr")!=1) {
	Session("isagr")=0
%>
      <p align="left"><b>��� ������� ���� ��� ����������� ������ �������:</b></p>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;1. </font></b>��� 
        ���������� ������ ����� � ������� 72RUS, ������������ � ��������� ����������:</p>
      <p align="left">&nbsp;</p>
      <p align="left">������������� 72RUS ��������� � ���������� ����� ��������������� 
        ������������ ��������� ��������, ������� ��������� � ��������� ����� ��������, 
        ������� ���������� � �������� ����������.</p>
      <p align="left">&nbsp;</p>
      <p align="left">��� ����, ����� �������� ���� � ������� ��� ���������� ������ 
        <a href="regmemurl.asp">����������� ������ ������������</a>. ����� ����� 
        �� ��������� ��� ������������ � ������ � ������� � ����� ����� ��������� 
        ������� �����, �������� �� ��������, ����� � ������ ��� ���������. </p>
      <p align="left">&nbsp;</p>
      <p align="left">���������� ������ ����� ���������� � ��� �������� ��������, 
        � ������� �� ������ ��� ����������. </p>
      <p align="left">&nbsp; </p>
      <p align="left">��������� �������� ��������� �� ����� ����� �� ��������� 
        ������ �� �����, ������� �� �������� �����, �������� ��������, ���������� 
        ����� � ������ �����������. ��� ����� ����������� �� ��������� ��������: 
        �������������� ���� �� ������������� ���������� ���� ��������, ���� �� 
        ����� ��������� ����������, ���������� ����� ������������ ���������������� 
        ��, ������ ������, ���������� ����� � ����������� ������� ���, URL ����� 
        �� ��������.</p>
      <p align="left">&nbsp;</p>
      <p align="left">��������� ������ �������������� ����� ��������������� �� 
        ��������, �����������, ���������� ���������� ������.</p>
      <p align="left">&nbsp;</p>
      <p align="left">������������� 72RUS �� ����������� ��������� �����-���� 
        ������ ���������� ������ � ������������� ��������, ��������������� ������ 
        ��������. ����������, �������, ���������, ������ �������� ����� ���� �������� 
        � ����� ����� ��� ����������� ���������� ������.</p>
      <p align="left">���� �� �� �������� � ���������, ���������� �� ��������� 
        ���� ������ � ��� �������.</p>
      <p align="left">���� �� �������� � ���������, �� ��������� ������ � ��������������� 
        ���� � ������� ������ &quot;����������&quot;.</p>
      <form name="form1" method="post" action="addurl.asp">
        <input type="hidden" name="hid" value="<% =hid %>">
        <p align="left"> 
          <input type="checkbox" name="agr" value="1">
          ��, � �������� � ��������� ����������� ����� � ������� 72RUS � ���� 
          ���������� ����� ������ � ��������� <b><font color="#FF0000" size="3"><%=tit%></font></b> ������� ��������!</p>
        <p>&nbsp;</p>
        <p align="left"> 
          <input type="submit" name="next1" value="����������">
        </p>
      </form>
      <%} else {%>
      <%
if (Session("is_rite_connect_url")!=1) {
	Session("is_rite_connect_url")=0
%>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;2. </font></b>��������� 
        ��� - ��� ���������� ����� � ������� ��� ����� ������ � �������, ��� ����� 
        ��������� ��������������� ���� � ������� ������ &quot;�����&quot;</p>
      <p align="left">���� � ��� ��� ����� � ������ � ����� ������� (������� 72RUS) 
        �� �� ������ ������������������. ����� ����������� ������� ���������� 
        ����������� ������ �������.</p>
      <form name="form2" method="post" action="addurl.asp">
        <input type="hidden" name="hid" value="<% =hid %>">
        <div align="center"> 
          <table width="90%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="47%">&nbsp;</td>
              <td width="6%">&nbsp;</td>
              <td>&nbsp; </td>
            </tr>
            <tr> 
              <td width="47%" height="18"> 
                <div align="right"> 
                  <p><b>��� �����</b></p>
                </div>
              </td>
              <td width="6%" height="18"> 
                <p align="center">-</p>
              </td>
              <td valign="top" height="18"> 
                <p> 
                  <input type="text" name="logname">
                </p>
              </td>
            </tr>
            <tr> 
              <td width="47%"> 
                <div align="right"> 
                  <p><b>��� ������</b></p>
                </div>
              </td>
              <td width="6%"> 
                <p align="center">-</p>
              </td>
              <td valign="top"> 
                <p> 
                  <input type="password" name="psw">
                </p>
              </td>
            </tr>
            <tr valign="bottom"> 
              <td width="47%" height="40"> 
                <div align="right"> 
                  <p>���� �� ��� �� ������������������</p>
                </div>
              </td>
              <td width="6%" height="40"> 
                <p align="center">-</p>
              </td>
              <td height="40"> 
                <p>�� ������ ������������������ <a href="regmemurl.asp">�����</a></p>
              </td>
            </tr>
          </table>
        </div>
        <p align="center"> 
          <input type="submit" name="Enter" value="�����">
        </p>
      </form>
      <%} else {%>
      <%
if (ShowForm) {
%>
      <p>&nbsp;</p>
      <p align="left"><b><font size="3" color="#FF0000">&nbsp;&nbsp;3. </font></b>������� 
        � ��������� ����� ����� ���������� ����������, ����������� ����, ����� 
        ����������� ������ �������:</p>
      <p align="left">��� ���� ����������� � ����������!!</p>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#FFFFFF">
        <tbody> 
        <tr> 
          <td bgcolor=#EEEEEE width="100%" bordercolor="#999999"> 
            <div align="left"> 
              <p><font size="1"><b>����������� ����� � ������</b></font> <b><font size="1"> 
                <%
	  tekhia=hid
	  while (tekhia!=0) {
	     sql="Select * from catarea where id="+tekhia
		 Records.Source=sql
		 Records.Open()
		 if (!Records.EOF) {
			hname=String(Records("NAME").Value)
			hiadr="catarea.asp?hid="+tekhia
			tekhia=Records("HI_ID").Value
		%>
                &lt; <a href="<%=hiadr%>"><%=hname%></a> 
                <%	
	  }  Records.Close()
	  }
	  if (hid!=0) { %>
                &lt; <a href="catarea.asp?hid=0">�������</a> 
                <% } else {%>
                &lt; <a href="index.asp"> 72RUS.RU</a> 
                <%} %>
                </font></b> </p>
            </div>
          </td>
        </tr>
        </tbody> 
      </table>
      <form name="form3" method="post" action="addurl.asp">
        <input type="hidden" name="hid" value="<% =hid %>">
        <div align="center"> 
          <table width="100%" border="1" cellspacing="2" cellpadding="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
            <tr> 
              <td width="37%" bgcolor="#0066CC" bordercolor="#0000CC"> 
                <div align="center"> 
                  <p><b><font color="#FFFFFF">���������</font></b><font color="#FFFFFF">:</font></p>
                </div>
              </td>
              <td bgcolor="#0066CC" bordercolor="#0000CC"> 
                <div align="center"> 
                  <p><b><font color="#FFFFFF">��������</font></b></p>
                </div>
              </td>
            </tr>
            <tr bordercolor="#999999"> 
              <td width="37%" bgcolor="#EEEEEE"> 
                <div align="right"> 
                  <p>������������ �����&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; 
                    <input type="text" name="name" size="35" maxlength="99" value="<%=name%>">
                    <font size="1">(�� 100 ��������)</font></p>
                </div>
              </td>
            </tr>
            <tr bordercolor="#999999"> 
              <td width="37%" bgcolor="#EEEEEE"> 
                <div align="right"> 
                  <p>URL �����&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; http:// 
                    <input type="text" name="url" size="27" maxlength="96" value="<%=url%>">
                  </p>
                </div>
              </td>
            </tr>
            <tr bordercolor="#999999"> 
              <td width="37%" bgcolor="#EEEEEE"> 
                <div align="right"> 
                  <p>������� �������� �����&nbsp;&nbsp;</p>
                </div>
              </td>
              <td bgcolor="#EEEEEE"> 
                <div align="left"> 
                  <p>&nbsp;&nbsp;&nbsp; 
                    <input type="text" name="about" size="35" maxlength="250" value="<%=about%>">
                    <font size="1">(�� 250 ��������)</font></p>
                </div>
              </td>
            </tr>
            <tr> 
              <td width="37%"> 
                <div align="right">&nbsp;</div>
              </td>
              <td>&nbsp;</td>
            </tr>
          </table>
          <input type="submit" name="reg" value="����������������">
          <hr size="4" noshade>
        </div>
      </form>
      <%} else {%>
      <p>&nbsp;</p>
      <p align="center"><font color="#FF0000"> <b><font color="#009966">��� ������ 
        �������� � ������� ����������� ������� �������� 72RUS.</font></b></font></p>
      <p align="center"><font color="#009966"> <b>����� �������� �� ����� �������� 
        � ��������.</b></font></p>
      <p align="center">&nbsp;</p>
      <h1><font color="#0000FF">��� ����������� ����� ���������� ���������� ������ 
        ����� � ��������,</font></h1>
      <h1><font color="#0000FF">������ ��� �������� �� ����� �� ������������� 
        ������ ����� ���������</font></h1>
      <% } } } %>
      <p>&nbsp; </p>
    </td>
    <td width="150" align="center"> 
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
      <p><b><font size="3"><%=blokname%></font></b> </p>
      <table width="120" border="1" cellspacing="0" cellpadding="0" bordercolor="#003366">
        <tr> 
          <td align="CENTER" bordercolor="#EBF5ED"> 
            <script language="javascript" src="banshow.asp?rid=5"></script>
          </td>
        </tr>
      </table>
      <p align="CENTER"></p>
    </td>
  </tr>
</table>
<!-- start Link.Ru -->
<div align="center">
  <script language="JavaScript">
// <!--
var LinkRuRND = Math.round(Math.random() * 1000000000);
document.write('<iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF&r='+LinkRuRND+' frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> </iframe>');
// -->
</script>
  <noscript> <iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> 
  </iframe> </noscript> 
  <!-- end Link.Ru -->
</div>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=6"></script>
  <hr size="1">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#FF0000"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>�������������� 
          ������ 72RUS - ��������� ������� </b></font><font color="#FFFFFF" size="1"><b>- 
          ���������������� � ������</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">��� 
          ��������</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
      </td>
    </tr>
  </table>
  <hr size="1">
  <p align="center">| <a href="http://auto.72rus.ru">���� ������</a> | <a href="http://www.auction.72rus.ru/">�������</a> 
    | <a href="messages.asp">����������</a> | <a href="Rail_roads.asp">����������</a> 
    | <a href="catarea.asp">��������� �������</a> | <br>
    � 2002 <a href="http://www.rusintel.ru">Rusintel Company</a> 
</div>
<p align="center"> 
  <!--RAX logo-->
  <a href="http://www.rax.ru/click" target=_blank><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="rax.ru: �������� ����� ����� �� 24 ����, ����������� �� 24 ���� � �� �������"></a> 
  <!--/RAX-->
  <!-- HotLog -->
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
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"> <img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
</Body>
</Html>
