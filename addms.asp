<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->

<%
var tip=0
var subj_id=parseInt(Request("subj"))
var namarea=""
var name=""
var ErrorMsg=""
var ShowForm=true
var tit="" 
var path=""
var sbj=0
var nm=""
var leftit=""
var ritit=""
var canaddmsg=0
var mid=0
var about=""
var sw=""
var nikname=""
var phone=""
var email=""
var dase=""
var sale=""
var price=""
var sex=""
var city=""
var de=""
var sql=""
var sql1=""
var cid=0
var ts=""
var filename=""
var sumdat=Server.CreateObject("datesum.DateSummer")
var dats = new Date()

var smi_id=1
var news_bl=""
var ishtml2=0
var ishtml=0
var puid=0
var filnam=""
var path=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""

var locks=String(Math.random()).substr(3,6)
if(locks.length<6){locks+=String(Math.random()).substr(3,(6-locks.length))}
var cw=String(Math.random()).substr(3,20)
if(cw.length<20){cw+=String(Math.random()).substr(3,(20-cw.length))}

if (isNaN(subj_id)) {Response.Redirect("messages.asp")}
if (subj_id==0) {Response.Redirect("messages.asp")}

if (Session("test_subj")!=subj_id) {Response.Redirect("messages.asp")}

sbj=subj_id

while (sbj>0) {
	Records.Source="Select * from trade_subj where id="+sbj+" and marketplace_id="+market
	Records.Open()
	if (Records.EOF){
		Records.Close()
		Response.Redirect("messages.asp")
	}
	nm=String(Records("NAME").Value)
	if (sbj==subj_id) {
		leftit=String(Records("IN_NAME").Value)
		ritit=String(Records("OUT_NAME").Value)
		namarea=String(Records("NAME").Value)
		tip=Records("MSG_TYPE").Value
	}
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> / "+path
	sbj=Records("HI_ID").Value
	tit=nm+" : "+tit
	Records.Close()
}

Records.Source="Select * from trade_subj where hi_id="+subj_id+" and marketplace_id="+market
Records.Open()
if (Records.EOF) {canaddmsg=1}
Records.Close()

path="<a href=\"messages.asp\">����������</a>  / "+path

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && canaddmsg==0) {Response.Redirect("messages.asp?subj="+subj_id)}

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){

     name=TextFormData(Request.Form("name"),"")
	 sw=TextFormData(Request.Form("sw"),"")
	 nikname=TextFormData(Request.Form("nikname"),"")
	 phone=TextFormData(Request.Form("phone"),"")
	 email=TextFormData(Request.Form("email"),"")
	 dase=Request.Form("dase")
	 sex=Request.Form("sex")
	 sale=Request.Form("sale")
	 price=TextFormData(Request.Form("price"),"")
	 city=TextFormData(Request.Form("city"),"")
	 about=TextFormData(Request.Form("txt"),"")
	 while (about.indexOf("<")>=0) {about=about.replace("<","&lt;")}
	 
	 
	 if (name.length<3) {ErrorMsg+="������� �������� ������������.<br>"}
	 if (nikname.length==0) {ErrorMsg+="������ ��� ��� ���������.<br>"}
	 if (city.length<3) {ErrorMsg+="������� �������� ������������ ������.<br>"}
	 if (about.length<15) {ErrorMsg+="������� �������� ����������.<br>"}
	 if (about.length>4000) {ErrorMsg+="������� ������� ����������.<br>"}
	 if (sw!=Session("lcode")) {ErrorMsg+="����������� ������ ���.<br>"}
	 if (dase>365) {dase=365}
	 if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && dase>30 && tip!=10) {dase=30}
	 if ((tip>=10 && tip<20) && (sex!=0 && sex!=1)) {ErrorMsg+="�� ������ ��� ���.<br>"}
	 if (email.length>0) {	 
	     if (!/(\w+)@((\w+).)*(\w+)$/.test(email)) {ErrorMsg=ErrorMsg+"�������� ������ ���� 'E-mail'.<br>"}}
	if (email=="autoservis@list.ru") {ErrorMsg+="����������� ������ ���.<br>"} 
	if (email=="asido@chat.ru") {ErrorMsg+="����������� ������ ���.<br>"}
	Records.Source="Select * from TRADEMSG where date_create>='TODAY'-3 and TRADE_SUBJ_ID="+subj_id+" and name='"+name+"' and NIKNAME='"+nikname+"'"
	Records.Open()
	if (!Records.EOF) {ErrorMsg+="����������� ������ ���.<br>"}
	Records.Close()
	
	sql="Select * from CITY where Upcase(name)=Upcase('"+city+"')"
	Records.Source=sql
    Records.Open()
   if (!Records.EOF) {cid=Records("ID").Value} else {cid=-1}
   Records.Close()
	 
	  if (ErrorMsg==""){
	  		mid=NextID("TRADEMSGID")
			filename=MsFilePath+mid+".ms"
			var fs= new ActiveXObject("Scripting.FileSystemObject")
			if (cid<0) {
			cid=NextID("cityid")
			sql1="Insert into CITY (ID,NAME) values ( %ID, '%NAME' )"
			sql1=sql1.replace("%ID",cid)
			sql1=sql1.replace("%NAME",city)
			}
			de=String(dats.getDate())+"."+String(dats.getMonth()+1)+"."+String(dats.getYear())
			de=sumdat.SummToDate(de,dase)
			sql="Insert into TRADEMSG (ID, NAME, DATE_CREATE, PHONE, E_MAIL, CITY_ID, TRADE_SUBJ_ID,CODEWORD,IS_FOR_SALE, PRICE, DATE_END, MSG_TYPE, STATE,NIKNAME,COUNTRY_ID,AREA_ID) "
			sql+=" values (%ID,'%NAME','TODAY','%PH','%EML',%CITY,%SUBJ,'%CW',%IS,'%PRICE','%DE',%MT,0,'%NIK',%LND,%AR)"
			sql=sql.replace("%ID",mid)
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%PH",phone)
			sql=sql.replace("%EML",email)
			sql=sql.replace("%CITY",cid)
			sql=sql.replace("%SUBJ",subj_id)
			sql=sql.replace("%CW",cw)
			if (tip==10) {sql=sql.replace("%IS",sex)}
			else {sql=sql.replace("%IS",sale)}
			sql=sql.replace("%PRICE",price)
			sql=sql.replace("%DE",de)
			sql=sql.replace("%MT",tip)
			sql=sql.replace("%NIK",nikname)
			sql=sql.replace("%LND",Request.Form("country"))
			sql=sql.replace("%AR",Request.Form("region"))
			Connect.BeginTrans()
			try{
			if (sql1 != "") {Connect.Execute(sql1)}
			Connect.Execute(sql)
			ts=fs.OpenTextFile(filename,2,true)
			ts.Write(about)
			ts.Close()
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="������ �������.<br>"
			}
			if (ErrorMsg==""){
		   Connect.CommitTrans()
		   ShowForm=false
		   Session("lcode")="undefined"
		   Session("id_ms")=mid
			}
	  }

}

if (ShowForm) {Session("lcode")=locks}

%>

<html>
<head>
<title>���������� ���������� �: (<%=tit%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        72RUS.RU ����� ���������� - ������, ��������� �������</font> </p>
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
<table width="100%" border="0" cellspacing="0" bordercolor="#003366">
  <tr bgcolor="#FBF8D7"> 
    <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle" align="center"> 
      <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
        <a href="index.asp">72RUS.RU</a> / <%=path%></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
  <tr> 
    <td bgcolor="#FF9900" align="center" background="images/fon_menu08.gif"> 
      <h1><b><font color="#FFFFFF">�� ���������� ���������� � ������: 
        <%if (subj_id>0) {%>
        <%=namarea%> 
        <%} else {%>
        ��������� ����� ���������� 
        <%}%>
        </font></b>&nbsp;</h1>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
  <p> <font color="#FF3300" size="2"><b><i>������!</i></b></font> <br>
    <%=ErrorMsg%></p>
</center>
<%}%>
<%if(ShowForm){%>
<table width="92%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td> 
      <p>&nbsp;</p>
      <p><b>�������� ���������, ����������� ��� �����-������� �� &quot;����� ���������� 
        72RUS&quot;</b></p>
      <p> ������ ��������, �� ������ ������� ����������� � ������� 
        ���� ������� �� ������������ ������� �� ���������������� ���������� ���������, 
        �������, �� �� ������������� ���, ���������� � ������� �������� ����������, 
        ����������, �������� � �����������, ��������������������� � ������ ������� 
        �������� � ��������, ������������� ���������, � ����� ������������� �������� 
        � ������������ ��������. ��������, ��������� ������� � ��� ��� ��������� 
        �����.</p>
      <p> ����������, ������� ����� ���������� ������ 
        � ��������������� �������������. ����� ��������, ���������� ��� ������������ 
        ���������� ���� ������� ���. ����� ������������� ����, ������������ ����� 
        � �������. �������� ��������. ��������������� �������. </p>
      <p>����������� �������� ����� � ����� �������� 
        ������.</p>
      <p>������ � ���� ������ ������.</p>
      <p>������������ ������ � ���� �������� �������� 
        (�� ����������� �� �����-������� ��� ��������������� �����).</p>
      <p>����������������� � ����������� ����������� 
        �������� � ������������ � ��������, ������������ �������������� ������������� 
        �� �� 02.02.98 � 111.</p>
      <p>������������ ��� ������� ���� ��������������� 
        ������������� ��������, ��������, ���������� � ��������. </p>
      <p>������������ ������, ����� ������� ����������.</p>
      <p>������ ��������, �� ��������������� ������� 
        ������������ � �������������� � ���������� ���������. </p>
      <p>������������� ����� ��������� �� ����� ����� 
        ������� ����� ��������� � ��������� ���������, � ����� ����� � ��� ���������������� 
        �����������. �������� ����������, �� ������������� ���� �������� � ���� 
        �������� ���������, � ����� ��������� � ������������ �� ��� �������, ��� 
        ��������� ����, ��� � ���������������. 
      <p>������������� ����� �� ����� ��������������� 
        �� ���������� � ����������� ����������� ���� ����������. 
      <p align="center">&nbsp; 
      <p align="center"><b>���������� ����������� � ���������� ������, ����������� 
        � ����������������� �������, � ������������ ��������� ����������� ������ 
        ����� ������� ���������������.</b> 
    </td>
  </tr>
</table>
<form name="form1" method="post" action="addms.asp">
  <hr size="1" noshade width="95%">
  <p align="center">��� ���������� ����������, ����������, ��������� ���� �����<br>
    <font color="#FF0000">��������� ���������� ������� ������ ����������� � ����������!</font> 
    <input type="hidden" name="subj" value="<%=subj_id%>">
    <%if (tip>=30 && tip<40) {%>
    <input type="hidden" name="sale" value="1">
    <input type="hidden" name="price" value="">
    <%}%>
  </p>
  <table width="95%" border="1" bordercolor="#FFFFFF" align="center">
    <tr background="HeadImg/shadow.gif"> 
      <td width="294" height="21" bordercolor="#003366" background="HeadImg/shadow.gif"> 
        <div align="center"> 
          <h1><b><font color="#003366">���������:</font></b></h1>
        </div>
      </td>
      <td height="21" width="466" bordercolor="#003366" background="HeadImg/shadow.gif"> 
        <div align="center"> 
          <h1><b><font color="#003366">��������:</font></b></h1>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="294" height="2" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">��������� ���������� (��������, 
            ������������):&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td height="2" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="100" size="45">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="294" height="18" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">������� ���, ������� ����� �� 
            ����:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td align="left" valign="top" height="18"> 
        <p><img src="im.asp"> 
          <input type="text" name="sw" maxlength="6">
          <font size="1">(���� ��� ����������� - �������� ��� ���������)</font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" height="11" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">���� ��� (���������) 
            <%if (tip<10 || tip>=20) {%>
            ��� ����������� 
            <%}%>
            :&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td height="11" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="nikname" value="<%=isFirst?"":Request.Form("nikname")%>" maxlength="50">
          </font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest">��� �������:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="text" name="phone" value="<%=isFirst?"":Request.Form("phone")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest">��� E-mail:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="text" name="email" value="<%=isFirst?"":Request.Form("email")%>" maxlength="80">
          </font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">���� ���������� ����������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="dase" value="1" <%=Request.Form("dase")==1?"checked":""%>>
          1 ���� 
          <input type="radio" name="dase" value="2" <%=Request.Form("dase")==2?"checked":""%>>
          2 ��� 
          <input type="radio" name="dase" value="7" <%=Request.Form("dase")==7?"checked":""%>>
          7 ���� 
          <input type="radio" name="dase" value="10" <%=isFirst?"checked":""%> <%=Request.Form("dase")==10?"checked":""%>>
          10 ���� 
          <input type="radio" name="dase" value="20" <%=Request.Form("dase")==20?"checked":""%>>
          20 ���� 
          <input type="radio" name="dase" value="30" <%=Request.Form("dase")==30?"checked":""%>>
          30 ���� 
          <%if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog || tip==10) {%>
          <input type="radio" name="dase" value="365" <%=Request.Form("dase")==365?"checked":""%>>
          ��� 
          <%}%>
          </font></p>
      </td>
    </tr>
    <% if (tip<10) { %>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">������� ��� �������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="sale" value="1" <%=isFirst?"checked":""%> <%=Request.Form("sale")==1?"checked":""%>>
          ������� 
          <input type="radio" name="sale" value="0" <%=Request.Form("sale")==0?"checked":""%>>
          �������</font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest">����:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="text" name="price" value="<%=isFirst?"":Request.Form("price")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <%} 
if (tip>=10 && tip<20) {%>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">�� ������� ��� �������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top" height="23"> 
        <p><font size="2"> 
          <input type="radio" name="sex" value="1" <%=Request.Form("sex")==1?"checked":""%>>
          � ������� 
          <input type="radio" name="sex" value="0" <%=Request.Form("sex")==0?"checked":""%>>
          � �������</font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest">������� / ���� / ���:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="text" name="price" value="<%=isFirst?"":Request.Form("price")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <%}%>
    <% if (tip>=20 && tip<30) { %>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">��� ����������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="sale" value="1" <%=isFirst?"checked":""%> <%=Request.Form("sale")==1?"checked":""%>>
          ��� 
          <input type="radio" name="sale" value="0" <%=Request.Form("sale")==0?"checked":""%>>
          ���������</font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle"> 
        <div align="right"> 
          <p class="digest">��������:&nbsp;&nbsp;</p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="text" name="price" value="<%=isFirst?"":Request.Form("price")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <%} %>
    <tr> 
      <td width="294" valign="middle" height="2"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <select name="country">
            <%
				sql="Select * from  country where regionality=2 and rus_name is not null order by rus_name"
				Records.Source=sql
				Records.Open()
				while(!Records.EOF){
				%>
            <option value="<%=Records("ID").Value%>"
				<%=isFirst&&(Records("ID").Value==190)?"selected":""%>
				<%=!isFirst&&(Records("ID").Value==Request.Form("country"))?"selected":""%>> 
            <%=Records("RUS_NAME").Value%> </option>
            <%
				Records.MoveNext()
				}
				Records.Close()
				%>
          </select>
          </font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle" height="2"> 
        <div align="right"> 
          <p class="digest"><font color="#FF0000">������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <select name="region">
            <%
				sql="Select * from  area where state=0 order by id"
				Records.Source=sql
				Records.Open()
				while(!Records.EOF){
				%>
            <option value="<%=Records("ID").Value%>"
				<%=isFirst&&(Records("ID").Value==54)?"selected":""%>
				<%=!isFirst&&(Records("ID").Value==Request.Form("region"))?"selected":""%>> 
            <%=Records("NAME").Value%> </option>
            <%
				Records.MoveNext()
				}
				Records.Close()
				%>
          </select>
          </font></p>
      </td>
    </tr>
    <tr> 
      <td width="294" valign="middle" height="2"> 
        <div align="right"> 
          <p><font color="#FF0000">�����, ���������� �����:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <input type="text" name="city" value="<%=isFirst?"":Request.Form("city")%>" maxlength="100">
          </font></p>
      </td>
    </tr>
    <tr valign="top"> 
      <td width="294"> 
        <div align="right"> 
          <p><font color="#FF0000">����� ����������:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td valign="top"> 
        <p><font size="2"> 
          <textarea name="txt" cols="40" rows="5"><%=isFirst?"":Request.Form("txt")%></textarea>
          </font></p>
      </td>
    </tr>
  </table>
  <p align="center"> 
    <input type="submit" name="Submit" value="���������� ����������">
  </p>
</form>
<%} 
else 
{%>
<center>
  <h1>&nbsp;</h1>
  <h1><font color="#3333FF">�������, ��� ��������������� ������ ���������� 72RUS!</font></h1>
  <h1><font color="#3333FF"> ���� ���������� ���������!</font></h1>
  <h1>&nbsp;</h1>
  <p><font color="#3333FF"><a href="addmsimg.asp?ms=<%=mid%>">����� �� ������ 
    �������� ���������� ��� ����������� � ����������</a></font></p>
  <p>&nbsp;</p>
  <p><font color="#FF0000">�������� ��� ������� ��� ��������� ��� ���������� �������� 
    ������ ����������:</font></p>
  <p><b><font size="3" color="#FF0000"><%=cw%></font></b></p>
  <hr noshade size="1" width="70%">
  <h1><font color="#0000FF">��� ����������� ����� ���������� ���������� ���������� 
    �� ����� �����,</font></h1>
  <h1><font color="#0000FF">������ ��� �������� �� ����� �� ������������� ������ 
    ����� ���������</font></h1>
</center>
<%}%>
<div align="center"> 
 <!-- start Link.Ru -->

<script language="JavaScript">
// <!--
var LinkRuRND = Math.round(Math.random() * 1000000000);
document.write('<iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF&r='+LinkRuRND+' frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> </iframe>');
// -->
</script>
<noscript>
<iframe src=http://link.link.ru/show?squareid=1389&showtype=2&cat_id=100080&tar_id=1&sc=3&bg=FFFFFF frameborder=0 vspace=0 hspace=0 marginwidth=0 marginheight=0 scrolling=no width=100% height=150> </iframe>
</noscript>

<!-- end Link.Ru -->
 <hr size="1">
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr> 
      <td height="1" bgcolor="#666666"></td>
    </tr>
    <tr> 
      <td height="19" bgcolor="#AFC0D0" align="center">
        <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">�� ������� 
          ��������</a> | </font><font face="Arial, Helvetica, sans-serif"><a href="messages.asp">� 
          �����������</a></font> </p>
      </td>
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
    <div align="center"></div>
</div>
</body>
</html>
