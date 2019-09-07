<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->

<%

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="admarea.asp"
Response.Redirect("login.asp")
}

// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!

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
%>

<html>
<head>
<title><%=tit%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" align="center">
  <tr bgcolor="#CCCCCC"> 
    <td bgcolor="#003366"> 
      <p><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><%=sminame%> :: <%=nm%></b></font> <font face="Arial, Helvetica, sans-serif" size="2"><b><font face="Arial, Helvetica, sans-serif" size="1"> 
        <%if (tpm<4) {%>
        <a href="headimglst.asp?hid=<%=hid%>"><font color="#FFFFFF">Добавить 
        Имидж раздела</font></a></font></b><b><font face="Arial, Helvetica, sans-serif" size="1"> 
        <%}%>
        </font></b></font></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="60">
  <tr>
    <td>
<%if (hidimg != "") {%><img src="<%=hidimg%>" border="0"><%} else {%><font face="Arial, Helvetica, sans-serif" size="2"><b><%=sminame%></b></font><%}%></td>
    <td width="468" align="center">
      <table width="120" border="0" cellspacing="0" height="60" align="center">
        <tr> 
          <%
// В переменной bk содержится код блока новостей
var bk=29
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
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
          <td align="CENTER"> <%=news%> </td>
          <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
        </tr>
      </table>
    </td></tr></table>
<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="#666666" bordercolor="#FFFFFF"> 
    <td bgcolor="#FFCC66" bordercolor="#FFCC66"> 
      <p><b>:: <%=path%> </b> <FONT COLOR="#000000"><B><FONT SIZE="2" COLOR="#FFFFFF"> 
        <%if (isnews==1) {%>
        </FONT></B> <B>></B></FONT><FONT COLOR="#000000"><B> <A HREF="archive.asp?hid=<%=hid%>&pg=<%=pg%>">Архив 
        рубрики</A></B> <A HREF="archive.asp?hid=<%=hid%>&pg=<%=pg%>"> 
        </A> 
        <%}%>
        </FONT> 
        <%if (usok) {%>
        <font size="1"><b> | <a href="ednewsheading.asp?hid=<%=hid%>">Редактировать 
        рубрику</a> | <a href="delpubheading.asp?hid=<%=hid%>">Удалить 
        рубрику</a> | <a href="addnewsheading.asp?hid=<%=hid%>">Добавить 
        подраздел рубрики</a> | 
        <%}%>
        <%if (tpm<7) {%>
       <a href="addpub.asp?hid=<%=hid%>">Добавить 
        публикацию</a></b></font> 
        <%}%>
      </p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" HEIGHT="350">
  <tr valign="top"> 
    <td WIDTH="150" height="369"> 
      <table width="100%" border="1" cellspacing="0" cellpadding="0" bgcolor="#E6F2EC" align="CENTER">
        <tr> 
          <td bgcolor="#FFCC66" bordercolor="#FFCC66"><font face="Arial, Helvetica, sans-serif" size="1"> 
            <%
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
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
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
            &nbsp; - <a  href="<%=url%>"><b><font size="2"><%=hname%></font></b></a> <br>
            <%
} 
Records.Close()
delete recs
%>
            </font></td>
        </tr>
      </table>
      <p align="CENTER"><font face="Arial, Helvetica, sans-serif" size="1"><b> 
        <%
isnews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
        </b></font> </p>
      <div align="CENTER"> 
        <table width="100%" border="1" cellspacing="0" bordercolor="#FFFFFF" align="center">
          <tr bordercolor="#CC3333"> 
            <td bordercolor="#CC3333"> 
              <p><b> 
                <%if (hdd != hid){%>
                <a  class=globalnav href="<%=url%>"> 
                <%}else{%>
                <%}%>
                <%=hname%> 
                <%if (hdd != hid){%>
                </a> 
                <%} else {%>
                <%}%>
                </b></p>
            </td>
          </tr>
        </table>
        <%
} 
Records.Close()
%>
        <%
isnews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
        <table width="100%" border="1" cellspacing="0" bordercolor="#FFFFFF" align="center">
          <tr bordercolor="#CC3333"> 
            <td> 
              <p><b> 
                <%if (hdd != hid){%>
                <a  class=globalnav href="<%=url%>"> 
                <%}else{%>
                <%}%>
                <%=hname%> 
                <%if (hdd != hid){%>
                </a><font color="#FFCC66"> 
                <%} else {%>
                </font> 
                <%}%>
                </b></p>
            </td>
          </tr>
        </table>
        <p> 
          <%
} 
Records.Close()
%>
        </p>
        <p align="CENTER">Подрубрики</p>
        <p><font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
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
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
          &nbsp;<a  href="<%=url%>"><b><font size="2"><%=hname%></font></b></a></font></p>
        <p><font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
} 
Records.Close()
delete recs
%>
          </font></p>
        
      </div>
    </td>
    <td height="369"> 
      <TABLE WIDTH="100%" BORDER="1" CELLSPACING="0" CELLPADDING="0" BGCOLOR="#E6F2EC" ALIGN="CENTER">
          <TR> 
            <TD BGCOLOR="#E6F2EC" BORDERCOLOR="#E6F2EC"><FONT FACE="Arial, Helvetica, sans-serif" SIZE="1"> 
              <%
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
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
		recs.Source="Select count_pub  from get_count_pub_show("+hdd+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
              &nbsp;<A  HREF="<%=url%>"><B><FONT SIZE="2"><%=hname%></FONT></B></A> / 
              <%
} 
Records.Close()
delete recs
%>
              </FONT></TD>
          </TR>
        </TABLE>
       
      <h1 align="center"><font size="3"><%=nm%></font></h1>
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
%>
        

      <table width="100%" border="0" bordercolor="#FFFFFF" align="center" cellspacing="0">
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td width="150" align="center" valign="middle"> 
            <%if (imgname != "") {%>
            <img src="<%=imgname%>" BORDER="1" > 
            <%}else{%>
            &nbsp; 
            <%}%>
          </td>
          <td bordercolor="#FFFFFF"> 
            <p><font size="2" face="Arial, Helvetica, sans-serif"> <%=pdat%>&nbsp;<a href="<%=url%>"><%=pname%></a> <%=digest%></font>| 
              <%if (usok) {%>
              <br>
              <a href="pubresume.asp?pid=<%=pid%>&st=0">Остановить 
              публикацию</a> | <a href="delpub.asp?pid=<%=pid%>">Удалить 
              публикацию</a>|<br>
              <a href="bloknews.asp?pid=<%=pid%>">Разместить 
              в блок</a> | <a href="edpub.asp?pid=<%=pid%>">Редактировать</a> 
              <%}%>
              <% if (tpm<4) {%>
            </p>
            <p><a href="addpubimg.asp?pid=<%=pid%>">Добавить 
              фотографию</a></p>
            <%}%>
          </td>
        </tr>
      </table>
      <div align="center"> 
        <%
	}
	Records.MoveNext()
	pos+=1
} 
Records.Close()
%>
      </div>
      <table width="100%" border="1" cellspacing="0" bordercolor="#FFFFFF" align="center">
        <tr bgcolor="#CCCCCC" bordercolor="#CC3333"> 
          <td bgcolor="#FFFFFF" align="CENTER" bordercolor="#FFFFFF" height="23"> 
            <p><font face="Arial, Helvetica, sans-serif" size="2"><b> 
              <%if (pg>0) {%>
              <a href="pubheading.asp?hid=<%=hid%>&pg=<%=pg-1%>">Предыдущая 
              страница раздела</a> 
              <%
  } 
  if (pos>lpag*(1+pg)) {
  %>
              <a href="pubheading.asp?hid=<%=hid%>&pg=<%=pg+1%>">Следующая 
              страница раздела</a> </b><b> 
              <%}%>
              </b></font></p>
          </td>
        </tr>
      </table>
    </td>
    <td WIDTH="150" height="369"> 

    </td>
  </tr>
</table>
<hr size="1">
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="150">&nbsp;</td>
      <td align="center"><font face="Arial, Helvetica, sans-serif" size="1"> 
        <%
isnews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
        <%if (hdd != hid){%>
        <a  href="<%=url%>"> 
        <%}else{%>
        <%}%>
        <%=hname%> 
        <%if (hdd != hid){%>
        </a> 
        <%} else {%>
        <%}%>
        | 
        <%
} 
Records.Close()
%>
        <%
isnews=0
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
        </font> <font face="Arial, Helvetica, sans-serif" size="1"> 
        <%if (hdd != hid){%>
        <a  href="<%=url%>"> 
        <%}else{%>
        <%}%>
        <%=hname%> 
        <%if (hdd != hid){%>
        </a> 
        <%} else {%>
        <%}%>
        | 
        <%
} 
Records.Close()
%>
        </font></td>
      <td width="150">&nbsp;</td>
    </tr>
  </table>
  <font face="Arial, Helvetica, sans-serif" size="1"> </font> 
  <hr size="1">
</div>
<p align="CENTER">&nbsp; </p>
<div align="CENTER"> 
  <table width="468" border="0" cellspacing="0" height="60">
    <tr> 
      <%
// В переменной bk содержится код блока новостей
var bk=29
// Не забывать его менять!!
var recs=CreateRecordSet()
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
imgLname=""
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


%><td align="CENTER">
        <%=news%> 
        </td><%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </tr>
  </table>
  <p>&nbsp; </p>
</div>
</body>
</html>
