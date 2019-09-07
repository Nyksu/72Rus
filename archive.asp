<%@LANGUAGE="JScript"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\url.inc" -->


<%
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
var isnewspb=1
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
isnewspb=Records("ISNEWS").Value
Records.Close()

if (isnewspb==0) {Response.Redirect("pubheading.asp?hid="+hid)}

var hidimg=""
hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	hadr=TextFormData(Records("URL").Value,"archive.asp")
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

<html>
<head>
<title><%=tit%> :: Тюмень и Тюменская Область</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="72RUS.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}--></style>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0">
  <tr> 
    <td width="16" bgcolor="#003366">&nbsp;</td>
    <td bgcolor="#003366"> 
      <p><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Тюмень 
        и Тюменская Область :: <%=sminame%> :: <%=nm%></b></font> <font face="Arial, Helvetica, sans-serif" size="2"><b><font face="Arial, Helvetica, sans-serif" size="1"> 
        <%if (tpm<4) {%>
        <a href="headimglst.asp?hid=<%=hid%>"><font color="#FFFFFF">Добавить 
        Имидж раздела</font></a></font></b><b><font face="Arial, Helvetica, sans-serif" size="1"> 
        <%}%>
        </font></b></font><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>:: 
        </b></font></p>
    </td>
    <td bgcolor="#003366" width="23"><a href="admarea.asp"><img src="HeadImg/round_inv.gif" width="23" height="23" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="300" align="center" valign="middle"><%if (hidimg != "") {%><a href="/"><img src="<%=hidimg%>" border="0" width="300" height="60" alt="На главную страницу"></a><%} else {%><a href="/"><img src="HeadImg/logo_72.gif" width="300" height="60" border="0" alt="На главную страницу"></a><%}%></td>
    <td align="left" width="468"> 
      <table width="120" border="0" cellspacing="0" height="60">
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
          <td align="CENTER"><%=news%></td>
          <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
        </tr>
      </table>
    </td>
    <td align="right"><font color="#FFFFFF"><font face="Arial, Helvetica, sans-serif" size="-2"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
      <!--begin of Rambler's Top100 code -->
      <a href="http://top100.rambler.ru/top100/"> <img src="http://counter.rambler.ru/top100.cnt?388244" alt="" width=1 height=1 border=0></a> 
      <!--end of Top100 code-->
      </font></b></font></font> </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="8" bgcolor="#003366">
  <tr> 
    <td></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="8" bgcolor="#003366">
  <tr> 
    <td width="150" align="center"> 
      <p align="center"> <font size="3"><b><font color="#FFFFFF"> Тюмень <a href="/"><font color="#FFFFFF">72rus.ru</font></a></font></b></font></p>
    </td>
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="23" valign="top" bgcolor="#FF0000"><img src="HeadImg/round_blu_red.gif" width="23" height="23"></td>
          <td width="285"> 
            <table border="0" cellspacing="0" cellpadding="0" width="285" bordercolor="#FFFFFF">
              <tr> 
                <td align="left" width="23" bgcolor="#FF0000">&nbsp;</td>
                <td bgcolor="#FF0000" align="center"> 
                  <p><font size="3"><font color="#FFFFFF"><b><%=nm%></b></font></font> 
                </td>
                <td bgcolor="#FF0000" align="left" width="23"> <img src="HeadImg/round_red_up_r.gif" width="23" height="23"></td>
              </tr>
            </table>
          </td>
          <td align="left" valign="middle">&nbsp;</td>
          <td align="center" width="150">
            <p><b><font size="3"> 
              <%
// В переменной bk содержится код блока новостей
var bk=35
// Не забывать его менять!!
Records.Source="Select * from block_news where id="+bk+" and smi_id="+smi_id
Records.Open()
if (!Records.EOF ) {
blokname=TextFormData(Records("SUBJ").Value,"")
}
Records.Close()
%>
              <%=blokname%></font></b> </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" HEIGHT="350">
  <tr valign="top"> 
    <td WIDTH="150" height="369"> 
      <%
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id="+hid+" and smi_id="+smi_id+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"archive.asp")
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
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tr> 
          <td bgcolor=#FF0000 width="100%" bordercolor="#FF0000"> 
            <p><font  face=Verdana size=1><b> <img src="HeadImg/arrow2.gif" width="11" height="10" align="middle">&nbsp;<a  href="<%=url%>"><font color="#FFFFFF"><%=hname%></font></a> </b></font> 
          </td>
          <td bgcolor=#003366 width="2" valign="top" bordercolor="#FF0000"></td>
        </tr>
        <tbody> 
        <tr> 
          <td bgcolor=#003366 width="100%" height="2" bordercolor="#FF0000"></td>
          <td bgcolor=#003366 width="2" valign="top" height="2"></td>
        </tr>
        </tbody> 
      </table>
      <%
} 
Records.Close()
delete recs
%>
      <%
isnews=1
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hdd=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"archive.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><b><font  face=Verdana size=1><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              <%if (hdd != hid){%>
              <a  class=globalnav href="<%=url%>"> 
              <%}else{%>
              <%}%>
              <%=hname%> 
              <%if (hdd != hid){%>
              </a> 
              <%} else {%>
              <%}%>
              </font></b></p>
          </td>
        </tr>
        </tbody> 
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
	url=TextFormData(Records("URL").Value,"archive.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><b><font  face=Verdana size=1><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
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
              </font> </b></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <%
} 
Records.Close()
%>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              </font><font 
                  face=Verdana size=1><b><a href="messages.asp">Объявления</a> 
              </b>(<%=msgcount%>) </font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table bgcolor=#ffd34e border=1 cellpadding=0 cellspacing=0 
            width="100%" bordercolor="#003366">
        <tbody> 
        <tr> 
          <td bgcolor=#EBF5ED width="100%"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1"><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
              </font><font 
                  face=Verdana size=1><b><a href="catarea.asp">Каталог сайтов</a></b> 
              (<%=urlcount%>)</font></p>
          </td>
        </tr>
        </tbody> 
      </table>
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td bgcolor=#003366 >&nbsp; </td>
          <td bgcolor=#003366 width="23"> <img src="HeadImg/round_inv.gif" width="23" height="23"></td>
        </tr>
        </tbody> 
      </table>
      <%
// В переменной bk содержится код блока новостей
var bk=38
// Не забывать его менять!!
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
		news+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
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
          <td align="center" ><%=news%> </td>
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
    <td height="369"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="23"><font size="1"><img src="HeadImg/round_red_up_r.gif" width="23" height="23"></font></td>
                <td> 
                  <p><font size="1">:: <%=path%> <font> 
                    </font> > <a href="pubheading.asp?hid=<%=hid%>">Активные публикации</a>
                    </font></p>
                </td>
              </tr>
            </table>
            
              <%if (usok) {%>
              <p align="center"> <font size="1"><b> | <a href="ednewsheading.asp?hid=<%=hid%>">Редактировать 
              рубрику</a> | <a href="delarchive.asp?hid=<%=hid%>">Удалить 
              рубрику</a> | <a href="addnewsheading.asp?hid=<%=hid%>">Добавить 
              подраздел рубрики</a> | 
              <%}%>
              <%if (tpm<7) {%>
              <a href="addpub.asp?hid=<%=hid%>">Добавить 
              публикацию</a></b></font> | 
                          </p><%}%>

          </td>
        </tr>
      </table>
      <%
if (isnewspb==1){
Records.Source="Select * from publication where heading_id="+hid+"and state=1 and public_date<'TODAY'-"+period+" and public_date<='TODAY' order by public_date desc, id desc"
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
      <table width="100%" border="0" bordercolor="#FFFFFF" cellspacing="0">
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td width="50" align="center" valign="middle"> </td>
          <td bordercolor="#FFFFFF"> 
            <dl>
              <dd>
                <li>
                  <p>
                    <%if (imgname != "") {%>
                    <img class="digest" src="<%=imgname%>" BORDER="1" align="left" > 
                    <%}else{%>
                    &nbsp; 
                    <%}%>
                    &nbsp;<a href="<%=url%>"><b><font size="3"><%=pname%></font></b></a><br>
                    [<%=pdat%>] <%=digest%> 
                    <%if (usok) {%>
                    <br>
                    <a href="pubresume.asp?pid=<%=pid%>&st=0">Остановить 
                    публикацию</a> | <a href="delpub.asp?pid=<%=pid%>">Удалить 
                    публикацию</a>|<br>
                    <a href="bloknews.asp?pid=<%=pid%>">Разместить 
                    в блок</a> | <a href="edpub.asp?pid=<%=pid%>">Редактировать</a> 
                    <%}%>
                    <% if (tpm<4) {%>
                    | <a href="addpubimg.asp?pid=<%=pid%>">Добавить 
                    фотографию</a></p>
                  <%}%>
              </dd>
            </dl>
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
              <font color="#FF0000">&lt;&lt;</font> <a href="archive.asp?hid=<%=hid%>&pg=<%=pg-1%>">Предыдущая 
              страница раздела</a> 
              <%
  } 
  if (pos>lpag*(1+pg)) {
  %>
              <a href="archive.asp?hid=<%=hid%>&pg=<%=pg+1%>">Следующая 
              страница раздела</a> </b><b> <font color="#FF0000">&gt;&gt;</font> 
              <%}%>
              </b></font></p>
          </td>
        </tr>
      </table>
    </td>
    <td WIDTH="150" height="369" align="center"> 
      <%

// В переменной bk содержится код блока новостей
var bk=35
// Не забывать его менять!!
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
	hadr=TextFormData(recs("URL").Value,"archive.asp")
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
      <b><font size="3"> </font></b> 
      <table width="120" border="0" cellspacing="0" height="60">
        <tr> 
          <td align="CENTER"><%=news%></td>
        </tr>
      </table>
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <table width="708" border="0" cellspacing="0" height="60" align="center" cellpadding="0">
    <tr> 
      <%
// В переменной bk содержится код блока новостей
var bk=36
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
      <td align="CENTER" width="150"> 
        <div align="center"><%=news%></div>
      </td>
      <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#FF0000"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
          портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
          Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
          Русинтел</font></a> <font color="#FFFFFF">&copy; 2002 - 2003</font></font></b></p>
      </td>
    </tr>
  </table>
  <hr size="1">
  <p align="center"><font face="Arial, Helvetica, sans-serif"> 
    <%
// маркек признака новостей
isnews=1
// если необходимо вывести рубрики не новостей то установить в ноль

var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="archive.asp"}
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
// маркек признака новостей
isnews=0
// если необходимо вывести рубрики не новостей то установить в ноль
var recs=CreateRecordSet()
Records.Source="Select * from heading where hi_id=0 and smi_id="+smi_id+" and isnews="+isnews+" order by name"
Records.Open()
while (!Records.EOF)
{
	hid=String(Records("ID").Value)
	hname=String(Records("NAME").Value)
	per=Records("PERIOD").Value
	url=TextFormData(Records("URL").Value,"")
	if (url=="") {url="archive.asp"}
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
    </font><br>
    | <a href="http://auto.72rus.ru">Авто.72rus</a> | <a href="http://www.auction.72rus.ru/">Аукцион</a> 
    | <a href="messages.asp">Объявления</a> | <a href="Rail_roads.asp">Расписание</a> 
    | <a href="catarea.asp">Тюменский Каталог</a> | <a href="terms.html">Условия 
    использования</a> | <br>
    © 2002 - 2003 <a href="http://www.rusintel.ru">Rusintel Company</a> 
  <p align="center"></p>
</div>
<div align="CENTER"> 
  <p> 
    <!-- HotLog -->
<script language="javascript">
hotlog_js="1.0";
hotlog_r=""+Math.random()+"&s=46088&im=105&r="+escape(document.referrer)+"&pg="+
escape(window.location.href);
document.cookie="hotlog=1; path=/"; hotlog_r+="&c="+(document.cookie?"Y":"N");
</script><script language="javascript1.1">
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
  </p>
</div>
</body>
</html>
