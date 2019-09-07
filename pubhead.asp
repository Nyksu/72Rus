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
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#CCCCCC" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="catarea.asp">Каталог 72RUS.RU</a> 
        | <a href="http://chat.72rus.ru/">Чат</a> | <font size="3"><font color="#FFFFFF"><b><%=nm%></b></font></font></p>
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
%><table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td width="100%"> 
            <p>&gt;&nbsp;<a  href="<%=url%>"><%=hname%></a> 
               
          </td>
          <td bgcolor=#FF0000 width="23" valign="top" bordercolor="#FF0000">&nbsp;</td>
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
	url=TextFormData(Records("URL").Value,"pubheading.asp")
	url+="?hid="+hdd
	Records.MoveNext()
%>
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td width="100%"> 
            <p>&gt;
              <%if (hdd != hid){%>
              <a  class=globalnav href="<%=url%>"> 
              <%}else{%>
              <%}%>
              <%=hname%> 
              <%if (hdd != hid){%>
              </a> 
              <%} else {%>
              <%}%>
              </p>
          </td>
        </tr>
        </tbody> 
      </table>

<%
} 
Records.Close()
%><%
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
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
          <td width="100%"> 
            <p> &gt; 
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
            </p>
          </td>
        </tr>
        </tbody> 
      </table>
<%
} 
Records.Close()
%>
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
            
          <td width="100%"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1"> &gt; </font><a href="messages.asp">Объявления</a><b> 
              </b>(<%=msgcount%>) </p>
            </td>
          </tr>
          </tbody> 
        </table>
        
      <table border=0 cellpadding=0 cellspacing=0 
            width="100%">
        <tbody> 
        <tr> 
            
          <td width="100%" bordercolor="#EBF5ED"> 
            <p><font face="Arial, Helvetica, sans-serif" size="1">&gt; </font><a href="catarea.asp">Каталог 
              сайтов</a><font 
                  face=Verdana size=1> (<%=urlcount%>)</font></p>
            </td>
          </tr>
          </tbody> 
        </table>
        
      
    </td>
    <td height="369"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="23">&nbsp;</td>
                <td>
                  <p><font size="1">:: <%=path%> <font> 
                    <%if (isnews==1) {%>
                    </font> > <a href="archive.asp?hid=<%=hid%>&pg=<%=pg%>">Архив 
                    рубрики</a> 
                    <%}%>
                    </font></p>
                </td>
              </tr>
            </table>
            <p align="center"> 
              <%if (usok) {%>
              <font size="1"><b> | <a href="ednewsheading.asp?hid=<%=hid%>">Редактировать 
              рубрику</a> | <a href="delpubheading.asp?hid=<%=hid%>">Удалить 
              рубрику</a> | <a href="addnewsheading.asp?hid=<%=hid%>">Добавить 
              подраздел рубрики</a> | 
              <%}%>
              <%if (tpm<7) {%>
              <a href="addpub.asp?hid=<%=hid%>">Добавить 
              публикацию</a></b></font> | 
              <%}%>
            </p>
          </td>
        </tr>
      </table>
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
          <td width="50" align="center" valign="middle"> </td>
          <td bordercolor="#FFFFFF">
		  <dl><dd><li><p><%if (imgname != "") {%>
                  <img src="<%=imgname%>" BORDER="1" align="left" > 
                  <%}else{%>
          &nbsp; 
          <%}%>
                    &nbsp;<a href="<%=url%>"><b><font size="3"><%=pname%></font></b></a><br>
                      [<%=pdat%>] <%=digest%> 
                  <%if (usok) {%><br>
             
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
          </dd></dl></td>
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
              </b><a href="pubheading.asp?hid=<%=hid%>&pg=<%=pg-1%>">Предыдущая 
              страница раздела</a> 
              <%
  } 
  if (pos>lpag*(1+pg)) {
  %>
              <a href="pubheading.asp?hid=<%=hid%>&pg=<%=pg+1%>">Следующая страница 
              раздела</a> 
              <%}%>
              </font></p>
          </td>
        </tr>
      </table>
    </td>
    <td WIDTH="150" height="369" align="center">&nbsp; </td>
  </tr>
</table>
<hr size="1">
<div align="center">
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
    </font>
</div>
</body>
</html>
