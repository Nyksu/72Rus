<%@LANGUAGE="JScript"%>

<%
var getData = Server.CreateObject("SOFTWING.ASPtear")
var Request_GET = 2
var dirData = ""
var startchar=0
var endchar=0
var startstr="<CENTER>Search"
var strres=""
var strpar=Request.QueryString()
var tit=Request.form("search")
var ss=""
var s1=""

dirData = getData.Retrieve("http://search.dmoz.org/cgi-bin/search?"+strpar, Request_GET, "", "", "")

startchar=dirData.indexOf("<CENTER>")
endchar=dirData.indexOf("<p><CENTER>")

strres=dirData.substring(startchar,endchar)

while (strres.indexOf("–<b>∫</b>")>-1) {strres=strres.replace("–<b>∫</b>","–∫")}
while (strres.indexOf("–<b>ê</b>")>-1) {strres=strres.replace("–<b>ê</b>","–ê")}
while (strres.indexOf("–<b>®</b>")>-1) {strres=strres.replace("–<b>®</b>","–®")}
while (strres.indexOf("–<b>û</b>")>-1) {strres=strres.replace("–<b>û</b>","–û")}
while (strres.indexOf("–<b>ö</b>")>-1) {strres=strres.replace("–<b>ö</b>","–ö")}
while (strres.indexOf("search?")>-1) {strres=strres.replace("search?","copypage.asp?")}
while (strres.indexOf("<a href=\"http://dmoz.org/")>-1) {strres=strres.replace("<a href=\"http://dmoz.org/","<a href=\"/directory.asp?cat=/")}

//startchar=dirData.indexOf("<center>&nbsp;")
//endchar=dirData.indexOf("Next</a>")
//ss=strres.substring(startchar,endchar)
//startchar=dirData.indexOf("?search=")
//endchar=dirData.indexOf("&utf8")
//ss=ss.substring(startchar+1,endchar)
//while (strres.indexOf(ss)>-1) {strres=strres.replace(ss,strpar)}

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html"; charset=UTF-8>
<link rel="stylesheet" href="72.css" type="text/css">
</HEAD>
<Body>
<form accept-charset="UTF-8" action="copypage.asp" method="get" target=_top>
  <table width="467" border="0" cellspacing="0" cellpadding="0" height="60"">
    <tr> 
      <td bgcolor="#FFCC00"> 
        <div align="left">
<input accept-charset="UTF-8" type="text" name="search" size="25">
          <input type="submit" name="Submit" value="–ü–æ–∏—Å–∫">
        </div>
      </td>
    </tr>
  </table width="100%" border="0" cellspacing="0" cellpadding="0">
  </form>

<%=strres%>

</Body>
</HTML>