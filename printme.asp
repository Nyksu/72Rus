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

var pid=parseInt(Request("pid"))
if (isNaN(pid)) {Response.Redirect("index.asp")}

var hid=0
var sminame=""
var tit=""
var hdd=0
var nm=""
var hiname=""
var period=0
var path=""
var hadr=""
var pname=""
var pdat=""
var autor=""
var news=""
var imgname=""
var imgLname=""
var imgRname=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var sos=""
var ishtml=0
var isnews=1
var lgok=false
var usok=false
var bnm=""
var bpos=""
var bid=0

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
	if (Session("tip_mem_pub")<4) {lgok=true}
} else {
	if ((Session("is_adm_mem")!=1) && (Session("is_host")!=1)) {
		sql="Select * from smi where users_id="+Session("id_mem")+"and id="+smi_id
		Records.Source=sql
		Records.Open()
		if (!Records.EOF) {
			usok=true
			lgok=true
		}
		Records.Close()
	} else {
		usok=true
		lgok=true
	}
}


Records.Source="Select t1.* from publication t1, heading t2 where t1.state=1 and t1.id="+pid+" and t1.heading_id=t2.id and t2.smi_id="+smi_id
Records.Open()
if (Records.EOF) {
	Records.Close()
	Response.Redirect("index.asp")
}
hid=Records("heading_id").Value
pname=String(Records("NAME").Value)
pdat=Records("PUBLIC_DATE").Value
autor=TextFormData(Records("AUTOR").Value,"")
ishtml=TextFormData(Records("ISHTML").Value,"0")
Records.Close()

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()

tit=sminame

hdd=hid
while (hdd>0) {
	Records.Source="Select * from heading where id="+hdd
	Records.Open()
	nm=String(Records("NAME").Value)
	isnews=Records("ISNEWS").Value
	hadr=TextFormData(Records("URL").Value,"pubheading.asp")
	if (hdd==hid) {
		hiname=String(Records("NAME").Value)
		period=Records("PERIOD").Value
	}
	path="<a href=\""+hadr+"?hid="+hdd+"\"><font color=\"#FFFFFF\">"+nm+"</font></a> > "+path
	hdd=Records("HI_ID").Value
	Records.Close()
}

var ddt = new Date()
var dt=""
var str=""
var sumdat=Server.CreateObject("datesum.DateSummer")
if ( isnews ) {
str=String(ddt.getMonth()+1)
if (str.length==1) {str="0"+str}
dt="."+str+"."+ddt.getYear()
str=String(ddt.getDate())
if (str.length==1) {str="0"+str}
dt=str+dt
dt=sumdat.SummToDate(dt,"-"+period)
if (sumdat.DateComparing(dt,pdat) > 0) {sos="В архиве"} else {sos="В публикации"}
}

path="<a href=\"index.asp\"><font color=\"#FFFFFF\">"+sminame+"</font></a> > "+path
tit+=" | "+hiname

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

imgname=PubImgPath+pid+".gif"
if (!fs.FileExists(PubFilePath+pid+".gif")) { imgname="" }
if (imgname=="") {
	imgname=PubImgPath+pid+".jpg"
	if (!fs.FileExists(PubFilePath+pid+".jpg")) { imgname="" }
}

imgLname=PubImgPath+"l"+pid+".gif"
if (!fs.FileExists(PubFilePath+"l"+pid+".gif")) { imgLname="" }
if (imgLname=="") {
	imgLname=PubImgPath+"l"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"l"+pid+".jpg")) { imgLname="" }
}

imgRname=PubImgPath+"r"+pid+".gif"
if (!fs.FileExists(PubFilePath+"r"+pid+".gif")) { imgRname="" }
if (imgRname=="") {
	imgRname=PubImgPath+"r"+pid+".jpg"
	if (!fs.FileExists(PubFilePath+"r"+pid+".jpg")) { imgRname="" }
}

%>

<html>
<head>
<title><%=pname%> - <%=tit%> -  г. Тюмень</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
<style type=text/css> 
a:hover { color: #0066FF; text-decoration: underline; color: #0000FF;font: 12px Arial, Helvetica, sans-serif}
a.globalnav:visited { font-size: 12px; font-family: Arial, Helvetica, Verdana, sans-serif; text-decoration: none; line-height: 11px ; color: #0000FF}
a.globalnav:link { font-size: 12px; font-family: Arial, Helvetica, Verdana, sans-serif; text-decoration: none; line-height: 11px; color: #0000FF}
a.globalnav:hover { color: #FF0000; font-size: 12px; font-family: Arial, Helvetica, Verdana, sans-serif; text-decoration: none; line-height: 11px}
 .secondarystories { font-family: Arial, Helvetica, Verdana, sans-serif; font-size: 11px; font-style: normal; line-height: normal; font-weight: normal; font-variant: normal; text-transform: none; color: #000000; text-decoration: none}a {color:  #0000FF;font: 12px Arial, Helvetica, sans-serif}
h1 { font-family: Arial, Helvetica, sans-serif; font-size: 17px; font-style: normal; line-height: 20px; font-weight: bold; font-variant: normal; color: #123D87; text-decoration: none ; margin-left: 5px; vertical-align: middle; margin-top: 3px; margin-bottom: 2px; text-transform: uppercase}
p { font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 11pt; margin: 0px 4px 0px 0px; padding: 3px 3px 3px 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 15px; font-style: normal; line-height: 18px; font-weight: bold; font-variant: normal; color: #123D87; text-decoration: none ; margin-left: 5px; vertical-align: middle; margin-top: 3px; margin-bottom: 2px }
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="35" marginwidth="35">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left">&nbsp;</td>
    <td align="center"> 
      <script language="javascript" src="banshow.asp?rid=4"></script>
    </td>
    <td width="150" align="center"> 
      <!--RAX counter-->
      <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
      <!--/RAX-->
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#666666">
  <tr> 
    <td height="16" width="350"> 
      <p><%=path%> </p>
    </td>
    <td height="16" align="right"> 
      <p><font color="#FFFFFF"><%=tit%></font></p>
    </td>
    <form name="form1" method="post" action="search.asp">
    </form>
  </tr>
</table>


<table width="100%" border="0" cellpadding="0" align="center" cellspacing="0" bgcolor="#E1F4FF" height="300">
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"> 
      <table width="98%" border="0" bordercolor="#FFFFFF" align="center" cellspacing="0" class="base_text">
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td bordercolor="#FFFFFF"> 
            <h1 align="center"><font color="#000000"><%=pname%></font></h1>
            <p align="right"><b>&copy; <%=pdat%>&nbsp; <%=autor%> </b><br>
              Версия для печати: 
              <script>
	id='';
	date=new Date();
	s=date.getDate()+" ";
	
	month=date.getMonth();
	if(month==0)s+="Января";
	else if(month==1)s+="Февраля";
	else if(month==2)s+="Марта";
	else if(month==3)s+="Апреля";
	else if(month==4)s+="Мая";
	else if(month==5)s+="Июня";
	else if(month==6)s+="Июля";
	else if(month==7)s+="Августа";
	else if(month==8)s+="Сентября";
	else if(month==9)s+="Октября";
	else if(month==10)s+="Ноября";
	else if(month==11)s+="Декабря";
	
	s+=" 2003"+"г"+". ";
	
	document.write(s);
</script>
            </p>
          </td>
        </tr>
        <tr valign="top" bordercolor="#FFFFFF"> 
          <td height="55"> 
            <p> 
              <%if (imgLname != "") {%>
              <img src="<%=imgLname%>" align="left" border="1" > 
              <%}else{%>
              &nbsp; 
              <%}%>
              <font face="Times New Roman, Times, serif"><%=news%> </font></p>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellspacing="0">
        <tr bordercolor="#CCCCCC"> 
          <td valign="top" height="40"> 
            <div align="center"> 
              <%if (imgRname != "") {%>
              <img src="<%=imgRname%>" border="1" > 
              <%}else{%>
              &nbsp; 
              <%}%>
            </div>
          </td>
        </tr>
      </table>
      <table border="0" cellspacing="2" cellpadding="2" width="100%">
        <tr> 

        </tr>
      </table>
      
      
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=6"></script>
  <hr size="1">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#666666"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
          портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
          Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
          Русинтел</font></a> <font color="#FFFFFF">&copy; 2002 - 2003</font></font></b></p>
      </td>
    </tr>
  </table>
  <hr size="1">
  <p align="center"><a href="/">На главную</a> | <a href="http://auto.72rus.ru">Авто 
    Тюмень</a> | <a href="http://www.auction.72rus.ru/">Аукцион</a> | <a href="messages.asp">Объявления</a> 
    | <a href="Rail_roads.asp">Расписание</a> | <a href="catarea.asp">Тюменский 
    Каталог</a> | <br>
    Улуги хостинга, создания интернет сайтов, рекламы в интернет от компании РУСИНТЕЛ 
    <a href="http://www.rusintel.ru">www.rusintel.ru</a> 
</div>
<p align="center"> 
  <!--RAX logo-->
  <a href="http://www.rax.ru/click" target=_blank><img src="http://counter.yadro.ru/logo?16.1" border=0 width=1 height=1></a> 
  <!--/RAX-->
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
hotlog_r+"&' border=0 width=1 height=1></a>")</script>
  <noscript><a href=http://click.hotlog.ru/?46088 target=_top><img
src="http://hit3.hotlog.ru/cgi-bin/hotlog/count?s=46088&im=105" border=0 
width="1" height="1" alt="HotLog"></a></noscript> 
  <!-- /HotLog -->
  <!--Begin of HMAO RATINGS-->
  <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"><img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="1" height="1"></a> 
  <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
  <!--End of HMAO RATINGS-->
</body>
</html>
