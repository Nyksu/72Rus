<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->

<%
var tip=0
var subj_id=parseInt(Request("subj"))
var namarea=""
var name=""
var inname=""
var outname=""
var ErrorMsg=""
var hid=0
var ShowForm=true
var tit="" 
var path=""
var sbj=0
var nm=""
var id=0
var dase=""
var sql=""

if (isNaN(subj_id)) {Response.Redirect("messages.asp")}
if (subj_id==0) {Response.Redirect("messages.asp")}

if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog ) {Response.Redirect("messages.asp?subj="+subj_id)}

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
		dase=tip
		hid=Records("HI_ID").Value
	}
	else {path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> | "+path}
	sbj=Records("HI_ID").Value
	tit=nm+" : "+tit
	Records.Close()
}

path="<a href=\"messages.asp\">Все темы</a> | "+path

isFirst=String(Request.Form("Submit"))=="undefined"

if(!isFirst){

     name=TextFormData(Request.Form("name"),"")
	 inname=TextFormData(Request.Form("inname"),"")
	 outname=TextFormData(Request.Form("outname"),"")
	 dase=Request.Form("dase")
	 
	 if (name.length<3) {ErrorMsg+="Слишком короткое наименование.<br>"}
 	 if (inname.length<2) {ErrorMsg+="Слишком короткое наименование.<br>"}
	 if (outname.length<2) {ErrorMsg+="Слишком короткое наименование.<br>"}
 
	  if (ErrorMsg==""){
	  		sql="Update TRADE_SUBJ set NAME='%NAME', HI_ID=%SUBJ, OUT_NAME='%ONAM' ,IN_NAME='%INAM' ,MARKETPLACE_ID=%MAR ,MSG_TYPE=%TP  where ID="+subj_id
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%TP",dase)
			sql=sql.replace("%MAR",market)
			sql=sql.replace("%INAM",inname)
			sql=sql.replace("%SUBJ",hid)
			sql=sql.replace("%ONAM",outname)
			Connect.BeginTrans()
			try{
			Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="Ошибка изменения.<br>"
			}
			if (ErrorMsg==""){
		   Connect.CommitTrans()
		   ShowForm=false
			}
	  }

}

%>

<html>
<head>
<title>Изменение раздела : (<%=tit%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="/72RUS.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00">
  <tr> 
    <td width="150" valign="middle" background="HeadImg/yellow.jpg" height="19"> 
      <div align="center"> <a href="index.asp"><img src="HeadImg/home.gif" width="14" height="14" align="absbottom" border="0" alt="Home"></a> 
        <font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНСКИЙ РЕГИОН</b></font></div>
    </td>
    <td bordercolor="#FFCC00" background="HeadImg/yellow.jpg" height="19"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>WWW.72RUS.RU</b></font></div>
    </td>
    <td width="468" background="HeadImg/yellow.jpg" height="19"> 
      <table width="410" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="HeadImg/tmn.gif" width="82" height="20"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2"><b>ТЮМЕНЬ</b></font></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b><a href="messages.asp">ОБЪЯВЛЕНИЯ</a></b></font></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><a href="catarea.asp"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>КАТАЛОГ</b></font></a></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><a href="Rail_roads.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>РАСПИСАНИЕ</b></font></a></div>
          </td>
          <td width="82" height="20" background="HeadImg/tmn.gif"> 
            <div align="center"><a href="marketBuilder.html"><font face="Arial, Helvetica, sans-serif" size="-2" color="#000000"><b>АНАЛИТИКА</b></font></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFCC00" height="60">
  <tr> 
    <td bgcolor="#ffd34e" width="150" align="left" valign="middle" background="HeadImg/runline.gif"> 
      <div align="center"><a href="index.asp"><img src="Img/%F2%FE%EC%E5%ED%FC%2072rus.gif" alt="www.72rus.ru" border="0"></a></div>
    </td>
    <td bordercolor="#FFCC00" bgcolor="#ffd34e" background="HeadImg/runline.gif"> 
      <div align="center"><object type="text/x-scriptlet" width=171 height="60" data="searchbox.html">
        </object> </div>
    </td>
    <td bgcolor="#ffd34e" align="right" background="HeadImg/runline.gif" width="468"><a href="http://www.auction.72rus.ru/"><img src="bannes/auc_face_anim.gif" width="468" height="60" alt="--&gt; Сибирский Аукцион --&gt; Интернет торги для сибиряков!" border="0"></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0" background="HeadImg/line.gif">
  <tr bgcolor="#996600" background="HeadImg/line.gif"> 
    <td width="150" background="HeadImg/line.gif"> 
      <p align="left"><b><font color="#FFCC33">- Тюмень -</font></b></p>
    </td>
    <td width="490" bgcolor="#996600" background="HeadImg/line.gif"> 
      <p>
    </td>
    <td width="213" background="HeadImg/line.gif"> 
      <p> 
    </td>
    <td width="150" bgcolor="#996600" background="HeadImg/line.gif"> 
      <p align="right"><a href="index.asp"><b><font color="#FFCC33">Home</font></b></a></p>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0" background="HeadImg/line.gif">
  <tr bgcolor="#996600"> 
    <td width="150" bgcolor="#ffd34e" align="right" valign="top"> 
      <p><img src="HeadImg/arrow2.gif" width="11" height="10" align="middle"> 
        <font face=Verdana size=1><b> <a href="index.html">72RUS.RU</a></b></font> 
      </p>
    </td>
    <td bgcolor="#ffd34e" valign="top"> 
      <p><b><img src="HeadImg/arrow2.gif" width="11" height="10" align="absmiddle"> 
        <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=path%></font></b></p>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 

<%if(ShowForm){%>
<p align="center"><b><font size="3" color="#0000CC">Изменение реквизитов раздела: 
  <font color="#990000"><%=namarea%></font></font></b></p>
 
<form name="form1" method="post" action="edsubjmsg.asp">
  <p align="center">
<input type="hidden" name="subj" value="<%=subj_id%>">
  </p>
  <table width="80%" border="1" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="380" bgcolor="#FFCC33"> 
        <div align="center"> 
          <h1><b>Параметры:</b></h1>
        </div>
      </td>
      <td width="20">&nbsp;</td>
      <td bgcolor="#FFCC33" width="582"> 
        <div align="center"> 
          <h1><b>Значения:</b></h1>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#0099CC" height="31" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Имя раздела:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20" height="31"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#0099CC" width="582" height="31" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?namarea:Request.Form("name")%>" maxlength="100" size="45">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#0099CC" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Правое наименование:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#0099CC" width="582" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="outname" value="<%=isFirst?ritit:Request.Form("outname")%>" maxlength="20">
          (Продам, Ищу, Объявления юношеи, и т.д.)</font></p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#0099CC" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Левое наименование:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#0099CC" width="582" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="inname" value="<%=isFirst?leftit:Request.Form("inname")%>" maxlength="20">
          (Куплю, Требуется, Объявления девушек)</font></p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#0099CC" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Тип объявлений:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#0099CC" width="582" valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="dase" value="0" <%=dase==1?"checked":""%>>
          Комерческое объявление<br>
          </font><font size="2"> 
          <input type="radio" name="dase" value="10" <%=dase==10?"checked":""%>>
          Объявления знакомства<br>
          <input type="radio" name="dase" value="20" <%=dase==20?"checked":""%>>
          Объявления работы<br>
          <input type="radio" name="dase" value="30" <%=dase==30?"checked":""%>>
          Одна колонка (правая)</font></p>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#FF0000">Параметры выделенные красным цветом 
    обязательны к заполнению!</font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="Сохранить">
  </p>
</form>
<%} 
else 
{%>
<center>
  <h2><font color="#3333FF">Раздел ИЗМЕНЕН!</font></h2>
  <p><font face="Arial, Helvetica, sans-serif"><a href="index.asp">На главную 
    страницу</a></font></p>
  <p><font face="Arial, Helvetica, sans-serif"><a href="messages.asp">К объявлениям</a></font></p>
</center>
<%}%>
<div align="center"> 
  <table width="100%" border="0" bgcolor="#CCCC99" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
    <tr bgcolor="#996600" bordercolor="#996600"> 
      <td width="150" height="23" background="HeadImg/line.gif"> 
        <div align="center"> 
          <p align="left"><img src="HeadImg/home.gif" width="14" height="14" align="absbottom"> 
            <a href="index.asp"><b><font color="#FFCC33">Home</font></b></a></p>
        </div>
      </td>
      <td background="HeadImg/line.gif">&nbsp; </td>
      <td height="23" background="HeadImg/line.gif"> 
        <div align="right"> 
          <p> <a href="usrarea.asp"><font color="#FFCC33"><b> </b></font></a><a href="admarea.asp"><font color="#FFCC33"><b><font size="1">Вход 
            для Администратора</font> </b></font></a></p>
        </div>
      </td>
    </tr>
  </table>
  <table width="100%" cellspacing="0" cellpadding="0" border="1" bordercolor="#ffd34e">
    <tr bgcolor="#EBF5ED" bordercolor="#EBF5ED"> 
      <td bordercolor="#EBF5ED" width="150" valign="middle" align="center"> 
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
      </td>
      <td bordercolor="#EBF5ED" bgcolor="#EBF5ED"> 
        <div align="center"> 
          <script language="JavaScript" src="http://vbn.tyumen.ru/cgi-bin/hints.cgi?vbn&auto72">
</script>
        </div>
      </td>
      <td width="150" valign="middle" align="center"> 
        <div align="left"> 
          <!--Begin of HMAO RATINGS-->
          <center>
            <a href="http://www.isurgut.ru/Spravka/ResHMAO/stat.asp"> <img src="http://www.isurgut.ru/spravka/top100hmao/StatCounter1.gif" border="0" width="88" height="31"></a> 
            <img src="http://www.isurgut.ru/spravka/top100hmao/counter.asp?Resource_id=1119" border="0" height="1" width="1" > 
          </center>
          <!--End of HMAO RATINGS-->
        </div>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF" align="center">
    <tr bordercolor="#FFFFFF" align="center" bgcolor="#3399FF"> 
      <td valign="middle" bgcolor="#996600" height="13"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Информационный 
          портал 72RUS - Тюменская Область </b></font><font color="#FFFFFF" size="1"><b>- 
          Программирование и дизайн</b></font><b><font size="1"> <a href="http://www.rusintel.ru/" target="_blank"><font color="#FFFFFF">ЗАО 
          Русинтел</font></a> <font color="#FFFFFF">&copy; 2002</font></font></b></p>
      </td>
    </tr>
  </table>
</div>
<p align="center">| <a href="terms.html">Условия использования</a> | <a href="adv.html">Реклама 
  на сервере</a> 
  <p align="center"> © 2002 <a href="http://www.rusintel.ru">Rusintel Company</a> 
<p align="center"></p>
</body>
</html>
