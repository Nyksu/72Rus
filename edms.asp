<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\next_id.inc" -->
<!-- #include file="inc\Creaters.inc" -->
<!-- #include file="inc\path.inc" -->

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

var locks=String(Math.random()).substr(3,6)
if(locks.length<6){locks+=String(Math.random()).substr(3,(6-locks.length))}
var cw=String(Math.random()).substr(3,20)
if(cw.length<20){cw+=String(Math.random()).substr(3,(20-cw.length))}

if (isNaN(subj_id)) {Response.Redirect("messages.asp")}
if (subj_id==0) {Response.Redirect("messages.asp")}

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
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> <img src=\"Img/arrow2.gif\" width=\"11\" height=\"10\" align=\"middle\"> "+path
	sbj=Records("HI_ID").Value
	tit=nm+" : "+tit
	Records.Close()
}

Records.Source="Select * from trade_subj where hi_id="+subj_id+" and marketplace_id="+market
Records.Open()
if (Records.EOF) {canaddmsg=1}
Records.Close()

path="<a href=\"messages.asp\">Объявления</a>  <img src=\"Img/arrow2.gif\" width=\"11\" height=\"10\" align=\"middle\"> "+path

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
	 
	 
	 if (name.length<3) {ErrorMsg+="Слишком короткое наименование.<br>"}
	 if (nikname.length==0) {ErrorMsg+="Пустое имя или псевдоним.<br>"}
	 if (city.length<3) {ErrorMsg+="Слишком короткое наименование города.<br>"}
	 if (about.length<15) {ErrorMsg+="Слишком короткое объявление.<br>"}
	 if (about.length>1000) {ErrorMsg+="Слишком длинное объявление.<br>"}
	 if (sw!=Session("lcode")) {ErrorMsg+="Неправильно введен код.<br>"}
	 if (dase>365) {dase=365}
	 if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && dase>30 && tip!=10) {dase=30}
	 if ((tip>=10 && tip<20) && (sex!=0 && sex!=1)) {ErrorMsg+="Не указан ваш пол.<br>"}
	 if (email.length>0) {	 
	     if (!/(\w+)@((\w+).)*(\w+)$/.test(email)) {ErrorMsg=ErrorMsg+"Неверный формат поля 'E-mail'.<br>"}}
	
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
				ErrorMsg+="Ошибка вставки.<br>"
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
<title>Размещение объявления в: (<%=tit%>)</title>
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
<p align="center"><b><font size="3" color="#0000CC">Вы добавляете объявление в 
  раздел: <font color="#990000"><%=namarea%></font></font></b></p>
 
<form name="form1" method="post" action="addms.asp">
  <p align="center">Для добавления объявления, пожалуйста, заполните поля формы: 
    <input type="hidden" name="subj" value="<%=subj_id%>">
    <%if (tip>=30 && tip<40) {%>
    <input type="hidden" name="sale" value="1">
    <input type="hidden" name="price" value="">
    <%}%>
  </p>
  <table width="90%" border="0" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="38%" height="21" bgcolor="#ffd34e" bordercolor="#FFCC33"> 
        <div align="center"> 
          <h1><b>Параметры:</b></h1>
        </div>
      </td>
      <td width="2%" height="21">&nbsp;</td>
      <td height="21" width="60%" bgcolor="#ffd34e" bordercolor="#FFCC33"> 
        <div align="center"> 
          <h1><b>Значения:</b></h1>
        </div>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" height="2" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Заголовок объявления:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%" height="2"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" height="2" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="100" size="45">
        </p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" height="11" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Ваше имя (псевдоним) 
            <%if (tip<10 || tip>=20) {%>
            или предприятие 
            <%}%>
            :</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%" height="11"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" height="11" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="nikname" value="<%=isFirst?"":Request.Form("nikname")%>" maxlength="50">
          </font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Ваш телефон:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="phone" value="<%=isFirst?"":Request.Form("phone")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Ваш E-mail:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="email" value="<%=isFirst?"":Request.Form("email")%>" maxlength="80">
          </font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Срок размещения объявления:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="dase" value="1" <%=isFirst?"checked":""%> <%=Request.Form("dase")==1?"checked":""%>>
          1 день 
          <input type="radio" name="dase" value="2" <%=Request.Form("dase")==2?"checked":""%>>
          2 дня 
          <input type="radio" name="dase" value="7" <%=Request.Form("dase")==7?"checked":""%>>
          7 дней 
          <input type="radio" name="dase" value="10" <%=Request.Form("dase")==10?"checked":""%>>
          10 дней 
          <input type="radio" name="dase" value="20" <%=Request.Form("dase")==20?"checked":""%>>
          20 дней 
          <input type="radio" name="dase" value="30" <%=Request.Form("dase")==30?"checked":""%>>
          30 дней 
          <%if (Session("is_adm_mem")==1 || Session("cataloghost")==catalog || tip==10) {%>
          <input type="radio" name="dase" value="365" <%=Request.Form("dase")==365?"checked":""%>>
          год 
          <%}%>
          </font></p>
      </td>
    </tr>
    <% if (tip<10) { %>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Продажа или покупка:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="sale" value="1" <%=isFirst?"checked":""%> <%=Request.Form("sale")==1?"checked":""%>>
          Продажа 
          <input type="radio" name="sale" value="0" <%=Request.Form("sale")==0?"checked":""%>>
          Покупка</font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Цена:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="price" value="<%=isFirst?"":Request.Form("price")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <%} 
if (tip>=10 && tip<20) {%>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Вы мужчина или женщина:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%" height="23"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top" height="23"> 
        <p><font size="2"> 
          <input type="radio" name="sex" value="1" <%=Request.Form("sex")==1?"checked":""%>>
          Я мужчина 
          <input type="radio" name="sex" value="0" <%=Request.Form("sex")==0?"checked":""%>>
          Я женщина</font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Возраст / рост / вес:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="price" value="<%=isFirst?"":Request.Form("price")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <%}%>
    <% if (tip>=20 && tip<30) { %>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Тип объявления:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="sale" value="1" <%=isFirst?"checked":""%> <%=Request.Form("sale")==1?"checked":""%>>
          Ищу 
          <input type="radio" name="sale" value="0" <%=Request.Form("sale")==0?"checked":""%>>
          Требуется</font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle"> 
        <div align="right"> 
          <p><font size="2">Зарплата:&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="price" value="<%=isFirst?"":Request.Form("price")%>" maxlength="40">
          </font></p>
      </td>
    </tr>
    <%} %>
    <tr> 
      <td width="38%" valign="middle" height="2"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Страна:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
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
      <td width="38%" valign="middle" height="2"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Регион:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
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
    <tr bordercolor="#FFFFFF"> 
      <td width="38%" valign="middle" height="2"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Город:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td width="60%" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="city" value="<%=isFirst?"":Request.Form("city")%>" maxlength="100">
          </font></p>
      </td>
    </tr>
    <tr bordercolor="#FFFFFF" valign="top"> 
      <td width="38%"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Текст объявления:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="2%"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td valign="top" width="60%"> 
        <p><font size="2"> 
          <textarea name="txt" cols="40" rows="5"><%=isFirst?"":Request.Form("txt")%></textarea>
          </font></p>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#FF0000">Параметры выделенные красным цветом 
    обязательны к заполнению!</font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="Разместить объявление">
  </p>
</form>
<%} 
else 
{%>
<center>
  <h1><font color="#3333FF">Объявление размещено!</font></h1>
  <h1>&nbsp;</h1>
  <p><font color="#3333FF"><a href="addmsimg.asp?ms=<%=mid%>">Здесь 
    Вы можете добавить фотографию или иллюстрацию к объявлению</a></font></p>
  <p>&nbsp;</p>
  <p><font color="#FF0000">Запишите код доступа для изменения или досрочного удаления 
    вашего объявления:</font></p>
  <p><b><font size="3" color="#FF0000"><%=cw%></font></b></p>
  <hr noshade size="1" width="70%">
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
  на сервере</a> | 
  <%if (Session("is_adm_mem")!=1 && Session("cataloghost")!=catalog && canaddmsg==1) {%>
  <a href="addms.asp?subj=<%=subj_id%>">Добавить 
  объявление</a> 
  <%}%>
<p align="center"> © 2002 <a href="http://www.rusintel.ru">Rusintel Company</a> 
<p align="center"></p>
</body>
</html>
