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
var inname=""
var outname=""
var ErrorMsg=""
var ShowForm=true
var tit="" 
var path=""
var sbj=0
var nm=""
var id=0
var dase=""
var sql=""

if (isNaN(subj_id)) {Response.Redirect("messages.asp")}

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
	}
	path="<a href=\"messages.asp?subj="+sbj+"\">"+nm+"</a> | "+path
	sbj=Records("HI_ID").Value
	tit=nm+" : "+tit
	Records.Close()
}

Records.Source="Select * from trade_subj where hi_id="+subj_id+" and marketplace_id="+market
Records.Open()
if (Records.EOF) {canaddmsg=1}
Records.Close()

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
	  		id=NextID("TRADESUBJID")
			sql="Insert into TRADE_SUBJ (ID, NAME,HI_ID,OUT_NAME,IN_NAME,MARKETPLACE_ID,MSG_TYPE) "
			sql+=" values (%ID,'%NAME',%SUBJ,'%ONAM','%INAM',%MAR,%TP)"
			sql=sql.replace("%ID",id)
			sql=sql.replace("%NAME",name)
			sql=sql.replace("%TP",dase)
			sql=sql.replace("%MAR",market)
			sql=sql.replace("%INAM",inname)
			sql=sql.replace("%SUBJ",subj_id)
			sql=sql.replace("%ONAM",outname)
			Connect.BeginTrans()
			try{
			Connect.Execute(sql)
			}
			catch(e){
				Connect.RollbackTrans()
				ErrorMsg+=ListAdoErrors()
				ErrorMsg+="Ошибка вставки.<br>"
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
<title>Добавление раздела в: (<%=tit%>)</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<link rel="stylesheet" href="style.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr> 
    <td bgcolor="#F2F2F2" bordercolor="#333333"> 
      <p><a href="/"><b>72RUS.RU</b></a> | <a href="admarea.asp">Кабинет</a> |</p>
    </td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<table width="100%" border="0" cellspacing="0" align="center" cellpadding="0">
  <tr> 
    <td width="150" align="right" valign="top"> 
      <p>&nbsp;</p>
    </td>
    <td valign="top"> 
      <p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=path%></font></b></p>
    </td>
  </tr>
</table>
<%if(ErrorMsg!=""){%>
<center>
<p> <font color="#FF3300" size="2"><b>Ошибка!</b></font> <br><%=ErrorMsg%></p>
</center>
<%}%> 

<%if(ShowForm){%>
<p align="center"><b><font size="3" color="#0000CC">Вы добавляете новый раздел 
  в : <font color="#990000"><%=namarea%></font></font></b></p>
 
<form name="form1" method="post" action="addsubjms.asp">
  <p align="center">Для добавления раздела, пожалуйста, заполните поля формы: 
    <input type="hidden" name="subj" value="<%=subj_id%>">
  </p>
  <table width="80%" border="1" bordercolor="#FFFFFF" align="center">
    <tr> 
      <td width="380" bgcolor="#F2F2F2" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>Параметры:</b></p>
        </div>
      </td>
      <td width="20">&nbsp;</td>
      <td bgcolor="#F2F2F2" width="582" bordercolor="#333333"> 
        <div align="center"> 
          <p><b>Значения:</b></p>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#333333" height="31" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Имя раздела:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20" height="31"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" width="582" height="31" valign="top"> 
        <p> 
          <input type="text" name="name" value="<%=isFirst?"":Request.Form("name")%>" maxlength="100" size="45">
        </p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Правое наименование:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" width="582" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="outname" value="<%=isFirst?"":Request.Form("outname")%>" maxlength="20">
          (Продам, Ищу, Объявления юношеи, и т.д.)</font></p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Левое наименование:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" width="582" valign="top"> 
        <p><font size="2"> 
          <input type="text" name="inname" value="<%=isFirst?"":Request.Form("inname")%>" maxlength="20">
          (Куплю, Требуется, Объявления девушек)</font></p>
      </td>
    </tr>
    <tr> 
      <td width="380" bordercolor="#333333" valign="middle"> 
        <div align="right"> 
          <p><font size="2" color="#FF0000">Тип объявлений:</font><font size="2">&nbsp;&nbsp;</font></p>
        </div>
      </td>
      <td width="20"> 
        <div align="center"> 
          <p>-</p>
        </div>
      </td>
      <td bordercolor="#333333" width="582" valign="top"> 
        <p><font size="2"> 
          <input type="radio" name="dase" value="0" <%=isFirst?"checked":""%> <%=Request.Form("dase")==1?"checked":""%>>
          Комерческое объявление<br>
          </font><font size="2"> 
          <input type="radio" name="dase" value="10" <%=Request.Form("dase")==10?"checked":""%>>
          Объявления знакомства<br>
          <input type="radio" name="dase" value="20" <%=Request.Form("dase")==20?"checked":""%>>
          Объявления работы<br>
          <input type="radio" name="dase" value="30" <%=Request.Form("dase")==30?"checked":""%>>
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
  <p><b><font color="#3333FF">Раздел добавлен!</font></b></p>
  <p>&nbsp;</p>
  <p><a href="admarea.asp">В кабинет</a> | <font face="Arial, Helvetica, sans-serif"><a href="index.asp"> 
    На главную страницу</a> | </font><font face="Arial, Helvetica, sans-serif"><a href="messages.asp">К 
    объявлениям</a></font></p>
  </center>
<%}%>
<div align="center"> </div>
</body>
</html>
