<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\records.inc" -->
<!-- #include file="inc\getform.inc" -->
<!-- #include file="inc\creaters.inc" -->
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\path.inc" -->
<!-- #include file="inc\url.inc" -->



<%

// тут запишем код СМИ... Не забыть изменить его в других сайтах!!
var smi_id=1
// +++  smi_id - код СМИ в таблице SMI !!

var hid=0
var hname=""
var url=""
var nid=0
var name=""
var ndat=""
var nadr=""
var per=0
var kvopub=0
var pname=""
var pdat=""
var autor=""
var digest=""
var imgLname=""
var imgname=""
var path=""
var hdd=0
var hadr=""
var nm=""
var filnam=""
var fs= new ActiveXObject("Scripting.FileSystemObject")
var ts=""
var isnews=1
var blokname=""
var tpm=1000
var usok=false
var ishtml=0
var urlname=""
var urlid=0
var urlabout=""
var daterenew=""
var urladr=""
var urlcount=0
var msgcount=0
var sminame=""
var pid=0
var news=""


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

var getData = Server.CreateObject("SOFTWING.ASPtear")
var dirData=""

sql="Select t1.* from url t1, catarea t2 where t1.recl_id=1 and t1.catarea_id=t2.id and t2.catalog_id="+catalog
Records.Source=sql
Records.Open()
if (!Records.EOF) {
	urlname=String(Records("NAME").Value)
	urlid=Records("ID").Value
	urlabout=String(Records("ABOUT").Value)
	urladr=String(Records("URL").Value)
}
Records.Close()
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

Records.Source="Select * from smi where  id="+smi_id
Records.Open()
sminame=String(Records("NAME").Value)
Records.Close()


var sens=parseInt(Request("sensation"))
if (isNaN(sens)) {sens=0}
var sch=TextFormData(Request("sch"),"")
var wrds=parseInt(Request("wrds"))
if (isNaN(wrds)) {wrds=0}
var pg=0
var lp=15 // Длина страницы
var tps=1 // Поиск в публикациях (2 - объявления, 3 - сайты)
var sstr=""
var ssql=""
var qsql=""
var  tsql=""
var wsql=""
var wsql2=""
var wsql3=""
var ii=0
var name=""
var id=0
var coment=""
var kvo=0
var tkvo=0
var dat=""
var url=""
var ll=0

if (sch!="") {
	// Какой-то запрос
	sstr=sch
	// while (sch.indexOf(".")>=0) {sch=sch.replace(".","")}
	while (sch.indexOf(",")>=0) {sch=sch.replace(","," ")}
	while (sch.indexOf(" - ")>=0) {sch=sch.replace(" - "," ")}
	while (sch.indexOf(" -")>=0) {sch=sch.replace(" -"," ")}
	while (sch.indexOf("- ")>=0) {sch=sch.replace("- "," ")}
	while (sch.indexOf(";")>=0) {sch=sch.replace(";"," ")}
	while (sch.indexOf(":")>=0) {sch=sch.replace(":"," ")}
	while (sch.indexOf("\"")>=0) {sch=sch.replace("\"","")}
	while (sch.indexOf("'")>=0) {sch=sch.replace("'","")}
	while (sch.indexOf("(")>=0) {sch=sch.replace("(","")}
	while (sch.indexOf(")")>=0) {sch=sch.replace(")","")}
	while (sch.indexOf("<")>=0) {sch=sch.replace("<"," ")}
	while (sch.indexOf(">")>=0) {sch=sch.replace(">"," ")}
	while (sch.indexOf("?")>=0) {sch=sch.replace("?"," ")}
	while (sch.indexOf("=")>=0) {sch=sch.replace("="," ")}
	while (sch.indexOf(" % ")>=0) {sch=sch.replace(" % "," ")}
	while (sch.indexOf(" _ ")>=0) {sch=sch.replace(" _ "," ")}
	while (sch.indexOf("+")>=0) {sch=sch.replace("+"," ")}
	while (sch.indexOf("  ")>=0) {sch=sch.replace("  "," ")}
	ll=sch.length
	if (ll<3) {sch=""}
	if (sch.indexOf(" ")==0) {sch=sch.substring(1,ll)}
	ll=sch.length
	if (ll<3) {sch=""}
	if (sch.indexOf(" ")==ll-1) {sch=sch.substring(0,ll-1)}
	if (wrds==0) {sch=""}
	//
	// Публикации
	if (sch!="") {
		ssql=sch
		while (ssql.indexOf(" ")>=0) {ssql=ssql.replace(" ","+")}
		if (sens==1) {
			if (wrds==1) { // С учетом регистра букв ВСЕ слова
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%' and t1.name like  '%")}
				ssql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.name like '%"+ssql+"%' "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%' and t1.digest like  '%")}
				qsql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.digest like '%"+qsql+"%'"
			}
			if (wrds==2) { // С учетом регистра букв Хотябы одно слово
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%' or t1.name like  '%")}
				ssql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (t1.name like '%"+ssql+"%') "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%' or t1.digest like  '%")}
				qsql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (t1.digest like '%"+qsql+"%')"
			}
			if (wrds==3) { // С учетом регистра букв Фраза целиком
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+"," ")}
				ssql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.name like '%"+ssql+"%' "
				qsql=sch
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+"," ")}
				qsql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and t1.digest like '%"+qsql+"%'"
			}
		} else {
			if (wrds==1) { // Без учета регистра букв ВСЕ слова
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%') and UpCase(t1.name) like  UpCase('%")}
				ssql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.name) like UpCase('%"+ssql+"%') "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%') and UpCase(t1.digest) like  UpCase('%")}
				qsql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.digest) like UpCase('%"+qsql+"%')"
			}
			if (wrds==2) { // Без учета регистра букв Хотябы одно слово
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+","%') or UpCase(t1.name) like  UpCase('%")}
				ssql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (UpCase(t1.name) like UpCase('%"+ssql+"%')) "
				qsql=sch
				while (qsql.indexOf(" ")>=0) {qsql=qsql.replace(" ","+")}
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+","%') or UpCase(t1.digest) like  UpCase('%")}
				qsql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and (UpCase(t1.digest) like UpCase('%"+qsql+"%'))"
			}
			if (wrds==3) { // Без учета регистра букв Фраза целиком
				while (ssql.indexOf("+")>=0) {ssql=ssql.replace("+"," ")}
				ssql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.name) like UpCase('%"+ssql+"%') "
				qsql=sch
				while (qsql.indexOf("+")>=0) {qsql=qsql.replace("+"," ")}
				qsql="Select distinct t1.*, t2.name as hnam, t2.url as hurl from publication t1, heading t2 where t1.heading_id=t2.id and t2.smi_id="+smi_id+" and t1.state=1 and UpCase(t1.digest) like UpCase('%"+qsql+"%')"
			}
		}
		ssql=ssql+" UNION "+qsql+" order by 2"
	}
	
// объявления	
	if (sch!="") {
		tsql=sch
		while (tsql.indexOf(" ")>=0) {tsql=tsql.replace(" ","+")}
		if (sens==1) {
			if (wrds==1) { // С учетом регистра букв ВСЕ слова
				while (tsql.indexOf("+")>=0) {tsql=tsql.replace("+","%' and t1.name like  '%")}
				tsql="Select distinct t1.*, t2.name as snam, t2.id as sid, t2.out_name, t2.in_name from trademsg t1, trade_subj t2 where t1.trade_subj_id=t2.id and t2.marketplace_id="+market+" and t1.state=0 and t1.DATE_END>='TODAY' and t1.name like '%"+tsql+"%' "
			}
			if (wrds==2) { // С учетом регистра букв Хотябы одно слово
				while (tsql.indexOf("+")>=0) {tsql=tsql.replace("+","%' or t1.name like  '%")}
				tsql="Select distinct t1.*, t2.name as snam, t2.id as sid, t2.out_name, t2.in_name from trademsg t1, trade_subj t2 where t1.trade_subj_id=t2.id and t2.marketplace_id="+market+" and t1.state=0 and t1.DATE_END>='TODAY' and (t1.name like '%"+tsql+"%') "
			}
			if (wrds==3) { // С учетом регистра букв Фраза целиком
				while (tsql.indexOf("+")>=0) {tsql=tsql.replace("+"," ")}
				tsql="Select distinct t1.*, t2.name as snam, t2.id as sid, t2.out_name, t2.in_name from trademsg t1, trade_subj t2 where t1.trade_subj_id=t2.id and t2.marketplace_id="+market+" and t1.state=0 and t1.DATE_END>='TODAY' and t1.name like '%"+tsql+"%' "
			}
		} else {
			if (wrds==1) { // Без учета регистра букв ВСЕ слова
				while (ssql.indexOf("+")>=0) {tsql=tsql.replace("+","%') and UpCase(t1.name) like  UpCase('%")}
				tsql="Select distinct t1.*, t2.name as snam, t2.id as sid, t2.out_name, t2.in_name from trademsg t1, trade_subj t2 where t1.trade_subj_id=t2.id and t2.marketplace_id="+market+" and t1.state=0 and t1.DATE_END>='TODAY' and UpCase(t1.name) like UpCase('%"+tsql+"%') "
			}
			if (wrds==2) { // Без учета регистра букв Хотябы одно слово
				while (tsql.indexOf("+")>=0) {tsql=tsql.replace("+","%') or UpCase(t1.name) like  UpCase('%")}
				tsql="Select distinct t1.*, t2.name as snam, t2.id as sid, t2.out_name, t2.in_name from trademsg t1, trade_subj t2 where t1.trade_subj_id=t2.id and t2.marketplace_id="+market+" and t1.state=0 and t1.DATE_END>='TODAY' and (UpCase(t1.name) like UpCase('%"+tsql+"%')) "
			}
			if (wrds==3) { // Без учета регистра букв Фраза целиком
				while (tsql.indexOf("+")>=0) {tsql=tsql.replace("+"," ")}
				tsql="Select distinct t1.*, t2.name as snam, t2.id as sid, t2.out_name, t2.in_name from trademsg t1, trade_subj t2 where t1.trade_subj_id=t2.id and t2.marketplace_id="+market+" and t1.state=0 and t1.DATE_END>='TODAY' and UpCase(t1.name) like UpCase('%"+tsql+"%') "
			}
		}
		tsql=tsql+" order by 2"
	}

	// Сайты
	if (sch!="") {
		wsql=sch
		while (wsql.indexOf(" ")>=0) {wsql=wsql.replace(" ","+")}
		if (sens==1) {
			if (wrds==1) { // С учетом регистра букв ВСЕ слова
				while (wsql.indexOf("+")>=0) {wsql=wsql.replace("+","%' and t1.name like  '%")}
				wsql="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and t1.name like '%"+wsql+"%' "
				wsql2=sch
				while (wsql2.indexOf(" ")>=0) {wsql2=wsql2.replace(" ","+")}
				while (wsql2.indexOf("+")>=0) {wsql2=wsql2.replace("+","%' and t1.about like  '%")}
				wsql2="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and t1.about like '%"+wsql2+"%' "
				wsql3=sch
				while (wsql3.indexOf(" ")>=0) {wsql3=wsql3.replace(" ","+")}
				while (wsql3.indexOf("+")>=0) {wsql3=wsql3.replace("+","%' and t1.url like  '%")}
				wsql3="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and t1.url like '%"+wsql3+"%' "
			}
			if (wrds==2) { // С учетом регистра букв Хотябы одно слово
				while (wsql.indexOf("+")>=0) {wsql=wsql.replace("+","%' or t1.name like  '%")}
				wsql="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and (t1.name like '%"+wsql+"%') "
				wsql2=sch
				while (wsql2.indexOf(" ")>=0) {wsql2=wsql2.replace(" ","+")}
				while (wsql2.indexOf("+")>=0) {wsql2=wsql2.replace("+","%' or t1.about like  '%")}
				wsql2="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and (t1.about like '%"+wsql2+"%') "
				wsql3=sch
				while (wsql3.indexOf(" ")>=0) {wsql3=wsql3.replace(" ","+")}
				while (wsql3.indexOf("+")>=0) {wsql3=wsql3.replace("+","%' or t1.url like  '%")}
				wsql3="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and (t1.url like '%"+wsql3+"%') "
			}
			if (wrds==3) { // С учетом регистра букв Фраза целиком
				while (wsql.indexOf(" ")>=0) {wsql2=wsql2.replace("+"," ")}
				wsql="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and t1.name like '%"+wsql+"%' "
				wsql2=sch
				while (wsql2.indexOf(" ")>=0) {wsql2=wsql2.replace("+"," ")}
				wsql2="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and t1.about like '%"+wsql2+"%' "
				wsql3=sch
				while (wsql3.indexOf(" ")>=0) {wsql3=wsql3.replace("+"," ")}
				wsql3="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and t1.url like '%"+wsql3+"%' "
			}
		} else {
			if (wrds==1) { // Без учета регистра букв ВСЕ слова
				while (wsql.indexOf("+")>=0) {wsql=wsql.replace("+","%') and UpCase(t1.name) like  UpCase('%")}
				wsql="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and UpCase(t1.name) like UpCase('%"+wsql+"%') "
				wsql2=sch
				while (wsql2.indexOf(" ")>=0) {wsql2=wsql2.replace(" ","+")}
				while (wsql2.indexOf("+")>=0) {wsql2=wsql2.replace("+","%') and UpCase(t1.about) like  UpCase('%")}
				wsql2="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and UpCase(t1.about) like UpCase( '%"+wsql2+"%' )"
				wsql3=sch
				while (wsql3.indexOf(" ")>=0) {wsql3=wsql3.replace(" ","+")}
				while (wsql3.indexOf("+")>=0) {wsql3=wsql3.replace("+","%') and UpCase(t1.url) like UpCase('%")}
				wsql3="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and UpCase(t1.url) like UpCase('%"+wsql3+"%') "
		}
			if (wrds==2) { // Без учета регистра букв Хотябы одно слово
				while (wsql.indexOf("+")>=0) {wsql=wsql.replace("+","%') or UpCase(t1.name) like  UpCase('%")}
				wsql="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and (UpCase(t1.name) like UpCase('%"+wsql+"%')) "
				wsql2=sch
				while (wsql2.indexOf(" ")>=0) {wsql2=wsql2.replace(" ","+")}
				while (wsql2.indexOf("+")>=0) {wsql2=wsql2.replace("+","%') or UpCase(t1.about) like  UpCase('%")}
				wsql2="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and (UpCase(t1.about) like UpCase( '%"+wsql2+"%' ))"
				wsql3=sch
				while (wsql3.indexOf(" ")>=0) {wsql3=wsql3.replace(" ","+")}
				while (wsql3.indexOf("+")>=0) {wsql3=wsql3.replace("+","%') or UpCase(t1.url) like UpCase('%")}
				wsql3="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and (UpCase(t1.url) like UpCase('%"+wsql3+"%')) "
			}
			if (wrds==3) { // Без учета регистра букв Фраза целиком
				while (wsql.indexOf(" ")>=0) {wsql2=wsql2.replace("+"," ")}
				wsql="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and UpCase(t1.name) like UpCase('%"+wsql+"%') "
				wsql2=sch
				while (wsql2.indexOf(" ")>=0) {wsql2=wsql2.replace("+"," ")}
				wsql2="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and UpCase(t1.about) like UpCase('%"+wsql2+"%') "
				wsql3=sch
				while (wsql3.indexOf(" ")>=0) {wsql3=wsql3.replace("+"," ")}
				wsql3="Select distinct t1.*, t2.name as cnam, t2.id as cid from url t1, catarea t2 where t1.catarea_id=t2.id and t2.catalog_id="+catalog+" and t1.state=1 and UpCase(t1.url) like UpCase('%"+wsql3+"%') "
			}
		}
		wsql=wsql+" UNION "+wsql2+" UNION "+wsql3+" order by 2"
	}
	
}

%>
<html>
<head>
<title><%=sch%> :: 72rus.ru - Тюмень, Тюменский регион</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="style1.css" type="text/css">
<style><!--p {  font-family: Arial, Helvetica, sans-serif; font-size: 9pt; line-height: 12pt; font-weight: 400; margin:  3px 3px 3px 4px}
h1 {color: #0000CC; font-family: Arial, Helvetica, sans-serif; font-size: 16px; line-height: 17px; margin-top: 3px; margin-right: 3px; margin-bottom: 3px; margin-left: 5px}
h2 { font-family: Arial, Helvetica, sans-serif; font-size: 7pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.text { font: 10px Arial, Helvetica, sans-serif; color: #003300;}.digest { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
.bar { color: #FFCC00}.digesta { font-family: Arial, Helvetica, sans-serif; font-size: 8.5pt; line-height: 10pt; font-weight: 400; margin: 3px 3px 3px 4px }
--></style>
<SCRIPT language=JavaScript>


DOM=(document.getElementById)?1:0;
NS4=(document.layers)?1:0;
IE4=(document.all)?1:0;

function CHange_city(nomer){
        if(NS4){document.images['weather'].src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
        if(IE4){document.all['weather'].src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
        if(DOM){getElementsByName('weather').src="http://img.gismeteo.ru/informer/"+nomer+"-6.GIF";return;}
}

</SCRIPT>
<script language="JavaScript">
 function goToUrl(){
  if(NS4)location.href="http://www.gismeteo.ru/weather/towns/"+document.select["City_pogoda"].value+".htm";
  if(IE4)location.href="http://www.gismeteo.ru/weather/towns/"+document.all["City_pogoda"].value+".htm";
  if(DOM)location.href="http://www.gismeteo.ru/weather/towns/"+getElementById("City_pogoda").value+".htm";
 }
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellspacing="1" width="100%" cellpadding="0">
  <tr> 
    <td> 
      <p class="menu01"> <font color="#333333"> 
        <!--LiveInternet counter-->
        <script language="JavaScript">document.write('<img src="http://counter.yadro.ru/hit?r' + escape(document.referrer) + ((typeof(screen)=='undefined')?'':';s'+screen.width+'*'+screen.height+'*'+(screen.colorDepth?screen.colorDepth:screen.pixelDepth)) + ';' + Math.random() + '" width=1 height=1 alt="">')</script>
        <!--/LiveInternet-->
        <%=sminame%> : Поиск на сайте</font></p>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td background="images/fon02.gif" height="87" align="center" width="170"> 
      <a href="/"><img src="images/72rus.gif" width="170" height="87" alt="72RUS.RU Тюменский Регион - информационный портал " border="0"></a> 
    </td>
    <td background="images/fon02.gif" height="87" align="center"> 
      <script language="javascript" src="banshow.asp?rid=4"></script>
    </td>
    <td background="images/fon02.gif" height="87" align="center" width="170">
<table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="164" bgcolor="#F3F3F3" bordercolor="#E5E5E5"> 
            <p class="digest"> <a href="admarea.asp"><img src="images/e06.gif" width="16" height="9" alt="" border="0"></a> 
              <a href="#" class="digest">Посетителей сейчас:</a> <%=Application("visitors")%><br>
              <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="http://www.72rus.ru/newshow.asp?pid=728" class="digest">Реклама 
              на 72RUS.RU</a><br>
              <img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
              <a href="#" onClick="window.external.AddFavorite(parent.location,document.title)" class="digest">Добавить 
              в избранное</a><br>
			  <img src="images/e06.gif" width="16" height="9" alt="" border="0">
			  <a href="searchall.asp" class="digest">Поиск на сайте</a>
            </p>
</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#FF6600"> 
    <td colspan="4" height="1"></td>
  </tr>
</table>
<table width="173" border="0" cellspacing="0" cellpadding="0" align="left">
  <tr> 
    <td width="172" align="left" valign="top"> 
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
	kvopub=0
	if (name!="") {
		recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
		recs.Open()
		kvopub=recs("COUNT_PUB").Value
		recs.Close()
	}
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%> 
              [<%=kvopub%>]</a></p>
          </td>
        </tr>
      </table>
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
	kvopub=0
	recs.Source="Select count_pub  from get_count_pub_show("+hid+")"
	recs.Open()
	kvopub=recs("COUNT_PUB").Value
	recs.Close()
	Records.MoveNext()
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="<%=url%>"><%=hname%></a></p>
          </td>
        </tr>
      </table>
      <%
} Records.Close()
delete recs
%>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="messages.asp">Объявления [<%=msgcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="lentamsg.asp">Объявления 
              [новые]</a> </p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" width="170" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="catarea.asp">WEB 
              Каталог [<%=urlcount%>]</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auto.72rus.ru">Авто 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="http://auction.72rus.ru">Аукцион 
              72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="air_russia.asp">АВИА 
              Расписание</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a href="Rail_roads.asp">Расписание 
              поездов</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="170">
        <tr> 
          <td background="images/fon_menu01.gif" height="30" bgcolor="#6699CC"> 
            <p class="menu01"><img src="images/arrow.gif" align="absmiddle"> <a 
            href="http://bn.72rus.ru">Баннерообмен 72rus</a></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td valign="top"><img src="images/top01.gif" width="170" height="19" alt="" border="0"> 
            <img src="images/fon_menu04.gif" width="172" height="1"></td>
        </tr>
      </table>
      <font color="#000000"> </font> 
      <table width="100%" border="0" cellspacing="0" height="60" align="center" cellpadding="0">
        <tr> 
          <td align="center"> 
            <%
// В переменной bk содержится код блока новостей
var bk=38
// Не забывать его менять!!
Records.Source="Select t1.*, t2.posit from publication t1, news_pos t2 where t1.state=1 and t1.id=t2.publication_id and t2.block_news_id="+bk+" order by t2.posit"
Records.Open()
while (!Records.EOF )
{
	puid=String(Records("ID").Value)
	ishtml2=TextFormData(Records("ISHTML").Value,"0")
	filnam=PubFilePath+puid+".pub"
	if (!fs.FileExists(filnam)) { filnam="" }

	if (filnam != "") {
		ts= fs.OpenTextFile(filnam)
		if (ishtml2==0){
			while (!ts.AtEndOfStream){
				news_bl+="<p style='text-align:justify'>"+ts.ReadLine()+"</p>"
			}
		} else {news_bl=ts.ReadAll()}
		ts.Close()
	}

%>
            <%=news_bl%> 
            <%
Records.MoveNext()
} 
Records.Close()
%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="120" border="0" cellspacing="0" cellpadding="0" align="right" height="750">
  <tr> 
    <td valign="top" width="1" rowspan="2"></td>
    <td width="1" valign="top" bgcolor="#003366" rowspan="2"></td>
    <td width="150" valign="top" align="center"> 
      <table border="0" cellpadding="0" cellspacing="0" width="120">
        <tr> 
          <td valign="top" background="images/top01.gif" align="center" bgcolor="#FF9900"> 
            <p class="menu01"> 
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
              <%=blokname%></p>
          </td>
        </tr>
      </table>
      <table border="0" cellpadding="0" cellspacing="0" width="156">
        <tr> 
          <td align="center" height="10"></td>
        </tr>
      </table>
      <p> 
        <script language="javascript" src="banshow.asp?rid=5"></script>
    </td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" height="900" align="center">
  <tr> 
    <td valign="top" width="1"></td>
    <td valign="top" align="center"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bordercolor="#CCCCFF" align="center">
        <tr bordercolor="#CCCCFF" valign="middle" align="center"> 
          <td bordercolor="#CCCCFF" height="11"> 
            <table width="100%" border="0" cellspacing="0" bordercolor="#003366">
              <tr bgcolor="#FBF8D7"> 
                <td height="35" bgcolor="#FFFFFF" bordercolor="#FFFFFF" valign="middle"> 
                  <p class="menu02"><img src="images/e06.gif" width="16" height="9" alt="" border="0"> 
                    <a href="index.asp">72RUS.RU</a> / Поиск</p>
                </td>
              </tr>
            </table>
            <table width="97%" border="0" cellspacing="0" cellpadding="0" height="23" align="center">
              <tr> 
                <td bgcolor="#FF9900" align="left" background="images/fon_menu08.gif"> 
                  <h1><b><font color="#FFFFFF"> Поиск <%=sch%> </font></b></h1>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td class=n3 valign="top"> </td>
        </tr>
      </table>
      <table width="99%" border="0" bordercolor="#FFFFFF" align="center" cellspacing="0" height="350">
        <tr> 
          <td bordercolor="#CCCCCC" width="778" valign="TOP"> 
            <table width="100%" border="0" cellpadding="0" align="center" cellspacing="0" bgcolor="#E1F4FF">
              <tr> 
                <td valign="top" bgcolor="#FFFFFF" align="center"> 
                  <table border="0" cellspacing="0" cellpadding="0" align="center" width="97%">
                    <form name="form" method="get" action="searchall.asp">
          <tr> 
                        <td align="center" height="45"> 
                          <input type="text" name="sch" size="40" value="<%=sch%>" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 275px" >
              <input type="image" src="images/search.gif" width="55" height="20" alt="Найти на сайте" border="0" hspace="10" align="absbottom" name="Findit">
              <select name="wrds" style="BACKGROUND-COLOR: #FFFFFF; BORDER-BOTTOM: #000000 1px solid; BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; COLOR: #303030; FONT-FAMILY: tahoma; FONT-SIZE: 11px; WIDTH: 100px">
                <option value="1" <%=wrds==1?"selected":""%>>Все слова</option>
                <option value="2" <%=wrds==2?"selected":""%>>Одно из слов</option>
                <option value="3" <%=wrds==3?"selected":""%>>Фраза целиком</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td align="center"> 
              <p class="digest">Везде | <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=2" class="digest">В 
                объявлениях</a> | <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=3" class="digest">В 
                каталоге сайтов</a> | <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=1" class="digest">В 
                публикациях</a> | 
                <input type="checkbox" name="sensation" value="1" <%=sens==1?"checked":""%>>
                Учесть регистр</p>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"></td>
          </tr>
        </form>
      </table>
    
				  <p>&nbsp;</p>

<%
if (wsql!="") {
//Records.CursorType=3
Records.Source=wsql
Records.Open()
if (!Records.EOF) {
kvo=Records.RecordCount
if ((pg+1)*lp > kvo) {tkvo=kvo} else {tkvo=(pg+1)*lp}
%>
				  <p align="left"><b><font color="#EA7500">Найдено сайтов: <%=kvo%></font></b></p>
<%
	ii=0
	sid=0
	snam=""
	while ((!Records.EOF) && (ii<tkvo)) {
		ii+=1
		id=String(Records("ID").Value)
		name=String(Records("NAME").Value)
		coment=String(Records("ABOUT").Value)
		url=String(Records("URL").Value)
		sid=String(Records("CID").Value)
		snam=String(Records("CNAM").Value)
		if (ii>=(pg*lp+1)) {
%>
                  <table width="100%" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#F8F8E9"> 
                      <td width="6%" valign="top"> 
                        <div align="center"> 
                          <p><%=ii%>.</p>
                        </div>
                      </td>
                      <td valign="top"> 
                        <div align="left"> 
                          <p><b></b> <font color="#999999"><a href="<%=url%>" target="_blank"><%=name%></a> [ <%=url%> ]</font></p>
                        </div>
                      </td>
                    </tr>
                    <tr bgcolor="#FBFDFD"> 
                      <td>&nbsp;</td>
                      <td> 
                        <p><%=coment%></p>
                      </td>
					 </tr>
					 <tr>
					  <td>&nbsp;</td>
					  <td> 
                        <p class="digest"><a href="catarea.asp?hid=<%=sid%>" class="digest"><%=snam%></a></p>
                      </td>
					</tr>
                    <tr> 
                      <td colspan="2" height="6"></td>
                    </tr>
                  </table>
<%
		}
		Records.MoveNext()
	}
}
Records.Close()
%>
                  <p align="left"><font size="1">Страницы: 
<%
for ( ii=1; ii<(kvo/lp + 1) ; ii++) {
%>
                    <% if (ii==(pg+1)) { %>
                    <%=ii%> | 
                    <%} else {%>
                    <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=3&pg=<%=ii-1%>"><%=ii%></a> 
                    | 
                    <%}%>
                    <%
}
%>
			</font></p>
            <hr noshade size="1">
<%
}
%>

<%
if (ssql!="") {
//Records.CursorType=3
Records.Source=ssql
Records.Open()
if (!Records.EOF) {
kvo=Records.RecordCount
if ((pg+1)*lp > kvo) {tkvo=kvo} else {tkvo=(pg+1)*lp}
%>
                  <p align="left"><font color="#003399"><b>Найдено публикаций</b>: 
                    <%=kvo%></font></p>
<%
	ii=0
	while ((!Records.EOF) && (ii<tkvo)) {
		ii+=1
		id=String(Records("ID").Value)

		name=String(Records("NAME").Value)
		coment=TextFormData(Records("DIGEST").Value,"")
		dat=Records("PUBLIC_DATE").Value
		url=TextFormData(Records("URL").Value,"newshow.asp")
		url+="?pid="+id
		hhd=Records("HEADING_ID").Value
		hhname=Records("HNAM").Value
		hurl=TextFormData(Records("HURL").Value,"")
		if (hurl=="") {hurl="pubheading.asp"}
		hurl+="?hid="+hhd
		if (ii>=(pg*lp+1)) {
%>
                  <table width="100%" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#EBF3F5"> 
                      <td width="6%" bgcolor="#EBF3F5" valign="top"> 
                        <div align="center"> 
                          <p><%=ii%>.</p>
                        </div>
                      </td>
                      <td valign="top"> 
                        <div align="left"> 
                          <p><b><a href="<%=url%>" target="_blank"><%=name%></a></b> 
                            <font color="#999999">[ <%=dat%> ]</font></p>
                        </div>
                      </td>
                    </tr>
                    <tr bgcolor="#FBFDFD">
                      <td></td>
                      <td><p><font size="2"><%=coment%></font></p></td>
                    </tr>
                    <tr bgcolor="#FBFDFD"> 
                      <td> </td>
                      <td>
                        <p><a href="<%=hurl%>" class="digest"><%=hhname%></a></p>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="2" height="6"></td>
                    </tr>
                  </table>
<%
		}
		Records.MoveNext()
	}
}
Records.Close()
%>
                  <p align="left"><font size="1">Страницы: 
<%
for ( ii=1; ii<(kvo/lp + 1) ; ii++) {
%>
                    <% if (ii==(pg+1)) { %>
                    <%=ii%> | 
                    <%} else {%>
                    <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=1&pg=<%=ii-1%>"><%=ii%></a> 
                    | 
                    <%}%>
                    <%
}
%>
                    </font></p>
                  <hr noshade size="1">
<%
}
%>
 
<%
if (tsql!="") {
Records.Source=tsql
Records.Open()
if (!Records.EOF) {
kvo=Records.RecordCount
if ((pg+1)*lp > kvo) {tkvo=kvo} else {tkvo=(pg+1)*lp}
%>               
                  <p align="left"><font color="#003399"><b><font color="#006633">Найдено 
                    объявлений</font></b><font color="#006633">: <%=kvo%></font></font></p>
<%
	ii=0
	snam=""
	onm=""
	cinam=""
	sid=0
	var recs=CreateRecordSet()
	while ((!Records.EOF) && (ii<tkvo)) {
		ii+=1
		id=String(Records("ID").Value)
		name=String(Records("NAME").Value)
		sid=String(Records("SID").Value)
		snam=String(Records("SNAM").Value)
		if (Records("IS_FOR_SALE").Value==0) {onm=Records("IN_NAME").Value} else {onm=String(Records("OUT_NAME").Value)}
		recs.Source="Select * from city where id="+Records("CITY_ID").Value
		recs.Open()
		if (!recs.EOF) {
			cinam=String(recs("NAME").Value)
		}
		recs.Close()
		delete recs
		if (ii>=(pg*lp+1)) {
%>
                  <table width="100%" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#ECF4ED"> 
                      <td width="6%" valign="top"> 
                        <div align="center"> 
                          <p><%=ii%>.</p>
                        </div>
                      </td>
                      <td valign="top"> 
                        <div align="left"> 
                          <p><b></b> <font color="#999999"><a href="msg.asp?ms=<%=id%>" target="_blank"><%=name%></a> [ <%=cinam%> ]</font></p>
                        </div>
                      </td>
                    </tr>
                    <tr bgcolor="#FBFDFD"> 
                      <td> </td>
                      <td> 
                        <p class="digest"><a href="messages.asp?subj=<%=sid%>" class="digest"><%=snam%></a> 
						<img src="HeadImg/arr_gr.gif" width="5" height="5" align="middle"> <%=onm%></p>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="2" height="6"></td>
                    </tr>
                  </table>
 <%
		}
		Records.MoveNext()
	}
}
Records.Close()
%>                
                  <p align="left"><font size="1">Страницы: 
<%
for ( ii=1; ii<(kvo/lp + 1) ; ii++) {
%>
                    <% if (ii==(pg+1)) { %>
                    <%=ii%> | 
                    <%} else {%>
                    <a href="search.asp?sch=<%=sch%>&wrds=<%=wrds%>&sensation=<%=sens%>&tps=2&pg=<%=ii-1%>"><%=ii%></a> 
                    | 
                    <%}%>
                    <%
}
%>
				  </font></p>
            <hr noshade size="1">
<%
}
%>
			
				</td>
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr> 
    <td height="1" bgcolor="#666666"></td>
  </tr>
  <tr> 
    <td height="19" bgcolor="#AFC0D0" align="center">&nbsp; </td>
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
      <p class="menu02">| <font face="Arial, Helvetica, sans-serif"> 
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
        </font><a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <font face="Arial, Helvetica, sans-serif"> 
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
        </font><a href="<%=url%>"><%=hname%></a> | 
        <%
} Records.Close()
delete recs
%>
        <br>
        | <a href="http://auto.72rus.ru">Авто.72rus</a> | <a href="http://www.auction.72rus.ru/">Аукцион</a> 
        | <a href="messages.asp">Объявления</a> | <a href="Rail_roads.asp">Расписание</a> 
        | <a href="catarea.asp">Тюменский Каталог</a> | 
    </td>
    <td width="180"> 
      <p>Сделано в <a href="http://www.rusintel.ru">ЗАО Русинтел</a> <br>
        WWW.72RUS.RU © 2002-2004 </p>
    </td>
  </tr>
</table>
<hr size="1">
<div align="center"> 
  <script language="javascript" src="banshow.asp?rid=6"></script>
</div>
<hr size="1">
<div align="center"> <a href="http://www.rax.ru/click" target=_blank> 
  <%
// В переменной bk содержится код блока новостей
var bk=33
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
  </a><a href="http://www.rax.ru/click" target=_blank><%=news%></a> 
  <%
Records.MoveNext()
} 
Records.Close()
delete recs
%>
  <!--LiveInternet logo-->
  <a href="http://www.liveinternet.ru/click" target=liveinternet><img src="http://counter.yadro.ru/logo?16.1" border=0 width=88 height=31 alt="liveinternet.ru: показано число хитов за 24 часа, посетителей за 24 часа и за сегодня"></a> 
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
</body>
</html>
