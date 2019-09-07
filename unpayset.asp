<%@LANGUAGE="JAVASCRIPT"%>
<!-- #include file="inc\err.inc" -->
<!-- #include file="inc\email.inc" -->
<!-- #include file="inc\records.inc" -->

<%
var email=""
var name=""
var isOk=false
var id=parseInt(Request("url"))
var smi_id=1
var nikname=""
var uid=0
var urlname=""
var url=""

if (isNaN(id)) {id=0}
if (id==0) {Response.Redirect("admarea.asp")}

if ((Session("is_adm_mem")!=1)&&(Session("cataloghost")!=catalog)){
Session("backurl")="unpayset.asp?url="+id+"&st="+setst
Response.Redirect("login.asp")
}

sql="Select t1.name, t1.email, t1.nikname, t2.url, t2.name as urnam from HOST_URL t1, url t2 where t2.id="+id+" and t1.id=t2.host_url_id"
Records.Source=sql
Records.Open()
if (Records.EOF){
	Records.Close()
	Response.Redirect("admarea.asp")
}
name=Records("NAME").Value
nikname=Records("NIKNAME").Value
uid=Records("ID").Value
email=Records("EMAIL").Value
url=Records("URL").Value
urlname=Records("URNAME").Value
Records.Close()

if (email != "") {isOk=true} else {Response.Redirect("admarea.asp")}

var eml=Server.CreateObject("JMail.Message")
eml.Logging=false
eml.From="support@72rus.ru"
eml.AddRecipient(email)
eml.Charset=characterset
eml.ContentTransferEncoding = "base64"

var isSending=false

if(!isOk){
	
		//-------------input validation-----------
		try{
				eml.FromName="Служба поддержки портала 72RUS.RU"
				eml.Subject=tit+"Уведомление о вашем заказе на услугу каталога сайтов портала 72RUS.RU"
				eml.AppendText(" Здравствуйте, "+name+"!\n")
				eml.AppendText(" Вами был добавлен ресурс в каталог сайтов нашего портала : "+urlname+"\n")
				eml.AppendText(" с адресом : "+url+"\n")
				eml.AppendText(" \n")
				eml.AppendText("Это письмо отправлено Вам, по причине того, что Ваш сайт был предложен для включения в каталог Ё72RUS.RU - Тюменский регионЁ. \n")
				eml.AppendText("Единовременная стоимость размещения одного интернет ресурса составляет 150 рублей. \n")
				eml.AppendText("Оплата производится либо банковским перечеслением (отправить заявку на выставление счета:  \n")
				eml.AppendText("http://rusintel.ru/message.asp), либо через систему WebMoney перечислением 5.00 WMZ на кошелек  \n")
				eml.AppendText("Z200556237621 в комментарии к платежу укажите адрес и название сайта. \n")
				eml.AppendText(" \n")
				eml.AppendText("Если в течении 20 дней оплата не будет произведена - Ваша заявка будет удалена из очереди. \n")
				eml.AppendText(" \n")
				eml.AppendText("Правила размещения сайтов в каталог 72RUS.RU см: http://www.72rus.ru/addurl.asp \n")
				eml.AppendText(" \n")
				eml.AppendText("Надеемся на дальнейшее сотрудничество \n")
				eml.AppendText("С наилучшими пожеланиями, служба поддержки портала 72RUS.RU \n")
				isSending=eml.Send(servsmtp)
	 			if (isSending) {Response.Redirect("urlresume.asp?url="+id+"&st=-3")} else {Response.Redirect("admarea.asp")}
			}
			catch(e){
				Response.Redirect("admarea.asp")
			}
}
%>