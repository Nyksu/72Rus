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
				eml.FromName="������ ��������� ������� 72RUS.RU"
				eml.Subject=tit+"����������� � ����� ������ �� ������ �������� ������ ������� 72RUS.RU"
				eml.AppendText(" ������������, "+name+"!\n")
				eml.AppendText(" ���� ��� �������� ������ � ������� ������ ������ ������� : "+urlname+"\n")
				eml.AppendText(" � ������� : "+url+"\n")
				eml.AppendText(" \n")
				eml.AppendText("��� ������ ���������� ���, �� ������� ����, ��� ��� ���� ��� ��������� ��� ��������� � ������� �72RUS.RU - ��������� ������. \n")
				eml.AppendText("�������������� ��������� ���������� ������ �������� ������� ���������� 150 ������. \n")
				eml.AppendText("������ ������������ ���� ���������� ������������� (��������� ������ �� ����������� �����:  \n")
				eml.AppendText("http://rusintel.ru/message.asp), ���� ����� ������� WebMoney ������������� 5.00 WMZ �� �������  \n")
				eml.AppendText("Z200556237621 � ����������� � ������� ������� ����� � �������� �����. \n")
				eml.AppendText(" \n")
				eml.AppendText("���� � ������� 20 ���� ������ �� ����� ����������� - ���� ������ ����� ������� �� �������. \n")
				eml.AppendText(" \n")
				eml.AppendText("������� ���������� ������ � ������� 72RUS.RU ��: http://www.72rus.ru/addurl.asp \n")
				eml.AppendText(" \n")
				eml.AppendText("�������� �� ���������� �������������� \n")
				eml.AppendText("� ���������� �����������, ������ ��������� ������� 72RUS.RU \n")
				isSending=eml.Send(servsmtp)
	 			if (isSending) {Response.Redirect("urlresume.asp?url="+id+"&st=-3")} else {Response.Redirect("admarea.asp")}
			}
			catch(e){
				Response.Redirect("admarea.asp")
			}
}
%>