<%
function AddShows(isclick){
	var newid=0
if (String(Session.SessionID())!="undefined") {
	var ipadr=""
	var urlall=""
	var qqstr=""
	var sqltext=""
	var cw1=""
	var rs=CreateRecordSet()
	var newrec=false
	var newses=false
	var newuser=false
	var cocr=""
	var ddd = new Date()
	var ss=""
	var s1=""
	var dd=0
	var mm=0
	
	if (isclick!=1) {isclick=0}
	
	cw1=String(Math.random()).substr(3,20)
	if(cw1.length<20){cw1+=String(Math.random()).substr(3,(20-cw1.length))}
	ipadr=Request.ServerVariables("REMOTE_ADDR")
	qqstr=Request.ServerVariables("QUERY_STRING")
	urlall="http://"+Request.ServerVariables("SERVER_NAME")+Request.ServerVariables("URL")
	if (qqstr!="") {urlall=urlall+"?"+qqstr}
	
	if (isclick==1) {
		if (String(Session("urlstat"))=="undefined") {Session("urlstat")=""}
		rs.Source="select gen_id(BAN600,1) as id from dual"
		rs.Open()
		newid=rs("ID").Value
		rs.Close()
		sqltext="Insert into BAN600(id,dt,ip_adr,is_click,url,dtr) values ("+newid+",'now','"+ipadr+"',"+isclick+",'"+Session("urlstat")+"','TODAY')"
		Connect.BeginTrans()
		try{
			Connect.Execute(sqltext)
			Connect.CommitTrans()
		}
		catch(e){
			Connect.RollbackTrans()
			newid=0
		}
	} else {
		Session("urlstat")=urlall
		rs.Source="Select * from ban600st where dt='TODAY'"
		rs.Open()
		if (rs.EOF) {newrec=true}
		else {
			newid=rs("ID").Value
			cw1=rs("CW").Value
		}
		rs.Close()
		cocr=Request.Cookies("ban600")
		//ss=Math.round(Date.UTC(ddd)/100) + 60*60*24
		//mm=ddd.getMonth()+1
		//dd=ddd.getDate()+1
		//if (dd==32) {
		//	dd=1
		//	mm=mm+1
		//}
		//ss=String(mm)
		//s1=String(dd)
		//if (ss.length<2) {ss="0"+String(ss)}
		//if (s1.length<2) {s1="0"+String(s1)}
		//ss=ss+" "+s1+" "+String(ddd.getFullYear())
		ss="01 01 "+String(ddd.getFullYear()+1)
		if (String(cocr)!=String(cw1)) {
			newuser=true
			Response.Cookies("ban600")
			Response.Cookies("ban600")=cw1
			Response.Cookies("ban600").Expires = ss
		}
		if (String(Session("ban600"))=="undefined") {
			newses=true
			Session("ban600")=1
		}
		if (newrec) {
			sqltext="Insert into BAN600ST(id,dt,allshow,seshow,unishow,cw) values(gen_id(BAN600,1),'TODAY',1,1,1,'"+cw1+"')"
		} else {
			sqltext="Update BAN600ST set allshow=allshow+1, "
			sqltext+="seshow=seshow+"
			if (newses) {sqltext+="1, "} else {sqltext+="0, "}
			sqltext+="unishow=unishow+"
			if (newuser) {sqltext+="1 "} else {sqltext+="0 "}
			sqltext+="where id="+newid
		}
		Connect.BeginTrans()
		try{
			Connect.Execute(sqltext)
			Connect.CommitTrans()
		}
		catch(e){
			Connect.RollbackTrans()
			newid=-1
		}
	}
	
	delete rs
}
	retval=newid//String(ss)+" - "+cocr
	return retval
}
%>