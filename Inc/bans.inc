<%
function AddClik(bid,ura){
	if (String(Session.SessionID())!="undefined") {
		var sqltext=""
		var rs=CreateRecordSet()
		var ipadr=Request.ServerVariables("REMOTE_ADDR")
		var newid=0
		
		rs.Source="select gen_id(BAN600,1) as id from dual"
		rs.Open()
		newid=rs("ID").Value
		rs.Close()
		sqltext="Insert into BANCLIKS(id,dt,ip_adr,url,dtr,baner_id) values ("+newid+",'now','"+ipadr+"','"+ura+"','TODAY', "+bid+")"
		Connect.BeginTrans()
		try{
			Connect.Execute(sqltext)
			Connect.CommitTrans()
		}
		catch(e){
			Connect.RollbackTrans()
			newid=0
		}
		delete rs
	}
	retval=newid
	return retval
}

function AddShows(rid){
	var cw1=""
	var bid=0
	var newrec=false
	var newban=false
	var newid=0
	var banst=9
	var recst=9
	var sqltext=""
	var sqltext1=""
	var cocr=""
	var cocr1=""
	var newuser=false
	var newusr=false
	var newbid=0
	var newblok=false
	var cw2=""
	if (String(Session.SessionID())!="undefined") {
		var rs=CreateRecordSet()
		var ddd = new Date()

		cw1=String(Math.random()).substr(3,20)
		if(cw1.length<20){cw1+=String(Math.random()).substr(3,(20-cw1.length))}
		cw2=cw1

		rs.Source="Select t1.*, t3.state as bstat from banblock t1, baner t3 where t1.state=0 and t1.stopdate > 'TODAY' and t1.startdate <= 'TODAY' and ((t1.maxshows > (Select sum(t2.allshow) from banblockst t2 where t1.id=t2.banblock_id and t1.baner_id=t2.baner_id)) or (not exists (Select * from banblockst t2 where t1.id=t2.banblock_id and t1.baner_id=t2.baner_id) )) and t1.baner_id=t3.id and t1.id="+rid
		rs.Open()
		if (!rs.EOF) {
			recst=rs("STATE").Value
			bid=rs("BANER_ID").Value
			banst=rs("BSTAT").Value
		}
		rs.Close()		
if (banst==0 && recst==0) {
		rs.Source="Select * from banblockst where dt='TODAY' and baner_id="+bid+" and banblock_id="+rid
		rs.Open()
		if (rs.EOF) {newban=true}
		else {
			newid=rs("ID").Value
			cw1=rs("CW").Value
			}
		rs.Close()

		rs.Source="Select * from blockbanst where dt='TODAY' and banblock_id="+rid
		rs.Open()
		if (rs.EOF) {newblok=true}
		else {
			cw2=rs("CW").Value
			newbid=rs("ID").Value
		}
		rs.Close()

		if (newban) {
			rs.Source="select gen_id(BAN600,1) as id from dual"
			rs.Open()
			newid=rs("ID").Value
			rs.Close()
			sqltext="Insert into banblockst(id,dt,baner_id,banblock_id,allshow,unishow,cw) values("+newid+",'TODAY',"+bid+","+rid+",1,1,'"+cw1+"')"
		} else {sqltext="Update banblockst set allshow=allshow+1, unishow=unishow+%UN where id="+newid}

		if (newblok) {
			rs.Source="select gen_id(BAN600,1) as id from dual"
			rs.Open()
			newbid=rs("ID").Value
			rs.Close()
			sqltext1="Insert into blockbanst(id,dt,banblock_id,allshow,unishow,cw) values("+newid+",'TODAY',"+rid+",1,1,'"+cw2+"')"
		} else {sqltext1="Update blockbanst set allshow=allshow+1, unishow=unishow+%UN where id="+newbid}

		cocr=Request.Cookies("bn72rus")("R"+String(rid)+"B"+String(bid))
		ss="01 01 "+String(ddd.getFullYear()+1)
		if (String(cocr)!=String(cw1)) {
			newuser=true
			Response.Cookies("bn72rus")("R"+String(rid)+"B"+String(bid))=cw1
			Response.Cookies("bn72rus").Expires = ss
		}
		if (newuser) {sqltext=sqltext.replace("%UN",1)} else {sqltext=sqltext.replace("%UN",0)}

		cocr1=Request.Cookies("bn72rus")("R"+String(rid))
		if (String(cocr1)!=String(cw2)) {
			newusr=true
			Response.Cookies("bn72rus")("R"+String(rid))=cw2
		}
		if (newusr) {sqltext1=sqltext1.replace("%UN",1)} else {sqltext1=sqltext1.replace("%UN",0)}

		Connect.BeginTrans()
		try{
			Connect.Execute(sqltext)
			Connect.Execute(sqltext1)
			Connect.CommitTrans()
		}
		catch(e){
			Connect.RollbackTrans()
			newid=0
		}
		delete rs
	}
	}
	retval=newid
	return retval
}
%>