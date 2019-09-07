<%
function Count_url(tp,hiid){
	var recs=CreateRecordSet()
	if (tp==5) {recs.Source="Select * from get_count_url("+hiid+")"}
	else {recs.Source="Select * from get_count_url_st("+hiid+","+tp+")"}
	recs.Open()
	retval=recs("COUNT_URL").Value
	recs.Close()
	delete recs
	return retval
}
%>