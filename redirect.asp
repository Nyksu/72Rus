<%
Dim vist, currentcount
visit = Application("visitors")
currentcount = Application("myCounter")
%> <%=currentcount & "/" & visit%>