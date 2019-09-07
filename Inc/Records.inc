<%
var company=1
var catalog=1
var market=1

var Connect=Server.CreateObject("ADODB.Connection")
Connect.Open("DSN=SB;UID=siteb;PWD=Ghjyby")


var Records=Server.CreateObject("ADODB.Recordset")
Records.ActiveConnection=Connect
Records.CursorType = 3
Records.CursorLocation = 3
Records.LockType = 3

%>