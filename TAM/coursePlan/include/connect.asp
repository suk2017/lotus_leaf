<%
Set conn=Server.CreateObject("ADODB.Connection")
connstr1="Provider=SQLOLEDB;Data Source=127.0.0.1;User ID=user1;Password=123456;Initial Catalog=ATADB"
conn.Open connstr1
%>
