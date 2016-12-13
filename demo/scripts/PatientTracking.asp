<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("/data/SenderDB.mdb")
Set rs = Server.CreateObject("ADODB.recordset")
Set encounterN =  Request.Form("txtQuery")
rs.Open  "SELECT * FROM prescription WHERE PHN=" & encounterN , conn

for each x in rs.Fields
    Response.Write (x.name & ": "& x.value & "<br/>")
next

conn.close

%>
