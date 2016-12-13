<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("/data/SenderDB.mdb")

'Connect to Office 2007 Access DB'
' conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&Server.MapPath("Office 2007 Access DB format")'

Set rs = Server.CreateObject("ADODB.recordset")
Set root =  Request.Form("txtQuery")
%>

<%
'Test if table exists: flag=0'
    ExistFlag = 0
    Set adoxConn = CreateObject("ADOX.Catalog")
    adoxConn.activeConnection = conn

    for each table in adoxConn.tables
        if lcase(table.name) = root then
            ExistFlag =1
            exit for
        end if
    next
%>

<% If (ExistFlag) then %>
	<% Set file = Server.CreateObject("Scripting.FileSystemObject") %>
	<% Set toText = file.CreateTextFile("C:\HINF200-Demo\XML\"&root&".xml", True) %>
	<% rs.Open  "SELECT * FROM " & root, conn %>
	<% Response.Write("<?xml version=""1.0"" encoding=""ISO-8859-1""?>")%>
	<% toText.Write("<?xml version=""1.0"" encoding=""ISO-8859-1""?>")%>
	<% toText.Write(VbCrLf)%>
	<% toText.Write(VbCrLf)%>

	<<% Response.Write(root)%>>
	<% toText.Write("<")%>
	<% toText.Write(root)%>
	<% toText.Write(">")%>
	<% toText.Write(VbCrLf)%>

	<% Do until rs.EOF %>
	   <Record>
	   <% toText.Write("  <Record>")%>
	   <% toText.Write(VbCrLf)%>
	   <% for each x in rs.Fields %>
		  <<% Response.Write (x.name)%>>
		  <% toText.Write("    <")%>
		  <% toText.Write(x.name)%>
		  <% toText.Write(">")%>
			  <% Response.Write (x.value) %>
			  <% toText.Write(x.value)%>
		  </<% Response.Write (x.name)%>>
		  <% toText.Write("</")%>
		  <% toText.Write(x.name)%>
		  <% toText.Write(">")%>
		  <% toText.Write(VbCrLf)%>

		<% next %>
	  <% rs.MoveNext %>
	  </Record>
	   <% toText.Write("  </Record>")%>
	   <% toText.Write(VbCrLf)%>
	<%Loop %>
	</<% Response.Write(root)%>>
	<% toText.Write("</")%>
	<% toText.Write(root)%>
	<% toText.Write(">")%>
	<% toText.Write(VbCrLf)%>

	<% rs.close %>
<% else Response.Write (root &" doesn't exist in SenderDB!") %>
<% end if %>
<% conn.close %>