<html>
<body>

<%

    Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("/data/newDB.mdb")
	Set rs = Server.CreateObject("ADODB.recordset")
    Set adoxConn = CreateObject("ADOX.Catalog")
    adoxConn.activeConnection = conn



  found = false
    for each table in adoxConn.tables
        if lcase(table.name) = "unit" then
            found = true
            exit for
        end if
    next
    conn.close: set conn = nothing
    set adoxConn = nothing


    if found then
        response.write("Table exists.")
    else
        response.write("Table does not exist.")
    end if
%>

</body>
</html>
