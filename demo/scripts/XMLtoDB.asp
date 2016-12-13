<%
'Connect to database'
  Set conn = Server.CreateObject("ADODB.Connection")
  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("/data/ReceiverDB.mdb")

'If you use Office 2007 Access DB, then you use following staement to connect to data base:'
  'conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&Server.MapPath("Office 2007 Access DB format")'

'create ADODB object'
  Set rs = Server.CreateObject("ADODB.recordset")

'Create DOM'
  Set doc= Server.CreateObject("Microsoft.XMLDOM")
  doc.async = false

'Load xml document'
  Set XMLfile =  Request.Form("txtQuery")
  Response.Write("Loading: " & XMLfile & "... <p/>")
  doc.Load(Server.MapPath("/XML/" & XMLfile))
  Set root=  doc.documentElement

'If table exists, delete it'
    Set adoxConn = CreateObject("ADOX.Catalog")
    adoxConn.activeConnection = conn

    for each table in adoxConn.tables
        if lcase(table.name) = lcase(root.nodeName) then
            rs.Open "drop table " & root.nodeName, conn
            exit for
        end if
    next

'Create table'
  rs.Open "create table " & root.nodeName, conn
  Set table = root.childNodes(0).childNodes

  For each x in table
    rs.Open "alter table " & root.nodeName & " add " & x.nodeName & " varchar(150)", conn
  Next

Set Xnodelist = root.childNodes
For each x In Xnodelist

Set Ynodelist = x.childNodes

Dim i
i = 1


For each y In Ynodelist
    If i=1 then
        'Response.Write("Moo")
        i=0
        rs.Open "INSERT INTO " & root.nodeName & " (" & y.nodeName & ") VALUES('" & y.text & "')", conn
        Dim namer, id
        namer = y.nodeName
        id = y.text
     Else
        rs.Open "UPDATE " & root.nodeName & " SET " & y.nodeName & "= '" & y.text & "' WHERE " & namer & "='" & id & "'", conn
       'Response.Write(namer)
        'Response.write(id)

     End If
Next

Next

Response.Write ("<br /> Successfully loaded " & XMLfile & " into Database (newDB)!")
 %>

<% 'rs.close %>
<% conn.close %>
<% 'doc.close %>