<%
Set XMLfile =  Request.Form("XMLname")
Response.Write ("Loading  and validating [" & XMLfile & "] ...  <p/>")
Set doc= Server.CreateObject("Microsoft.XMLDOM")
doc.async = False
doc.Load (Server.MapPath("/XML/" & XMLfile))
Set SCHEMAfile =  Request.Form("SCHEMAName")
Set root=  doc.documentElement
root.setAttribute "xmlns", "x-schema:" & Server.MapPath("/XML/" & SCHEMAfile)
Set ValDoc = CreateObject("Microsoft.XMLDOM")
ValDoc.validateOnParse = True
ValDoc.async = False

ValDoc.load doc

If (ValDoc.parseError.errorCode = 0 ) Then
	 Response.Write ("<FONT color=#ff0000> *** Validation successful ! </FONT><p/>")
	 Response.Write (root.nodeName & "<p/>")

	 Set Xnodelist = root.childNodes
	 For each x In Xnodelist
		Response.Write ("- " & x.nodeName & "<br/>")
	   Set Ynodelist = x.childNodes
	   For each y In Ynodelist
		  Response.Write ("-- " & y.nodeName &": "& y. text & "<br/>")
	   Next

	   Response.Write ("<br/>")
	 Next
Else
	 Response.Write ("<FONT color=#ff0000>*** Validation fail ! </FONT><p/>")
End If

%>
