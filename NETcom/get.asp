<%
  Response.Buffer = True
  Dim objXMLHTTP, xml

  ' Create an xmlhttp object:
  Set xml = Server.CreateObject("Microsoft.XMLHTTP")
  ' Or, for version 3.0 of XMLHTTP, use:
  ' Set xml = Server.CreateObject("MSXML2.ServerXMLHTTP")

  ' Opens the connection to the remote server.
  xml.Open "GET", "https://e-services.clalit.org.il/Zimunet/Content/Zimun/doctors.asp", False
	
  ' Actually Sends the request and returns the data:
  xml.Send

  'Display the HTML both as HTML and as text
  Response.Write "<h1>The HTML text</h1><xmp>"
  Response.Write xml.responseText
  Response.Write "</xmp><p><hr><p><h1>The HTML Output</h1>"

  'Response.Write xml.responseText
 
  
  Set xml = Nothing
%>

