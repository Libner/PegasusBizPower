<%SERVER.ScriptTimeout=3000%>
<%'on error resume next
   PRODUCT_ID=trim(Request.QueryString("PRODUCT_ID"))
   F=trim(Request.QueryString("F"))   

   set con_image=Server.CreateObject("adodb.connection")
   con_image.open "FILE NAME="& server.MapPath("../../netReply.udl")
   
   If Request.QueryString("delLogo") <> nil Then
		con_image.execute("UPdate PRODUCTS set PRODUCT_LOGO = null where PRODUCT_ID = " & PRODUCT_ID)
   Else 
		set pr=Server.CreateObject("ADODB.Recordset")
		pr.Open "SELECT * FROM PRODUCTS where PRODUCT_ID="&PRODUCT_ID&" ",con_image,2,3
		set upl=Server.CreateObject("SoftArtisans.FileUp") 

		upl.SaveAsBlob pr.Fields(F)
		pr.Update
		pr.Close
		set pr = Nothing
		End If
   Response.Redirect "addquestions.asp?prodID="&PRODUCT_ID
 
Set con=Nothing
%>
