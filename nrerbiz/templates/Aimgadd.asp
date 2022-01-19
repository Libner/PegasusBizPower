<%SERVER.ScriptTimeout=3000%>
<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="..\..\netcom/connect.asp"-->
<%'on error resume next
   templateId=trim(Request.QueryString("templateId"))
   F=trim(Request.QueryString("F"))   

   set con_image=Server.CreateObject("adodb.connection")
   con_image.open "FILE NAME="& server.MapPath("..\..\netcom/netReply.udl")
   set pr=Server.CreateObject("ADODB.Recordset")
   pr.Open "SELECT * FROM Templates where Template_ID="&templateId&" ",con_image,2,3
   set upl=Server.CreateObject("SoftArtisans.FileUp") 

   upl.SaveAsBlob pr.Fields(F)
   pr.Update
   pr.Close
   set pr = Nothing
  
   'Response.Write "ok"
   'Response.End
   Response.Redirect "editScreen.asp?templateId="&templateId
 
Set con=Nothing
%>
