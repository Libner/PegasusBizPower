<%SERVER.ScriptTimeout=3000%>
<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="..\..\netcom/connect.asp"-->
<%'on error resume next
   ORGANIZATION_ID=trim(Request.QueryString("ORGANIZATION_ID"))
   F=trim(Request.QueryString("F"))   

   set con_image=Server.CreateObject("adodb.connection")
   con_image.open "FILE NAME="& server.MapPath("..\..\netcom/netReply.udl")
   set pr=Server.CreateObject("ADODB.Recordset")
   pr.Open "SELECT * FROM ORGANIZATIONS where ORGANIZATION_ID="&ORGANIZATION_ID&" ",con_image,2,3
   set upl=Server.CreateObject("SoftArtisans.FileUp") 

   upl.SaveAsBlob pr.Fields(F)
   pr.Update
   pr.Close
   set pr = Nothing
  
   'Response.Write "ok"
   'Response.End
   Response.Redirect "addOrg.asp?ORGANIZATION_ID="&ORGANIZATION_ID
 
Set con=Nothing
%>
