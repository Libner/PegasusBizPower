<%@ LANGUAGE=VBScript %>
<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
   If (Trim(Request.QueryString("callback")) <> "") Then

       If IsNumeric(Request.QueryString("ContactId")) Then
             ContactId = cLng(Request.QueryString("ContactId"))
             isDouble = cLng(Request.QueryString("isDouble"))

             con.ExecuteQuery "UPDATE CONTACTS Set is_double = "  & isDouble & " WHERE CONTACT_ID=" & ContactId
             Set con = Nothing      
                
        End If

End If       %>