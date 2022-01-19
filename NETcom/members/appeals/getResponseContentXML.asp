<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<% Response.Expires = -1
 Response.Buffer = TRUE
 Response.Clear
 Response.ContentType = "text/xml"
 Response.CharSet="windows-1255"%>
 <root>
<%
 indexResponse=request.QueryString("indexResponse")
 quest_id=request.QueryString("quest_id")
 OrgID=request.QueryString("OrgID")
 Response_Content=""
 if indexResponse<>"" then
	if cint(indexResponse)>0 then
    
		sqlStr = "Select Responses.Response_Title,Responses.Response_Content,Responses.Response_File_1,Responses.Response_File_2,Responses.Response_File_3 from " &_
	" (Select Response_Title,Response_Content,Response_File_1,Response_File_2,Response_File_3,ROW_NUMBER() OVER(ORDER BY Response_Title ASC) AS RowIndex from Product_Responses " & _
	"    WHERE ORGANIZATION_ID= "& OrgID & " And Product_ID = " & quest_id & ") as Responses WHERE  RowIndex=" & indexResponse
		'Response.Write sqlStr
		'response.end
		set rs_Responses = con.GetRecordSet(sqlStr)
		If not rs_Responses.eof Then
			Response_Title = trim(rs_Responses("Response_Title"))
			Response_Content = trim(rs_Responses("Response_Content"))
			Response_File_1=trim(rs_Responses("Response_File_1"))
			Response_File_2=trim(rs_Responses("Response_File_2"))
			Response_File_3=trim(rs_Responses("Response_File_3"))
  		End If
  		set rs_Responses = nothing			
 	End If
End If
response.write "<Response_Title>" & Response_Title & "</Response_Title>"
response.write "<Response_Content>" & Response_Content & "</Response_Content>"
response.write "<Response_File_1>" & Response_File_1 & "</Response_File_1>"
response.write "<Response_File_2>" & Response_File_2 & "</Response_File_2>"
response.write "<Response_File_3>" & Response_File_3 & "</Response_File_3>"
%>
</root>