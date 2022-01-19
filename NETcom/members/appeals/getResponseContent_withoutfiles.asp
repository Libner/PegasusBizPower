<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
 indexResponse=request.QueryString("indexResponse")
 quest_id=request.QueryString("quest_id")
 OrgID=request.QueryString("OrgID")
 Response_Content=""
 if indexResponse<>"" then
	if cint(indexResponse)>0 then
    
    sqlStr = "Select Responses.Response_Content from " &_
 " (Select Response_Content,ROW_NUMBER() OVER(ORDER BY Response_Title ASC) AS RowIndex from Product_Responses " & _
 "    WHERE ORGANIZATION_ID= "& OrgID & " And Product_ID = " & quest_id & ") as Responses WHERE  RowIndex=" & indexResponse
	'Response.Write sqlStr
	'response.end
	set rs_Responses = con.GetRecordSet(sqlStr)
	If not rs_Responses.eof Then
	Response_Content = rs_Responses("Response_Content")
  		End If
  		set rs_Responses = nothing			
 		End If
End If
response.write (Response_Content)%>