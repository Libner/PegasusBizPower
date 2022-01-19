<!--#include file="../../connect.asp"-->
<%	sqlstr = "SET DATEFORMAT DMY; Select Max(DATE_SEND_END) From PRODUCTS Where " & _
	" (PRODUCT_NUMBER Like '111') And (DateDiff(mi, DATE_SEND_END, GetDate()) <= 1)"
	Set rs_tmp = con.getRecordSet(sqlstr)
	If Not rs_tmp.Eof Then
		max_date_send_end = trim(rs_tmp(0))
	Else
		max_date_send_end = null
	End If
	Set rs_tmp = Nothing
	
	IsOkToSend = 1
	If IsDate(max_date_send_end) Then
		IsOkToSend = 0
	End If    
    
 Response.Write IsOkToSend
 set con = Nothing%>    