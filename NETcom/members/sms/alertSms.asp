<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
	msg=trim(Request.QueryString("msg"))
	tid=trim(Request.QueryString("tid"))
	info=trim(Request.QueryString("info"))
	sms_status = 1
	
	Select Case UCase(trim(msg))
		Case "ERROR" : error = msg :  sms_status = 1 	'- ����� ����� ������ ����� �������.
		Case "START" : error = "" :  sms_status = 2  '- ������ ����� �������.
		Case "END" : error = "" :  sms_status = 3  '- ������� ����� �������, ������ info ���� �� ���� ������� ������.
		Case "STOP" : error = "" :  sms_status = 3  '- ����� ����� ����� �� ����� �������, ������ info ���� �� ���� ������� ������.
		Case "NO_SMS_LEFT" : error = msg :  sms_status = 1  '- ������ ������ ���� ��� ����� SMS ������ ������, ����� �� ����� ������ ������ ������ ����.
		Case "NO_USERS_FOUND" : error = msg :  sms_status = 1  '- �� ����� ������� ����� ����� ����� ��� ����� (����� ���, �� �� �������� ����� �������� ������� ����� ����).
		Case "UNKNOWN_ERROR" : error = msg :  sms_status = 1   '- ����� ����� �� �����.
		Case "CP_SEND_ERROR" : error = msg :  sms_status = 1  '- ����� ����� ������ ���� �� �� �������
	End Select
	
	If IsNumeric(tid) And trim(tid) <> "" Then	
	 sqlstr = "UPDATE sms_sends SET error = '" & vFix(error) & "', status = '" & vFix(sms_status) & "', info = '" & vFix(info) & "' WHERE tid =  " & tid
	 con.ExecuteQuery(sqlstr)
   End If	 
	 
   Set con = Nothing %>