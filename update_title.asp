<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%barID = 10

	sqlstr = "DELETE FROM bar_organizations WHERE bar_id = " & barID & ";" & _
	"DELETE FROM bar_users WHERE bar_id = " & barID
	con.executeQuery(sqlstr)
		
	sqlstr = "Select ORGANIZATION_ID,LANG_ID FROM ORGANIZATIONS Order BY ORGANIZATION_ID"
	set rs_org = con.getRecordSet(sqlstr)
	while not rs_org.eof
		orgID = trim(rs_org(0))
		langID = trim(rs_org(1))
		If trim(langID) =  "2" Then
			lang_ = "_eng"
		Else
			lang_ = ""
		End If	
		sqlstr = "Select bar_id, bar_title" & lang_ & " FROM bars WHERE bar_id = " & barID
		set rs_bars = con.getRecordSet(sqlstr)
		while not rs_bars.eof
			barID = rs_bars(0)
			barTitle = rs_bars(1)	
			barVisible = "1"
			sqlstr = "Insert Into bar_organizations values (" & barID & "," & orgID & ",'" & sFix(barTitle) & "','" & barVisible & "')"
			'Response.Write sqlstr
			'Response.End
			con.executeQuery(sqlstr)
			rs_bars.moveNext
			Wend
			set rs_bars = nothing
			
			sqlstr = "Select User_ID FROM USERS WHERE ORGANIZATION_ID = " & OrgID
			set rs_users = con.getRecordSet(sqlstr)
			while not rs_users.eof
			    userID = trim(rs_users(0))			  				
				sqlstr = "Select bar_id FROM bar_organizations WHERE organization_id = " & OrgID & " AND  bar_id = " & barID
				set rs_bars = con.getRecordSet(sqlstr)
				while not rs_bars.eof
					barID = rs_bars(0)
					barVisible = "1"
					
					sqlstr = "Insert Into bar_users values (" & barID & "," & orgID & "," & userID & ",'" & barVisible & "')"
					con.executeQuery(sqlstr)
					rs_bars.moveNext
				Wend
				set rs_bars = nothing			    	
			    						
				rs_users.moveNext
			wend
			set rs_users = nothing

	rs_org.moveNext
	Wend
	set rs_org = Nothing
%>
