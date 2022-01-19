<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<% 
	UserID=trim(Request.Cookies("bizpegasus")("UserID"))
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	perSize = trim(Request.Cookies("bizpegasus")("perSize"))
	is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
		self_name = "Self"
	Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
		self_name = "עצמי"
	End If		  	
  
	start_date = trim(Request("dateStart"))		
	end_date = trim(Request("dateEnd"))
	start_date_ =  Day(start_date) & "/" & Month(start_date) & "/" & Year(start_date)
    end_date_ =  Day(end_date) & "/" & Month(end_date) & "/" & Year(end_date)
    CompanyID = trim(Request("company_id"))
	ProjectID = trim(Request("project_id"))
	mechanismID = trim(Request("mechanism_id")) 
  
	Function GetFieldGraph(pProduct_Id, pField_Id, pField_Title, pField_Type, pField_Scale, pLanguage, pType_ID)
	Dim tmpStr , selValue, td_align
	Dim selAll
	Dim colind, wd, i, td_width, ScQuant
	Set rsSelect = Server.CreateObject("AdoDB.RecordSet")
	
	'col = Array("#65A825", "#19B4D0", "#E87117", "#6E15C3", "#44B4AD", "#011D66", "#F84637", "#1DF3FD", "#EBB20D")
	col = Array("color1.gif", "color2.gif", "color3.gif", "color4.gif", "color5.gif", "color6.gif", _
	"color7.gif", "color8.gif", "color9.gif", "color10.gif", "color11.gif", "color12.gif", "color13.gif", _
	"color14.gif", "color15.gif", "color16.gif")
	colind = 0
    If pLanguage = "eng" Then
        td_align = "left"
    Else
        td_align = "right"
        string_sort = "DESC"
    End If
    Select Case pField_Type  
    Case "3", "8", "11" 'select
        sqlstr = "Exec dbo.FIELD_STATISTICS_PROC '" & pProduct_Id & "','" & pField_Id & "','" & pType_ID &_
        "','" & start_date_ & "','" & end_date_ & "','" & UserID & "','" & is_groups & "','" & CompanyID &_
        "','" & ProjectID & "','" & mechanismID & "'"
        'Response.Write sqlstr
        'Response.End
        Set rsSelect = con.GetRecordSet(sqlstr)
        'Response.Write rsSelect.RecordCount 
        'Response.End
        selAll = 0
        tmpStr = "<tr><td class='card5' style='font-size:11pt;padding-right:5px;padding-left:5px' dir=""rtl"" align=" & td_align & " valign=top colspan=" & rsSelect.RecordCount & ">" & trim(pField_Title) & "</td></tr>" & vbCrLf
        If rsSelect.RecordCount <> 0 Then
            td_width = cInt(100 / rsSelect.RecordCount)
        End If
        
        If Not rsSelect.EOF Then
            tmpStr = tmpStr & "<tr>" & vbCrLf
            arrSelect = rsSelect.GetRows()
            Set rsSelect = Nothing
            
            If IsArray(arrSelect) Then
            For ss=0 To Ubound(arrSelect,2)
                selAll = selAll + CInt(arrSelect(2, ss))            
            Next
            
            For ss=0 To Ubound(arrSelect,2)
                If selAll <> 0 Then
                    wd = CInt((CInt(arrSelect(2, ss)) / selAll) * 100)
                    If arrSelect(2, ss) <> 0 Then
                        str_FIELD_QUANTITY = " &nbsp;<span style=color:blue;font-size:10pt>(" & arrSelect(2, ss) & ")</span>"
                    Else
                        str_FIELD_QUANTITY = ""
                    End If
                Else
                    wd = 0
                End If
                tmpStr = tmpStr & "<td align=center valign=bottom width='" & td_width & "%' height=120>" & vbCrLf
                tmpStr = tmpStr & "<TABLE width='100%' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
                tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf                
                tmpStr = tmpStr & "<tr><td align=center valign=bottom><img src='../../images/"& col(colind Mod 16) &"' style='vertical-align: bottom' border=0 hspace=0 vspace=0 width='70%' height='" & wd & "'>" & vbCrLf
                tmpStr = tmpStr & "</td></tr></table></td>" & vbCrLf
            
            colind = colind + 1
            Next
            tmpStr = tmpStr & "</tr><tr>" & vbCrLf
            
            For ss=0 To Ubound(arrSelect,2)
                tmpStr = tmpStr & "<td align=center dir=rtl class=card5 style='font-size:11pt;padding-right:5px;padding-left:5px'>" & trim(arrSelect(1, ss)) & "</td>"
            Next
            tmpStr = tmpStr & "</tr>"
            End If            
        End If  'not rsSelect.EOF
        Set arrSelect = Nothing        
       
        GetFieldGraph = tmpStr
    Case "9", "12" 'radio
        sqlstr = "Exec dbo.QUESTION_VALUES_PROC '" & pProduct_Id & "','" & pField_Id & "','" & pType_ID &_
        "','" & start_date_ & "','" & end_date_ & "','" & UserID & "','" & is_groups & "','" & CompanyID &_
        "','" & ProjectID & "','" & mechanismID & "'"        
        Set rsSelect = con.GetRecordSet(sqlstr)
        selAll = rsSelect.RecordCount
            
        tmpStr = "<tr><td class='card5' style='font-size:11pt;padding-right:5px;padding-left:5px' dir=""rtl"" align=" & td_align & " valign=top colspan=" & pField_Scale & ">" & trim(pField_Title) & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr>" & vbCrLf
        
        If pField_Scale <> 0 Then
            td_width = 100 / pField_Scale
        End If
        
        If pLanguage = "heb" Then
            fori = CInt(pField_Scale)
            fotto = 1
            forstep = -1
        Else
            fori = 1
            fotto = CInt(pField_Scale)
            forstep = 1
        End If
        
        For i = fori To fotto Step forstep
            rsSelect.Filter = "FIELD_VALUE='" & i & "'"
            ScQuant = rsSelect.RecordCount
            If ScQuant <> 0 Then
                wd = CInt((ScQuant / selAll) * 100)
                str_FIELD_QUANTITY = "&nbsp;<span style=color:blue;font-size:10pt>(" & ScQuant & ")</span>"
            Else
                wd = 0
                str_FIELD_QUANTITY = ""
            End If
            tmpStr = tmpStr & "<td align=center valign=bottom width='" & td_width & "%' height=120>" & vbCrLf
            tmpStr = tmpStr & "<TABLE width='100%' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
            tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
            tmpStr = tmpStr & "<tr><td align=center valign=bottom><img src='../../images/"& col(colind Mod 16) &"' style='vertical-align: bottom' border=0 hspace=0 vspace=0 width='70%' height='" & wd & "'>" & vbCrLf
            tmpStr = tmpStr & "</td></tr></table></td>" & vbCrLf
        colind = colind + 1
        Next
        Set rsSelect = Nothing
        tmpStr = tmpStr & "</tr><tr>" & vbCrLf
        For i = fori To fotto Step forstep
            tmpStr = tmpStr & "<td align=center dir=rtl class=card5 style='font-size:11pt'>&nbsp;" & i & "&nbsp;</td>"
        Next
        tmpStr = tmpStr & "</tr>"
        
        GetFieldGraph = tmpStr
    Case "5" 'check
        sqlstr = "Exec dbo.QUESTION_VALUES_PROC '" & pProduct_Id & "','" & pField_Id & "','" & pType_ID &_
        "','" & start_date_ & "','" & end_date_ & "','" & UserID & "','" & is_groups & "','" & CompanyID &_
        "','" & ProjectID & "','" & mechanismID & "'"        
        Set rsSelect = con.GetRecordSet(sqlstr)
        selAll = rsSelect.RecordCount
        tmpStr = "<tr><td class='card5' style='font-size:11pt;padding-right:5px;padding-left:5px' dir=""rtl"" align=" & td_align & " valign=top colspan=2>" & trim(pField_Title) & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr>" & vbCrLf
        rsSelect.Filter = "FIELD_VALUE='on'"
        ScQuant = rsSelect.RecordCount
        Set rsSelect = Nothing
        If ScQuant <> 0 Then
            wd = CInt((ScQuant / selAll) * 100)
            str_FIELD_QUANTITY = "&nbsp;<span style=color:blue;font-size:10pt>(" & ScQuant & ")</span>"
        Else
            wd = 0
            str_FIELD_QUANTITY = ""
        End If
        tmpStr = tmpStr & "<td align=center valign=bottom width='320' height=120>" & vbCrLf
        tmpStr = tmpStr & "<TABLE width='320' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center valign=bottom><img src='../../images/"& col(colind Mod 16) &"' style='vertical-align: bottom' border=0 hspace=0 vspace=0 width='70%' height='" & wd & "'>" & vbCrLf
        tmpStr = tmpStr & "</td></tr></table></td>" & vbCrLf        
        
        ScQuant = selAll - ScQuant
        If ScQuant <> 0 Then
            wd = CInt((ScQuant / selAll) * 100)
            str_FIELD_QUANTITY = "&nbsp;<span style=color:blue;font-size:10pt>(" & ScQuant & ")</span>"
        Else
            wd = 0
            str_FIELD_QUANTITY = ""
        End If
        colind = colind + 1
        tmpStr = tmpStr & "<td align=center valign=bottom height=120>" & vbCrLf
        tmpStr = tmpStr & "<TABLE width='320' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf        
        tmpStr = tmpStr & "<tr><td align=center valign=bottom><img src='../../images/"& col(colind Mod 16) &"' style='vertical-align: bottom' border=0 hspace=0 vspace=0 width='70%' height='" & wd & "'>" & vbCrLf
        tmpStr = tmpStr & "</td></tr></table></td></tr>" & vbCrLf        
        tmpStr = tmpStr & "<tr><td align=center class=card5 style='font-size:11pt' nowrap>&nbsp;כן&nbsp;</td><td align=center class=card5 nowrap>&nbsp;לא&nbsp;</td></tr>"
         
        GetFieldGraph = tmpStr
   
    Case Else
    
        GetFieldGraph = ""    
        
    End Select
   
End Function
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
	prodId=trim(Request.Form("quest_id"))
	type_reports=trim(Request.QueryString("type_reports"))
	If trim(type_reports) = "cl" Then
		type_id = 2
	Else
	    type_id = 1
	End If
	
	strTitle = ""

	If trim(CompanyID) <> "" Then
		sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
		set rs_com = con.getRecordSet(sqlstr)
		if not rs_com.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#666699"">" & trim(rs_com(0)) & "</font>"
		end if
		set rs_com = nothing
	End If

	If trim(ProjectID) <> "" Then
		sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
		set rs_project = con.getRecordSet(sqlstr)
		if not rs_project.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#666699"">" & trim(rs_project(0)) & "</font>"
		end if
		set rs_project = nothing
	End If

	If trim(mechanismID) <> "" Then
		sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & mechanismID
		set rs_mech = con.getRecordSet(sqlstr)
		if not rs_mech.eof then
			If lang_id = 1 Then
				strTitle = strTitle & "<br>&nbsp;מנגנון:&nbsp;<font color=""#666699"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			Else
				strTitle = strTitle & "<br>&nbsp;Sub-Project:&nbsp;<font color=""#666699"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			End If			
		end if
		set rs_mech = nothing
	End If		
%>
<table border="0" width="640" cellspacing="0" cellpadding="0" align=center>
<tr>
<td width="640" align="center">
<!--#INCLUDE FILE="logo_top.asp"-->
</td></tr> 
<tr><td align="center">
<table border="0" width="640" cellspacing="0" cellpadding="0" align=center>
<!-- start code --> 
	<tr>
		<td class="card5" style="font-size:16pt" align=center width=640 dir="<%=dir_obj_var%>"><b>			
         <%If trim(lang_id) = "1" Then%>&nbsp;דוח התפלגות שאלות (עמודות)&nbsp;
         <%Else%>&nbsp;Bar distribution report&nbsp;<%End If%></b></td>
	</tr>
	<tr>
		<td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" width=640 align="center"><b>
<%
			sqlStr = "Select Product_Name,QUESTIONS_ID,Langu  from products where product_id=" & prodId
			''Response.Write sqlStr
			''Response.End			
			set prod = con.GetRecordSet(sqlStr)
					if not prod.eof then
					productName=prod("Product_Name")					
					quest_id = prod("QUESTIONS_ID")
					if prod("Langu") = "eng" then
						dir_align = "ltr"
						td_align = "left"
						pr_language = "eng"
					else
						dir_align = "rtl"
						td_align = "right"
						pr_language = "heb"
					end if
				end if
				set prod = nothing%>
			&nbsp;<%=productName%>&nbsp;<%=strTitle%></b></td>
	</tr>
	<tr><td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" width=100% align="center"><b><%=start_date%>  -  <%=end_date%></b></td></tr>
	<!--#INCLUDE FILE="graph_fields.asp"-->
    <!-- end code --> 
	<tr><td width="650" height=5></td></tr>			
	<!--#INCLUDE FILE="bottom_inc.asp"-->	
</table>
</td></tr></table>
</body>
<%set con = nothing%>
</html>