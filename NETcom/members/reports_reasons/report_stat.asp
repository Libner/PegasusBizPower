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
	start_date =  Day(start_date) & "/" & Month(start_date) & "/" & Year(start_date)
    end_date =  Day(end_date) & "/" & Month(end_date) & "/" & Year(end_date)
	prodId = trim(Request.Form("quest_id"))
	categID = trim(Request.Form("categID"))	
  
reasons=""
	'only for tofes mitan'en
	
			'הגורם ליצירת הטופס
			'---------------P2932 - type of form- REASONS ---------------
				'select creation reason-------------------------------
				'added by Mila 22/10/2019-----------------------------	
			reasons=Request.Form("reasons")
			reasons=replace(reasons," ","")
	'-----------------------------------------------------------
		   
	sqlStr = "Select Product_Name,QUESTIONS_ID,Langu From Products Where Product_id=" & prodId
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
	set prod = nothing
	
	if prodId=16504 and reasons<>"" then
	'---------------P2932 - type of form- REASONS ---------------
				'select creation reason-------------------------------
				'added by Mila 22/10/2019-----------------------------	
		sql = "SET DATEFORMAT DMY; SELECT DISTINCT St1.category_ID, " &_
		" (Select referer from statistic_from_banner where category_ID = St1.category_ID), " &_
		" (Select TOP 1 St2.Appeal_ID From Statistic St2 inner join appeals on appeals.APPEAL_ID=St2.Appeal_ID  Where St1.category_ID = St2.category_ID " &_
		" AND DateDiff(d, St2.[Date], '" & start_date & "') <= 0 " &_
		" AND DateDiff(d, St2.[Date], '" & end_date & "') >= 0 And St2.Product_ID = " & prodId & " And appeals.Reason_Id in(" & reasons & "))" &_
		" FROM Statistic St1 WHERE St1.Product_ID = " & prodId
	else
		sql = "SET DATEFORMAT DMY; SELECT DISTINCT St1.category_ID, " &_
		" (Select referer from statistic_from_banner where category_ID = St1.category_ID), " &_
		" (Select TOP 1 St2.Appeal_ID From Statistic St2 Where St1.category_ID = St2.category_ID " &_
		" AND DateDiff(d, St2.[Date], '" & start_date & "') <= 0 " &_
		" AND DateDiff(d, St2.[Date], '" & end_date & "') >= 0 And St2.Product_ID = " & prodId & ")" &_
		" FROM Statistic St1 WHERE St1.Product_ID = " & prodId
	end if
	
	If Len(categID) > 0 Then
		sql = sql & " And St1.category_ID IN (" & categID & ")" 
	End If
	'Response.Write sql
	'Response.End 
	set pr=con.getRecordSet(sql)
	if not pr.EOF then
		arr_cat = pr.getRows()
	end if
	Set pr = Nothing	
	%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
</head>
<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0">
<table border="0" width="750" cellspacing="0" cellpadding="0" align=center>
<tr>
<td nowrap align="center">
<!--#INCLUDE FILE="logo_top.asp"-->
</td>
<td width="750">
<table border="0" width="750" cellspacing="0" cellpadding="0" align=center>
<!-- start code --> 
	<tr>
		<td class="card5" style="font-size:16pt" align="right" dir="<%=dir_obj_var%>"><b><%If trim(lang_id) = "1" Then%>התפלגות מכירות עפ"י מקורות פרסום
         <%Else%>Bar distribution report<%End If%></b></td>
	</tr>
	<tr>
		<td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" align="right"><b><%=productName%></b></td>
	</tr>	
	<tr><td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" align="right"><b><%=start_date%>  -  <%=end_date%></b></td></tr>
	<%'---------------P2932 - type of form- REASONS ---------------
   	If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			%>
			<tr ><td  class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" align="right"><b> הגורם ליצירת הטופס:
 				<%For mm=0 To Ubound(arr_Reason,2)
					if mm>0 then%>
					,&nbsp;
					<%end if%>
					<%= trim(arr_Reason(0,mm))%>
				<%next%>
				</b></td></tr>
<%
			end if
	end if
	%>
    </table>
</td>
</tr> 
<tr><td align="center" colspan=2>
<table border="0" width="750" cellspacing="0" cellpadding="0" align=center>
	<tr><td height=10 nowrap></td></tr>
	<%If isArray(arr_cat) Then
	  For cc=0 To Ubound(arr_cat,2)	  
		catID = trim(arr_cat(0,cc))
		If isNumeric(catID) Then	
			catName = trim(arr_cat(1,cc))		
			If Len(catName) = 0 Or isNULL(catName) Then
				catName = "ללא מקור פרסום"
			End If
		Else
			catName = "סה''כ"
		End If
		IsAppeals = trim(arr_cat(2,cc))	%>
	<tr><td style="font-size:12pt;color:#736BA6" class="tab"><%=catName%></td></tr> 
	<tr><td><%If trim(IsAppeals) <> "" And Not isNULL(IsAppeals) Then%>
	<img src="combine_graph.asp?prodId=<%=prodId%>&catID=<%=catID%>&start_date=<%=Server.HTMLEncode(start_date)%>&end_date=<%=Server.HTMLEncode(end_date)%>">
	<%Else%><div align="center" style="padding:10px;" >אין טפסים מלאים</div>
	<%End If%>
	</td></tr>	
	<%Next
	End If%>
	<tr><td style="font-size:12pt;color:#736BA6" class="tab">סה"כ</td></tr> 
	<tr><td><img src="combine_graph.asp?prodId=<%=prodId%>&catID=&start_date=<%=Server.HTMLEncode(start_date)%>&end_date=<%=Server.HTMLEncode(end_date)%>"></td></tr>
	<tr><td><img src="product_graph_all.asp?prodId=<%=prodId%>&start_date=<%=Server.HTMLEncode(start_date)%>&end_date=<%=Server.HTMLEncode(end_date)%>"></td></tr>
    <!-- end code --> 
	<!--#INCLUDE FILE="bottom_inc.asp"-->	
</table>
</td></tr></table>
</body>
<%set con = nothing%>
</html>