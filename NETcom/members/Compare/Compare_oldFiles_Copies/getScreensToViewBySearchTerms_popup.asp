<!--#include file="../../connect.asp"-->
<%  

	If Request.form("tourId") <> nil Then
		if isnumeric(trim(Request.form("tourId"))) then
			if trim(Request.form("tourId"))<>"0" then
				tourId = trim(Request.form("tourId"))
			end if
		end if
	end if
	If Request.form("competitorId") <> nil Then
		if isnumeric(trim(Request.form("competitorId"))) then		
			if trim(Request.form("competitorId"))<>"0" then
				competitorId = trim(Request.form("competitorId"))
			end if
		end if
	end if
'response.Write(	"<br>selTour=" & Request.form("selTour"))
'response.Write(	"<br>selCompetitor=" & Request.form("selCompetitor"))
'response.Write(	"<br>Start_Date=" & Request.form("Start_Date"))
'response.Write(	"<br>End_Date=" & Request.form("End_Date"))


	
'response.Write(	"<br>tourId=" & tourId)
'response.Write(	"<br>competitoId=" & competitoId)
'response.Write(	"<br>startDate=" & startDate)
'response.Write(	"<br>endDate=" & endDate)

	If Request.querystring("tourId") <> nil Then
		if isnumeric(trim(Request.querystring("tourId"))) then
			tourId = trim(Request.querystring("tourId"))
		end if
	end if
	If Request.querystring("competitorId") <> nil Then
		if isnumeric(trim(Request.querystring("competitorId"))) then
			competitorId = trim(Request.querystring("competitorId"))
		end if
	end if

	
'response.Write(	"<br>screenId=" & Request.querystring("screenId"))
'response.Write(	"<br>tourId=" & Request.querystring("tourId"))
'response.Write(	"<br>competitorId=" & Request.querystring("competitorId"))
'response.Write(	"<br>startDate=" & Request.querystring("startDate"))
'response.Write(	"<br>endDate=" & Request.querystring("endDate"))


	
'response.Write(	"<br>screenId=" & screenId)
'response.Write(	"<br>tourId=" & tourId)
'response.Write(	"<br>competitorId=" & competitorId)
'response.Write(	"<br>startDate=" & startDate)
'response.Write(	"<br>endDate=" & endDate)

		dir_var = "ltr"  
		  align_var = "right"  
		    dir_obj_var = "rtl"
	
	'if tourId<>"" or competitorId<>"" then %>
	
		<table width="100%" cellspacing="1" cellpadding="2" border="0" bgcolor="#ffffff" ID="Table3">
										<%		'get tours' names from Pegasusisrael site
									sqlstr="SELECT distinct Compare_Screens.Tour_Id, " & pegasusDBName & ".dbo.Tours.Tour_Name, "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours_Categories.Category_Name,  " & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Name, "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									sqlstr=sqlstr & " FROM   Compare_Screens INNER JOIN "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours ON Compare_Screens.Tour_Id = " & pegasusDBName & ".dbo.Tours.Tour_Id "
									sqlstr=sqlstr & " INNER JOIN " & pegasusDBName & ".dbo.Tours_Categories ON " & pegasusDBName & ".dbo.Tours.Category_Id = " & pegasusDBName & ".dbo.Tours_Categories.Category_Id INNER JOIN "
										sqlstr=sqlstr & pegasusDBName & ".dbo.Tours_SubCategories ON " & pegasusDBName & ".dbo.Tours.SubCategory_Id = " & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Id "
										'SEARCH
										if tourId<>"" then
											sqlstr=sqlstr & " and Compare_Screens.Tour_Id=" & tourId
										end if
										if competitorId<>"" then
											sqlstr=sqlstr & " and Compare_Screens.Competitor_Id=" & competitorId
										end if
									sqlstr=sqlstr & " and datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
										sqlstr=sqlstr & "order by " & pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									set rs_Tours = con.GetRecordSet(sqlstr)
									do while not rs_Tours.eof
    										Tour_Id = trim(rs_Tours("Tour_Id"))
												Category_Name = trim(rs_Tours("Category_Name"))
												SubCategory_Name = trim(rs_Tours("SubCategory_Name"))
												Tour_Name = trim(rs_Tours("Tour_Name"))
											
										%>
										<tr>
								<td  align="<%=align_var%>" class="titleList">
									<b>
										<%=Category_Name%>
										- <font color="navy">
											<%=Tour_Name%>
										</font></b>&nbsp;</td>
							</tr>
							<%
											sqlstr= "Select Screen_Id,Competitor_Name,Start_Date,End_Date  from Compare_Screens "
											sqlstr=sqlstr & "  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id where Tour_Id=" & Tour_Id 
											'SEARCH
											if competitorId<>"" then
												sqlstr=sqlstr & " and Compare_Screens.Competitor_Id=" & competitorId
											end if
											sqlstr=sqlstr & " and datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
											sqlstr=sqlstr & " order by Competitor_Name,End_Date desc, Start_Date desc"
											set rs_Comp = con.GetRecordSet(sqlstr)
											do while not rs_Comp.eof
    											Screen_Id = trim(rs_Comp("Screen_Id"))
												Competitor_Name = trim(rs_Comp("Competitor_Name"))
												Start_Date = trim(rs_Comp("Start_Date"))
												End_Date = trim(rs_Comp("End_Date"))
										%>
							<tr>
								<td align="<%=align_var%>"  class="itemList">
									<a href="viewScreen_popup.asp?ScreenId=<%=Screen_Id%>" target="_blank" class="itemList">
										<b>
											<%=Competitor_Name%>
										</b></a></td>
							</tr>
							<%rs_Comp.MoveNext
										loop
										rs_Comp.close
										set rs_Comp=nothing
									rs_Tours.MoveNext
								loop
							set rs_Tours=nothing%>
						</table>
<%	'End If
%>