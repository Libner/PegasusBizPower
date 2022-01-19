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

	If Request.querystring("stourId") <> nil Then
		if isnumeric(trim(Request.querystring("stourId"))) then
			tourId = trim(Request.querystring("stourId"))
		end if
	end if
	If Request.querystring("scompetitorId") <> nil Then
		if isnumeric(trim(Request.querystring("scompetitorId"))) then
			competitorId = trim(Request.querystring("scompetitorId"))
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
										<tr>
											<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word3" name="word3">מחיקה</span>&nbsp;</td>
											<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name="word4">עדכון</span>&nbsp;</td>
											<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="Span1" name="word4">עד תאריך</span>&nbsp;</td>
											<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="Span2" name="word4">מתאריך</span>&nbsp;</td>
											<td width="100%"  align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name="word5"><!--מתחרה-->מתחרה</span>&nbsp;</td>
										</tr>
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
										
										sqlstr=sqlstr & "order by " & pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									set rs_Tours = con.GetRecordSet(sqlstr)
									do while not rs_Tours.eof
    										Tour_Id = trim(rs_Tours("Tour_Id"))
												Category_Name = trim(rs_Tours("Category_Name"))
												SubCategory_Name = trim(rs_Tours("SubCategory_Name"))
												Tour_Name = trim(rs_Tours("Tour_Name"))
											
										%>
										<tr>
											<th nowrap colspan="2" align="center" bgcolor="#e6e6e6">
															<a class="button_edit_1" style="width:100;" href="addScreen.asp?tourId=<%=Tour_Id%>">
																<span id="Span3" name="word6"><!--הוספת מתחרה-->הוספת מסך השוואה</span></a></th>
														<th colspan="3" align="<%=align_var%>" bgcolor="#e6e6e6">
															<b><%=Tour_Name%>
															</b>&nbsp;</th>
										</tr>
										<%
											sqlstr= "Select Screen_Id,Competitor_Name,Start_Date,End_Date  from Compare_Screens "
											sqlstr=sqlstr & "  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id where Tour_Id=" & Tour_Id 
											'SEARCH
											if competitorId<>"" then
												sqlstr=sqlstr & " and Compare_Screens.Competitor_Id=" & competitorId
											end if
											 sqlstr=sqlstr & " order by Competitor_Name,End_Date desc, Start_Date desc"
											set rs_Comp = con.GetRecordSet(sqlstr)
											do while not rs_Comp.eof
    											Screen_Id = trim(rs_Comp("Screen_Id"))
												Competitor_Name = trim(rs_Comp("Competitor_Name"))
												Start_Date = trim(rs_Comp("Start_Date"))
												End_Date = trim(rs_Comp("End_Date"))
										%>
										<tr>
											<td align="center" bgcolor="#e6e6e6" nowrap><a href="screensAdmin.asp?deleteId=<%=Screen_Id%><%if TourID<>"" then%>&stourId=<%=TourID%><%end if%><%if competitorId<>"" then%>&scompetitorId=<%=competitorID%><%end if%>" ONCLICK="return checkDelete(<%=countComp %>)"><IMG SRC="../../images/delete_icon.gif" BORDER="0"></a></td>
											<td align="center" bgcolor="#e6e6e6" nowrap><a href="addScreen.asp?TourID=<%=Tour_Id%>&screenID=<%=Screen_Id%><%if TourID<>"" then%>&stourId=<%=TourID%><%end if%><%if competitorId<>"" then%>&scompetitorId=<%=competitorID%><%end if%>"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>
											<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap <%if datediff("d",date(),Start_Date)>0 or datediff("d",date(),End_Date)<0 then%>style="color:#a9a9a9"<%end if%>><b><%= cdate(End_Date)%></b>&nbsp;</td>
											<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap <%if datediff("d",date(),Start_Date)>0 or datediff("d",date(),End_Date)<0 then%>style="color:#a9a9a9"<%end if%>><b><%=cdate(Start_Date)%></b>&nbsp;</td>
											<td align="<%=align_var%>" bgcolor="#e6e6e6" width="100%" <%if datediff("d",date(),Start_Date)>0 or datediff("d",date(),End_Date)<0 then%>style="color:#a9a9a9"<%end if%>><b><%=Competitor_Name%></b>&nbsp;</td>
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