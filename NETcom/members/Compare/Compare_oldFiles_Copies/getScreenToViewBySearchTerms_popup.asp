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
			dir_var = "ltr"  
		  align_var = "right"  
		    dir_obj_var = "rtl"
		    
		'get screen  
		Screen_Id=0
		if  tourId<>"" and competitorId<>"" then
		
			sqlstr= "Select top 1 Screen_Id,Tour_Id,Competitor_Name from Compare_Screens  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id "
				sqlstr=sqlstr & " where Tour_Id=" & tourId & " and Compare_Screens.Competitor_Id=" & competitorId
			sqlstr=sqlstr & " and datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
			sqlstr=sqlstr & " order by End_Date desc, Start_Date desc"
			set rs_Comp = con.GetRecordSet(sqlstr)
			if not rs_Comp.eof then
				Screen_Id=rs_Comp("Screen_Id")
				Tour_Id=rs_Comp("Tour_Id")
				Competitor_Name=rs_Comp("Competitor_Name")
			end if
			rs_Comp.close
			set rs_Comp=nothing
			
			if Tour_Id<>0 then
				sqlstr= " Select Tour_Name from  Tours  where Tour_Id=" & Tour_Id
				set rs_Comp = conPegasus.Execute(sqlstr)
				if not rs_Comp.eof then
					Tour_Name=rs_Comp("Tour_Name")
				end if
				rs_Comp.close
				set rs_Comp=nothing
			end if
			
			if Screen_Id>0 then%>
<table width="780" cellspacing="1" cellpadding="2" align="center" border="0" ID="Table3">
	<tr>
		<td  dir="<%=dir_obj_var%>" class="form_header"><h3><div id="tourTitle"><%=Tour_Name%></div>
			</h3>
		</td>
	</tr>
	<tr>
		<td align="right" nowrap bgcolor="#ffffff" valign="top" colspan="2">
			<table border="0" cellpadding="7" cellspacing="3" ID="Table6" dir="<%=dir_obj_var%>" width="100%" style="border-collapse:separate">
				<tr>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>">
						השירות בטיול</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" style="background-color:#99cc00;min-width:100px">
						פגסוס</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" style="background-color:#ffcc33;min-width:100px">
						<!--מתחרה--><div id="competitorTitle"><%=Competitor_Name%></div>
					</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>">
						&nbsp;יתרון?&nbsp;</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>">
						הסבר על היתרון</th>
				</tr>
				<%'get parameters
												dim numParameters
												numParameters=0
												dim numFixedParameters
												numFixedParameters=0
													sqlstr="Select  Parameter_Id, Parameter_Name, Field_Type, isFixedName,Parameter_Icon from Compare_Parameters order by Parameter_Order"
													set rssub = con.getRecordSet(sqlstr)
													
												isVariabledParameters=false
												do while not rssub.eof
													numParameters=numParameters + 1
													
													Parameter_Id = cint(rssub("Parameter_Id"))
													Parameter_Name = trim(rssub("Parameter_Name"))
													Field_Type = cint(rssub("Field_Type"))
													isFixedName = cbool(rssub("isFixedName"))
													Parameter_Icon=  trim(rssub("Parameter_Icon"))
													if Screen_Id>0 then
														sqlstr="SELECT  Parameter_Value, Parameter_Advantages, Advantages_Description "
														sqlstr=sqlstr & " FROM  Compare_Parameters_ToScreen where Screen_Id=" & Screen_Id & " and Parameter_Id=" & Parameter_Id
														set rssubSCR = con.getRecordSet(sqlstr)
														if not	rssubSCR.eof then
															Parameter_Value_Screen=	trim(rssubSCR("Parameter_Value"))
															Parameter_Advantages=	trim(rssubSCR("Parameter_Advantages"))	
															isAdvantages=cbool(Parameter_Advantages)
															Advantages_Description=	trim(rssubSCR("Advantages_Description"))
															isParameterInScreen=true												
														else
															Parameter_Value_Screen=""	
															Parameter_Advantages=""	
															isAdvantages=true	
															Advantages_Description=""										
															isParameterInScreen=false
														end if		
														rssubSCR.close
														set rssubSCR =nothing	
													else								
														isParameterInScreen=false
													end if
													Parameter_Value_Tour=""
														Parameter_Name_Variabled=""
													if Tour_Id>0  then
														sqlstr="SELECT  Parameter_Value, Parameter_Name_Variabled  FROM  Compare_Parameters_ToTour where Tour_Id=" & Tour_Id & " and Parameter_Id=" & Parameter_Id
														set rssubTR = con.getRecordSet(sqlstr)	
														if not	rssubTR.eof then
															Parameter_Value_Tour=	trim(rssubTR("Parameter_Value"))
															if not isFixedName then
																Parameter_Name_Variabled=	trim(rssubTR("Parameter_Name_Variabled"))
															end if
														end if
														rssubTR.close
														set rssubTR =nothing			
													end if
													%>
				<%if Parameter_Value_Screen<>"" or Advantages_Description<>"" then%>
				<tr style="background-color:#f5f5f5" onmouseover="this.style.backgroundColor ='#e5e5e5'"
					onmouseout="this.style.backgroundColor='#f5f5f5'">
					<td>
						<table width="100%" cellpadding="0" cellspacing="0" ID="Table1">
							<tr>
								<td align="right" dir="rtl" <%if Parameter_Advantages<>"" and isAdvantages then%> style="color:green"<%else%> style="color:red"<%end if%>>
									<%=Parameter_Icon%>
								</td>
								<td align="right" dir="rtl">
									<%if isFixedName then%>
									<%=Parameter_Name%>
									<%else%>
									<%=Parameter_Name_Variabled%>
									<%end if%>
								</td>
							</tr>
						</table>
					</td>
					<td align="<%=align_var%>"  dir="<%=dir_obj_var%>"   <%if Parameter_Advantages<>"" and isAdvantages then%> style="background-color:#f1ffef"<%else%> style="background-color:#fff0e2"<%end if%>>
						<%=Parameter_Value_Tour%>
					</td>
					<td align="<%=align_var%>"  dir="<%=dir_obj_var%>"  style="background-color:#fffff1">
						<%=Parameter_Value_Screen%>
					</td>
					<td align="center"  <%if Parameter_Advantages<>"" and isAdvantages then%> style="background-color:#e5ffe2"<%else%> style="background-color:#fff0e2"<%end if%>>
						<%if Parameter_Advantages<>"" and isAdvantages then%>
						<i class="fa fa-thumbs-up fa-flip-horizontal" aria-hidden="true" style="color:green">
						</i>
						<%end if%>
						<%if Parameter_Advantages<>"" and not isAdvantages then%>
						<i class="fa fa-thumbs-down fa-flip-horizontal" aria-hidden="true" style="color:red">
						</i>
						<%end if%>
					</td>
					<td align="<%=align_var%>"  dir="<%=dir_obj_var%>" >
						<%=Advantages_Description%>
					</td>
				</tr>
				<%end if%>
				<%rssub.movenext
												loop
												set rssub=Nothing
												
												%>
			</table>
		</td>
	</tr>
</table>
<%end if%>
<%end if%>
