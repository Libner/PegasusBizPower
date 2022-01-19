<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
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
		
			sqlstr= "Select top 1 Screen_Id,Tour_Id,Competitor_Name,Competitor_Logo from Compare_Screens  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id "
				sqlstr=sqlstr & " where Tour_Id=" & tourId & " and Compare_Screens.Competitor_Id=" & competitorId
			sqlstr=sqlstr & " and datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
			sqlstr=sqlstr & " order by End_Date desc, Start_Date desc"
		
			set rs_Comp = con.GetRecordSet(sqlstr)
			if not rs_Comp.eof then
				Screen_Id=rs_Comp("Screen_Id")
				Tour_Id=rs_Comp("Tour_Id")
				Competitor_Name=rs_Comp("Competitor_Name")
				Competitor_Logo=rs_Comp("Competitor_Logo")
				if Competitor_Logo<>"" then
					set fso=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
					fileString= Server.MapPath("../../../download/competitors/"& Competitor_Logo ) 'deleting of existing file
					if not fso.FileExists(fileString) then
						Competitor_Logo=""
					end if
				end if	
			end if
			rs_Comp.close
			set rs_Comp=nothing
			
			strUrlTour=""
			if TourId<>0 then
				sqlstr= " Select Tour_Name,rewrite_friendlyUrl from  Tours  where Tour_Id=" & TourId
				set rs_Comp = conPegasus.Execute(sqlstr)
				if not rs_Comp.eof then
					Tour_Name=rs_Comp("Tour_Name")
					rewrite_friendlyUrl=rs_Comp("rewrite_friendlyUrl")					
				    If trim(rewrite_friendlyUrl)<>"" Then
						strUrlTour = Application("SiteUrl") & "/tour/" & Replace(Trim(rewrite_friendlyUrl), " ", "_") & ".aspx"
					else
                        strUrlTour = Application("SiteUrl") & "/tours/tour.aspx?tour=" & Trim(tourId)
                    End If
				end if
				rs_Comp.close
				set rs_Comp=nothing
			end if
			%>
<%if strUrlTour<>"" then%>
<table border="0" width="100%" cellpadding="0" cellspacing="3" dir="ltr" style="margin-right:10px"
	ID="Table2">
	<tr>
		<td colspan="2" align="<%=align_var%>" dir="<%=dir_obj_var%>"  class="tourTitle" style="color:navy">
			<a href="<%=strUrlTour%>" target="_blank" title="לדף הטיול">
				<%=Tour_Name%>
			</a>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="5"></td>
	</tr>
</table>
<%end if%>
<%if Screen_Id>0 then%>
<table border="0" cellpadding="7" cellspacing="3" ID="Table6" dir="<%=dir_obj_var%>" width="100%" style="border-collapse:separate"  style="background-color:#ffffff">
	<tr>
		<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
			השירות בטיול</th>
		<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam">
			פגסוס</th>
		<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
			<!--מתחרה--><div id="competitorTitle"><%=Competitor_Name%></div>
		</th>
		<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
			&nbsp;יתרון?&nbsp;</th>
		<th  dir="<%=dir_obj_var%>" align="<%=align_var%>"class="thParam"  width="410" nowrap>
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
															if Parameter_Advantages<>"" then
																isAdvantages=cbool(Parameter_Advantages)
															else
																isAdvantages=true
															end if
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
	<%if Parameter_Value_Tour<>"" or Parameter_Value_Screen<>"" or Parameter_Advantages<>"" or Advantages_Description<>"" then%>
	<tr onmouseover="markRow($(this))" onmouseout="cancelRowMarking($(this))">
		<td class="paramTitle" dir="ltr" nowrap="true">
			<table cellpadding="0" cellspacing="0" ID="Table1">
				<tr>
					<td class="paramName" align="right" dir="rtl" width="150" nowrap="true">
						<span style="white-space:normal">
						<%if isFixedName then%>
						<%=Parameter_Name%>
						<%else%>
						<%=Parameter_Name_Variabled%>
						<%end if%>
						</span>
					</td>
					<td class="paramIcon <%if Parameter_Advantages<>"" then%><%if isAdvantages then%>greenFNT<%else%>redFNT<%end if%><%end if%>" align="right" dir="rtl" width="30" nowrap=true >
						<%=Parameter_Icon%>
					</td>
				</tr>
			</table>
		</td>
		<td id="ParamValTour_<%=Parameter_Id%>" width="50%" class="paramVal  <%if Parameter_Advantages<>"" then%><%if isAdvantages then%>greenBG<%else%>redBG<%end if%><%end if%>" isAdvantages="<%=isAdvantages%>" align="<%=align_var%>"  dir="<%=dir_obj_var%>" >
			<%=Parameter_Value_Tour%>
		</td>
		<td id="ParamValScreen_<%=Parameter_Id%>" width="50%" class="paramVal <%if Parameter_Advantages<>"" then%><%if not isAdvantages then%>greenBG<%else%>redBG<%end if%><%end if%>" isAdvantages="<%=not (isAdvantages)%>" align="<%=align_var%>"  dir="<%=dir_obj_var%>" >
			<%=Parameter_Value_Screen%>
		</td>
		<td class="paramAdvantages" align="center">
			<%if Parameter_Advantages<>"" and isAdvantages then%>
			<i class="fa fa-thumbs-up fa-flip-horizontal <%if Parameter_Advantages<>"" then%><%if isAdvantages then%>greenFNT<%else%>redFNT<%end if%><%end if%>" aria-hidden="true" style="font-size:24px;">
			</i>
			<%end if%>
			<%if Parameter_Advantages<>"" and not isAdvantages then%>
			<i class="fa fa-thumbs-down fa-flip-horizontal <%if Parameter_Advantages<>"" then%><%if isAdvantages then%>greenFNT<%else%>redFNT<%end if%><%end if%>" aria-hidden="true" style="font-size:24px;">
			</i>
			<%end if%>
		</td>
		<td class="paramDesc" align="<%=align_var%>"  dir="<%=dir_obj_var%>" width="410" nowrap>
			<!--<div id="Desc_<%=Parameter_Id%>" style="position:relative;max-width:400px;min-height:28px;height:28px;overflow-y:hidden;white-space:normal;padding-left:10px"
																	onmouseover="openDesc(this)"><%=breaks(Advantages_Description)%>
																		<%if Advantages_Description<>"" then%>
																		<div id="iconDescOpen_<%=Parameter_Id%>" style="position:absolute;bottom:0px;left:0px;margin-bottom:3px"><i  class="fa fa-chevron-circle-down" aria-hidden="true" style="color:rgba( 0, 0, 0, 0.3 )"></i></div>
																		<%end if%>
																		<%if Advantages_Description<>"" then%>
																		<a id="iconDescClose_<%=Parameter_Id%>" style="position:absolute;bottom:0px;left:0px;display:none" href="" onclick="hideDesc($(this).parent());return false"><i  class="fa fa-chevron-circle-up" aria-hidden="true" style="color:rgba( 0, 0, 0 )"></i></a>
																		<%end if%>
																</div>
																-->
			<div style="position:relative;width:410px;max-width:410px;">
				<div id="paramDesc_<%=Parameter_Id%>" 
																	style="position:relative;width:400px;max-width:400px;min-height:26px;height:26px;overflow-y:hidden;white-space:normal;z-index:1">
					<%=breaks(Advantages_Description)%>
				</div>
				<%if Advantages_Description<>"" then%>
				<div class="moreDesc" style="position:absolute;width:auto;height:auto;bottom:-1px;left:-1px;white-space:normal;">...</div>
				<%end if%>
				<%if Advantages_Description<>"" then%>
				<div id="paramDescAll_<%=Parameter_Id%>" class="plata_description"
																	style="position:absolute;width:420px;height:auto;top:-10px;left:-10px;white-space:normal;padding-left:10px;display:none;z-index:1000"><%=breaks(Advantages_Description)%>
				</div>
				<%end if%>
			</div>
		</td>
	</tr>
	<%end if%>
	<%rssub.movenext
												loop
												set rssub=Nothing
												
												%>
</table>
<%end if%>
<%end if%>
<%if Competitor_Logo<>"" then%>
<script>
$("#CompetitorLogo").attr("src","../../../download/competitors/<%=Competitor_Logo%>")
$("#CompetitorLogo").css("display","block")
</script>
<%else%>
<script>
$("#CompetitorLogo").attr("src","")
$("#CompetitorLogo").css("display","none")
</script>
<%end if%>
