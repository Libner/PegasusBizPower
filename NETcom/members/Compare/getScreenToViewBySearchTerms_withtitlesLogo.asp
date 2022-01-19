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
<table cellspacing="1" cellpadding="2" width="100%" align="center" border="0" ID="Table3">
	<tr>
		<td  dir="<%=dir_obj_var%>"  align="<%=align_var%>">
			<table border="0" width="100%" cellpadding="0" cellspacing="0" dir="ltr" style="padding-right:10px">
				<tr>
					<td colspan="2" align="<%=align_var%>" dir="<%=dir_obj_var%>"  class="tourTitle" style="color:navy"><%=Tour_Name%></td>
				</tr>
				<tr>
					<td colspan="2" height="5"></td>
				</tr>
				<tr>
					<td  dir="<%=dir_obj_var%>" width="100%" align="<%=align_var%>" class="tourTitle"><%=Competitor_Name%></td>
					<td align="<%=align_var%>"><%if Competitor_Logo<>"" then%>
						<img src="../../../download/competitors/<%=Competitor_Logo%>" border="0"  style="margin-left:10px;max-height:40px">
						<%end if%>
					</td>
				</tr>
				<tr>
					<td colspan="2" height="5"></td>
				</tr>
			
			</table>
		</td>
	</tr>
	<tr>
		<td align="right" nowrap bgcolor="#ffffff" valign="top" colspan="2">
			<table border="0" cellpadding="7" cellspacing="3" ID="Table6" dir="<%=dir_obj_var%>" width="100%" style="border-collapse:separate">
				<tr>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
						השירות בטיול</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" style="background-color:#99cc00;">
						פגסוס</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" style="background-color:#ffcc33;">
						<!--מתחרה--><div id="competitorTitle"><%=Competitor_Name%></div>
					</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
						&nbsp;יתרון?&nbsp;</th>
					<th  dir="<%=dir_obj_var%>" align="<%=align_var%>"class="thParam" >
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
								<td  class="paramName"align="right" dir="rtl" <%if Parameter_Advantages<>"" and isAdvantages then%> style="color:green"<%else%> style="color:red"<%end if%>>
									<%=Parameter_Icon%>
								</td>
								<td class="paramName" align="right" dir="rtl">
									<%if isFixedName then%>
									<%=Parameter_Name%>
									<%else%>
									<%=Parameter_Name_Variabled%>
									<%end if%>
								</td>
							</tr>
						</table>
					</td>
					<td class="paramVal" align="<%=align_var%>"  dir="<%=dir_obj_var%>"   <%if Parameter_Advantages<>"" and isAdvantages then%> style="background-color:#f1ffef"<%else%> style="background-color:#fff0e2"<%end if%>>
						<%=Parameter_Value_Tour%>
					</td>
					<td class="paramVal" align="<%=align_var%>"  dir="<%=dir_obj_var%>"  style="background-color:#fffff1">
						<%=Parameter_Value_Screen%>
					</td>
					<td class="paramVal" align="center"  <%if Parameter_Advantages<>"" and isAdvantages then%> style="background-color:#e5ffe2"<%else%> style="background-color:#fff0e2"<%end if%>>
						<%if Parameter_Advantages<>"" and isAdvantages then%>
						<i class="fa fa-thumbs-up fa-flip-horizontal" aria-hidden="true" style="font-size:24px;color:green">
						</i>
						<%end if%>
						<%if Parameter_Advantages<>"" and not isAdvantages then%>
						<i class="fa fa-thumbs-down fa-flip-horizontal" aria-hidden="true" style="font-size:24px;color:red">
						</i>
						<%end if%>
					</td>
					<td class="paramDesc" align="<%=align_var%>"  dir="<%=dir_obj_var%>" >
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
						<div style="position:relative">
							<div id="paramDesc_<%=Parameter_Id%>" 
																	style="position:relative;max-width:400px;min-height:28px;height:28px;overflow-y:hidden;white-space:normal;padding-left:10px;z-index:1"
																	onmouseover="openDesc($(this))"><%=breaks(Advantages_Description)%>
							</div>
							<%if Advantages_Description<>"" then%>
							<div id="paramDescAll_<%=Parameter_Id%>" class="plata_description"
																	style="position:absolute;width:420px;height:auto;top:-10px;left:-10px;white-space:normal;padding-left:10px;display:none;z-index:1000"
																	onmouseout="hideDesc($(this))"><%=breaks(Advantages_Description)%>
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
		</td>
	</tr>
</table>
<%end if%>
<%end if%>
