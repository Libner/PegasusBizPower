<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%DepartureId = trim(Request.QueryString("dID"))
	urlSort="GuideMessages.asp?DeparureId="&DepartureId

	

	
	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_defaultClientScript content="JavaScript">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name=ProgId content=VisualStudio.HTML>
<meta name=Originator content="Microsoft Visual Studio .NET 7.1">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
	function checkAddTasks()
	{
		
	}	
	function report_open(fname){
		form1.action=fname;
		form1.target = '_blank';
		form1.submit();
	
}
//-->
</script>

</head>
<body style="margin: 0px;" onload="self.focus();"  >

<%
DepartureId=request.QueryString("dID")
 sort_app = Request.QueryString("sort_app")	
 if trim(sort_app)="" then  sort_app=7 end if  

urlSort="GuideMessages.asp?1=1&dID="&DepartureId '?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;Message_status="&Message_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;MessagetypeID=" & MessagetypeID & "&amp;T=" & T
dim sortby(14)	
sortby(0) = "Departure_Code"
sortby(1) = "Departure_Code desc"
sortby(2)="Departure_Code"
sortby(3) = "Departure_Code desc"


sortby(5) = "Guide_LName"
sortby(4) = "Guide_LName Desc"

sortby(6) = "Messages_Date"
sortby(7) = "Messages_Date Desc"
sortby(8) = "convert(datetime,Messages_Date,8)"
sortby(9) = "convert(datetime,Messages_Date,8) Desc"

sortby(10) = "LASTNAME"
sortby(11) = "LASTNAME Desc" 'user name
	'sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','" & sortby_app(sort_app) & "','','','','" & DeparureId & "','','','" & productID & "','" & UserID & "','','','" & is_groups & "'"
  	'sqlstr = "Exec dbo.[get_GuideMessagesByDep] '1','100','','"& DepartureId  &"','', '1'"
  		sqlstr = "Exec dbo.[get_GuideMessages] '1','100','','"& DepartureId &"','', '" & sortby(sort_app) & "','"& StrSeriasId  & "','"& pSeries & "','"& pUsers & "','"& pGuide  & "','"& FromDate & "','"& ToDate &"'"


    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount	
	if Request("page_app")<>"" then
		page_app=request("page_app")
	else
		page_app=1
	end if
	if not app.eof then
		app.PageSize = 15
		app.AbsolutePage=page_app
		recCount=app.RecordCount 		
		NumberOfPagesApp = app.PageCount		
		i=1
		j=0
		ids = "" 'list of appeal_id
	end if
	if not app.eof Or search = true then %>
  <form id="form1" name="form1" action="contact_appeals.asp?DeparureId=<%=DeparureId%>" method="post">
  <table cellpadding="0" cellspacing="0" dir="<%=dir_var%>" width="100%" border="0" ID="Table1">  
  <tr><td width="100%"><A name="table_appeals"></A>  
  <table cellpadding=0 cellspacing=0 width="100%" dir="<%=dir_var%>" border="0" ID="Table2">
  <tr >
  <td align =left><a href="#" onclick="report_open('reportExcelMessagesByDepId.asp?dID=<%=DepartureId%>');" class="button_edit_1" style="direction:rtl;width:95;background-color:#FFD011;color:#000000;width:150px">דו"ח EXCEL שיחות</a></td>
    <td valign="top" class="title_form" width="100%" align="<%=align_var%>" dir=<%=dir_obj_var%>>&nbsp;&nbsp;<font color="#E6E6E6">תיעוד שיחות</font>&nbsp;</td>
  </tr>
 <tr><td height=20 colspan=2></td></tr>
  </table></td></tr>  	  	
  <tr>
	<td width="100%" align="center" valign="top">
	<table BORDER="0" CELLSPACING="2" CELLPADDING="2" bgcolor="#FFFFFF" dir="<%=dir_var%>" ID="Table4" width="80%">	
    <tr style="height:35px">
	    <td width="100%" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center">פרוט</td>
			<td width="110"  nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=10 or sort_app=11 then%>_act<%end if%>"><a class="title_sort" name=word25 title="מיון" HREF="<%=urlSort%>&sort_app=<%if sort_app=11 then%>10<%elseif sort_app=10 then%>11<%else%>10<%end if%>" target="_self">&nbsp;שם משתמש&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="10" then%>bot<%elseif trim(sort_app)="11" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
	
		<td width="110"  nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=8 or sort_app=9 then%>_act<%end if%>"><a class="title_sort" name=word25 title="מיון" HREF="<%=urlSort%>&sort_app=<%if sort_app=9 then%>8<%elseif sort_app=8 then%>9<%else%>9<%end if%>" target="_self">&nbsp;<!--עובד-->שעת השיחה&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="8" then%>bot<%elseif trim(sort_app)="9" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=6 or sort_app=7 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=7 then%>6<%elseif sort_app=6 then%>7<%else%>7<%end if%>#table_appeals" target="_self">&nbsp;תאריך שיחה&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="6" then%>bot<%elseif trim(sort_app)="7" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>				
		<td width="80" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=4 or sort_app=5 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=5 then%>4<%elseif sort_app=4 then%>5<%else%>5<%end if%>#table_appeals" target="_self">&nbsp;מדריך&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="4" then%>bot<%elseif trim(sort_app)="5" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="80"  align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=2 or sort_app=3 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=3 then%>2<%elseif sort_app=2 then%>3<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<!--תאריך-->קוד טיול&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="2" then%>bot<%elseif trim(sort_app)="3" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>			
		<td width="80" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=0 or sort_app=1 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=1 then%>0<%elseif sort_app=0 then%>1<%else%>1<%end if%>#table_appeals" target="_self">&nbsp;<!--'סט-->תאריך&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="0" then%>bot<%elseif trim(sort_app)="1" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="40" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center">טיול</td>
	</tr>
	<%Do while (Not app.eof And j<app.PageSize)
		If j Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If		
		Messages_Id = trim(app("Messages_Id"))
				
		Departure_Code = app("Departure_Code")
	Departure_Date= app("Departure_Date")
		Messages_Date = app("Messages_Date")
		Series_Name=app("Series_Name")
		Messages_Content = app("Messages_Content")
		GuideName = app("GuideName")
User_Name = app("User_Name")
	'	if len(Messages_Content) > 30 then
	'		Messages_Content = left(Messages_Content,27) & "..."		
	'	end if
	

	%>
		<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
	    <td width="40" nowrap  dir="rtl" align="right"><%=Messages_Content%></td>
		<td width="60" nowrap  dir="rtl" align="center"><%=User_Name%></td>
		<td width="95"  nowrap align="center" >&nbsp;<%if IsDate(Messages_Date) then%><%=Hour(Messages_Date)%>:<%= Right("00" & Minute(Messages_Date), 2)%><%end if%>&nbsp;</td>	
		<td align="<%=align_var%>" nowrap>&nbsp;<%=FormatDateTime(Messages_Date,2)%>&nbsp;</td>				
		<td width="50" nowrap align="center" dir="rtl"  nowrap><%=GuideName%></td>
		<td width="60"  align="center"  nowrap>&nbsp;<%=Departure_Code%>&nbsp;</td>			
		<td width="50" nowrap align="center" nowrap><%= Right("0" & Month(Departure_Date), 2) & right("0" & Day(Departure_Date), 2) %>&nbsp;</td>
		<td width="20" nowrap  align="center"><%=Series_Name%></td>
		</tr>
<%		app.movenext
		j=j+1
		if not app.eof and j <> app.PageSize then
		ids = ids & ","
		end if
		loop 
		%>
		</table>		
				
		<input type="hidden" name="DeparureId" value="0" ID="DeparureId">			
		</td></tr>		
		<% if cInt(NumberOfPagesApp) > 1 then%>
		<tr class="card">
		<td width="100%" align="center" nowrap class="card" dir=ltr>
			<table border="0" cellspacing="0" cellpadding="2" ID="Table5">               
	        <% If NumberOfPagesApp > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPagesApp / 10)
	           else num = NumberOfPagesApp : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowApp") <> nil Then
	               numOfRowApp = Request.QueryString("numOfRowApp")
	           Else numOfRowApp = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRowApp <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp-1)-9%>&numOfRowApp=<%=numOfRowApp-1%>#table_appeals" name=word56 title="<%=arrTitles(56)%>" target="_self"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowApp-1)) <= NumberOfPagesApp Then
	                  if CInt(page_app)=CInt(i+10*(numOfRowApp-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowApp-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=i+10*(numOfRowApp-1)%>&numOfRowApp=<%=numOfRowApp%>#table_appeals" target="_self"><%=i+10*(numOfRowApp-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesApp > cint(num * numOfRowApp) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp) + 1%>&numOfRowApp=<%=numOfRowApp+1%>#table_appeals" name=word57 title="<%=arrTitles(57)%>" target="_self"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<%End If%>	
	<%If app.recordCount = 0 Then%>
	<tr><td align="center" class="card1">&nbsp;</td></tr>									
	<%End If%>		
	</form>		 			 
	</table>
	<%Else%><p class="titlep" align="center" >לא נמצאו שיחות מצורפים לטיול</p>
	<%End If%>	
<%set app = nothing	%>
</body>
</html>
<%Set con = Nothing%>