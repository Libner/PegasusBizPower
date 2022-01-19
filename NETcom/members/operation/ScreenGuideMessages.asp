<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->

<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../../tooltip.js"></script>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
	<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
<script>
function openSendMessage()
{
newWin=window.open("AddMessage.asp", "newWin", "toolbar=0,menubar=0,width=600,height=500,top=100,left=5,scrollbars=auto");
	newWin.opener=window
	}
function ClearAll()
{window.location ="ScreenGuideMessages.asp"}
function Search()
{
form1.submit()
}


</script>
</head>

<body>
<div id="ToolTip"></div>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>" ID="Table1">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
<%
qrystring = Request.ServerVariables("QUERY_STRING")

if IsNumeric(Request.Form("selUsers")) and Request.Form("selUsers")>0 then
pUsers=Request.Form("selUsers")
elseif IsNumeric(Request.QueryString("pUsers")) then
pUsers=Request.QueryString("pUsers")
else
pUsers=0
end if
if IsNumeric(Request.Form("selGuide"))and Request.Form("selGuide")>0 then
pGuide=Request.Form("selGuide")
elseif IsNumeric(Request.QueryString("pGuide")) then
 pGuide=Request.QueryString("pGuide")
else
pGuide=0
end if
if IsNumeric(Request.Form("selSerias")) and Request.Form("selSerias")>0 then
pSeries=Request.Form("selSerias")
elseif IsNumeric(Request.QueryString("pSeries")) then
pSeries=Request.QueryString("pSeries")
else
pSeries=0
end if

  If Request.Form("codeTiul") <> "" Then
            pcodeTiul = Request.Form("codeTiul")
        elseif Request.QueryString("pcodeTiul")<>"" then
                  pcodeTiul=Request.QueryString("pcodeTiul")
        Else
            pcodeTiul = ""
        End If
        If Request.Form("dateTiul") <> "" Then
            pdateTiul = Request.Form("dateTiul")
            elseif Request.QueryString ("pdateTiul")<>"" then
             pdateTiul=Request.QueryString ("pdateTiul")
        Else
            pdateTiul = ""
        End If
'response.Write "pdateTiul="&pdateTiul
if Request.Form("sPayFromDate")<>"" then
FromDate=Request.Form("sPayFromDate")
else
FromDate=Request.QueryString("FromDate")
end if
if Request.Form("sPayToDate")<>"" then
ToDate=Request.Form("sPayToDate")
else
ToDate=Request.QueryString("ToDate")
end if

%>

  <%numOftab = 82%>
    <%numOfLink = 1
      If Request.Cookies("bizpegasus")("Chief") <> "1" Then
                sqlstr = "Select Series_Id  from Series where User_Id=" & Request.Cookies("bizpegasus")("UserId")		
	  	set rs_Message = con.getRecordSet(sqlstr)
	
		 do while not rs_Message.EOF
		   If StrSeriasId = "" Then
                    StrSeriasId = rs_Message("Series_Id")
                Else
                    StrSeriasId = StrSeriasId & "," & rs_Message("Series_Id")
                End If
   rs_Message.MoveNext
    loop
	
	set rs_Message = Nothing	 			
            End If

     sort_app = Request.QueryString("sort_app")	
 if trim(sort_app)="" then  
 sort=0
' sort_app=9
  end if  
    urlSort="ScreenGuideMessages.asp?1=1&pSeries="& pSeries &"&amp;pdateTiul="& Server.URLEncode(pdateTiul) &"&amp;pcodeTiul="&Server.URLEncode(pcodeTiul) &"&amp;pGuide="&pGuide &"&amp;pUsers="&pUsers & "&amp;FromDate="&FromDate & "&amp;ToDate=" & ToDate ' & "&amp;sender_id=" & sender_id & "&amp;MessagetypeID=" & MessagetypeID & "&amp;T=" & T
dim sortby(14)	
sortby(0) = "Departure_Code"
sortby(1) = "Departure_Code desc"
sortby(2)="Departure_Code"
sortby(3) = "Departure_Code desc"
sortby(4)="Departure_Code"
sortby(5) = "Departure_Code desc"

sortby(6) = "Guide_LName"
sortby(7) = "Guide_LName Desc"

sortby(8) = "Messages_Date"
sortby(9) = "Messages_Date Desc"

sortby(12) = "LASTNAME"
sortby(13) = "LASTNAME Desc" 'user name
    %>
 
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>		   
<tr><td width=100%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellpadding=0 cellspacing=0 dir="<%=dir_var%>" ID="Table2">
   <tr>    
    <td width="100%" valign="top" align="center">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table3">
    <tr>
    <td bgcolor=#FFFFFF align="left" width="100%" valign=top>
    <table width="100%" cellspacing="1" cellpadding="0" border=0 ID="Table4">      
  <tr><td width=100%>
  <form id="form1" name="form1" action="ScreenGuideMessages.asp?<%=qrystring%>" method="post">
  <table cellpadding="0" cellspacing="0" dir="<%=dir_var%>" width="100%" border="0" ID="Table5">  
	
  <tr>
	<td width="100%" align="right" valign="top">
	<table BORDER="0" CELLSPACING="2" CELLPADDING="2" bgcolor="#FFFFFF" dir="<%=dir_var%>" ID="Table8" width="100%">	
    <tr  style="height:35px;bgcolor:#D8D8D8">
    <td class="title_sort" align=right>
    	<table border=0 cellpadding=2 cellspacing=2 ID="Table7">
    	<tr>
    	<td  class="td_admin_5" align="center" width="70" valign=middle><a href="#" OnClick ="javascript:ClearAll();" Class="button_small1">הצג הכל</a>
    	<td class="td_admin_5" align="center"  valign=bottom><a href="#" OnClick ="javascript:Search();" Class="button_small1">&nbsp;חפש&nbsp;
    	</tr>
    	</table>
    	</td>
    <td class="title_sort">
    <select name="selUsers" id="selUsers" dir="rtl" class="norm" style="width:100%" >
    <option value="0">הכל</option>
    <%set SerUser=con.GetRecordSet("SELECT User_Id,FIRSTNAME +' ' + LASTNAME as  UserName  from Users where ACTIVE=1 order by LASTNAME")
    
    do while not SerUser.EOF
    selUserID=SerUser(0)
    selUserName=SerUser(1)%>
    <option value="<%=selUserID%>" <%if trim(pUsers)=trim(selUserID) then%>selected<%end if%>><%=selUserName%></option>
    <%
    SerUser.MoveNext
    loop
    set SerUser=Nothing
        %>
    </td>
    <td class="title_sort">&nbsp;</td>
    <td class="title_sort"nowrap>	<a href="" onclick="cal1xx.select(document.getElementById('sPayFromDate'),'AsPayFromDate','dd/MM/yy'); return false;"
											id="AsPayFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input type="text" id="sPayFromDate" readonly class="searchList" style="WIDTH:70px"
											NAME="sPayFromDate" value="<%=FromDate%>"> מי  <br>
	<a href="" onclick="cal1xx.select(document.getElementById('sPayToDate'),'AsPayToDate','dd/MM/yy'); return false;"
											id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input  type="text" id="sPayToDate" readonly class="searchList" style="WIDTH:70px"
											NAME="sPayToDate" value="<%=ToDate%>"> עד </td>
    <td class="title_sort"><select name="selGuide" id="selGuide" dir="rtl" class="norm" style="width:100%" id="selGuide">
    <option value="0">הכל</option>
    <%set SerGuide=conPegasus.Execute("SELECT Guide_Id,Guide_FName +' ' + Guide_LName as  GuideName  from Guides where Guide_Vis=1 order by Guide_FName,Guide_LName")
    
    do while not SerGuide.EOF
    selGuideID=SerGuide(0)
    selSGuideName=SerGuide(1)%>
    <option value="<%=selGuideID%>" <%if trim(pGuide)=trim(selGuideID) then%>selected<%end if%>><%=selSGuideName%></option>
    <%
    SerGuide.MoveNext
    loop
    set SerGuide=Nothing
    conPegasus.close()
    %>
    </td>
   	<td class="title_sort"><input type="text" id="codeTiul" name="codeTiul" style="width:82px" value="<%=pcodeTiul%>"></td>
	<td class="title_sort" align=center><input type="text" id="dateTiul" name="dateTiul" style="width:32px"  value="<%=pdateTiul%>"></td>
						
    <td class="title_sort"><select name="selSerias" id="selSerias" dir="rtl" class="norm" style="width:100%" >
    <option value="0">הכל</option>
    <%set SerList=con.GetRecordSet("SELECT Series_Id,Series_Name  from Series order by Series_Name")
    
    do while not SerList.EOF
    selSerID=SerList(0)
    selSerName=SerList(1)%>
    <option value="<%=selSerID%>" <%if trim(pSeries)=trim(selSerID) then%>selected<%end if%>><%=selSerName%></option>
    <%
    SerList.MoveNext
    loop
    set SerList=Nothing
        %>
    </td>
    </tr>
  <!--include table-->
  <%
  ' cmdSelect.Parameters.Add("@pcodeTiul", SqlDbType.VarChar, 100).Value = pcodeTiul
  ' cmdSelect.Parameters.Add("@pdateTiul", SqlDbType.VarChar, 4).Value = pdateTiul



	'sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','" & sortby_app(sort_app) & "','','','','" & DeparureId & "','','','" & productID & "','" & UserID & "','','','" & is_groups & "'"
'  	sqlstr = "Exec dbo.[get_GuideMessages] '1','100','','"& DepartureId &"','', '" & sortby(sort_app) & "','"& StrSeriasId &"'"
  '	sqlstr = "Exec dbo.[get_GuideMessages] '1','100','','"& DepartureId &"','', '" & sortby(sort_app) & "','"& StrSeriasId  & "','"& pSeries & "','"& pUsers & "','"& pGuide  & "','"& FromDate & "','"& ToDate &"'"
  	sqlstr = "Exec dbo.[get_GuideMessages] '1','100','','"& DepartureId &"','', '" & sortby(sort_app) & "','"& StrSeriasId  & "','"& pSeries & "','"& pUsers & "','"& pGuide  & "','"& FromDate & "','"& ToDate &"','"& pcodeTiul &"','"& pdateTiul &"'"
  	
'Response.Write (sortby(sort_app) &":"& sort_app &":"& sort)

 'Response.Write sqlStr
' response.end
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount	
	if Request("page_app")<>"" then
		page_app=request("page_app")
	else
		page_app=1
	end if
	if not app.eof then
		app.PageSize = 50
		app.AbsolutePage=page_app
		recCount=app.RecordCount 		
		NumberOfPagesApp = app.PageCount		
		i=1
		j=0
		ids = "" 'list of appeal_id
	end if
	if not app.eof Or search = true then %>
  
    <tr style="height:35px">
    <!--td  class="title_sort">עדכון </td-->
	    <td width="100%" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right">פרוט</td>
	     <td width="110" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_app)="12" OR trim(sort_app)="13" then%>_act<%end if%>"><%if trim(sort_app)="12" then%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=<%=sort_app+1%>" name=word25 title=""><%elseif trim(sort_app)="13" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=12"  title=""><%end if%><!--תאריך יעד-->שם משתמש<img src="../../images/arrow_<%if trim(sort_app)="12" then%>bot<%elseif trim(sort_app)="13" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	
	
	
	
		<td width="110"  nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<!--עובד-->שעת השיחה&nbsp;</td>	
		 
	    <td width="150" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_app)="8" OR trim(sort_app)="9" then%>_act<%end if%>"><%if trim(sort_app)="8" then%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=<%=sort_app+1%>" name=word25 title=""><%elseif trim(sort_app)="9" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=8"  title=""><%end if%><!--תאריך יעד-->תאריך שיחה<img src="../../images/arrow_<%if trim(sort_app)="8" then%>bot<%elseif trim(sort_app)="9" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	   
	   
	     <td width="80" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_app)="6" OR trim(sort_app)="7" then%>_act<%end if%>"><%if trim(sort_app)="6" then%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=<%=sort_app+1%>" name=word25 title=""><%elseif trim(sort_app)="7" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=6"  title=""><%end if%><!--תאריך יעד-->מדריך<img src="../../images/arrow_<%if trim(sort_app)="6" then%>bot<%elseif trim(sort_app)="7" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	
	
	
	
				  <td width="80" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_app)="4" OR trim(sort_app)="5" then%>_act<%end if%>"><%if trim(sort_app)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=<%=sort_app+1%>" name=word25 title=""><%elseif trim(sort_app)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=4"  title=""><%end if%><!--תאריך יעד-->קוד טיול<img src="../../images/arrow_<%if trim(sort_app)="4" then%>bot<%elseif trim(sort_app)="5" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a>

	
	
		
		  <td width="80" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_app)="2" OR trim(sort_app)="3" then%>_act<%end if%>"><%if trim(sort_app)="2" then%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=<%=sort_app+1%>" name=word25 title=""><%elseif trim(sort_app)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=2"  title=""><%end if%><!--תאריך יעד-->תאריך<img src="../../images/arrow_<%if trim(sort_app)="2" then%>bot<%elseif trim(sort_app)="3" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a>
		  <td width="80" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_app)="0" OR trim(sort_app)="1" then%>_act<%end if%>"><%if trim(sort_app)="0" then%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=<%=sort_app+1%>" name=word25 title=""><%elseif trim(sort_app)="1" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title=""><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort_app=0"  title=""><%end if%><!--תאריך יעד-->טיול<img src="../../images/arrow_<%if trim(sort_app)="0" then%>bot<%elseif trim(sort_app)="1" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	
	
		
		</tr>
	<%Do while (Not app.eof And j<app.PageSize)
		If j Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If		
		Messages_Id = trim(app("Messages_Id"))
				Departure_Id= trim(app("Departure_Id"))
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
	<!--td><a href="#"  onclick="window.open('AddGuideMessages.asp?dID=<%=Departure_Id%>&Messages_Id=<%=Messages_Id%>','winCA','top=20, left=10, width=1000, height=350, scrollbars=1');"" target="_blank"><img src="../../images/edit_icon.gif" border="0" alt=""></a></td-->
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
		<% if cInt(NumberOfPagesApp) > 1 then	   %>
		<tr class="card">
		<td width="100%" align="center" nowrap class="card" dir=ltr>
			<table border="0" cellspacing="0" cellpadding="2" ID="Table9">               
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
			<SCRIPT type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.setYearSelectStartOffset(120);
	cal1xx.setYearSelectEndOffset(2);
	//cal1xx.setYearSelectStart(1910); //Added by Mila
                cal1xx.showNavigationDropdowns();
                cal1xx.offsetX = -50;
                cal1xx.offsetY = -40;
          
                //-->
			</SCRIPT>
			<DIV ID='CalendarDiv' STYLE='POSITION:absolute;VISIBILITY:hidden;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>	
	</form>		 			 
	</table>
	<%Else%><p class="titlep" align="center" >לא נמצאו שיחות מצורפים לטיול</p>
	<%End If%>	
<%set app = nothing	%>
  
  
  
  
  
      <!--include table-->
  </td></tr>
</table><input type="hidden" id="ids" value="<%=ids%>" NAME="ids">
<input type="hidden" name="trapp" value="" ID="trapp">
</td>
<td width=80 nowrap valign="top" class="td_menu" style="border: 1px solid #808080; border-top: 0px">
<table cellpadding="1" cellspacing="0" width="100%" ID="Table6">
<tr><td align="<%=align_var%>" colspan="2" height=20 class="title_search">&nbsp;<!--חיפוש--></td></tr>
<tr><td colspan="2" height=30 nowrap></td></tr>

<tr><td colspan="2" height=10 nowrap>
<a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openSendMessage();">הוסף שיחה</a></td></tr>

<tr><td colspan=2 height=10 nowrap></td></tr>
</table>
</td></tr></table>
</td></tr></table>
</td></tr></table>

</body>
</html>
<%set con=Nothing%>