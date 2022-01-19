<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckFields()
{	
	if ( document.frmMain.Messages_Content.value==''
	     )		   
		{
				alert("'*' אנא מלאו כל השדות הצוינות בסימן")
				return false;				
		}
	else
		{
			document.frmMain.submit();
			return true;
		}		
}

function exit_()
	{
		this.close();	
		return false;
	}

//-->
</script>  

<%

departureId=request.querystring("dID")
if departureId>0 then
sqlDep="SELECT Departure_Code,Guide_FName +' ' + Guide_LName as Guide_Name from  Tours_Departures "& _
" left join Guides on  Tours_Departures.Guide_Id=Guides.Guide_Id "& _
" WHERE Departure_Id = "  & departureId

 set rs_Dep =  conPegasus.Execute(sqlDep)
'response.Write "bgg"
'response.end
	if not rs_Dep.eof then
		departure_code = rs_Dep("Departure_Code")
		Guide_Name=rs_Dep("Guide_Name")
end if
end if

errName = false
 UserID=trim(Request.Cookies("bizpegasus")("UserID"))
Messages_Id=request.querystring("Messages_Id")
if request.form("Messages_Content")<>nil then 'after form filling
	 Messages_Content = sFix(request.form("Messages_Content"))
	 
	 if Messages_Id=nil or Messages_Id="" then 'new record in DataBase
		   
		con.executeQuery("SET DATEFORMAT dmy")
		sqlStr = "insert into GuideMessages (Departure_Id,User_Id,Messages_Content) "
		sqlStr=sqlStr& " values (" & departureId &",'"& UserID & "','" & Messages_Content &"')"
		'Response.Write sqlStr
		con.GetRecordSet(sqlStr)					
		   
			 
		    
	 else			   
			
		
			con.executeQuery("SET DATEFORMAT dmy")
			sqlStr = "Update GuideMessages set Messages_Content = '" & Messages_Content &_
			"' where Messages_Id=" & Messages_Id
			con.GetRecordSet (sqlStr)
			
    end if  %>
     <SCRIPT LANGUAGE=javascript>
	 <!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	 //-->
	</SCRIPT>      
  <%  end if%>
<body>


<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title"><%if Messages_Id<>nil then%>עדכון<%else%>הוספת<%end if%>&nbsp;שיחה&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>  

<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="500">    
<%
if Messages_Id<>nil and Messages_Id<>"" then
  sqlStr = "select Departure_Id, User_Id, Messages_Content,  Messages_Date from GuideMessages where Messages_Id=" & Messages_Id  
  set rs_GuideMessages = con.GetRecordSet(sqlStr)
	if not rs_GuideMessages.eof then
	
	
		Messages_Content = rs_GuideMessages("Messages_Content")
		dateStart = rs_GuideMessages("Messages_Date")
		
	end if
	set rs_GuideMessages = nothing
	
else
	dateStart = Day(date()) & "/" & Month(date()) & "/" & Year(date())
end if
'Response.Write company_id
%>
<FORM name="frmMain" ACTION="AddGuideMessages.asp?Messages_Id=<%=Messages_Id%>&dID=<%=departureId%>" METHOD="post" onSubmit="return CheckFields()">
<tr>
	<td align=right><%=Guide_Name%> &nbsp;&nbsp;&nbsp;<%=vfix(departure_code)%></td>
	<td align="right" nowrap width="100">קוד טיול&nbsp;</td>
</tr>
<tr>
	<td align=right><%=vfix(dateStart)%></td>
	<td align="right" nowrap width="100">תאריך שיחה&nbsp;</td>
</tr>
<tr>
	<td align=right><textarea dir=rtl name="Messages_Content" style="width:300px;height:200px;font-family:arial" ID="Messages_Content"><%=vfix(Messages_Content)%></textarea></td>
	<td align="right" nowrap width="100" valign=top>פרוט שיחה&nbsp;</td>
</tr>




<tr><td colspan="2" height=10></td></tr>

<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>
<td width=50% align=right><input type=button class="but_menu" style="width:90px" onclick="return exit_();" value="ביטול"></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="אישור"></td></tr>
</table></td></tr>
</form>
<tr><td colspan="2" height=10></td></tr>
</table>
</td></tr></table>
</body>
<%
set con = nothing
%>
</html>
