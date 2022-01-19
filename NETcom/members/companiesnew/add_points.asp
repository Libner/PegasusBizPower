<% Server.ScriptTimeout=10000 %>
<% Response.Buffer = False %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%ContactId = trim(Request.QueryString("ContactId"))%>
<%If isNumeric(trim(UserID)) = true And IsNumeric(trim(OrgID)) Then%>
<%If trim(ContactId)<>"" then   
	 sqlStr = "SELECT CONTACT_NAME, dbo.fnc_bonus_points_client(CONTACT_ID) as total_points, " & _
	 " dbo.fnc_realized_points_client(CONTACT_ID) as realized_points FROM CONTACTS WHERE CONTACT_ID="& ContactId 
	 set pr=con.GetRecordSet(sqlStr)
	 If not pr.EOF then	
		CONTACT_NAME = trim(pr("CONTACT_NAME"))
		total_points = trim(pr("total_points"))
		realized_points = trim(pr("realized_points"))
	 End If
	 set pr = Nothing
	 End If	
	 
	 If trim(total_points) = "" Or IsNull(total_points) Then
		total_points = 0.0
	Else
		total_points	 = cDbl(total_points)
	 End If
	 
	 If trim(realized_points) = "" Or IsNull(realized_points)  Then
		realized_points = 0.0
	Else
		realized_points	 = cDbl(realized_points)
	 End If	 
	 
	 total_points = (total_points - realized_points)
	
	If Request.QueryString("add") <> nil then	
		SaleId = cInt(Request.Form("cmbSales"))	
		sqlstr = "INSERT INTO Contacts_To_Sales (contact_id, sale_id, insert_date, user_id) VALUES ('"&_
		ContactId & "','" & SaleId &"', getDate(), '"& UserId &"')"
		'Response.Write sqlstr
		'Response.End
		con.executeQuery(sqlstr)
		Set rs_tmp = Nothing	  		
		Set con = Nothing 	 %>
	 <script language="javascript" type="text/javascript">
	 <!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	 //-->
	</script>
	 <%	
	End If 
  End If	   
%>	
<%sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
	function exit_()
	{
		this.close();	
		return false;
	}
//-->
</script>
</head>
<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" >
<tr>
<td width="100%" class="page_title" dir="<%=dir_obj_var%>" colspan="2">מימוש נקודות&nbsp;&nbsp;<font color="#6F6DA6"><%=CONTACT_NAME%></font>&nbsp;</td>
</tr> 
<tr><td height=20 nowrap></td></tr>        
<%  sqlstr = "SELECT SaleCode, SaleName, Points FROM pegasus.dbo.Sales WHERE (Sale_Visible = 1) " & _
		" AND (Points < "  & total_points & ") " & " ORDER BY SaleOrder"
		Set rs_list = con.getRecordSet(sqlstr)
		If not rs_list.EOF Then%>
	<tr>
		<td align="<%=align_var%>" width="100%">
		<form name="Form1" action="add_points.asp?add=1&ContactId=<%=ContactId%>" method="post" target="_self" ID="Form1">
		<table border="0" cellpadding="1" cellspacing="3" width="100%" align="<%=align_var%>" dir="<%=dir_var%>" >        
			<TR>
				<td align="<%=align_var%>" width="100%">		
				<select id="cmbSales" name="cmbSales" dir="rtl" style="width: 100%;" class="sel">
			<%	While not rs_list.eof		%>
					<option value="<%=trim(rs_list(0))%>" ><%=trim(rs_list(1))%> - <%=trim(rs_list(2))%> נקודות</option>
			<%	rs_list.moveNext
				Wend		%>				
				</select>	
				</td>
				<td width="80" nowrap align="right" class="card">&nbsp;מבצע&nbsp;</td> 
			</TR>
			<tr><td colspan="2" height="15" nowrap></td></tr>
			<tr>
            	<td colspan="2" align="center" nowrap dir=<%=dir_var%> class="card"><INPUT type="button" class="but_menu" 
            	onclick="return exit_();" value="<%=arrButtons(2)%>" style="width:100px" ID="btnClose" NAME="btnClose">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<INPUT type="submit" class="but_menu" value="<%=arrButtons(1)%>" style="width:100px" ID="btnSave" 
				NAME="btnSave"></td>
			</tr>
		</table>
	 </form>								
	</td>
	</tr>
<%Else%>	
<tr><td class="form_header" >המערכת לא מאפשרת לבחור מבצע שיש לו יותר נקודות מכמות נקודות הזכות של הלקוח</td></tr>
<%End If	
	Set rs_list = Nothing	%>
</table>
</body>
</html>	
<%Set con = Nothing%>