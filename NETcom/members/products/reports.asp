<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%  numOftab = 1
    numOfLink = 1
	
topLevel2 = 14 'current bar ID in top submenu - added 03/10/2019

	prodId = trim(Request("prodId"))
	start_date = "1/" & Month(Date()) & "/" & Year(date())
	end_date = DateAdd("d",30,start_date)
	If Request.Form("dateStart") <> nil Then
		start_date = Request.Form("dateStart")
	End If
	If Request.Form("dateEnd") <> nil Then
		end_date = Request.Form("dateEnd")
	End If		
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 25 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
		dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	End If
	set rstitle = Nothing%> 	
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function report_open(fname){
	if(isNaN(eval(window.document.all("quest_id").value)) == false)
	{
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();

	}else{
   <%
     If trim(lang_id) = "1" Then
        str_alert = "! נא לבחור טופס"
     Else
		str_alert = "Please choose a form!"
     End If   
    %>
	window.alert("<%=str_alert%>");
	window.document.all("quest_id").focus();
	}
}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}
//-->
</SCRIPT>
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width=100% align="<%=align_var%>"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</td></tr>
<tr><td>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;</td></tr>
<tr><td height=15 nowrap></td></tr>
 <tr>    
    <td width="100%" valign="top" align="center">
    <table width="100%" cellspacing="0" cellpadding="3" align=center border="0" bgcolor="#ffffff" >
      <tr>
        <td width="100%" align=center valign=top >
<!-- start code -->
	<FORM action="" method=POST id=formsearch name=formsearch target=_self>
	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>">
		<tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2" height="15" nowrap>
				<input type="hidden" name="report" value="yes">
			</td>
		</tr>
		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_obj_var%>">				
			<select dir="<%=dir_obj_var%>" name="quest_id" class="app" ID="quest_id" style="width:320" onchange="document.location.href='reports.asp?prodId='+this.value">
			<option value=""><%=arrTitles(3)%></option>
			<%	sqlstr = "Select product_id, product_name from Products Where "&_
				" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
				'Response.Write sqlstr
				'Response.End
				set rs_products = con.GetRecordSet(sqlstr)
				if not rs_products.eof then 
					ProductsList = rs_products.getRows()		
				end if
				set rs_products=nothing														
				
				If IsArray(ProductsList) Then
					For i=0 To Ubound(ProductsList,2)
						prod_Id = ProductsList(0,i)   	
						product_name = ProductsList(1,i)
				%>
				<option value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </option>
				<%	Next	
				End If	
				%>
				</select>
			</td>
			<td width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<%=arrTitles(4)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_obj_var%>">				
			<select dir="<%=dir_obj_var%>" name="categID" class="app" ID="categID" style="width:320" size=5 multiple>
			<option value="" selected><%=String(25,"-") & arrTitles(17) & String(25,"-")%></option>			
			<%	If Len(prodId) > 0 Then
			    sqlstr = "Select DISTINCT category_ID, (Select referer from statistic_from_banner where category_ID = Statistic.category_ID) from Statistic Where "&_
				" Product_ID =" & prodId & " order by category_ID"
				'Response.Write sqlstr
				'Response.End
				set rs_categs = con.GetRecordSet(sqlstr)
				if not rs_categs.eof then 
					CategsList = rs_categs.getRows()		
				end if
				set rs_categs=nothing														
				
				If IsArray(CategsList) Then
				For i=0 To Ubound(CategsList,2)
					catID = CategsList(0,i)   	
					referer = CategsList(1,i)
					If Len(referer) = 0 Or isNULL(referer) Then
						referer = "ללא מקור פרסום"
					End If%>
				<option value="<%=catID%>"> <%=referer%> </option>
				<%Next	
				End If
				End If%>
				</select>
			</td>
			<td width="20%" class="subject_form" nowrap valign="top" align="<%=align_var%>">&nbsp;<%=arrTitles(17)%>&nbsp;</td>
		</tr>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateStart);' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateStart" name="dateStart" value="<%=start_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7><!--- תאריך מ--><%=arrTitles(7)%></span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateEnd);' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8><!--- תאריך עד--><%=arrTitles(8)%></span>&nbsp;</TD>
		</TR>								
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('../reports/report_stat.asp');"  style="width:250"><!--התפלגות מכירות עפ"י מקורות פרסום--><%=arrTitles(15)%></a></TD>
		</TR>
	</TABLE>			
</FORM>
<!-- End main content -->
</td></tr>
</table>	 
</body>
<%set con=nothing%>
</html>