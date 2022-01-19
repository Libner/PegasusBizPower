<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%  numOftab = 2
    numOfLink = 0
	topLevel2 = 18 'current bar ID in top submenu - added 03/10/2019
	
	pageID = trim(Request("pageID"))
	
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
	if(isNaN(eval(window.document.all("pageID").value)) == false)
	{
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();

	}else{
   <%
     If trim(lang_id) = "1" Then
        str_alert = "! נא לבחור דף מעוצב"
     Else
		str_alert = "Please choose a page!"
     End If   
    %>
	window.alert("<%=str_alert%>");
	window.document.all("pageID").focus();
	}
}
//-->
</SCRIPT>
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" >
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
			<select dir="<%=dir_obj_var%>" name="pageID" class="app" ID="pageID" style="width:320" onchange="document.location.href='reports.asp?pageID='+this.value">
			<option value=""><%=String(30,"-") & arrTitles(18) & String(30,"-")%></option>
			<%	sqlstr = "Select Page_ID, Page_Title from Pages Where "&_
				" ORGANIZATION_ID=" & OrgID & " order by Page_Title"
				'Response.Write sqlstr
				'Response.End
				set rs_pages = con.GetRecordSet(sqlstr)
				if not rs_pages.eof then 
					PagesList = rs_pages.getRows()		
				end if
				set rs_pages=nothing														
				
				If IsArray(PagesList) Then
				For i=0 To Ubound(PagesList,2)
					page_id = PagesList(0,i)   	
					Page_Title = PagesList(1,i)	%>
				<option value="<%=page_id%>" <%If trim(pageID) = trim(page_id) Then%> selected <%End If%>> <%=Page_Title%> </option>
				<%	Next	
				End If%>
				</select>
			</td>
			<td width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<%=arrTitles(18)%>&nbsp;</td>
		</tr>
		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_obj_var%>">				
			<select dir="<%=dir_obj_var%>" name="categID" class="app" ID="categID" style="width:320" size=5 multiple>
			<option value="" selected><%=String(25,"-") & arrTitles(17) & String(25,"-")%></option>
			<option value="0">ללא מקור פרסום</option>
			<%	If Len(pageID) > 0 Then
			    sqlstr = "Select category_ID, referer from statistic_from_banner Where "&_
				" Page_ID=" & pageID & " order by referer"
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
					referer = CategsList(1,i)	%>
				<option value="<%=catID%>"> <%=referer%> </option>
				<%Next	
				End If
				End If%>
				</select>
			</td>
			<td width="20%" class="subject_form" nowrap valign="top" align="<%=align_var%>">&nbsp;<%=arrTitles(17)%>&nbsp;</td>
		</tr>						
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_stat.asp');"  style="width:250"><!--דוח כניסות--><%=arrTitles(16)%></a></TD>
		</TR>
	</TABLE>			
</FORM>
<!-- End main content -->
</td></tr>
</table>	 
</body>
<%set con=nothing%>
</html>