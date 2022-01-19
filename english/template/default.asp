<!--#include file="../include/connect.asp"-->
<!--#include file="../include/reverse.asp"-->
<%
	pageId=request("pageId")
	parentId = request("parentId")

	if trim(parentId) <> "" then
		current_page = parentId
	else
		current_page = pageid
	end if

	if trim(pageid) <> "" then
		sqlStr = "SELECT Publication_Categories.Publication_Category_Name, PagesTav.Page_Title FROM Publication_Categories INNER JOIN  PagesTav ON Publication_Categories.Publication_Category_ID = PagesTav.Category_Id WHERE  Publication_Categories.Main_Category_Id=2 and (PagesTav.Page_Id = "& pageid &")"
		set rs_publicationPage = con.execute(sqlStr)
		if not rs_publicationPage.eof then
			Publication_Category_Name = rs_publicationPage("Publication_Category_Name")
			Page_Title = rs_publicationPage("Page_Title")	
		else
			Publication_Category_Name = ""
			Page_Title = ""
		end if
		set rs_publicationPage = nothing
	end if

	if trim(current_page)<>"" then
		if request.QueryString("show")<>nil then 'Preview
 			sql = "SELECT page_Width,  Page_content FROM pagesTav WHERE Page_ID = "& current_page 
		else
			sql = "SELECT page_Width,  Page_content FROM pagesTav WHERE Page_ID = "& current_page &" AND Page_Visible = 1"
		end if
  		set rs_Width = con.execute(sql)
		if not rs_Width.EOF then
			page_Width = rs_Width("page_Width")
			Page_content = rs_Width("Page_content")
		end if
		rs_Width.close
		set rs_Width=Nothing     
	end if          
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
<title>Bizpower<%=" - " & Publication_Category_Name & " - " & Page_Title%></title>
<!--#include file="../../include/title_meta_inc.asp"-->
</head>

<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="5" leftmargin="0"
		rightmargin="0" bgcolor="white">
<!--#include file="../include/top.asp"-->
<table dir="rtl" border="1" style="border-collapse:collapse" bordercolor="#999999" cellspacing="0" cellpadding="0" width="780" align="center" ID="Table1">
	<tr><td colspan=2 bgcolor="#FFFFFF" align="center" width="780">
	<div align="center">
					<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
						ID="Shockwaveflash1" VIEWASTEXT width=778 height="175">
						<PARAM NAME="movie" VALUE="../../flash/BP-flash-test.swf">
						<PARAM NAME="quality" VALUE="best">
						<PARAM NAME="wmode" VALUE="transparent">
						<PARAM NAME="bgcolor" VALUE="#EBEBEB">					
						<EMBED src="../../flash/BP-flash-test.swf" quality="best" wmode="transparent" bgcolor="#FFFFFF"
							TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
							width="778" height="175">
						</EMBED>
					</OBJECT>
				</div>				
		</td></tr>
		<tr>
			<td align="left" class="title_page" bgcolor="#EBEBEB" dir="rtl"
					style="padding:5px;padding-right:10px">

				<%if trim(Publication_Category_Name) <> "" or trim(page_Title) <> "" then%>
						<span style="font-size:17pt;color:#91A399"><%=Publication_Category_Name%> </span><span style="font-size:10pt;">>></span>
						<span style="font-size:13pt;"><%=page_Title%></span>
				<%end if%>															
			</td>											
			<td rowspan=2 dir=ltr align="left" valign="top" width="255" nowrap bgcolor="#FFFFFF">
			<!--#include file="../include/ashafim_inc.asp"-->
			</td>	
		</tr>			
   
	<tr>
		<td width="520" height="450" nowrap bgcolor="#F9F9F9" valign="top" align="left" style="padding-right:20px; padding-left:20px;">
        <table border="0" cellspacing="0" cellpadding="0" width="480" ID="Table2">       
		<%'//start of centent%>			
		<tr>
		<td width="480" valign="top" align="center">
		<%'//start of parent pages%>
		<%
		if trim(PageId) <> "" then
			sqlStr = "SELECT    Page_Id, Page_Title, Page_URL FROM  PagesTav WHERE     Page_Parent = "& pageid &" AND Page_Visible = 1 ORDER BY Page_Order"
			set rs_parent_Pages = con.execute(sqlStr)
			if not rs_parent_Pages.eof then
		%>
		<table border="0" width="480" cellspacing="0" cellpadding="0" ID="Table5">
			<tr>	
				<td align=right dir=rtl>
		<%
		do while not rs_parent_Pages.eof
			Page_Id = rs_parent_Pages("Page_Id")
			Page_Title = rs_parent_Pages("Page_Title")
			Page_URL = rs_parent_Pages("Page_URL")		
				%>
					<a dir="rtl" <%if trim(Page_URL)= "" then%>href="../template/default.asp?PageId=<%=pageid%>&parentId=<%=Page_Id%>&catId=<%=categoryId%>&maincat=<%=maincat%>"<%else%><%if instr(1,LCase(Page_URL),"http://")<>0 then%>href="<%=Page_URL%>" target="_new"<%else%>href="../<%=Page_URL%>"<%end if%><%end if%> class="parent_link" <%if cstr(Page_Id) = cstr(parentId) then%>style="color:#FFFFFF;BACKGROUND-COLOR:#9492C6;"<%end if%>><%=Page_Title%></a>&nbsp;&nbsp;
				<%
		rs_parent_Pages.movenext
		loop
		rs_parent_Pages.movefirst
		%>
			</td>
		</tr>	
	</table>
	<%
			end if
		set rs_parent_Pages = nothing
		end if	
	%>
	<%'//end of parent pages%>					
	</td>
	</tr>
	<tr>
		<td width="480" valign="top" align="left" height="16"></td>
	</tr>    
	<tr>
	<td align="left" dir="rtl">					
		<table width="<%if trim(page_Width)<>"" then %><%=page_Width%><%else%>100%<%end if%>"  align="left" border="0" cellspacing="0" cellpadding="0" ID="Table3">
			<tr>
				<td width="780">
				<%=Page_content%>
				</td>
			</tr>
		</table>												
	</td>
	</tr>
	<tr><td height="5"></td></tr>								
	</table>
	</td>
								
	</tr>					
</table>
<!--#include file="../include/bottom.asp"-->
</body>
</html>