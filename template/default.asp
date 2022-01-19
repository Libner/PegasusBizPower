<!--#include file="../include/connect.asp"-->
<!--#include file="../include/reverse.asp"-->  
<%  If trim(request("pageId")) <> "" Then
		pageId = cInt(trim(request("pageId")))
	Else
		pageId = ""
	End If
	
	If trim(request("parentId")) <> "" Then	
		parentId = cInt(trim(request("parentId")))
	Else
		parentId = ""	
	End If
	
	If trim(request("maincat")) <> "" Then
		maincat = cInt(trim(request("maincat")))	
	Else
		maincat = ""	
	End If

	If trim(request("catId")) <> "" Then	
		catId = cInt(trim(request("catId")))
	Else
		catId = ""	
	End If
	
	if request.QueryString("maincat")="2" then
		response.redirect(Application("VirDir") & "/english/template/default.asp?"&request.ServerVariables("QUERY_STRING"))
	end if

	if trim(parentId) <> "" then
		current_page = parentId
	else
		current_page = pageid
	end if

	if IsNumeric(pageId) then
		sqlStr = "SELECT Publication_Categories.Publication_Category_Name, PagesTav.Page_Title FROM Publication_Categories INNER JOIN  PagesTav ON Publication_Categories.Publication_Category_ID = PagesTav.Category_Id WHERE     (PagesTav.Page_Id = "& pageid &")"
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
 		sql = "SELECT page_Width, Page_content, Page_background FROM pagesTav WHERE Page_ID = "& current_page 
	else
		sql = "SELECT page_Width, Page_content, Page_background FROM pagesTav WHERE Page_ID = "& current_page &" AND Page_Visible = 1"
	end if
  		set rs_page = con.execute(sql)
		if not rs_page.EOF then
			page_Width = rs_page(0)		
			Page_content = rs_page(1)
			Page_background = rs_page(2)
		end if
		rs_page.close
		set rs_page=Nothing     
	end if          
%>			
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
<title>Bizpower<%=" - " & Page_Title & " - " & Publication_Category_Name%></title>
<!--#include file="../include/title_meta_inc.asp"-->
</head>
<body marginwidth="0" marginheight="0" topmargin="5" leftmargin="0" rightmargin="0" bgcolor="white">
<!--#include file="../include/top.asp"-->
<table dir="rtl" border="1" style="border-collapse:collapse" bordercolor="#999999" cellspacing="0" cellpadding="0" width="780" align="center" ID="Table1">
	<tr><td colspan=2 bgcolor="#FFFFFF" align="center" width="780">
	<%If trim(Page_background) = "" Or IsNull(Page_background) = true Then%>
	<div align="center">
			<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
				ID="Shockwaveflash1"  width=780 height="175" VIEWASTEXT>
				<PARAM NAME="movie" VALUE="<%=Application("VirDir")%>/flash/BP-flash-test.swf">
				<PARAM NAME="quality" VALUE="best">
				<PARAM NAME="wmode" VALUE="transparent">
				<PARAM NAME="bgcolor" VALUE="#EBEBEB">					
				<EMBED src="<%=Application("VirDir")%>/flash/BP-flash-test.swf" quality="best" wmode="transparent" bgcolor="#FFFFFF"
					TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
					width="780" height="175">
				</EMBED>
			</OBJECT>
	</div>
	<%Else%>
	<img src="<%=Application("VirDir")%>/download/page_backgrounds/<%=Page_background%>" border=0 align=center hspace=0 vspace=0>
	<%End If%>				
	</td></tr>
		<tr>
			<td align="right" class="title_page" bgcolor="#EBEBEB" dir="rtl" style="padding:5px;padding-right:10px">
				<%if trim(Publication_Category_Name) <> "" or trim(page_Title) <> "" then%>
				<%=Publication_Category_Name%>&nbsp;<span style="font-size:10pt;">>></span>&nbsp;<%=page_Title%>
				<%end if%>															
			</td>
			<td align="right" rowspan=2 valign="top" width="180" nowrap bgcolor="#FFFFFF">
			<!--#include file="../include/ashafim_inc.asp"-->
			</td>																
		</tr>			
   
	<tr>
		<td width="600" height=280 nowrap bgcolor="#F9F9F9" valign="top" align="right" style="padding-right:15px; padding-left:15px;">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">       
		<%'//start of centent%>			
		<tr>
		<td width="100%" valign="top" align="center">
		<%'//start of parent pages%>
		<%
		if trim(PageId) <> "" then
			sqlStr = "SELECT    Page_Id, Page_Title, Page_URL FROM  PagesTav WHERE     Page_Parent = "& pageid &" AND Page_Visible = 1 ORDER BY Page_Order"
			set rs_parent_Pages = con.execute(sqlStr)
			if not rs_parent_Pages.eof then
		%>
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>	
				<td align=right dir=rtl>
		<%
		do while not rs_parent_Pages.eof
			Page_Id = rs_parent_Pages("Page_Id")
			Page_Title = rs_parent_Pages("Page_Title")
			Page_URL = rs_parent_Pages("Page_URL")		
				%>
					<a dir="rtl" <%if trim(Page_URL)= "" then%>href="<%=Application("VirDir")%>/template/default.asp?PageId=<%=pageid%>&parentId=<%=Page_Id%>&catId=<%=categoryId%>&maincat=<%=maincat%>"<%else%><%if instr(1,LCase(Page_URL),"http://")<>0 then%>href="<%=Page_URL%>" target="_new"<%else%>href="<%=Application("VirDir")%>/<%=Page_URL%>"<%end if%><%end if%> class="parent_link" <%if cstr(Page_Id) = cstr(parentId) then%>style="color:#FFFFFF;BACKGROUND-COLOR:#9492C6;"<%end if%>><%=Page_Title%></a>&nbsp;&nbsp;
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
		<td valign="top" align="right" height="16"></td>
	</tr>    
	<tr>
	<td align="right" dir="rtl">					
		<table width="<%if trim(page_Width)<>"" then %><%=page_Width%><%else%>100%<%end if%>"  align="right" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="780">
				<%=Page_content%>
				</td>
			</tr>
		</table>												
	</td>
	</tr>
	<tr><td height="10" nowrap></td></tr>								
	</table></td></tr>					
</table>
<!--#include file="../include/bottom.asp"-->
</body>
</html>
<%Set con = Nothing%>