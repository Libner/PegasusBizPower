       	<table dir="ltr" border="0" bgcolor="#DDDDDD" style="border: solid 1px #999999;border-bottom:none" cellspacing="0" cellpadding="0" width="780" align="center">
        <tr><td width="780" bgcolor="#DDDDDD">
        <table cellpadding=0 cellspacing=0 width=780>
        <tr>                 
			<td nowrap class="titleB" width="223" align=right>כלים חזקים אונליין</td>
			<td width="100%" align="right" class="bar_top" valign="baseline">
			<table align=right bgcolor="#DDDDDD" cellpadding="0" cellspacing="0" border=0>
	        <tr style="height: 20px" valign=middle>		    
		    <!--td nowrap>
			<img src="<%=Application("VirDir")%>/images/shim.gif" width=0 height=3 border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class="top_link"  href="<%=Application("VirDir")%>/english/">English</a>&nbsp;<font color="#616D66">&#8226;</font>
			</td-->		
	<%
	sqlStr = "SELECT TOP 100 PERCENT Publication_Category_ID, Publication_Category_Name, Category_URL FROM         dbo.Publication_Categories WHERE   Main_Category_Id=1 and  (Category_Vis = 1) AND (Publication_Category_ID IN  (SELECT     Category_ID FROM          PagesTav WHERE      (Page_Parent = 0) AND (Page_Visible_Title = 1))) ORDER BY Category_Order desc"
	sqlStr = "SELECT Publication_Category_ID,Publication_Category_Name, Category_URL from Publication_Categories where  (Category_Vis = 1) AND Main_Category_ID = 1 order by Category_Order desc"
	set rs_Publication_Categories = con.execute(sqlStr)	
	i=1
	do while not rs_Publication_Categories.eof
		PubCategoryID = rs_Publication_Categories(0)
		PubCategoryName = rs_Publication_Categories(1)	
		Category_URL = rs_Publication_Categories(2)			  
		if trim(PubCategoryName) <> "" then
		  
			sqlStr = "Select Top 1 Page_Id,Page_Visible, Page_Visible_Title FROM PagesTav WHERE Category_Id="& PubCategoryID &" AND (Page_Visible_title=1 OR Page_Visible = 1 ) AND Page_Parent=0 order by page_order"
			''Response.Write sqlStr
			''Response.End
			set pag=con.Execute(sqlStr)
			if not pag.eof then
				currPageId = pag("Page_Id")
				PageVisible = pag("Page_Visible")
				PageVisible_Title = pag("Page_Visible_Title")
				
				'//START of show parent page if the page content not visible
				if PageVisible = "False" then
					sqlStr = "SELECT TOP 1 Page_Id FROM  PagesTav WHERE     Page_Parent = "& currPageId &" AND Page_Visible = 1 ORDER BY Page_Order"
					''Response.Write sqlStr
					''Response.End				
					set rs_parent_Pages = con.execute(sqlStr)
					if not rs_parent_Pages.eof then
						str_parentId = "&parentId=" & rs_parent_Pages("Page_Id")
					else
						str_parentId = ""	
					end if
					set rs_parent_Pages = nothing
				else
					str_parentId = ""	
				end if	
				'//END of show parent page if the page content not visible
			else
				currPageId="-1"	'if not exists page,that we can see on the first category page 
			end if
			set pag= nothing
			
			if trim(Category_URL) = "" then
				href_string = "../template/default.asp?PageId="& currPageId & str_parentId & "&amp;catId="& PubCategoryID & "&amp;maincat=1"
			elseif instr(1,LCase(Category_URL),"http://")<>0 then
				href_string="javascript:void(0)"
				onclick_string = "onclick=""javascript:window.open('"& Category_URL &"'); return false;""" ''Category_URL  ''" target='_new'"
			else
				href_string = Category_URL 
			end if		  
	%>
		<td nowrap>
			<%if currPageId="-1" then%>
				<img src="<%=Application("VirDir")%>/images/shim.gif" width=0 height=3 border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class="top_link"  href="<%=href_string%>" <%=onclick_string%> dir="rtl"><%=PubCategoryName%></a>&nbsp;<font color="#616D66">&#8226;</font>		
			<%else%>
				<img src="<%=Application("VirDir")%>/images/shim.gif" width=0 height=3 border="0" name="menuLayer<%=i%>" alt="menuLayer<%=i%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class="top_link"  href="<%=href_string%>"  dir="rtl" onMouseOver="popUp('elMenu<%=i%>',event);status='';return true" onMouseOut="popDown('elMenu<%=i%>');status=''"><%=PubCategoryName%></a>&nbsp;<font color="#616D66">&#8226;</font>
			<%i=i+1%>			
			<%end if%>
		</td>	
	<%
		end if
	rs_Publication_Categories.movenext
	loop
     set rs_Publication_Categories = nothing
     %>
		<td width="19" nowrap></td>
     
	</tr>
</table></td></tr>
</table></td></tr></table>