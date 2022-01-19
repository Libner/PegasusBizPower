
<table bgcolor="#DDDDDD" cellpadding="0" cellspacing="0" border="0">
<tr>
	<%
	sqlStr = "SELECT TOP 100 PERCENT Publication_Category_ID, Publication_Category_Name, Category_URL FROM  dbo.Publication_Categories WHERE  Main_Category_ID = 2 and  (Category_Vis = 1) AND (Category_URL is not null or Publication_Category_ID IN  (SELECT     Category_ID FROM          PagesTav WHERE      (Page_Parent = 0) AND (Page_Visible_Title = 1))) ORDER BY Category_Order"
	'sqlStr = "SELECT Publication_Category_ID,Publication_Category_Name, Category_URL from Publication_Categories where  (Category_Vis = 1) AND Main_Category_ID = 2 order by Category_Order desc"
	set rs_Publication_Categories = con.execute(sqlStr)	
	i=1
	do while not rs_Publication_Categories.eof
		PubCategoryID = rs_Publication_Categories("Publication_Category_ID")
		PubCategoryName = rs_Publication_Categories("Publication_Category_Name")	
		Category_URL = rs_Publication_Categories("Category_URL")		 
		if trim(PubCategoryName) <> "" then
		  
			sqlStr = "Select Top 1 Page_Id,Page_Visible, Page_Visible_Title FROM PagesTav WHERE Category_Id="& PubCategoryID &" AND (Page_Visible_title=1 OR Page_Visible = 1 ) AND Page_Parent=0 order by page_order"
			''Response.Write sqlStr
			''Response.End
			set pag=con.Execute(sqlStr)
			if not pag.eof then
				currPageId = pag("Page_Id")
				Page_Visible = pag("Page_Visible")
				Page_Visible_Title = pag("Page_Visible_Title")
				
				'//START of show parent page if the page content not visible
				if Page_Visible = "False" then
					sqlStr = "SELECT    TOP 1 Page_Id FROM  PagesTav WHERE     Page_Parent = "& currPageId &" AND Page_Visible = 1 ORDER BY Page_Order"
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
			
			if trim(Category_URL) = "" or IsNull(Category_URL) then
				href_string = "../template/default.asp?PageId="& currPageId & str_parentId & "&catId="& PubCategoryID & "&maincat=2"
			elseif instr(1,LCase(Category_URL),"http://")<>0 then
				href_string=""
				onclick_string = "onclick=""javascript:window.open('"& Category_URL &"'); return false;""" ''Category_URL  ''" target='_new'"
			else
				href_string = Category_URL 
			end if		  
	%>
		<td nowrap>
			<%if currPageId="-1" then%>
				&nbsp;&nbsp;&nbsp;<img src="<%=Application("VirDir")%>/images/shim.gif" width=0 height=3 border="0">&nbsp;&nbsp;<font color="#616D66">&#8226;</font>&nbsp;&nbsp;<a class="top_link"  href="<%=href_string%>" <%=onclick_string%> dir="rtl"><%=PubCategoryName%></a>		
			<%else%>
				&nbsp;&nbsp;&nbsp;<img src="<%=Application("VirDir")%>/images/shim.gif" width=0 height=3 border="0" name="menuLayer<%=i%>" alt="menuLayer<%=i%>">&nbsp;&nbsp;<font color="#616D66">&#8226;</font>&nbsp;&nbsp;<a class="top_link"  href="<%=href_string%>"  dir="rtl" onMouseOver="popUp('elMenu<%=i%>',event);status='';return true" onMouseOut="popDown('elMenu<%=i%>');status=''"><%=PubCategoryName%></a>
			<%i=i+1%>			
			<%end if%>
		</td>	
	<%
		end if
	rs_Publication_Categories.movenext
	loop
     set rs_Publication_Categories = nothing
     %>
		<td nowrap>
				&nbsp;&nbsp;&nbsp;<img src="<%=Application("VirDir")%>/images/shim.gif" width=0 height=3 border="0">&nbsp;&nbsp;<font color="#616D66">&#8226;</font>&nbsp;&nbsp;<a class="top_link"  href="<%=Application("VirDir")%>/../default.asp" >עברית</a>			
		</td>	
		<td width="19" nowrap></td>
     
	</tr>
</table>
