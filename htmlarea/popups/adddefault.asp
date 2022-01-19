<%mode=Request.QueryString("mode")
pageid = Request.QueryString("pageid")
if mode = "img" then
	htmlwidth = "432px"
	htmlheight = "310px"
	framesetrows = "310"
	main_frame_src = "insert_image.html"
	title_text = "Insert Image"
elseif mode = "file" then
	htmlwidth = "432px"
	htmlheight = "240px"
	framesetrows = "240"
	main_frame_src = "insert_file.html" 
	title_text = "Insert File"	
elseif mode = "table" then
	htmlwidth = "364px"
	htmlheight = "340px"
	framesetrows = "340"
	main_frame_src = "insert_table.html" 
	title_text = "Insert Table"	
elseif mode = "formtable" then
	htmlwidth = "364px"
	htmlheight = "435px"
	framesetrows = "4315"
	main_frame_src = "edit_table.html" 
	title_text = "Formating Table"
elseif mode = "flash" then
	htmlwidth = "364px"
	htmlheight = "369px"
	framesetrows = "369"
	main_frame_src = "insert_flash.html" 
	title_text = "Insert Flash"					
elseif mode = "form" then
	htmlwidth = "374px"
	htmlheight = "345px"
	framesetrows = "345"
	main_frame_src = "insert_form.asp?pageid="  & pageid
	title_text = "Insert Form"				
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML  id=FrameImage STYLE="width: <%=htmlwidth%>; height: <%=htmlheight%>; ">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="MSThemeCompatible" content="Yes">
<TITLE id=partitle><%=title_text%></TITLE>
</head>
  <frameset name="atraf_frameset" border="0" frameborder="0" framespacing="0" rows="<%=framesetrows%>,0">
	<frame name="main_frame" scrolling="no" noresize marginwidth="0" marginheight="0" src="<%=main_frame_src%>">
	<frame name="hidden_frame" scrolling="no" noresize marginwidth="0" marginheight="0" src="">
	<noframes>
	<body>
	</body>
	</noframes>
</frameset>
</html>
