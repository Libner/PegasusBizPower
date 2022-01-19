<%mode=Request.QueryString("mode")
if mode = "img" then
	htmlwidth = "432px"
	htmlheight = "290px"
	framesetrows = "290"
	main_frame_src = "insert_image.html"
	title_text = "Insert Image"
elseif mode = "file" then
	htmlwidth = "432px"
	htmlheight = "220px"
	framesetrows = "220"
	main_frame_src = "insert_file.html" 
	title_text = "Insert File"
elseif mode = "table" then
	htmlwidth = "364px"
	htmlheight = "320px"
	framesetrows = "320"
	main_frame_src = "insert_table.html" 
	title_text = "Insert Table"	
elseif mode = "formtable" then
	htmlwidth = "364px"
	htmlheight = "415px"
	framesetrows = "415"
	main_frame_src = "edit_table.html" 
	title_text = "Formating Table"	
elseif mode = "flash" then
	htmlwidth = "364px"
	htmlheight = "349px"
	framesetrows = "349"
	main_frame_src = "insert_flash.html" 
	title_text = "Insert Flash"				
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML  id=FrameImage STYLE="width: <%=htmlwidth%>; height: <%=htmlheight%>; ">
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="MSThemeCompatible" content="Yes">
<TITLE><%=title_text%></TITLE>
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
