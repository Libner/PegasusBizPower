<!--#INCLUDE FILE="../../reverse.asp"-->
<!--#include file="../../connect.inc"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<base target="middle">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
</head>

<body bgcolor="#FFFFFF" marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
prodId=Request.Form("product_Id")
%>
<table border="0" width="90%" cellspacing="0" cellpadding="0" align=center>
	<tr><td width="100%" height=10></td></tr>
<!-- start code --> 
	<tr>
		<th width="100%" bgcolor=Black>
<%				set prod = con_net.GetRecordSet("Select * from products where product_id=" & prodId)
					if not prod.eof then
					productName=prod("Product_Name")					
					quest_id = prod("QUESTIONS_ID")
					
					if prod("Langu") = "eng" then
						dir_align = "ltr"
						td_align = "left"
						pr_language = "eng"
					else
						dir_align = "rtl"
						td_align = "right"
						pr_language = "heb"
					end if
				end if
				set prod = nothing%>
			<font color=#FFFFFF size=4><%=productName%></font>
		</th>
	</tr>
	<tr><td width="100%" height=10></td></tr>
<%  ' level 1
	questsstr = "select * from  QUESTIONS where PRODUCT_ID = "&prodid&" and PARENT_ID = 0 order by Question_Order"
	set l1_quests = con_net.GetRecordSet(questsstr)
	DO WHILE not l1_quests.EOF 
		Question_Id = l1_quests("Question_Id")	
		QUESTION_TEXT = l1_quests("QUESTION_TEXT")
		QUESTION_TYPE = l1_quests("QUESTION_TYPE")
		QUESTION_DESCRIPTION = trim(l1_quests("QUESTION_DESCRIPTION"))
%>     
			<tr>
				<td align=<%=td_align%> class="subject_form">
					<TABLE align=center BORDER=0 CELLSPACING=0 CELLPADDING=1 width=100% bgcolor="#E2F1F8">
					<TR>
						<td class="subject_form" width="100%" align=<%=td_align%> nowrap>&nbsp;<%=QUESTION_TEXT)%>&nbsp;</td>
					</TR>
				<%if QUESTION_DESCRIPTION <> "" then%>
					<tr>
						<td class="form" width="100%" align=<%=td_align%> valign="top" nowrap>&nbsp;<%=QUESTION_DESCRIPTION%>&nbsp;</td>
					</tr>
				<%end if%>
					</TABLE>	
			 	</td>
			</tr> 
	<!--#INCLUDE FILE="proc_fields.asp"-->
	
	<%  ' level 2
	questsstr = "select * from  QUESTIONS where PRODUCT_ID = "&prodid&" and PARENT_ID = " & Question_Id & " order by Question_Order"
	set l2_quests = con_net.GetRecordSet(questsstr)
	DO WHILE not l2_quests.EOF 
		Question_Id = l2_quests("Question_Id")	
		QUESTION_TEXT = l2_quests("QUESTION_TEXT")
		QUESTION_DESCRIPTION = trim(l2_quests("QUESTION_DESCRIPTION"))
		QUESTION_TYPE = l2_quests("QUESTION_TYPE")
%> 	
			<tr>
				<td align=<%=td_align%> class="subsubj_form">
					<TABLE align=center BORDER=0 CELLSPACING=0 CELLPADDING=1 width=100% bgcolor="#E2F1F8">
					<TR>
						<td class="subsubj_form" width=10>&nbsp;</td>
						<td class="subsubj_form" align=<%=td_align%> nowrap><%=QUESTION_TEXT%></td>
						<td class="subsubj_form" width=10>&nbsp;</td>
					</TR>
				<%if QUESTION_DESCRIPTION <> "" then%>
					<tr>
						<td class="form" width=10>&nbsp;</td>
						<td class="form" align=<%=td_align%> valign="top" nowrap><%=QUESTION_DESCRIPTION%></td>
						<td class="form" width=10>&nbsp;</td>
					</tr>
				<%end if%>
					</TABLE>	
			 	</td>
			</tr> 
	<!--#INCLUDE FILE="proc_fields.asp"-->
	
	<%  ' level 3
	questsstr = "select * from  QUESTIONS where PRODUCT_ID = "&prodid&" and PARENT_ID = " & Question_Id & " order by Question_Order"
	set l3_quests = con_net.GetRecordSet(questsstr)
	DO WHILE not l3_quests.EOF 
		Question_Id = l3_quests("Question_Id")	
		QUESTION_TEXT = l3_quests("QUESTION_TEXT")
		QUESTION_DESCRIPTION = trim(l3_quests("QUESTION_DESCRIPTION"))
		QUESTION_TYPE = l3_quests("QUESTION_TYPE")
%> 
			<tr>
				<td align=<%=td_align%> class="quest_form">
					<TABLE align=center BORDER=0 CELLSPACING=0 CELLPADDING=1 width=100% bgcolor="#E2F1F8">
					<TR>
						<td class="quest_form" width=20>&nbsp;</td>
						<td class="quest_form" align=<%=td_align%> nowrap><%=QUESTION_TEXT%></td>
						<td class="quest_form" width=20>&nbsp;</td>
					</TR>
				<%if QUESTION_DESCRIPTION <> "" then%>
					<tr>
						<td class="form" width=20>&nbsp;</td>
						<td class="form" align=<%=td_align%> valign="top" nowrap>&nbsp;<%=QUESTION_DESCRIPTION%>&nbsp;</td>
						<td class="form" width=20>&nbsp;</td>
					</tr>
				<%end if%>
					</TABLE>	
			 	</td>
			</tr> 
	<!--#INCLUDE FILE="proc_fields.asp"-->
			
	<%		l3_quests.MoveNext
			loop
			set l3_quests = nothing	

		l2_quests.MoveNext
		loop
		set l2_quests = nothing
				
	l1_quests.MoveNext
	loop
	set l1_quests = nothing %>	
				
<!-- end code --> 
	<tr><td width="100%" height=20></td></tr>		
</TABLE>
</body>
<%
set con = nothing%>
</html>
