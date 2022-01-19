<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%	
	companyID = trim(Request.QueryString("companyID"))
	If trim(companyID)<>"" then   
	sqlStr = "SELECT company_name FROM companies WHERE company_id="& companyID 
	set pr=con.GetRecordSet(sqlStr)
	if not pr.EOF then	
		company_name = pr("company_name")
	end if
	set pr = Nothing
	End If	
	
	if Request.QueryString("add") <> nil then
		set upl=Server.CreateObject("SoftArtisans.FileUp")		
		If upl.TotalBytes > 6291456 Then 'cscript.exe adsutil.vbs SET w3svc/aspbufferinglimit 6291456
		%>
			<script language=javascript>
			<!--
				window.alert("ניתן לצרף מסמך עד 6 MB");
				window.history.back();
			//-->
			</script>
		<%
			Response.End
		Else
		str_mappath="../../../download/documents"
		File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
		extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
		name_without_extend = LCase(Mid(File_Name,1,Len(File_Name)-Len(extend)-1))
		NewFileName =  "company_doc_" & companyID
		new_name = NewFileName
		upload = true
		i = 0
		set fs = server.CreateObject("Scripting.FileSystemObject") 
		do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
		    i =  i + 1
			new_name = new_name & "_" & i
		loop
		newFileName = new_name & "." & extend
		set fs = Nothing
		
		upl.Form("document_file").SaveAs Server.Mappath("../../../download/documents/") & "/" & NewFileName
				
		sqlstr = "SET NOCOUNT ON; Insert INTO DOCUMENTS (document_name, document_file, document_desc) VALUES ('"&_
		sFix(trim(upl.Form("document_name")))&"','"&sFix(trim(NewFileName))&"','"&_
		sFix(trim(upl.Form("document_desc")))&"'); SELECT @@IDENTITY AS NewID"
		'Response.Write sqlstr
		'Response.End
		set rs_tmp = con.getRecordSet(sqlstr)
			document_id = rs_tmp.Fields("NewID").value
		set rs_tmp = Nothing	  		
		  		
		set upl = nothing	
	
		sqlstr = "Insert Into company_documents (company_id, document_id) values (" & companyID & "," & document_id & ")"
		'Response.Write sqlstr
		'Response.End
		con.ExecuteQuery(sqlstr)			
	
	 %>
	 <SCRIPT LANGUAGE=javascript>
	 <!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	 //-->
	</SCRIPT>
	 <%	
	End If 
  End If	   
%>	
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 16 Order By word_id"				
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
	  set rstitle = Nothing
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

<script LANGUAGE="javascript">
<!--
	
	function checkfile()
    {
	  if(document.AddAttach.document_name.value=='')
	  {
			<%If trim(lang_id) = "1" Then%>
			window.alert("נא למלא כותרת מסמך");
			<%Else%>
			window.alert("Please insert document title");
			<%End If%>
			document.AddAttach.document_name.focus();
			return false;		
	  }
	  if (document.AddAttach.document_file.value=='')
	  {
			<%If trim(lang_id) = "1" Then%>
			window.alert('! נא לבחור קובץ');
			<%Else%>
			window.alert('Please choose the document!');
			<%End If%>
			document.AddAttach.document_file.focus();
			return false;
	  }
	  else	
	  {
			var fname=new String();
			var fext=new String();
			var extfiles=new String();
			fname=document.AddAttach.document_file.value;
			indexOfDot = fname.lastIndexOf('.')
			fext=fname.slice(indexOfDot+1,-1)		
			fext=fext.toUpperCase();
			extfiles='HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT';		
			if(extfiles.indexOf(fext)>-1)
				return true;
			else
			    <%If trim(lang_id) = "1" Then%>
				window.alert(':סיומת הקובץ - אחת מרשימה' + '\n' + 'HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT');
				<%Else%>
				window.alert('The file ending should be one of the these:' + '\n' + 'HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT');
				<%End If%>
			return false;
	  }  
	   return true;
    }

	function exit_()
	{
		this.close();	
		return false;
	}
//-->
</script>
</head>
<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="100%" class="page_title" dir="<%=dir_obj_var%>" colspan=2><span id=word1 name=word1><!--הוסף מסמך--><%=arrTitles(1)%></span>&nbsp;&nbsp;<font color="#6F6DA6"><%=company_name%></font>&nbsp;</td>
</tr> 
<tr><td height=20 nowrap></td></tr>        
	<tr>
		<td align="<%=align_var%>" width="100%"> 
		<table border="0" cellpadding="1" cellspacing="3" width="100%" align="<%=align_var%>" dir="<%=dir_var%>">        
			<form name="AddAttach" action="adddoc.asp?add=1&companyID=<%=companyID%>" method="post" onsubmit="return checkfile();" target="_self" ENCTYPE="multipart/form-data">
			<TR>
				<TD align="<%=align_var%>" width=80%>					
				<INPUT dir="<%=dir_obj_var%>" type=text class="Form" name=document_name id="document_name" maxlength=100 style="width:250">
				</TD>
				<td width=20% align=left nowrap class="card">&nbsp;<span id=word2 name=word2><!--כותרת מסמך--><%=arrTitles(2)%></span>&nbsp;</td> 
			</TR>
			<TR>
				<TD align="<%=align_var%>" width=80%>					
				<textarea dir="<%=dir_obj_var%>" class="Form" name=document_desc id=document_desc maxlength=255 style="width:250"></textarea>
				</TD>
				<td width=20% align=left nowrap class="card" valign=top>&nbsp;<span id="word3" name=word3><!--תיאור מסמך--><%=arrTitles(3)%></span>&nbsp;</td> 
			</TR>
			<tr>
				<td align="<%=align_var%>" width="100%" colspan="2" class="card">
				<input dir="<%=dir_obj_var%>" type="file" name="document_file" style="width:340" maxlength=100>
				</td>
			</tr>
			<tr><td colspan=2 height=15 nowrap></td></tr>
			<tr>
            	<td colspan=2 align="center" nowrap dir=<%=dir_var%> class="card">
				<INPUT type=button class="but_menu" onclick="return exit_();" value="<%=arrButtons(2)%>" style="width:100px" ID="Button2" NAME="Button2">
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<INPUT type=submit class="but_menu" value="<%=arrButtons(1)%>" style="width:100px" ID="Button1" NAME="Button1">
				</td>
			</tr>
			</form>						
		</table>
	</td>
	</tr>
</table>
</body>
</html>	