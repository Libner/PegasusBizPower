<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--INCLUDE FILE="../checkWorker.asp"-->
<%OrgID=264
lang_id = 1
dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
%>
<html>
<head>
<!-- include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>

<%
     If Request.QueryString("delFile") <> nil Then
		delfile=trim(Request.QueryString("delFile"))
		if delfile<>"" then
			prodID = trim(Request.QueryString("prodID"))
			responseID= trim(Request.QueryString("responseID"))
			'NOT TO DELETE FILE FROM FOLDER - -REMIND FOR RESPONSES HISTORY
			'sqlstr="Select Response_File_" & delfile & " From Product_Responses Where Response_ID = " &_
			'responseID & " And Product_ID = " & prodID
			'set rssub = con.getRecordSet(sqlstr)
			'If not rssub.eof Then
            '     Response_File  = Trim(rssub("Response_File_" & delfile ))
            '     set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
            '     file_path="../../../download/files/" & Response_File 
   			'	if fs.FileExists(server.mappath(file_path)) then
			'		set a = fs.GetFile(server.mappath(file_path))
			'		a.delete
			'	end if 
			'end if   
			'set rssub = nothing	
			sqlstr="Update Product_Responses set Response_File_" & delfile & "=NULL Where Response_ID = " & responseID & " And Product_ID = " & prodID
			con.executeQuery(sqlstr)                   
		end if
     end if
     
set upl=Server.CreateObject("SoftArtisans.FileUp")
     If Request.QueryString("add") <> nil Then
     Response_File_1=""
     Response_File_2=""
     Response_File_3=""
       If  trim(upl.Form("fileuploadF1").UserFilename) <> "" Then
			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   			File_Name=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("fileuploadF1").UserFileName,InstrRev(upl.UserFilename,"\")+1) 
  		'response.Write 	File_Name
 		'response.end
   			file_path="../../../download/files/" & File_Name 
   			'if fs.FileExists(server.mappath(file_path)) then
			'	file_path=file_path &"_1"
				'set a = fs.GetFile(server.mappath(file_path))
				'a.delete			
			'end if			
			Response_File_1=File_Name				
			upl.Form("fileuploadF1").SaveAs server.mappath("../../../download/files/") & "/" & File_Name
			strUpd=strUpd & ",Response_File_1='" & sfix(Response_File_1) & "'"
			strInsFields=strInsFields & ",Response_File_1"
			strInsValues=strInsValues & ",'" & sfix(Response_File_1) & "'"
		End If	
       If  trim(upl.Form("fileuploadF2").UserFilename) <> "" Then
			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   			File_Name=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("fileuploadF2").UserFileName,InstrRev(upl.UserFilename,"\")+1) 
  		'response.Write 	File_Name
 		'response.end
   			file_path="../../../download/files/" & File_Name 
   			'Response.Write fs.FileExists(server.mappath(file_path))
			'Response.End
			'if fs.FileExists(server.mappath(file_path)) then
			'	file_path=file_path &"_1"
				'set a = fs.GetFile(server.mappath(file_path))
				'a.delete			
			'end if			
			Response_File_2=File_Name			
			upl.Form("fileuploadF2").SaveAs server.mappath("../../../download/files/") & "/" & File_Name
			strUpd=strUpd & ",Response_File_2='" & sfix(Response_File_2) & "'"
			strInsFields=strInsFields & ",Response_File_2"
			strInsValues=strInsValues & ",'" & sfix(Response_File_2) & "'"
		End If	
       If  trim(upl.Form("fileuploadF3").UserFilename) <> "" Then
			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   			File_Name=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("fileuploadF3").UserFileName,InstrRev(upl.UserFilename,"\")+1) 
  		'response.Write 	File_Name
 		'response.end
   			file_path="../../../download/files/" & File_Name 
   			'Response.Write fs.FileExists(server.mappath(file_path))
			'Response.End
			'if fs.FileExists(server.mappath(file_path)) then
			'	file_path=file_path &"_1"
				'set a = fs.GetFile(server.mappath(file_path))
				'a.delete			
			'end if			
			
			Response_File_3=File_Name						
			upl.Form("fileuploadF3").SaveAs server.mappath("../../../download/files/") & "/" & File_Name
			strUpd=strUpd & ",Response_File_3='" & sfix(Response_File_3) & "'"
			strInsFields=strInsFields & ",Response_File_3"
			strInsValues=strInsValues & ",'" & sfix(Response_File_3) & "'"
		End If	
       prodID = trim(upl.Form("prodID"))
		If trim(upl.Form("responseID")) = "" Then ' add type
			sqlstr = "Insert into Product_Responses (Product_ID,Organization_ID,Response_Title,Response_Content" & strInsFields & ") values (" &_
			prodID & "," & OrgID & ",'" & sFix(upl.Form("Response_Title")) & "','" & sFix(upl.Form("Response_Content")) & "'" & strInsValues & ")"
			'Response.Write sqlstr
			'Response.End
			con.executeQuery(sqlstr) 
			sqlstr="select TOP 1 Response_ID From Product_Responses ORDER BY Response_ID DESC"
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				responseID=cstr(rssub("Response_ID"))
			end if
			%>
		
	<%	Else ' update type
			responseID = trim(upl.Form("responseID"))
			sqlstr="Update Product_Responses set Response_Title = '" & sFix(upl.Form("Response_Title")) &_
			"', Response_Content = '" & sFix(upl.Form("Response_Content")) &_
			"' " & strUpd  &_
			" Where Response_ID = " & responseID & " And Product_ID = " & prodID
			'response.Write sqlstr
			con.executeQuery(sqlstr) %>
	<%	End If%>
	
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.reload(true);
				window.close();
			//-->
			</SCRIPT>
	<%End If
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 71 Order By word_id"				
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

<script>
function checkForm()
		{
			if(window.document.getElementById("Response_Title").value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס מענה "
				Else
					str_alert = "Please insert the reply !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.getElementById("Response_Title").focus();
				return false;
			}
			if(window.document.getElementById("Response_Content").value.length > 5000)
			{
				window.alert("התוכן שהזנת הינו גדול ממספר התוים המקסימלי");
				return false;
			}
			if(!checkFile("fileuploadF1"))
			{
					return false;
			}
			if(!checkFile("fileuploadF2"))
			{
					return false;
			}
			if(!checkFile("fileuploadF3"))
			{
					return false;
			}
			
			return true;			
		}

		function checkFile(fileId)
		{
		if (window.document.getElementById(fileId).value  != '')
			{
				var filename = new String(window.document.getElementById(fileId).value);
				filename = filename.toLowerCase();
				if (
				(filename.lastIndexOf(".doc") == -1) && (filename.lastIndexOf(".pdf") == -1)
				&& (filename.lastIndexOf(".rtf") == -1)  && (filename.lastIndexOf(".ppt") == -1)
				&& (filename.lastIndexOf(".docx") == -1) && (filename.lastIndexOf(".pptx") == -1)
				&& (filename.lastIndexOf(".gif") == -1) && (filename.lastIndexOf(".jpg") == -1)
				&& (filename.lastIndexOf(".jpeg") == -1)  && (filename.lastIndexOf(".bmp") == -1)
				&& (filename.lastIndexOf(".png") == -1)
				)
				{
					window.alert("סיומת קובץ אינה חוקית, אנא בחר קובץ חוקי");
					window.document.getElementById(fileId).focus();
					return false;
				}
			}
			return true				
		}
		
		function CheckDel(ifF,responseID,prodID)
		{
			if (confirm("?האם ברצונך למחוק את קובץ"))
			{      
				 window.document.location.href="addresponse.asp?responseID=" + responseID + "&prodID=" + prodID + "&delFile=" + ifF
			}
			return false
		}
		</script>
		
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
    prodID = trim(Request.QueryString("prodID"))
	If Request.QueryString("responseID") <> nil Then
		responseID = trim(Request.QueryString("responseID"))		
		If Len(responseID) > 0 Then
			sqlstr="Select Response_Title, Response_Content,Response_File_1,Response_File_2,Response_File_3 From Product_Responses Where Response_ID = " &_
			responseID & " And Product_ID = " & prodID
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				Response_Title = trim(rssub("Response_Title"))
				Response_Content = trim(rssub("Response_Content"))
                                Response_File_1  = Trim(rssub("Response_File_1"))
                                Response_File_2 = Trim(rssub("Response_File_2"))
                                Response_File_3  = Trim(rssub("Response_File_3"))
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table2">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(responseID) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word3 name=word3><!--מענה--><%=arrTitles(3)%></span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="100%" cellspacing="1" cellpadding="2" align=center border="0" ID="Table3">
<form name=form1 id=form1 action="addresponse.asp?add=1" target="_self" method="post"  enctype="multipart/form-data">
<input type=hidden name=responseID id=responseID value="<%=responseID%>">
<input type=hidden name=prodID id="prodID" value="<%=prodID%>">
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<input type="text" class="texts" name="Response_Title" id="Response_Title" value="<%=vFix(Response_Title)%>" dir="<%=dir_obj_var%>" style="width:450" maxLength=100>	
	</td>
	<td width="70" nowrap align="<%=align_var%>" valign=top>&nbsp;<b><span id=word4 name=word4><!--כותרת--><%=arrTitles(4)%></span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<textarea class="texts" name="Response_Content" id="Response_Content" dir="<%=dir_obj_var%>" style="width:450" rows=9><%=(Response_Content)%></textarea>
	</td>
	<td width="70" nowrap align="<%=align_var%>" valign=top>&nbsp;<b><span id="word5" name=word5><!--תוכן--><%=arrTitles(5)%></span></b>&nbsp;</td>	
</tr>

<tr><td  dir="<%=dir_obj_var%>">
<%if Response_File_1<>"" then%>
<a href="" onclick="CheckDel('1','<%=ResponseId%>','<%=prodID%>');return false"><img src="../../images/delete_icon.gif" border="0"></a>
<a ID="aFile1" href="../../../download/files/<%=Response_File_1%>" target="_blank"><%=Response_File_1%></a>
<br>
<%end if%>										
<input type="file" name="fileuploadF1" ID="fileuploadF1" size=33 value=""></td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>

<tr><td height=5 colspan="2" nowrap></td></tr>
<tr><td  dir="<%=dir_obj_var%>">
<%if Response_File_2<>"" then%>
<a href="" onclick="CheckDel('2','<%=ResponseId%>','<%=prodID%>');return false"><img src="../../images/delete_icon.gif" border="0"></a>
<a ID="aFile2" href="../../../download/files/<%=Response_File_2%>" target="_blank"><%=Response_File_2%></a>
<br>
<%end if%>										
<input type="file" name="fileuploadF2" ID="fileuploadF2" size=33 value=""></td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>
<tr><td height=5 colspan="2" nowrap></td></tr>

<tr><td  dir="<%=dir_obj_var%>">
<%if Response_File_3<>"" then%>
<a href="" onclick="CheckDel('3','<%=ResponseId%>','<%=prodID%>');return false"><img src="../../images/delete_icon.gif" border="0"></a>
<a ID="aFile3" href="../../../download/files/<%=Response_File_3%>" target="_blank"><%=Response_File_3%></a>
<br>
<%end if%>										
<input type="file" name="fileuploadF3" ID="fileuploadF3" size=33 value=""></td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>
<tr><td height=5 colspan="2" nowrap></td></tr>


<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</form>
</table>
</td></tr></table>
</BODY>
</HTML>
