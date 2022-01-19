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

<script LANGUAGE="JavaScript">
<!--
function getText(index,quest_id)
{
	var respText=''
	if(window.XMLHttpRequest && !(window.ActiveXObject))
	
	var xmlHTTP = new XMLHttpRequest();
    else
		if(window.ActiveXObject)
	
		var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	if (!xmlHTTP)
		var xmlHTTP = new ActiveXObject("Msxml2.XMLHTTP");
	if (xmlHTTP)
	{//console.log('getResponseContent.asp?OrgID=<%=OrgID%>&quest_id='+quest_id+'&indexResponse='+index)
		xmlHTTP.open("POST", 'getResponseContentXML.asp?OrgID=<%=OrgID%>&quest_id='+quest_id+'&indexResponse='+index, false);
   		xmlHTTP.send();	 
		result = new String(xmlHTTP.responseText);
		var parser=new DOMParser();
		var xmlDoc=parser.parseFromString(result,"text/xml")

		Response_Title=xmlDoc.getElementsByTagName("Response_Title")[0].textContent
		Response_Content=xmlDoc.getElementsByTagName("Response_Content")[0].textContent
		Response_File_1=xmlDoc.getElementsByTagName("Response_File_1")[0].textContent
		Response_File_2=xmlDoc.getElementsByTagName("Response_File_2")[0].textContent
		Response_File_3=xmlDoc.getElementsByTagName("Response_File_3")[0].textContent
		
		document.getElementById("SubjectText").value=Response_Title
		document.getElementById("ReplyText").value=Response_Content
		if (Response_File_1!=''){
		document.getElementById("txtFile1").value=Response_File_1
		document.getElementById("aFile1").href="../../../download/files/" + Response_File_1
		document.getElementById("aFile1").text=Response_File_1
		document.getElementById("aDelFile1").style.display="inline-block"
		}
		if (Response_File_2!=''){
	document.getElementById("txtFile2").value=Response_File_2
		document.getElementById("aFile2").href="../../../download/files/" + Response_File_2
		document.getElementById("aFile2").text=Response_File_2
		document.getElementById("aDelFile2").style.display="inline-block"
		}
		if (Response_File_3!=''){
	document.getElementById("txtFile3").value=Response_File_3
		document.getElementById("aFile3").href="../../../download/files/" + Response_File_3
		document.getElementById("aFile3").text=Response_File_3
		document.getElementById("aDelFile3").style.display="inline-block"
		}
	}
	 
}
function delTemplateFile(idF){
		document.getElementById("txtFile"+idF).value=''
		document.getElementById("aFile"+idF).href=''
		document.getElementById("aFile"+idF).text=''
		document.getElementById("aDelFile"+idF).style.display="none"
}

function CheckFields(action)
{	
	if (!checkEmail(document.frmMain.response_email.value) || document.frmMain.response_email.value == "")
	{
		<%
			If trim(lang_id) = "1" Then
				str_alert = "כתובת דואר  אלקטרוני לא חוקית"
			Else
				str_alert = "The email address is not valid!"
			End If	
		%>
		window.alert("<%=str_alert%>");
		document.frmMain.response_email.focus();
		return false;
	}
	
	if(window.document.frmMain.ReplyText.value.length > 5000)
	{
		window.alert("התוכן שהזנת הינו גדול ממספר התוים המקסימלי");
		return false;
	}
	  if(document.frmMain.attachment_file)
  {
   if (document.frmMain.attachment_file.value !='')
	{
		var fname=new String();
		var fext=new String();
		var extfiles=new String();
		fname=document.frmMain.attachment_file.value;
		indexOfDot = fname.lastIndexOf('.')
		fext=fname.slice(indexOfDot+1,-1)		
		fext=fext.toUpperCase();
		extfiles='HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT';		
		if ((extfiles.indexOf(fext)>-1) == false)
		{
		   <%
			If trim(lang_id) = "1" Then
				str_alert = ":סיומת הקובץ - אחת מרשימה \n HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT"
			Else
				str_alert = "The file ending should be one of the these \n HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT"
			End If	
	        %>	
			window.alert("<%=str_alert%>");
		    return false;
		}    
	  }   
   }
			
	if(action == "2") // שלח תגובה וסגור
	{
		document.frmMain.action = document.frmMain.action + "&close_flag=1";
		document.frmMain.submit();
		return true;
	}		
	
	else if(action == "1") //שלח תגובה
	{
		document.frmMain.submit();
		return true;
	}	
	return false;	
}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

function checkEmail(addr)
{
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}

//-->
</script>  

<%
  appID=request.querystring("appID")	
set upl=Server.CreateObject("SoftArtisans.FileUp")
	  if upl.Form("ReplyText")<>nil then 'after form filling
'TEMPLATES
File_Name1=""
File_Name2=""
File_Name3=""
  If  trim(upl.Form("attachment_file1").UserFilename) <> "" Then
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name1=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("attachment_file1").UserFileName,InstrRev(upl.Form("attachment_file1").UserFilename,"\")+1) 
  	'response.Write 	File_Name
 	'response.end
   		file_path="../../../download/send_attachments/" & File_Name1 
   		'Response.Write fs.FileExists(server.mappath(file_path))
        'Response.End
		'if fs.FileExists(server.mappath(file_path)) then
		'	file_path=file_path &"_1"
			'set a = fs.GetFile(server.mappath(file_path))
			'a.delete			
		'end if			
							
		upl.Form("attachment_file1").SaveAs server.mappath("../../../download/send_attachments/") & "/" & File_Name1
		if Err <> 0 Then
			 Response.Write("An error occurred when saving the file on the server.")			 
			 set upl = Nothing
			 Response.End
		end if
   		isfileTemplate1=false
   			strisfileTemplate1="0"
	else
		if upl.Form("txtFile1")<>"" then
			File_Name1=upl.Form("txtFile1")
   		
   			isfileTemplate1=true
   			strisfileTemplate1="1"
   		end if
	End If	


  If  trim(upl.Form("attachment_file2").UserFilename) <> "" Then
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name2=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("attachment_file1").UserFileName,InstrRev(upl.Form("attachment_file2").UserFilename,"\")+1) 
  	'response.Write 	File_Name
 	'response.end
   		file_path="../../../download/send_attachments/" & File_Name2 
   		'Response.Write fs.FileExists(server.mappath(file_path))
        'Response.End
		'if fs.FileExists(server.mappath(file_path)) then
		'	file_path=file_path &"_1"
			'set a = fs.GetFile(server.mappath(file_path))
			'a.delete			
		'end if			
							
		upl.Form("attachment_file2").SaveAs server.mappath("../../../download/send_attachments/") & "/" & File_Name2
		if Err <> 0 Then
			 Response.Write("An error occurred when saving the file on the server.")			 
			 set upl = Nothing
			 Response.End
		end if
   		isfileTemplate2=false
   			strisfileTemplate2="0"
	else
		if upl.Form("txtFile2")<>"" then
			File_Name2=upl.Form("txtFile2")
   			isfileTemplate2=true
   			
   			strisfileTemplate2="1"
   		end if
	End If	
	
  If  trim(upl.Form("attachment_file3").UserFilename) <> "" Then
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name3=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("attachment_file1").UserFileName,InstrRev(upl.Form("attachment_file3").UserFilename,"\")+1) 
  	'response.Write 	File_Name
 	'response.end
   		file_path="../../../download/send_attachments/" & File_Name3
   		'Response.Write fs.FileExists(server.mappath(file_path))
        'Response.End
		'if fs.FileExists(server.mappath(file_path)) then
		'	file_path=file_path &"_1"
			'set a = fs.GetFile(server.mappath(file_path))
			'a.delete			
		'end if			
							
		upl.Form("attachment_file3").SaveAs server.mappath("../../../download/send_attachments/") & "/" & File_Name3
		if Err <> 0 Then
			 Response.Write("An error occurred when saving the file on the server.")			 
			 set upl = Nothing
			 Response.End
		end if
   		isfileTemplate3=false
   			
   			strisfileTemplate3="0"
	else
		if upl.Form("txtFile3")<>"" then
			File_Name3=upl.Form("txtFile3")
   			isfileTemplate3=true  			
   			
   			strisfileTemplate3="1"
   		end if
	End If	
'============================================================================

  If  trim(upl.Form("attachment_file").UserFilename) <> "" Then
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.Form("attachment_file").UserFileName,InstrRev(upl.UserFilename,"\")+1) 
  	'response.Write 	File_Name
 	'response.end
   		file_path="../../../download/send_attachments/" & File_Name 
   		'Response.Write fs.FileExists(server.mappath(file_path))
        'Response.End
		'if fs.FileExists(server.mappath(file_path)) then
		'	file_path=file_path &"_1"
			'set a = fs.GetFile(server.mappath(file_path))
			'a.delete			
		'end if			
							
		upl.Form("attachment_file").SaveAs server.mappath("../../../download/send_attachments/") & "/" & File_Name
		if Err <> 0 Then
			 Response.Write("An error occurred when saving the file on the server.")			 
			 set upl = Nothing
			 Response.End
		end if			
	End If	
 'response.Write "gg"
 'response.end  	  
	  ReplyText = sFix(trim(upl.Form("ReplyText")))	 
	  SubjectText = sFix(trim(upl.Form("SubjectText")))	 
	  response_email = sFix(trim(upl.Form("response_email")))
	  sender_id = upl.Form("sender_id")
	  
	  'שמור תגובה בטבלת תגובות לטופס
	   sqlStr = "Insert Into appeal_responses (appeal_id ,response_content ,response_subject,response_date ,response_email, User_ID ,Organization_ID,response_attFile,response_attFile_1,response_attFile_2,response_attFile_3,isTemplateFile_1,isTemplateFile_2,isTemplateFile_3) "&_
	   " Values (" & appID & ",'" & ReplyText &"','" & SubjectText & "', GetDate(),'" & response_email & "','" & sender_id & "'," & OrgID & ",'" &  sFix(File_Name)& "','" &  sFix(File_Name1) & "','"&  sFix(File_Name2)& "','"&  sFix(File_Name3)& "'," & strisfileTemplate1 & "," & strisfileTemplate2 & "," & strisfileTemplate3 & ")"
	  'Response.Write sqlStr
	  'Response.End
	   con.ExecuteQuery(sqlStr)	    
	  
	  ' שליחת מייל ללקוח שהפניה נסגרה
	  If IsNumeric(appid) Then	    
	   
        sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
'	   Response.Write("sqlstr=" & sqlstr)
	'response.end
		set app = con.GetRecordSet(sqlstr)
		if not app.eof then
				
		appeal_date = FormatDateTime(app("appeal_date"), 2) & " " & FormatDateTime(app("appeal_date"), 4)	
		companyName = app("company_Name")
		contactName = app("contact_Name")
		projectName = app("project_Name")
		userName = app("user_Name")
		productName = app("product_Name")
		quest_id = app("questions_id")
		companyId = app("company_id")
		contactId = app("contact_id")
		projectId = app("project_id")
		appeal_status = app("appeal_status")	
		
		toMail = response_email
		
		If trim(sender_id) <> "" Then
			sqlstr = "Select EMAIL From USERS Where USER_ID = " & sender_id
			set rswrk = con.getRecordSet(sqlstr)
			If not rswrk.eof Then
			  fromMail = trim(rswrk("EMAIL"))
			End If
			set rswrk = Nothing	
		End If	
		
		sqlstr =  "Select Langu, Product_Name, RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		If not rsq.eof Then		
			Langu = trim(rsq(0))
			prodName = trim(rsq(1))			
			responsible_id = trim(rsq(2))
		End If
		set rsq = Nothing	
		if Langu = "eng" then
			td_align = "left"
			pr_language = "eng"
		else
			td_align = "right"
			pr_language = "heb"
		end if
		
		set app = nothing
		End If
		
	   If trim(Langu) = "heb" Then
			If len(companyName) > 0 Then
				mailSubject = " תגובה ל" & companyName & " - " & prodName
			Else
				mailSubject = " תגובה " & " - " & prodName
			End If
	   Else
   			If len(companyName) > 0 Then
				mailSubject = " Response to " & companyName & " - " & prodName
			Else
				mailSubject = " Response " & " - " & prodName
			End If
	   End If
	   
	   If trim(OrgID) <> "" Then
			sqlstring = "SELECT UseBizLogo FROM organizations WHERE ORGANIZATION_ID = " & OrgID
			'Response.Write sqlstring
			'Response.End
			set rs_tmp=con.GetRecordSet(sqlstring)
			if not rs_tmp.EOF  then				
				UseBizLogo = trim(rs_tmp("UseBizLogo"))
			else
				UseBizLogo = 0	
			end if
			set rs_tmp = Nothing	
	  End If
				
	  If Len(toMail) > 0 And IsNull(toMail) = false Then
	
	  If trim(Langu) = "heb" Then		
		strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & chr(10) & chr(13) &_
		"<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & chr(10) & chr(13)  &_			
		"<table border=0 width=100% cellspacing=0 cellpadding=0 align=right>"
        strBody=strBody & "<tr><td width=100% align=right valign=top height=20 nowrap></td></tr>"  & vbCrLf
     	If Len(ReplyText) > 0 Then
			strBody = strBody & "<tr><td align=""right"" width=100% ><span  dir=""rtl"">" & breaks(trim(ReplyText)) & "</span></td></tr>"
		End If	
		strBody = strBody & "<tr><td height=20></td></tr>" 
			'strBody = strBody & "<tr><td align=""center"" width=100% ><table border=0 cellpading=0 cellspacing=0><tr><td width=150>"
			'If trim(UseBizLogo) = "1" Then	
			'	strBody = strBody & "<a href=""http://pegasus.bizpower.co.il"" target=""_blank"">"
			'End If
			'strBody = strBody & "<img src=" & strLocal & "/images/Powered-By-BP2.gif border=0>"
			'If trim(UseBizLogo) = "1" Then	
			'	strBody = strBody & "</a>" 
			'End If					
			'		strBody = strBody & "</td><td align=right width='100%'></td></tr></table></td></tR>"  & vbCrLf
					If false then ' not (IsNull(File_Name) Or trim(File_Name) = "") Then
	        strBody = strBody & "<tr><td dir=ltr valign=top align=right colspan=2>" & vbCrLf &_
	        "<a class=""file_link"" href=""" & strLocal & "download/send_attachments/" & File_Name & """ target=_blank>" & File_Name & "</a>" & vbCrLf &_
	        "<b>מסמך</b></td></tr>" & vbCrLf
			End if

		'strBody = strBody &	"<tr><td height=30 nowrap colspan=2></td></tr>"  & chr(10) & chr(13) &_
		'			"<tr><td align=right dir=rtl width='100%' colspan=2>בברכה,<BR>צוות מכירות פגסוס ישראל<br>טל: 03-6374000</br><a href=https://www.pegasusisrael.co.il>www.pegasusisrael.co.il</a></td></tr>"& chr(10) & chr(13) &_
			strBody = strBody &	"</table></td></tr></table>"		
'response.Write strBody
'response.end
			Else
	    	'	strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1252"">" & chr(10) & chr(13) &_
			'	"<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & chr(10) & chr(13)  &_			
			'	"<table border=0 width=450 cellspacing=0 cellpadding=0 align=center>" & chr(10) & chr(13)  &_
			'	"<tr><td width=100% align=left valign=top height=20 nowrap></td></tr>" & chr(10) & chr(13)  &_
			'	"<tr><td width=100% style=""border: 1px solid #808080""><table cellpadding=0 cellspacing=0 width=100% >" & chr(10) & chr(13)  &_	
			'	"<tr><td width=100% align=left valign=top bgcolor=White>" & chr(10) & chr(13)  &_
			'	"<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=3>"
			'	If Len(ReplyText) > 0 Then
			'		strBody = strBody & "<tr><td align=""left"" bgcolor=#f0f0f0 width=100% class=""form"" style=""padding-left:15px""><span style=""width:470"" dir=""ltr"">" & breaks(trim(ReplyText)) & "</span></td></tr>"
			'	End If	
			'	strBody = strBody & "</table></td></tr></table></td></tr><tr><td height=20></td></tr>" & chr(10) & chr(13)  &_	
			'	"<tr bgcolor=#e6e6e6>" & chr(10) & chr(13)  &_	
			'	"<td class=title_form width=100% height=24 align=center dir=ltr> Response to your appeal from date " & appeal_date & " : " &_
			'	"</td></tr>" & chr(10) & chr(13) &_
			'	"<tr><td width=100% align=left valign=top bgcolor=#C9C9C9 style=""border: 1px solid #808080"">" & chr(10) & chr(13)  &_
			'	"<table cellpadding=0 cellspacing=0 width=100% >" & chr(10) & chr(13) &_
			'	"<tr><td width=100% align=left><TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=3 bgcolor=White>" & chr(10) & chr(13)
			'	sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','"& quest_id &"',''"
			'	set fields=con.GetRecordSet(sqlstr)
			'		do while not fields.EOF 
			'			Field_Value = trim(fields("Field_Value"))
			'			'Field_Id = fields("Field_Id")
			'			Field_Title = trim(breaks(fields("Field_Title")))
			'			Field_Type = fields("Field_Type")
			'			FIELD_EXCEPTION = fields("FIELD_EXCEPTION")
			'			Field_Align = trim(fields("Field_Align"))
			'			is_exception = false
			'			
			'			If pr_language = "eng" Then
			'				strBody = strBody & "<tr><td dir=ltr "
			'				If Field_Type <> "10" Then  'question
			'					strBody = strBody & " class=""form"""
			'				Else
			'					strBody = strBody & " class=""title_form"""
			'				End If
			'				strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px;""><b> " & breaks(fields("Field_Title")) & "</b></td></tr>" & chr(10) & chr(13)
			'			ElseIf pr_language = "heb" Then
			'				strBody = strBody & "<tr><td dir=ltr "
			'				If Field_Type <> "10" then  'question
			'					strBody = strBody & " class=""form"""
			'				Else
			'					strBody = strBody & " class=""title_form"""
			'				End If
			'				strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px;""><b> " & breaks(fields("Field_Title")) & "</b></td></tr>" & chr(10) & chr(13)
			'			End If
			'						
			'			If Field_Type <> "10" Then  'question
			'				strBody = strBody & "<tr><td width=""85%"" class=""form"" align=" & td_align & " bgcolor=#f0f0f0 style=""padding-left:15px;"""
			'				strBody = strBody & " dir=ltr"
			'				strBody = strBody & " >"
			'				If Field_Type = "1" Or Field_Type = "2" Or Field_Type = "3" Or  Field_Type = "6" Or Field_Type = "7" Or Field_Type = "8" Or Field_Type = "9" Or Field_Type = "11" Or Field_Type = "12" Then
			'					strBody = strBody & breaks(Field_Value)
			'		
			'				ElseIf Field_Type = "4" then  'scale with degree of importance
			'					If Field_Value <> "" then
			'					tmparr = Split(Field_Value," ")
			'					strValue = tmparr(0)
			'					If UBound(tmparr) > 0 Then
			'						strImportance = tmparr(1)
			'					Else
			'						strImportance = "H"
			'					End If
			'					select case strImportance
			'					case "H","ג"
			'						pf_exception = IMPORTANCE_H
			'					case "M","ב"
			'						pf_exception = IMPORTANCE_M
			'					case "L","נ"
			'						pf_exception = IMPORTANCE_L
			'					case "NA","לי"
			'						pf_exception = IMPORTANCE_NA
			'					case else
			'						pf_exception = IMPORTANCE_H
			''					end select
			''					is_exception = get_exception(pf_exception,strValue)
			'				end if
			'				if IsNumeric(strValue) then
			'					If pr_language = "eng" Then
			'					strBody = strBody & "Importance&nbsp;&nbsp;" & strImportance & "&nbsp;&nbsp;score&nbsp;&nbsp;"
			'					End If
			''					If is_exception=true Then
			'					strBody = strBody & " <font color=red>"
			'					end if
			'					strBody = strBody & "<b>" & strValue & "</b>"
			'					if is_exception=true then
			'					strBody = strBody & "</font>"
			'					end if
			'					if pr_language = "heb" then
			'					strBody = strBody & "ציון&nbsp;&nbsp;<b>" & strImportance & "</b>&nbsp;&nbsp;חשיבות&nbsp;&nbsp;"
			'					end if
			'				end if
			'			Elseif Field_Type = "5" Then  'check
			'				If Field_Value="on" Then
			'					strBody = strBody & "yes"						
			'				Else
			'					strBody = strBody & "no"
			''			End If					
			''		End If
			'		strBody = strBody & "</td></tr>" & chr(10) & chr(13)
			'		End If
			'		fields.moveNext()
			'		loop
			'		set fields=nothing
			'		strBody = strBody & "</TABLE></td></tr></TABLE></td></tr><tr><td align=""center"" width=100% height=10px></td></tr>" & chr(10) & chr(13) &_
			'		"<tr><td align=""center"" width=100% ><table cellpadding=0 cellspacing=0 border=0 align=center>" & chr(10) & chr(13) &_
			'		"<tr><td width=150>"
			'		If trim(UseBizLogo) = "1" Then	
			''			strBody = strBody & "<a href=""http://pegasus.bizpower.co.il"" target=""_blank"">"
			'		End If
			'		strBody = strBody & "<img src=" & strLocal & "/images/Powered-By-BP2.gif border=0>"
			'		If trim(UseBizLogo) = "1" Then	
			'			strBody = strBody & "</a>" 
			'		End If					
			'		strBody = strBody & "</td><td align=right width='100%'></td></tr>"  & chr(10) & chr(13) & _
			'		"<tr><td height=30 nowrap></td></tr>"  & chr(10) & chr(13) &_
			'		"<tr><td>בברכה,<BR>צוות מכירות פגסוס ישראל<br>טל: 03-6374000</br><a href=http://www.pegasusisrael.co.il>www.pegasusisrael.co.il</a></td></tr>"& chr(10) & chr(13) &_
			'		"</table></td></tr></table>"		
			End If			
		'response.Write strBody
		'response.end
			Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
				Msg.From = fromMail
				Msg.MimeFormatted = true
				if SubjectText<>"" then
				Msg.Subject=SubjectText
				else
				Msg.Subject = mailSubject
				end if
			
				Msg.To = toMail		
				Msg.HTMLBody = strBody
				Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"		
				If not (IsNull(File_Name) Or trim(File_Name) = "") Then
				'strAtt="http://bluto/bizpower_pegasus/download/send_attachments/1.txt"
				strAtt= strLocal & "download/send_attachments/"& File_Name
				'Msg.AddAttachment (strAtt)
				end if
				'add by Mila
				If not (IsNull(File_Name1) Or trim(File_Name1) = "") Then
					'strAtt="http://bluto/bizpower_pegasus/download/send_attachments/1.txt"
					if not isfileTemplate1 then
					strAtt= strLocal & "download/send_attachments/"& File_Name1
					else
					strAtt= strLocal & "download/files/"& File_Name1
					end if
					'Msg.AddAttachment (strAtt)
				end if
				'add by Mila
				If not (IsNull(File_Name2) Or trim(File_Name2) = "") Then
					'strAtt="http://bluto/bizpower_pegasus/download/send_attachments/1.txt"
					if not isfileTemplate2 then
					strAtt= strLocal & "download/send_attachments/"& File_Name2
					else
					strAtt= strLocal & "download/files/"& File_Name2
					end if
					'Msg.AddAttachment (strAtt)
				end if
				'add by Mila
				If not (IsNull(File_Name3) Or trim(File_Name3) = "") Then
					'strAtt="http://bluto/bizpower_pegasus/download/send_attachments/1.txt"
					if not isfileTemplate3 then
					strAtt= strLocal & "download/send_attachments/"& File_Name3
					else
					strAtt= strLocal & "download/files/"& File_Name3
					end if
					'Msg.AddAttachment (strAtt)
				end if
				Msg.Send()						
			Set Msg = Nothing		
		End If
   End If
	%>	
		<SCRIPT LANGUAGE=javascript>
			<!--			
			<% If trim(Request.QueryString("close_flag")) = "1" Then %>
			document.location.href = "appeal_close.asp?appId=<%=appId%>";
			<%Else%>
			window.close();
			<% End If%>
			window.opener.document.location.reload(true);
			//-->
		</SCRIPT>
   <%				
End If

if appID<>nil and appID<>"" then
  sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
'Response.Write sqlstr
'response.end

  set rs_appeals = con.GetRecordSet(sqlStr)

	if not rs_appeals.eof then
		companyId = rs_appeals("company_id")
		contactId = rs_appeals("contact_id")
		quest_id = rs_appeals("questions_id")	
		sqlstr =  "Select RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		If not rsq.eof Then						
			responsible_id = trim(rsq(0))
		End If
		set rsq = Nothing
		If IsNumeric(responsible_id) = false Then
			responsible_id = UserID
		End If					
	end if
	set rs_appeals = nothing	 	
	
	If trim(contactId) <> "" Then
		sqlstr = "Select email from contacts WHERE contact_Id = " & contactId
	
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			reponse_email = trim(rs_pr(0))				
		End If
		set rs_pr = Nothing
	End If
end if
'Response.Write company_id
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 63 Order By word_id"				
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
%>
<body style="margin:0px; background-color:#E6E6E6" onload="window.focus();">
<table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>" ID="Table1">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table2">
	  <tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;שליחת אימייל<!--שליחת תגובה--><%'=arrTitles(12)%></td></tr>		   
   </table>
</td></tr> 
<tr><td width=100%> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="98%" align=center ID="Table3">
<tr><td height=10></td></tr>          
<FORM name="frmMain" ACTION="send_replytest.asp?appID=<%=appID%>" METHOD="post"  enctype="multipart/form-data"  onSubmit="return CheckFields()" ID="Form1">
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<span id="word7" name=word7><!--כתובת אימייל למשלוח תגובה--><%=arrTitles(7)%></span>&nbsp;
	<input dir="ltr" type="text" class="texts" style="width:285" id="response_email" name="response_email" value="<%=reponse_email%>" maxlength=50>
	</td>
	<td align="<%=align_var%>" nowrap width="70">&nbsp;<span id="word8" name=word8><!--אל--><%=arrTitles(8)%></span>&nbsp;</td>
</tr>
<tr>
<td align="<%=align_var%>" width=100% valign=top>
<!--הכתובת שממנה יוצאת התגובה--><%=arrTitles(11)%>&nbsp;
<select name="sender_id" dir="ltr" class="texts" style="width:285;" ID="sender_id">   
    <%
 
    set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME,EMAIL FROM Users WHERE ORGANIZATION_ID = " & OrgID & "  AND ACTIVE = 1  ORDER BY FIRSTNAME + ' ' + LASTNAME")
      do while not UserList.EOF
    selUserID=trim(UserList(0))
    selUserName=trim(UserList(1))
    EMAIL=trim(UserList(2))
    %>
    <option value="<%=selUserID%>" <%if trim(responsible_id)=trim(selUserID) then%> selected <%end if%> dir="<%=dir_var%>"><%=EMAIL%>&nbsp;(<%=selUserName%>)</option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
</select>
</td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;<span id="word9" name=word9><!--מאת--><%=arrTitles(9)%></span>&nbsp;</td>
</tr>
<TR>	<td align="<%=align_var%>" width=100%>	
	<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=1 id="SubjectText" name="SubjectText"></textarea>
	</td>
	<td align="<%=align_var%>" nowrap width="70">&nbsp;<!--תוכן תגובה-->נושא&nbsp;</td>
</TR>
<tr>
<td align="<%=align_var%>" width=100% valign=top>

<select name="responses" dir="<%=dir_obj_var%>" class="texts" style="width:450;" ID="responses" onchange="getText(this.selectedIndex,<%=quest_id%>)">
<option value=""><!--בחר מענה אוטומטי--><%=arrTitles(10)%></option>   
    <%sqlStr = "Select Response_Id, Response_Title, Response_Content from Product_Responses WHERE ORGANIZATION_ID= "& OrgID &_
	" And Product_ID = " & quest_id & " Order By Response_Title"
	'Response.Write sqlStr
	'response.end
	set rs_Responses = con.GetRecordSet(sqlStr)
	If not rs_Responses.eof Then
	while not rs_Responses.eof
		Response_Id = rs_Responses("Response_Id")
		Response_Title = rs_Responses("Response_Title")
		Response_Content = rs_Responses("Response_Content")
    %>
    <option value="<%=vFix(Response_Content)%>"><%=Response_Title%></option>
    <%
		rs_Responses.movenext				
	Wend			
	set rs_Responses = nothing			
	End If%>
</select>
</td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;</td>
</tr>


<tr valign=top>
	<td align="<%=align_var%>" width=100%>	
	<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=9 id="ReplyText" name="ReplyText"></textarea>
	</td>
	<td align="<%=align_var%>" nowrap width="70">&nbsp;תוכן אימייל<!--תוכן תגובה--><%'=arrTitles(13)%>&nbsp;</td>
</tr>
<%'If IsNull(attachment) Or trim(attachment) = "" Then%>
<!--Added by Mila 02/11/2020 - files from template response-->

<tr><td  dir="<%=dir_obj_var%>">
<input type="file" name="attachment_file1" ID="attachment_file1" size=33 value="">
&nbsp;
<a href="" target="_blank" id="aFile1"></a><input type="hidden" id="txtFile1" name="txtFile1">
<a href="" target="_blank" id="aDelFile1" onclick="delTemplateFile('1');return false" style="display:none" title="מחק קובץ מצורף מובנה"><img src="../../images/delete_icon.gif" border="0"></a>

</td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>
<tr><td  dir="<%=dir_obj_var%>">
<input type="file" name="attachment_file2" ID="attachment_file2" size=33 value="">
&nbsp;
<a href="" target="_blank" id="aFile2"></a><input type="hidden" id="txtFile2" name="txtfile2">
<a href="" target="_blank" id="aDelFile2" onclick="delTemplateFile('2');return false" style="display:none" title="מחק קובץ מצורף מובנה"><img src="../../images/delete_icon.gif" border="0"></a>
</td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>
<tr><td  dir="<%=dir_obj_var%>">
<input type="file" name="attachment_file3" ID="attachment_file3" size=33 value="">
&nbsp;
<a href="" target="_blank" id="aFile3"></a><input type="hidden" id="txtFile3" name="txtFile3">
&nbsp;
<a href="" target="_blank" id="aDelFile3" onclick="delTemplateFile('3');return false" style="display:none" title="מחק קובץ מצורף מובנה"><img src="../../images/delete_icon.gif" border="0"></a>
</td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>
<tr><td colspan="2" height=10></td></tr>
<!--Added by Mila 02/11/2020 - files from template response-->

<tr><td  dir="<%=dir_obj_var%>"><input type="file" name="attachment_file" ID="attachment_file" size=33 value=""></td>
<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
</tr>
<tr><td colspan="2" height=10></td></tr>
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center width=90% align=left ID="Table4">
<tr valign=top>
<td width=28% align=center><A class="but_menu" href="#" style="width:100px" onclick="window.close();"><span id=word4 name=word4><!--ביטול--><%=arrTitles(4)%></span></A></td>
<td width=10 nowrap></td>
<td width=44% align="center"><A class="but_menu" style="width:180px;" href="#" onclick="return CheckFields(2);">שלח אימייל וסגור פניה<!--סגור ושלח תגובה--><%'=arrTitles(6)%></A></td>
<td width=10 nowrap></td>
<td width=28% align=center><A class="but_menu" style="width:120px" href="#" onclick="return CheckFields(1);">שלח אימייל<!--שלח תגובה--><%'=arrTitles(14)%></A></td>
</tr>
</table></td></tr>
</form>
<tr><td colspan="2" height=5></td></tr>
</table>
</td></tr></table>
</body> 
<%
set con = nothing
%>
</html>
