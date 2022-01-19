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
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</head>
	<script LANGUAGE="JavaScript">
<!--
function CheckFields(action)
{	
	if (!checkEmail(document.frmMain.sendermail.value) || document.frmMain.sendermail.value == "")
	{
		<%
			If trim(lang_id) = "1" Then
				str_alert = "כתובת דואר  אלקטרוני לא חוקית"
			Else
				str_alert = "The email address is not valid!"
			End If	
		%>
		window.alert("<%=str_alert%>");
		document.frmMain.sendermail.focus();
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
	//alert("2222")
		var fname=new String();
		var fext=new String();
		var extfiles=new String();
		//fname=document.frmMain.attachment_file.value;
		fname=document.getElementById("attachment_file").files[0].name; 
		//alert(fname)
		indexOfDot = fname.lastIndexOf('.')
		//alert(indexOfDot)	
		fext=fname.slice(indexOfDot+1)		
		fext=fext.toUpperCase();
		extfiles='HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT';	
		//alert(fext)	
		//alert(extfiles)
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
			
	 if(action == "1") //שלח תגובה
	{
	//alert("1")
	//	document.getElementById("frmMain").submit();
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
  'appID=request.querystring("appID")
  appealsId=	request.querystring("appealsId")
  quest_id=Request.QueryString("quest_id")
'  response.Write "quest_id="&quest_id
  
set upl=Server.CreateObject("SoftArtisans.FileUp")


	  if upl.Form("ReplyText")<>nil then 'after form filling


  If  trim(upl.UserFilename) <> "" Then
  response.Write "upl.UserFilename=" & upl.UserFilename
  'response.End 
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1) 
     		file_path="../../../download/send_attachments/" & File_Name 
  							
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
	  sender_name = sFix(trim(upl.Form("sender_name")))
	  sendermail = upl.Form("sendermail") 'mail
	  selUserID = Request.Cookies("bizpegasus")("UserId")
	  
 		strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & chr(10) & chr(13) &_
		"<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & chr(10) & chr(13)  &_			
		"<table border=0 width=100% cellspacing=0 cellpadding=0 align=right>"
        strBody=strBody & "<tr><td width=100% align=right valign=top height=20 nowrap></td></tr>"  & vbCrLf
     	If Len(ReplyText) > 0 Then
			strBody = strBody & "<tr><td align=""right"" width=100% ><span  dir=""rtl"">" & breaks(trim(ReplyText)) & "</span></td></tr>"
		End If	
		strBody = strBody & "<tr><td height=20></td></tr>" 
				
		'strBody = strBody &	"<tr><td height=30 nowrap colspan=2></td></tr>"  & chr(10) & chr(13) &_
		'			"<tr><td align=right dir=rtl width='100%' colspan=2>בברכה,<BR>צוות מכירות פגסוס ישראל<br>טל: 03-6374000</br><a href=http://www.pegasusisrael.co.il>www.pegasusisrael.co.il</a></td></tr>"& chr(10) & chr(13) &_
		'			"</table></td></tr></table>"		
					strBody = strBody &	"</table></td></tr></table>"		

	
	sqlstr = "Select appeal_id,appeals.Company_Id,appeals.Contact_Id,email from appeals left join Contacts on appeals.Contact_id=contacts.Contact_Id WHERE  len(email)>0 and appeal_id IN (" & appealsId & ")"
'response.Write sqlstr
'response.end
set rs_Contacts = con.getRecordSet(sqlstr)	
	do while not rs_Contacts.eof
	  companyID=rs_Contacts("Company_Id")
	contactID=rs_Contacts("contact_id")
	email=rs_Contacts("email")
	appID=rs_Contacts("appeal_id")

	  'שמור תגובה בטבלת תגובות לטופס
	   sqlStr1 = "Insert Into appeal_responses (appeal_id ,response_content ,response_subject,response_date ,response_email, FromName,FromEmail,User_ID ,Organization_ID,response_attFile,response_sendGroup) "&_
	   " Values (" & appID & ",'" & ReplyText &"','" & SubjectText & "', GetDate(),'" & email &"','"& sender_name &"','"& sendermail & "','" & selUserID & "'," & OrgID & ",'"&  sFix(File_Name)& "',1)"
	'Response.Write sqlStr1
	'Response.End
	   con.ExecuteQuery(sqlStr1)	    
	  
	  ' שליחת מייל ללקוח שהפניה נסגרה
	   
   	
			Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
				'response.Write "sendermail="& sendermail
				'response.end 
				'Msg.From = sendermail &""& sender_name
				Msg.From =sender_name &"<"&sendermail &">"
				Msg.MimeFormatted = true
			
				Msg.Subject=SubjectText
				
			
				'Msg.To = toMail		
				Msg.To = email '"faina@cyberserve.co.il"
			'	Msg.To = "faina@cyberserve.co.il"
				Msg.HTMLBody = strBody
				Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"		
'response.Write "strAtt="& strAtt
			'	response.end
				If not (IsNull(File_Name) Or trim(File_Name) = "") Then
				'response.Write "strLocal="& strLocal
				
				'strAtt="http://bluto/bizpower_pegasus/download/send_attachments/1.txt"
				strAtt= strLocal & "download/send_attachments/"& File_Name
				'response.Write strAtt
				'response.end
				'strAtt= http://pegasus.bizpower.co.il/download/send_attachments/"& File_Name
				'strAtt= http://pegasuslocal.bizpower.co.il/download/send_attachments/"& File_Name
				'http://pegasuslocal.bizpower.co.il:80/download/send_attachments/52_16_1.txt
'				response.Write "File_Name="& strAtt 
'				response.end
'response.Write "strAtt="& strAtt
'response.end
				Msg.AddAttachment (strAtt)
				end if
				Msg.Send()						
			Set Msg = Nothing		
			rs_Contacts.MoveNext
			loop
	%>
	<SCRIPT LANGUAGE="javascript">
			<!--			
		
		window.close();
		
		//	window.opener.document.location.reload(true);
			//-->
	</SCRIPT>
	<%				
End If


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
				<td align="left" valign="middle" nowrap>
					<table width="100%" border="0" cellpadding="0" cellspacing="0" ID="Table2">
						<tr>
							<td class="page_title" dir="<%=dir_obj_var%>">&nbsp;<!--שליחת תגובה--><%=arrTitles(12)%></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="100%">
					<table align="center" border="0" cellpadding="3" cellspacing="1" width="98%" align="center"
						ID="Table3">
						<tr>
							<td height="10"></td>
						</tr>
						<FORM name="frmMain" ACTION="sendMailGroup.asp?appealsId=<%=appealsId%>&quest_id=<%=quest_id%>" METHOD="post"  enctype="multipart/form-data"  onSubmit="return CheckFields()" ID="Form1">
							<%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME,EMAIL FROM Users WHERE User_id="& Request.Cookies("bizpegasus")("UserId"))
  if not UserList.eof then
     selUserID=trim(UserList(0))
    selUserName=trim(UserList(1))
    EMAIL=trim(UserList(2))
    end if
       set UserList=Nothing%>
							<tr>
								<td align="<%=align_var%>" width=100% valign=top>
									<span id="word7" name="word7"><!--כתובת אימייל למשלוח תגובה-->שם שולח אימייל</span>&nbsp;
									<input dir="ltr" type="text" class="texts" style="width:285" id="sender_name" name="sender_name" value="<%=selUserName%>" maxlength=50>
								<input type=hidden id="senderId" name="senderId" value="<%=Request.Cookies("bizpegasus")("UserId")%>">
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;<span id="word8" name="word8">מאת</span>&nbsp;</td>
							</tr>
							<tr>
								<td align="<%=align_var%>" width=100% valign=top>כתובת שממנה יוצאת אימייל&nbsp; <input id="sendermail" name="sendermail" dir="ltr" class="texts" style="width:285;" value="<%=EMAIL%>">
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;<span id="word9" name="word9"><!--מאת--><%=arrTitles(9)%></span>&nbsp;</td>
							</tr>
							<TR>
								<td align="<%=align_var%>" width=100%>
									<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=1 id="SubjectText" name="SubjectText"></textarea>
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;<!--תוכן תגובה--> נושא&nbsp;</td>
							</TR>
							<tr>
								<td align="<%=align_var%>" width=100% valign=top>
									<select name="responses" dir="<%=dir_obj_var%>" class="texts" style="width:450;" ID="responses" onchange="ReplyText.value=this.value;SubjectText.innerText=this.options[responses.selectedIndex].text;">
										<option value=""><!--בחר מענה אוטומטי--><%=arrTitles(10)%></option>
										<%sqlStr = "Select Response_Id, Response_Title, Response_Content from Product_Responses WHERE ORGANIZATION_ID= "& OrgID &_
	" And Product_ID = " & quest_id & " Order By Response_Title"
	'Response.Write sqlStr
'	response.end
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
							<tr valign="top">
								<td align="<%=align_var%>" width=100%>
									<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=9 id="ReplyText" name="ReplyText"></textarea>
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;תוכן אימייל&nbsp;</td>
							</tr>
							<%'If IsNull(attachment) Or trim(attachment) = "" Then%>
							<tr>
								<td  dir="<%=dir_obj_var%>"><input type="file" name="attachment_file" ID="attachment_file" size="33" value=""></td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
							</tr>
							<tr>
								<td colspan="2" height="10"></td>
							</tr>
							<tr>
								<td colspan="2">
									<table cellpadding="0" cellspacing="0" align="center" width="90%" align="left" ID="Table4">
										<tr valign="top">
											<td width="28%" align="center"><A class="but_menu" href="#" style="width:100px" onclick="window.close();"><span id="word4" name="word4"><!--ביטול--><%=arrTitles(4)%></span></A></td>
											<td width="10" nowrap></td>
											<td width="28%" align="center"><A class="but_menu" style="width:120px" href="#" onclick="return CheckFields(1);">שלח 
													אימייל</A></td>
										</tr>
									</table>
								</td>
							</tr>
						</FORM>
						<tr>
							<td colspan="2" height="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
	<%
set con = nothing
%>
</html>
