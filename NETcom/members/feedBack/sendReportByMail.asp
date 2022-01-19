<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--INCLUDE FILE="../checkWorker.asp"-->
<%
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
	
	if (window.document.getElementById("ReplyText").value == '') {
              window.alert("! נא להכניס תוכן ");
              window.document.getElementById("ReplyText").focus();
              return false;
          }
          if (window.document.getElementById("attachment_file").value == '') {
              window.alert("! נא להכניס  מסמך מצורף ");
              window.document.getElementById("attachment_file").focus();
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
			
	
			
	 if(action == "1") //שלח תגובה
	{
		document.frmMain.submit();
		return true;
	}	
	return false;	
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
  gId=	request.querystring("gId")
  gyear=Request.QueryString("y")
  sqlstrGuide = "SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name,Guide_Email  FROM Guides  where Guide_Id=" & gId

	set rs_Guide= conPegasus.Execute(sqlstrGuide)
   if not rs_Guide.EOF then
	
		Guide_Name =rs_Guide("Guide_Name")
		Guide_Email=rs_Guide("Guide_Email")
end if
set upl=Server.CreateObject("SoftArtisans.FileUp")


	  if upl.Form("ReplyText")<>nil then 'after form filling
'	  response.Write ("=="& upl.UserFilename)
'	  response.end
 If  trim(upl.UserFilename) <> "" Then
	    set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   		File_Name=Minute(Now()) & "_" & Second(Now()) & "_" & Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1) 
     		file_path="../../../download/send_feeadback/" & File_Name 
  							
		upl.Form("attachment_file").SaveAs server.mappath("../../../download/send_feeadback/") & "/" & File_Name
		if Err <> 0 Then
			 Response.Write("An error occurred when saving the file on the server.")			 
			 set upl = Nothing
			 Response.End
		end if			
	End If	

 
'response.Write "gg="&File_Name
' response.end  	  
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
				
		strBody = strBody &	"<tr><td height=30 nowrap colspan=2></td></tr>"  & chr(10) & chr(13) &_
					"</table></td></tr></table>"		
	
	

	   
   	
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
				'response.Write 	(upl.Form("Guide_Email"))
				'response.end
				Msg.To = upl.Form("Guide_Email") '"faina@cyberserve.co.il"
				'Msg.To ="faina@cyberserve.co.il"
				Msg.HTMLBody = strBody
				Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"	
			
				If not (IsNull(File_Name) Or trim(File_Name) = "") Then
				''strAtt="http://bluto/bizpower_pegasus/download/send_attachments/1.txt"
				'strAtt= strLocal & "download/send_attachments/"& File_Name
					Msg.AddAttachment Server.MapPath("../../../download/send_feeadback/") & "/" & File_Name
	
				'				Msg.AddAttachment (strAtt)
				end if
			
				Msg.Send()						
			Set Msg = Nothing		
			 If  trim(upl.UserFilename) <> "" Then		
		if fs.FileExists(Server.MapPath("../../../download/send_feeadback" & "/" & File_Name)) then
				set a = fs.GetFile(Server.MapPath("../../../download/send_feeadback" & "/" & File_Name))
			
			
			a.delete
			end if			
		end if		
	%>
	<SCRIPT LANGUAGE="javascript">
			<!--			
		
		window.close();
		
			window.opener.document.location.reload(true);
			//-->
	</SCRIPT>
	<%				
End If


'Response.Write company_id
%>
	<%
	
%>
	<body style="margin:0px; background-color:#E6E6E6" onload="window.focus();">
		<table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>" ID="Table1">
			<tr>
				<td align="left" valign="middle" nowrap>
					<table width="100%" border="0" cellpadding="0" cellspacing="0" ID="Table2">
						<tr>
							<td class="page_title" dir="<%=dir_obj_var%>">&nbsp;שליחת מייל למדריך&nbsp;<%=Guide_Name%> </td>
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
						<FORM name="frmMain" ACTION="sendreportByMail.asp?gId=<%=gId%>&y=<%=gyear%>" METHOD="post"  enctype="multipart/form-data"  onSubmit="return CheckFields()" ID="Form1">
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
								<td align="<%=align_var%>" nowrap width="70">&nbsp;<span id="word9" name="word9"><!--מאת-->מאת</span>&nbsp;</td>
							</tr>
							<tr>
								<td align="<%=align_var%>" width=100% valign=top>&nbsp; <input id="Guide_Email" name="Guide_Email" dir="ltr" class="texts" style="width:285;" value="<%=Guide_Email%>">
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;<span id="Span1" name="word9"><!--מאת-->אל</span>&nbsp;</td>
							</tr>
							
							<TR>
								<td align="<%=align_var%>" width=100%>
									<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=1 id="SubjectText" name="SubjectText">דוח שנתי למדרך <%=Guide_Name%></textarea>
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;<!--תוכן תגובה--> נושא&nbsp;</td>
							</TR>
						
							<tr valign="top">
								<td align="<%=align_var%>" width=100%>
									<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=9 id="ReplyText" name="ReplyText"></textarea>
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;תוכן אימייל&nbsp;</td>
							</tr>
						<tr>
								<td  dir="<%=dir_obj_var%>"><input type="file" name="attachment_file" ID="attachment_file" size="33" value=""></td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;מסמך מצורף</td>
							</tr>
							<tr>
								<td colspan="2" height="10"></td>
							</tr>	
							<tr>
								<td align=right>
									<table cellpadding="0" cellspacing="0"  width="50%" align="right" ID="Table4" border=0>
										<tr valign="top">
											<td width="28%" align="right"><A class="but_menu" href="#" style="width:100px" onclick="window.close();"><span id="word4" name="word4">ביטול</span></A></td>
											<td width="10" nowrap></td>
											<td width="28%" align="left" nowrap><A class="but_menu" style="width:120px" href="#" onclick="return CheckFields(1);">שלח 
													אימייל</A></td>
										</tr>
									</table>
								</td>
								<td align="<%=align_var%>" nowrap width="70">&nbsp;</td>
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
