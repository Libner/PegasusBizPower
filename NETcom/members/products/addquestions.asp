<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
    UserId = trim(Request.Cookies("bizpegasus")("UserId"))
	OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))	
	FILEUP = trim(Request.Cookies("bizpegasus")("FILEUP"))		
	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
	is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
	is_companies = trim(Request.Cookies("bizpegasus")("ISCOMPANIES"))
  
  sub copy_form(copyprodId,pprodId)				
  
	str="SELECT * FROM Form_Field WHERE Product_Id=" & copyprodId & " and ORGANIZATION_ID=" & trim(OrgID) & " order by Field_Order"
	set tmp_fields=con.GetRecordSet(str)
	rsCount = tmp_fields.RecordCount  
	ii=1    	
	DO WHILE not tmp_fields.eof
		Field_Id = tmp_fields("Field_Id")
		Field_Title = trim(sfix(tmp_fields("Field_Title")))
		Field_Type = tmp_fields("Field_Type")
		if isNull(tmp_fields("Field_Size")) then
			Field_Size = "Null"
		else
			Field_Size = tmp_fields("Field_Size")
		end if
		Field_Align = trim(tmp_fields("Field_Align"))
		if isNull(tmp_fields("Field_Scale")) then
			Field_Scale = "Null"
		else
			Field_Scale = tmp_fields("Field_Scale")
		end if
		Field_Order = tmp_fields("Field_Order")
		FIELD_EXCEPTION = tmp_fields("FIELD_EXCEPTION")
		FIELD_KEY = tmp_fields("FIELD_KEY")
		If trim(FIELD_KEY) = "" Or IsNull(FIELD_KEY) Then
			FIELD_KEY = 0
		Else		   
			FIELD_KEY = Abs(CInt(FIELD_KEY))
	    End If
		FIELD_MUST = tmp_fields("FIELD_MUST")
		If trim(FIELD_MUST) = "" Or IsNull(FIELD_MUST) Then
			FIELD_MUST = 0
		Else
			FIELD_MUST = Abs(CInt(FIELD_MUST))
		End If		
						
		sqlstring="INSERT INTO Form_Field (Product_ID,User_ID,Organization_ID,Field_Title,Field_Type,Field_Size,Field_Align,Field_Order,Field_Scale,FIELD_EXCEPTION,FIELD_KEY,FIELD_MUST) values ("&_
		pprodId & ","& trim(UserID) & "," & trim(OrgID) &",'" & Field_Title & "'," & Field_Type & "," & Field_Size & ",'" & Field_Align & "'," & Field_Order & "," & Field_Scale &",'" & FIELD_EXCEPTION & "'," & FIELD_KEY & "," & FIELD_MUST & ")"
		'Response.Write sqlstring
		'Response.End
		fieldins_ok = 0
		if con.ExecuteQuery(sqlstring) then
			fieldins_ok = 1
		end if	
		if fieldins_ok > 0 and (Field_Type = "3" or Field_Type = "8" Or Field_Type = "11") then 'select or radio
			set tmpId = con.GetRecordSet("Select Field_ID from Form_Field  WHERE ORGANIZATION_ID=" & trim(OrgID) & " ORDER BY Field_ID DESC")
				newFieldId=CStr(tmpId("Field_ID"))
			set tmpId = nothing 
			set sel = con.GetRecordSet("Select * from Form_select where Field_id="& Field_Id & " and ORGANIZATION_ID=" & trim(OrgID) & " Order by ID")
			do while not sel.eof
				selValue= sfix(sel("Field_Value"))
				strInsert="INSERT INTO Form_Select (Field_Id,ORGANIZATION_ID,Field_Value) VALUES ("& newFieldId &","& trim(OrgID) &",'" & selValue & "')"
				'Response.Write(strInsert)
			     con.ExecuteQuery strInsert
			sel.movenext
			loop
			set sel = nothing
							
		end if
	tmp_fields.MoveNext
	loop
	set tmp_fields = nothing
	
end sub
		
	set upl=Server.CreateObject("SoftArtisans.FileUp")
	If upl.Form("add") <> nil Then  
	
		If (upl.TotalBytes > 512000) Then
        %>
        <script language=javascript>
		<!--
			window.alert("ניתן לשלוח מסמך מצורף עד 500 KB");
			window.history.back();			
		//-->
		</script>
        <%
        Response.End
        End If
        
        File_Size = upl.TotalBytes
        prodId = upl.Form("prodId")
        copyprodId = upl.Form("copyprodId")    
		
		If upl.Form("attachment") <> nil And upl.TotalBytes > 0 Then   		
   			File_Name=trim(upl.Form("attachment"))
   			File_Title=trim(upl.Form("attachment_title"))
   		Else
   			File_Name=""
   			File_Size="NULL"	
   			File_Title=""
   		End if   
   		
   		If trim(upl.Form("deleteFile")) = "1" Then
   			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   			attachment_file = trim(upl.Form("attachment"))
	   				
			If trim(attachment_file) <> "" Then
			sqlstr = "Select TOP 1 product_id FROM products WHERE FILE_ATTACHMENT LIKE '" & sFix(attachment_file) & "' AND product_id <> " & prodId
			set rs_check = con.getRecordSet(sqlstr)
			if rs_check.eof then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!   			
   				file_path="../../../download/products/" & attachment_file
   				'Response.Write fs.FileExists(server.mappath(file_path))
				'Response.End
				if fs.FileExists(server.mappath(file_path)) then
					set a = fs.GetFile(server.mappath(file_path))
					a.delete			
				end if	
			end if
			set rs_check = nothing	
			End If				  			
   			
			sqlstr = "Update products set FILE_ATTACHMENT = NULL, ATTACHMENT_TITLE = NULL, ATTACHMENT_SIZE = NULL Where product_id = " & prodId
			'Response.Write sqlstr
			'Response.End	
			con.executeQuery(sqlstr)		
			Set fs = Nothing		
	   	
		ElseIf  trim(upl.UserFilename) <> "" And trim(upl.Form("add")) = "1" And upl.TotalBytes > 0 Then
			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   			File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
	   			
   			file_path="../../../download/products/" & File_Name
			if fs.FileExists(server.mappath(file_path)) then
				set a = fs.GetFile(server.mappath(file_path))
				a.delete			
			end if			
								
			upl.Form("attachment").SaveAs server.mappath("../../../download/products/") & "/" & File_Name
			
			If Len(File_Title) = 0 Then
				File_Title = File_Name
			End If
			
			if Err <> 0 Then
				Response.Write("An error occurred when saving the file on the server.")			 
				set upl = Nothing
				Response.End
			end if			
		End If
		
		If trim(upl.Form("add")) = "1" Then			
	
		If trim(upl.Form("prodOpen")) <> nil Then ' טופס פתוח
			prodOpen = "1"
	    Else
			prodOpen = "0"
	    End If
	    
	    If trim(upl.Form("FORM_MAIL")) <> nil Then ' טופס דיוור
			FORM_MAIL = "1"
	    Else
			FORM_MAIL = "0"
	    End If
	    
	    If trim(upl.Form("FORM_CLIENT")) <> nil Then ' טופס תיקי לקוחות
			FORM_CLIENT = "1"
	    Else
			FORM_CLIENT = "0"
	    End If
	    	    
   	    If trim(upl.Form("AddClientCheck")) <> nil Then ' קליטת נתוני לקוח למערכת 
			ADD_CLIENT = trim(upl.Form("ADD_CLIENT"))
	    Else
			ADD_CLIENT = ""
	    End If

	    if upl.Form("FORM_TO_RESP") <> nil then
			responsible = upl.Form("responsible")
			If isNULL(responsible) Or IsNumeric(responsible) = false Then
				responsible = "NULL"
			End If	
		else
			responsible = "NULL"
		end if
	    
		if prodId<>"" then
			sqlstr=	"UPDATE PRODUCTS SET PRODUCT_NAME='" & sFix(upl.Form("prodName")) &_
			 "', LANGU='" & upl.Form("langu") & "', PRODUCT_DESCRIPTION = '" & sFix(upl.Form("prodDesc")) &_
			 "', PRODUCT_THANKS = '" & sFix(upl.Form("prodThanks")) & "', OPEN_FORM = '" & prodOpen &_
			 "', FORM_MAIL = '" & FORM_MAIL & "', FORM_CLIENT = '" & FORM_CLIENT &_
			 "', RESPONSIBLE = " & RESPONSIBLE & ", ADD_CLIENT = '" & ADD_CLIENT &_
			 "', FILE_ATTACHMENT = '" & sFix(File_Name) & "', ATTACHMENT_TITLE ='" & sFix(File_Title) &_
			 "', ATTACHMENT_SIZE = " & File_Size & ", User_ID = " & trim(UserID) &_
			 " WHERE PRODUCT_ID = " & prodId & " AND ORGANIZATION_ID=" & trim(OrgID)
			'Response.Write("upd" & sqlstr)
			'Response.End
			con.ExecuteQuery(sqlstr)
			Response.Redirect "questions.asp"
		else
			sqlstring="INSERT INTO PRODUCTS (USER_ID,ORGANIZATION_ID,PRODUCT_NUMBER,PRODUCT_NAME,LANGU,"&_
			" PRODUCT_DESCRIPTION,PRODUCT_THANKS,OPEN_FORM,FORM_MAIL,FORM_CLIENT,RESPONSIBLE,ADD_CLIENT,"&_
			" FILE_ATTACHMENT,ATTACHMENT_TITLE,ATTACHMENT_SIZE) VALUES (" & trim(UserID) & "," & trim(OrgID) &_
			",'0','" & sFix(upl.Form("prodName")) & "','" & upl.Form("langu") &"','"& sFix(upl.Form("prodDesc")) &_
			"','" & sFix(upl.Form("prodThanks")) &	"','" & prodOpen & "','" & FORM_MAIL & "','" &_
			FORM_CLIENT & "'," & RESPONSIBLE & ",'" & ADD_CLIENT & "','" & sFix(File_Name) & "','" &_
			sFix(File_Title) & "'," & File_Size & ")"
			'Response.Write("ins " & sqlstring)
			if con.ExecuteQuery(sqlstring) then			
				set tId=con.GetRecordSet("SELECT PRODUCT_ID FROM PRODUCTS WHERE ORGANIZATION_ID=" & trim(OrgID) &" ORDER BY product_id DESC")
				prodId=tId("product_id")
				tId.close
				set tId=nothing
				' *** copying form 
				copyprodId = trim(upl.Form("copyprodId"))
				'Response.Write 		copyprodId
				'Response.End	
				if trim(copyprodId) <> "" then
					copy_form copyprodId,prodId				
				end if 'copyprodId <> ""
				' *** end copying form 						
			end if 'con.ExecuteQuery(sqlstring)	

			if trim(copyprodId) = "" And trim(is_groups) = "0" then		
				Response.Redirect "addfield.asp?prodId=" & prodId		
			elseif trim(is_groups) = "1" then		
				Response.Redirect "permitions.asp?prodId=" & prodId			
			else
				Response.Redirect "questions.asp"
			end if	
		  end if
	  Else
		 Response.Redirect "addquestions.asp?prodId=" & prodId
	  End If  	
	End If
	
	if Trim(request("prodId")) <> "" then
		prodId=request("prodId")
	else
		prodId=""
	end if	
	
	if trim(request("addcopy")) <> "" then
		copyprodId=request("addcopy")
		set qs=con.GetRecordSet("select product_name,langu from products where product_id=" & copyprodId & " and ORGANIZATION_ID=" & trim(OrgID) )
		copyprodName=qs("product_name")
		copyprodLangu = qs("langu")
		set qs=nothing
	else
		copyprodId=""		
	end if
%>	
<%
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 20 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
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
<meta charset="windows-1255">
<!--#include file="../checkWorker.asp"-->
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT language="javascript">
<!--
_editor_url = "../../../htmlarea/";                     // URL to htmlarea files
_view_form = false;
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
  document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
  document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }

function SetHTMLArea() {
	var config = new Object();    // create new config object
	config.bodyStyle = 'font-family: Arial; font-size: 12pt; color: #575757; PADDING-LEFT: 0px; PADDING-RIGHT: 0px';
	config.fontsize_default = "12";		
	config.toolbar = [
    ['fontname'],
    ['fontsize'],
    ['bold','italic','underline','separator'],
	['strikethrough','subscript','superscript','separator'],
    ['justifyleft','justifycenter','justifyright','justifyfull','justifynone','lefttoright','righttoleft','separator'],
    ['remove','removewordformat'],
    ['separator'],
	['OrderedList','UnOrderedList','separator'],
    ['forecolor','backcolor','separator'],
    ['HorizontalRule','Createlink','Removelink','separator'],
    ['insertbulletedlist','InsertTable','EditTable','separator','htmlmode']
    ];
	editor_generate("prodDesc", config);
}
	
function langu_onclick() {
	if (document.dataForm.langu.value == 'eng')
	{
		document.dataForm.prodName.dir = 'ltr';
		document.dataForm.prodDesc.dir = 'ltr';
		document.dataForm.prodThanks.dir = 'ltr';		
	}
	else
	{
		document.dataForm.prodName.dir = 'rtl';
		document.dataForm.prodDesc.dir = 'rtl';
		document.dataForm.prodThanks.dir = 'rtl';
	}	
}

//-->
</SCRIPT>
<script LANGUAGE="JavaScript">
<!--
function CheckField() {
  if (window.document.dataForm.prodName.value=='')
  {
  <%
     If trim(lang_id) = "1" Then
        str_alert = "! חובה למלא את שדה שם הטופס"
     Else
		str_alert = "Please, insert the form name!"
     End If   
  %>  
	window.alert("<%=str_alert%>");
	window.document.dataForm.prodName.focus();
	return false;
  } 
  if(window.document.dataForm.FORM_TO_RESP.checked && window.document.dataForm.responsible.selectedIndex == 0)
  {
	<%
		If trim(lang_id) = "1" Then
			str_alert = "! נא לבחור עובד אחראי טופס"
		Else
			str_alert = "Please, choose the responsible worker!"
		End If   
	%>    
	   window.alert("<%=str_alert%>");
	   window.document.dataForm.responsible.focus();
	   return false;
  }
  if(window.dataForm.AddClientCheck.checked == true && window.document.all("ADD_CLIENT")(0).checked == false && window.document.all("ADD_CLIENT")(1).checked == false)
  {
	<%
		If trim(lang_id) = "1" Then
			str_alert = "נא לבחור סוג לקוח להוספה עם הטופס"
		Else
			str_alert = "Please, choose the type of customer to add whith the form!"
		End If   
	%>      
	window.document.all("ADD_CLIENT")(0).focus();
	window.alert("<%=str_alert%>");
	return false;
  }
  <%If trim(FILEUP) = "1" Then%>
  if (document.dataForm.attachment.value !='')
  {
		var fname=new String();
		var fext=new String();
		var extfiles=new String();
		fname=document.dataForm.attachment.value;
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
	<%End If%>
  return true;
}
function DeleteAttachment(taskID)
{
    <%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק את הקובץ המצורף"
		Else
			str_confirm = "Are you sure want to delete the attached file?"
		End If	
	%>
	if (window.confirm("<%=str_confirm%>"))
	{
		window.document.all("add").value = "0";
		window.document.all("deleteFile").value = "1";
		window.dataForm.submit();
	}
	return false;
}

function CheckDel(field,prodID) 
//For Logo deleting
{
    <%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק את התמונה"
		Else
			str_confirm = "Are you sure want to delete the logo?"
		End If	
	%>
  if(confirm("<%=str_confirm%>") == true )
  {
	document.location.href='Aimgadd.asp?PRODUCT_ID='+prodID+'&'+field+'=1';
	return true;
  }
  return false;
}

function ifFieldEmpty(field)
{
	<%
		If trim(lang_id) = "1" Then
			str_alert = "!חובה לבחור את התמונה"
		Else
			str_alert = "Please, choose the logo!"
		End If   
	%>   
	if (field.value=='')
	{
		window.alert('<%=str_alert%>');
		return false;
	}
	return true;
}
//-->
</script>  
</head>
<%  	
	if prodId <> "" Or copyprodId <> "" then
		If trim(prodId) <> "" Then
			set qs=con.GetRecordSet("SELECT PRODUCT_NAME,PRODUCT_DESCRIPTION,PRODUCT_THANKS,LANGU,OPEN_FORM,FORM_CLIENT,FORM_MAIL,RESPONSIBLE,IsNull(ADD_CLIENT,0),FILE_ATTACHMENT,ATTACHMENT_TITLE FROM PRODUCTS WHERE PRODUCT_ID=" & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
		ElseIf trim(copyprodId) <> "" Then
			set qs=con.GetRecordSet("SELECT PRODUCT_NAME,PRODUCT_DESCRIPTION,PRODUCT_THANKS,LANGU,OPEN_FORM,FORM_CLIENT,FORM_MAIL,RESPONSIBLE,IsNull(ADD_CLIENT,0),FILE_ATTACHMENT,ATTACHMENT_TITLE FROM PRODUCTS WHERE PRODUCT_ID=" & copyprodId & " and ORGANIZATION_ID=" & trim(OrgID) )
		End If
		If not qs.eof Then
			prodName = trim(qs(0))
			prodDesc = trim(qs(1))
			prodThanks = trim(qs(2))
			prodLangu = trim(qs(3))
			prodOpen = trim(qs(4)) ' טופס פתוח
			prodClient = trim(qs(5))
			prodMail = trim(qs(6))
			prodResp = trim(qs(7))
			prodAddClient = trim(qs(8))
			attachment = trim(qs(9))
			attachment_title = trim(qs(10))
		End If
		set qs=nothing
	else
		If trim(lang_id) = "1" Then
			prodLangu = "heb"
		Else
			prodLangu = "eng"
		End If				
		prodClient = "1"	 : prodAddClient = "1" : prodMail = "1"	
	end if
	
	If isNumeric(trim(prodId)) Then
		sqlpg="SELECT PRODUCT_LOGO FROM PRODUCTS WHERE PRODUCT_ID="&prodId&""
		'Response.Write sqlpg
		'Response.End
		set pg=con.getRecordSet(sqlpg)	
		pictSize=pg.Fields("PRODUCT_LOGO").ActualSize
		set pg = Nothing
	End If
%>
<body onload="SetHTMLArea();">
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td width="100%">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 1%>
<%numOfLink = 1%>
<%topLevel2 = 14 'current bar ID in top submenu - added 03/10/2019%>
<%If trim(wizard_id) = "" Then%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<%End if%>
<tr><td class="page_title">&nbsp;</td></tr>
<tr><td width="100%" align="<%=align_var%>">
<% wizard_page_id = 1 %>
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<%If trim(wizard_id) <> "" Then%>
<tr>
<td width="100%" align="<%=align_var%>" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width="100%">
<tr><td class="explain_title">צור טופס רישום חדש!</td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td class=explain>
בשדה <b>שפת הטופס</b> , בחר את השפה בה ייכתב טופס ההרשמה.
</td></tr>
<tr><td class=explain>
בשדה <b>שם הטופס</b> הזן שם לטופס הרישום. שם זה לא יופיע בדיוור אלא ישמש לזיהוי הטופס במערכת ה-Bizpower.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט מקדים</b> , הקלד את הטקסט שיופיע בתחילת הטופס לפני שדות הרישום.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט תודה לאחר הרישום</b> , הקלד את הטקסט שברצונך שהנרשם יראה לאחר שליחת הטופס. <br>לדוגמה: "תודה! נציגינו יחזרו אליך ב-24 השעות הקרובות"
</td></tr>
</table>
</td>
</tr>
<tr>
    <td width="100%" height="2"></td>
</tr>
<%End If%>
<tr>    
    <td width="100%" valign="top" align="center" bgcolor="#E6E6E6">
	<table align=center border="0" cellpadding="2" cellspacing="1" width="70%" dir="<%=dir_var%>">
	<tr><td height=15 nowrap></td></tr>
	<FORM name="dataForm" ACTION="addquestions.asp" METHOD="post" onSubmit="return CheckField();" enctype="multipart/form-data">
	<input type=hidden name="deleteFile" id="deleteFile" value="0">
	<input type="hidden" name="prodId" value="<%=prodId%>" ID="prodId">
    <input type="hidden" name="copyprodId" value="<%=copyprodId%>" ID="copyprodId">
	<input type=hidden name="add" id="add" value="1">
	<tr>
		<td align="<%=align_var%>" width="100%"><input dir="<%if trim(prodLangu) = "eng" then%>ltr<%else%>rtl<%end if%>" type="text" class="texts" size=40 id="prodName" name="prodName" value="<%=vfix(prodName)%>"></td>
		<td align="<%=align_var%>" width=150 nowrap><b>&nbsp;<!--שם הטופס--><%=arrTitles(3)%>&nbsp;</b></td>
	</tr>
	<tr>
		<td align="<%=align_var%>">
		 <select name="langu" class="norm" dir="<%=dir_obj_var%>" LANGUAGE=javascript onclick="return langu_onclick()">
		  <option value="eng" <%if trim(prodLangu)="eng" then%> selected<%end if%> id=word16>&nbsp;&nbsp;<!--אנגלית--><%=arrTitles(16)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</OPTION>
		  <option value="heb" <%if trim(prodLangu)="heb" then%> selected<%end if%> id=word17>&nbsp;&nbsp;<!--עברית--><%=arrTitles(17)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</OPTION>
		 </select>
		</td>
		<td align="<%=align_var%>"><b>&nbsp;<!--שפת הטופס--><%=arrTitles(4)%>&nbsp;</b></td>
	</tr>
	<tr>
		<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>">
		<input <%if trim(prodClient) = "1" then%>checked<%end if%> type="checkbox" id="FORM_CLIENT" name="FORM_CLIENT"> 
		&nbsp;-&nbsp;<span id="word5" dir="<%=dir_var%>"><!--ניתן לצרף את הטופס לתיקי לקוחות--><%=arrTitles(5)%></span>
		</td>
		<td align="<%=align_var%>" width=150 nowrap><b>&nbsp;<!--טופס תיקי לקוחות--><%=arrTitles(6)%>&nbsp;</b></td>
	</tr>
	<!----------------------------------------------- קליטת נתוני לקוח ------------------------------------------->
	<tr>
		<td align="<%=align_var%>" width="100%"><!--קליטת נתוני לקוח למערכת--><%=arrTitles(18)%>
		&nbsp;-&nbsp;<input <%if trim(prodAddClient) <> "" then%>checked<%end if%> type="checkbox" onclick="tbClient.style.display=(this.checked)? 'block':'none'" name=AddClientCheck id=AddClientCheck> 
		</td>
		<td align="<%=align_var%>" width=150 nowrap><b>&nbsp;<!--נתוני לקוח--><%=arrTitles(19)%>&nbsp;</b></td>
	</tr>
	<tbody name="tbClient" id="tbClient" style="display:<%if trim(prodAddClient) <> "" then%>block<%else%>none<%end if%>">
	<tr>
		<td align="<%=align_var%>" width="100%"><!--לקוח פרטי--><%=arrTitles(25)%>
		&nbsp;-&nbsp;<input <%if trim(prodAddClient) = "1" then%>checked<%end if%> type="radio" value=1 name="ADD_CLIENT" ID="Radio1"> 
		</td>
		<td align="<%=align_var%>" width=150 nowrap></td>
	</tr>
	<%If trim(is_companies) = "1" Then%>
	<tr>
		<td align="<%=align_var%>" width="100%"><!--לקוח ארגוני--><%=arrTitles(24)%>
		&nbsp;-&nbsp;<input <%if trim(prodAddClient) = "2" then%>checked<%end if%> type="radio" value=2 name="ADD_CLIENT" ID="Radio2"> 
		</td>
		<td align="<%=align_var%>" width=150 nowrap>&nbsp;</td>
	</tr>
	<%End if%>
	</tbody>
	<!----------------------------------------------- קליטת נתוני לקוח ------------------------------------------->
	<tr>
		<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>">
		<input <%if trim(prodMail) = "1" then%>checked<%end if%> type="checkbox" id="FORM_MAIL" name="FORM_MAIL">
		&nbsp;-&nbsp;<span id="word7" dir="<%=dir_var%>"><!--ניתן לשלוח את הטופס בדיוור ולקבל משובים דרך האינטרנט--><%=arrTitles(7)%></span>
		</td>
		<td align="<%=align_var%>" width=150 nowrap><b>&nbsp;<span id="word8"><!--טופס דיוור--><%=arrTitles(8)%></span>&nbsp;</b></td>
	</tr>
	<tr>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><input <%if trim(prodOpen) = "1" then%>checked<%end if%> type="checkbox" id="prodOpen" name="prodOpen">&nbsp;-&nbsp;<span id="word9"><!--ניתן לפנות אל הטופס באמצעות כתובת URL--><%=arrTitles(9)%></span>
		</td>
		<td align="<%=align_var%>" nowrap><b>&nbsp;<!--טופס פניה--><%=arrTitles(10)%>&nbsp;</b></td>
	</tr>	
	<tr>
		<td align="<%=align_var%>" width="100%"><!--שלח נתוני טופס בדואר אלקטרוני--><%=arrTitles(11)%>
		&nbsp;-&nbsp;<input <%if trim(prodResp) <> "" then%>checked<%end if%> type="checkbox" id="FORM_TO_RESP" name="FORM_TO_RESP"> <!--onclick="selresponsible.style.display=(this.checked)? 'block':'none'"--> 
		</td>
		<td align="<%=align_var%>" width=150 nowrap><b>&nbsp;<!--חיווי קליטה--><%=arrTitles(20)%>&nbsp;</b></td>
	</tr>
	<TR>
		<TD align="<%=align_var%>">		
		<select name="responsible" id="responsible" class="norm" dir="<%=dir_obj_var%>" style="width:200">		
		<option value="" id=word13><%=arrTitles(13)%></option>		
		<%  sqlstr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME FROM USERS WHERE ORGANIZATION_ID = " & OrgID			
		    sqlstr = sqlstr & " Order BY FIRSTNAME + ' ' + LASTNAME "
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
		<option value="<%=rs_comp(0)%>" <%If trim(prodResp) = trim(rs_comp(0)) Then%> selected <%End if%>><%=rs_comp(1)%></option>			
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>
		</select>
		</TD>
		<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;<!--עובד אחראי--><%=arrTitles(12)%>&nbsp;</b></td>
	</TR>
	<%If trim(FILEUP) = "1" Then%>
	<tr>
	<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" name="attachment_title" ID="attachment_title" size=50 value="<%=vFix(attachment_title)%>" maxlength=100 dir="<%=dir_obj_var%>"></td>
	<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;<!--כותרת מסמך--><%=arrTitles(23)%>&nbsp;</b></td>
	</tr>
	<%If IsNull(attachment) Or trim(attachment) = "" Then%>
	<tr>
	<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="file" name="attachment" ID="attachment" size=33 value="" dir="<%=dir_obj_var%>"></td>
	<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b>&nbsp;מסמך&nbsp;</b></td>
	</tr>
	<%Else%>
	<tr>	
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>">
	<input type="hidden" name="attachment" ID="attachment" value="<%=vFix(attachment)%>">
	<a class="file_link" href="../../../download/products/<%=attachment%>" target=_blank style="position:relative;top:-4px"><%=attachment%></a>
	<input type=image src="../../images/delete_icon.gif" style="position:relative;top:2px" onclick="return DeleteAttachment('<%=prodId%>')">
	</td>
	<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><b>&nbsp;<!--מסמך--><%=arrTitles(22)%>&nbsp;</b></td>
	</tr>
	<%End If%>
	<%End If%>
	<tr>
		<td align="<%=align_var%>" valign=top><TEXTAREA dir="<%if trim(prodLangu) = "eng" then%>ltr<%else%>rtl<%end if%>" class="texts" style="width: 500px" rows=8 id="prodDesc" name="prodDesc"><%=(prodDesc)%></TEXTAREA></td>
		<td align="<%=align_var%>" nowrap valign="top"><b>&nbsp;<!--טקסט מקדים--><%=arrTitles(14)%>&nbsp;</b></td>
	</tr>
	<tr>
		<td align="<%=align_var%>" valign=top><TEXTAREA dir="<%if trim(prodLangu) = "eng" then%>ltr<%else%>rtl<%end if%>" class="texts" style="width: 500px" rows=4 id="prodThanks" name="prodThanks"><%=(prodThanks)%></TEXTAREA></td>
		<td align="<%=align_var%>" nowrap valign="top">&nbsp;<b><!--טקסט תודה<br>לאחר מילוי הטופס--><%=arrTitles(15)%>&nbsp;</b></td>
	</tr>
	<tr><td colspan=2 height="10"></td></tr>
	<tr><td colspan=2 align="center">
	<table cellpadding=0 cellspacing=0 width="100%">
<%If trim(wizard_id) <> "" Then%>		
	<tr>
		<td width="50%" align="<%=align_var%>" nowrap>
		<input style="width:90px" class="but_menu" type=submit value="<< המשך" ID="Submit1" NAME="Submit1">
		</td>
		<td width=50 align="<%=align_var%>" nowrap></td>
		<td width="50%" align="left" nowrap>
		<input type=button class="but_menu" value="חזור >>" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id%>.asp?pageId=<%=pageId%>'" style="width:90px" ID="Button1" NAME="Button1">
		</td>		
	</tr>	
<%Else%>
	<tr>
	<td width="50%" align="center" nowrap>
	<INPUT class="but_menu" type="button" style="width:90px" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="document.location.href='questions.asp'"></td>
	<td width=50 nowrap></td>
	<td width="50%" align="center" nowrap>
	<input style="width:90px" class="but_menu" type="submit" value="<%=arrButtons(1)%>" id=button1 name=button1>
	</td></tr>
<%End If%>	
</table>
</td></tr>
</FORM>
<tr><td colspan=2 height="5"></td></tr>
</table></td></tr>
<% If trim(prodID) <> "" Then %>
<tr><td height=1 bgcolor="#808080"></td></tr>
<tr><td width="100%" bgcolor="#CCCCCC">
<form ACTION="Aimgadd.asp?C=1&F=PRODUCT_LOGO&PRODUCT_ID=<%=prodID%>" METHOD="post" ENCTYPE="multipart/form-data" ID="Form1">						
<table cellpadding=2 cellspacing=0 width=70% border=0 align=center dir="<%=dir_var%>">
<tr><td colspan=2 align="<%=align_var%>" class="10normalB">&nbsp;<!--pxרוחב מקסימלי  350 X pxגודל מומלץ: גובה 60 --><%=arrTitles(26)%>&nbsp;</td></tr>
		<%If pictSize = 0 then %>		
		<tr><td align="<%=align_var%>" height=7 nowrap></td></tr>
		<tr valign="middle">
			<td align="<%=align_var%>" valign="middle">
			<table border="0" cellspacing="2" cellpadding="2" align="<%=align_var%>" dir="<%=dir_var%>">
				<tr valign="middle">				
				<td align="<%=align_var%>" class="td_admin_1" style="padding-right:2px;"><INPUT type="submit" class="but" onclick="return ifFieldEmpty('UploadFile2')" value="<%=arrTitles(27)%>"></td>
				<td align="<%=align_var%>" class="td_admin_1" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30></td>				
		        <td align="<%=align_var%>" nowrap valign="top"><b>&nbsp;Logo&nbsp;</b></td>
				</tr>
			</table>       
			</td>
		</tr>
		<%Else%>
		<tr>
			<td align="<%=align_var%>"><img id="imgPict" name="imgPict" src="../../GetImage.asp?DB=PRODUCT&amp;FIELD=PRODUCT_LOGO&amp;ID=<%=prodID%>" border="0" hspace=2 ></td>						
			<td align="<%=align_var%>" nowrap valign="top"><b>&nbsp;Logo&nbsp;</b></td>
	    </tr>
		<tr>
			<td align=right colspan=2>
			<table border="0" cellspacing="5" cellpadding="0" align="<%=align_var%>" dir="<%=dir_var%>">
				<tr valign="middle">        
				<td align="<%=align_var%>">
				<INPUT type="button" class="but" ONCLICK="return CheckDel('delLogo','<%=prodID%>');" value="<%=arrTitles(28)%>">
				</td>
				<td align="<%=align_var%>">
				<INPUT type="submit" class="but" onclick="return ifFieldEmpty('UploadFile2')" value="<%=arrTitles(29)%>">
				</td>
                <td align="<%=align_var%>"><INPUT TYPE="FILE" NAME="UploadFile2" size=30></td>
				</tr>
			</table>
			</td>
		</tr>
		<%End If%>			
</table>
</form></td></tr>
<tr><td height=1 bgcolor="#808080"></td></tr>
<%End If%>
<tr><td colspan=2 height="10"></td></tr>
</table>
<%set con=nothing%>
</body>
</html>

