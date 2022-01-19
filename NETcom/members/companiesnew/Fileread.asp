<%Response.Buffer = False%>
<!--#include file="../../adovbs.inc"-->
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
function CheckFields()
{
	objSel = window.document.getElementsByTagName("Select")
	count_selected = 0
	for(i=0;i<objSel.length;i++)
	{
		if(objSel[i].selectedIndex > 0)
			count_selected = count_selected + 1
	}
	//window.alert(count_selected);
	if(count_selected == 0)
	{
		window.alert("שים לב, על מנת להמשיך בייבוא נתונים אליך לשייך לפחות שדה אחד")
		return false;
	}	
	return true;
}

function OpenPreview(strUrl)
{
	File = window.open(strUrl,"File","");
}

function setColor(selObject, i, is_equal)
{
	selValue = parseInt(selObject.value);	 
	
	if(selValue == -1)
		window.document.all("tr"+i).style.background = "#E6E6E6";
	else if(is_equal == "-1")
		window.document.all("tr"+i).style.background = "#B3B2D0";	
	else
		window.document.all("tr"+i).style.background = "#ACE920";	
}
//-->
</script>
<style>
.normalB
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11pt;
    COLOR: #2e2380;
    FONT-FAMILY: Arial
}
</style>
</head>
<body  style="margin:0px;background-color:#E5E5E5">
<table cellpadding=0 cellspacing=0 width=100% border=0>
<%If Request.QueryString("import") <> nil Then%>
<div style="background-color:#E6E6E6" id="loadMessage">
<table border="0" width="400" cellspacing="1" cellpadding="0" align=center>
    <tr><td height="30"></td></tr>
	<tr>
		<td align="right" nowrap class="normalB" dir=rtl>תהליך טעינת הנתונים לאתר מתבצעת כעת, התהליך יארך מספר דקות</td>
	</tr>		
	<tr><td height="10"></td></tr>	
	<tr>
		<td  align="right" nowrap class="normalB" dir=rtl><font color="red">אין לסגור חלון זה ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</font></td>
	</tr>	
	<tr><td height="20"></td></tr>
	<tr><td align=center><img src="../../images/butt_act.gif" border=0 vspace=0 hspace=0></td></tr>
	<tr><td height="10"></td></tr>
</table>
</div>
<%End If%>
<%
	newFileName = trim(Request("newFileName"))
   
    Function IsEmailValid(strEmail) 
   
		Dim strArray 
		Dim strItem 
		Dim i 
		Dim c 
		Dim blnIsItValid 
	   
		' assume the email address is correct   
		blnIsItValid = True 
		
		If IsNull(strEmail) Then
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function
		End If
	     
		If InStr(strEmail,"@") < 0 Then
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function
		End If
		
		' split the email address in two parts: name@domain.ext 
		strArray = Split(strEmail, "@") 
	   
		' if there are more or less than two parts   
		If UBound(strArray) <> 1 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 
	   
		' check each part 
		For Each strItem In strArray 
			' no part can be void 
			If Len(strItem) <= 0 Then 
				blnIsItValid = False 
				IsEmailValid = blnIsItValid 
				Exit Function 
			End If 
	         
			' check each character of the part 
			' only following "abcdefghijklmnopqrstuvwxyz_-." 
			' characters and the ten digits are allowed 
			For i = 1 To Len(strItem) 
				c = LCase(Mid(strItem, i, 1)) 
				' if there is an illegal character in the part 
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then 
					blnIsItValid = False 
					IsEmailValid = blnIsItValid 
					Exit Function 
				End If 
			Next 
	    
		' the first and the last character in the part cannot be . (dot) 
			If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
			End If 
		Next 
	   
		' the second part (domain.ext) must contain a . (dot) 
		If InStr(strArray(1), ".") <= 0 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 
	   
		' check the length oh the extension   
		i = Len(strArray(1)) - InStrRev(strArray(1), ".") 
		' the length of the extension can be only 2, 3, or 4 
		' to cover the new "info" extension 
		If i <> 2 And i <> 3 And i <> 4 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 

		' after . (dot) cannot follow a . (dot) 
		If InStr(strEmail, "..") > 0 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 
	   
		' finally it's OK   
		IsEmailValid = blnIsItValid 
     
End Function  

	titlesArrayContact = Array("שם","טלפון איש קשר","פקס איש קשר","טלפון נייד","דואר אלקטרוני","תפקיד","כתובת איש קשר","עיר איש קשר","מיקוד איש קשר")
	titlesArrayCompany = Array("חברה","כתובת חברה","עיר חברה","מיקוד חברה","טלפון 1 חברה","טלפון 2 חברה","פקס חברה","דואר אלקטרוני חברה","אתר")	
	titlesArrayEContact = Array("Name","Phone","Fax","Cellular","Email","Position","Address","City","Postal code")	
	titlesArrayECompany = Array("Company","Business address","Business city","Business postal code","Business Phone 1","Business Phone 2","Business fax","Business email","Web Page")		
	columnArrayContact = Array("CONTACT_NAME","phone","fax","cellular","CONTACTS.email","messanger_name","contact_address","contact_city_Name","contact_zip_code")
	columnArrayCompany = Array("COMPANY_NAME","address","city_Name","zip_code","phone1","phone2","fax1","COMPANIES.email","url")
	columnWidthContact = Array(50,20,20,20,50,100,255,255,10)
	columnWidthCompany = Array(100,255,255,10,50,50,50,50,150)
	
	If Request.QueryString("import") = nil Then
	'Code to display data from the csv file
	Dim objFSO 

	'Create an instance of the FSO 
	Set objFSO = CreateObject("Scripting.filesystemObject") 
	strFilePath = Server.MapPath("../../../download/import/") & "/" & newFileName	
	strPath = Server.MapPath("../../../download/import/")
	'Check the file exists 
	If objFSO.FileExists(strFilePath) Then 	

		Set con_csv = Server.CreateObject("Adodb.Connection")
		con_csv.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&strPath&";Extended Properties=""text;HDR=No;FMT=Delimited""")
	
		Set rs_csv = Server.CreateObject("Adodb.RecordSet")
		rs_csv.Open "Select * From " & newFileName, con_csv, adOpenStatic, adLockReadOnly, adCmdText		
	
		If not rs_csv.eof Then
			arrData = rs_csv.getRows()
		End If
		set rs_csv = Nothing
		con_csv.Close()
		Set con_csv = Nothing
		
		'For k=0 To 0
		'	For y=0 To Ubound(arrData,2)
		'		Response.Write arrData(y,k) & "<br>"
		'	Next
		'Next	
		'Response.End
		
		session("arrData") = arrData	
	
		'Open a file for reading 
		Set oInStream = objFSO.OpenTextFile( Server.MapPath("../../../download/import/") & "/" & newFileName, 1, False ) 
		
		dim arrXlsTitles()
		
		If not oInStream.AtEndOfStream Then
		
			first_line = oInStream.readLine		
			fArr = Split( first_line, "," ) 
			If IsArray(fArr) Then
				numOfColumns = Ubound(fArr)
			End If
			
			For i=0 To numOfColumns				
				Redim Preserve arrXlsTitles(numOfColumns)
				If i <= Ubound(fArr) Then
					If Len(fArr(i)) > 0 Then
						strIndex = InStr(fArr(i), """")
						If 	strIndex = 1 Then
							strTmp = Replace(fArr(i),"""","",1,1)
						Else	
							strTmp = fArr(i)
						End If
						
						strTmp = Replace(strTmp,Chr(34)&chr(34),Chr(34))
					
						If InStrRev(strTmp,"""") > 0 Then					
							strTmp = Left(strTmp,Len(strTmp)-1)
						End If				
					
						arrXlsTitles(i) = strTmp				
					End If		
				End if			
			Next
			End If	
	
	Else 

	Response.Write "File not found!" 	
	Response.End

	End If 
	Set objFSO = Nothing 
	
	Else
	    arrData = session("arrData")
	    Dim arrColContact()
	    Redim arrColContact(Ubound(columnArrayContact))
	    
		For i=0 To Ubound(columnArrayContact)-1		
			If trim(Request.Form(columnArrayContact(0))) = "-1" Then
				is_contact = -1
			Else
				is_contact = Request.Form(columnArrayContact(0))	
			End If		
			If trim(Request.Form(columnArrayContact(i))) <> "-1" Then					
				arrColContact(i) = trim(Request.Form(columnArrayContact(i)))
			Else
				arrColContact(i) = ""
			End If
		Next
		
	    Dim arrColCompany()
	    Redim arrColCompany(Ubound(columnArrayCompany))

		For i=0 To Ubound(columnArrayCompany)-1
			If trim(Request.Form(columnArrayCompany(0))) = "-1" Then
				is_company = -1
			Else
				is_company = Request.Form(columnArrayCompany(0))	
			End If
			If trim(Request.Form(columnArrayCompany(i))) <> "-1" Then
				arrColCompany(i) = trim(Request.Form(columnArrayCompany(i)))
			Else
				arrColCompany(i) = ""
			End If
		Next

		
		sqlstr = "Select company_id From companies WHERE Organization_ID = " & orgID & " And private_flag = 1"
		set rs_private = con.getRecordSet(sqlstr)
		if not rs_private.eof then		
			compPrivID = rs_private(0)	
		end if
		set rs_private = Nothing	
		'Response.Write is_company
		'Response.End
		
		For i=1 To Ubound(arrData,2)
			companyID = ""
		   	
			If is_company > -1 Then
			  If Len(trim(arrData(is_company,i))) > 0 Then
				companyName = trim(arrData(is_company,i))
				sqlstr = "Select Top 1 Company_ID From Companies WHERE ltrim(rtrim(Company_Name)) Like '" &_
				sFix(companyName) & "' And Organization_ID = " & OrgID
				'Response.Write sqlstr
				'Response.End
				set rs_check = con.GetRecordSet(sqlstr)
				If not rs_check.eof Then
					companyID = rs_check(0)	
					arr_double_comp = arr_double_comp & companyID & ","									    
				Else		 			
					sqlstr = "SET NOCOUNT ON; Insert Into Companies ("
					For j=0 To Ubound(columnArrayCompany)-1	
						sqlstr = sqlstr & columnArrayCompany(j) & ","
					Next
					sqlstr = sqlstr & "User_Id,Organization_ID,date_update,private_flag,status) Values ("
					For kk=0 To Ubound(arrColCompany)-1
					fWidth = columnWidthCompany(kk)
					If IsNumeric(arrColCompany(kk)) Then
						sqlstr = sqlstr & "'" & Left(sFix(trim(arrData(arrColCompany(kk),i))), fWidth) & "', "	
					Else
						sqlstr = sqlstr & "'',"	
					End If	
					Next					
					sqlstr = sqlstr & UserID & "," & OrgID & ", getDate(),0,2); SELECT @@IDENTITY AS NewID"			
					'Response.Write sqlstr
					'Response.End
					set rs_tmp = con.getRecordSet(sqlstr)
						companyID = rs_tmp.Fields("NewID").value
					set rs_tmp = Nothing
					End If					
				
				set rs_check = Nothing
			  End If			 
				
			Else
				 companyID = compPrivID			'הוספת איש קשר לחברה פרטית קיימת
			End If 			
		
		  If trim(companyID) = "" Then
			companyID = compPrivID			'הוספת איש קשר לחברה פרטית קיימת
		  End If
			
		  If is_contact > -1 Then ' Insert contact			
			    
			    contactID = ""
				If Len(trim(arrData(is_contact,i))) > 0 Then
					contactName = trim(arrData(is_contact,i))				
					sqlstr = "Select Top 1 contact_ID From Contacts WHERE ltrim(rtrim(contact_Name)) Like '" &_
					sFix(contactName) & "' And Company_ID = " & companyID										
					'Response.Write sqlstr
					'Response.End
					set rs_check = con.GetRecordSet(sqlstr)
					If not rs_check.eof Then
						contactID = rs_check(0)	
						arr_double_cont = arr_double_cont & contactID & ","			    
					
					Else
						sqlstr = "Insert Into Contacts ("
						For p=0 To Ubound(columnArrayContact)-1	
							sqlstr = sqlstr & columnArrayContact(p) & ","
						Next						
						sqlstr = sqlstr & "Company_ID,date_update,Organization_ID) Values ("
						For kk=0 To Ubound(arrColContact)-1
						fWidth = columnWidthContact(kk)
						If IsNumeric(arrColContact(kk)) Then
							sqlstr = sqlstr & "'" & Left(sFix(trim(arrData(arrColContact(kk),i))), fWidth) & "',"
						Else
							sqlstr = sqlstr & "'',"
						End If						
						Next					
						sqlstr = sqlstr & companyID & ",getDate(),"&OrgID&");"						
						con.executeQuery(sqlstr)	
					End If
					set rs_check = Nothing				
					  	
				End If	
			End If									
		Next
		session("arrData") = ""
		
		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!  
		filestring="../../../download/import/double_" & trim(OrgID) & ".csv"
		'Response.Write server.mappath(filestring)
		'Response.End
		set XLSfile=fs.CreateTextFile(server.mappath(filestring)) ' Create text file as text file 
		strLine = "דוח כפולים מתאריך " & Date()
		XLSfile.writeline strLine	   
		XLSfile.writeline ""
		 
		If Len(arr_double_comp) > 1 Then
			strLine = "רשימת לקוחות כפולים"
			XLSfile.writeline strLine	    
			list_double_comp = Left(arr_double_comp, Len(arr_double_comp)-1)
			sqlstr = "Select Company_Name From Companies WHERE Company_ID IN (" & list_double_comp & ") And Organization_ID = " & OrgID
			set rs_names = con.getRecordSet(sqlstr)
			if not rs_names.eof then
				list_names = rs_names.getString(,,",",";")
			end if
			set rs_names = nothing
			If Len(list_names) > 0 Then
				list_names = Split(list_names, ";")
			End If
			If IsArray(list_names) Then	
			For i=0 To Ubound(list_names)			
			
			XLSfile.writeline list_names(i)
			Next 
			End If 
		End If 

		If Len(arr_double_cont) > 1 Then
			strLine = "רשימת אנשי קשר כפולים"
			XLSfile.writeline strLine		
			list_double_cont = Left(arr_double_cont, Len(arr_double_cont)-1)
			sqlstr = "Select Contact_Name From Contacts WHERE Contact_ID IN (" & list_double_cont & ")"
			set rs_names = con.getRecordSet(sqlstr)
			if not rs_names.eof then
				list_names = rs_names.getString(,,",",";")
			end if
			set rs_names = nothing
			If Len(list_names) > 0 Then
				list_names = Split(list_names, ";")
			End If
			If IsArray(list_names) Then	
				For i=0 To Ubound(list_names)			
					XLSfile.writeline list_names(i)
				Next 
			End If
		End If 

		
		%>
		<script language=javascript>
		<!--
			document.location.href = "Fileread.asp?end_import=1&newFileName=<%=newFileName%>";
		//-->
		</script>

		<%	
	End If
	
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 39 Order By word_id"				
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
<tr><td class="page_title" dir=rtl>&nbsp;שיוך שדות&nbsp;</td></tr>
<tr><td height=10 nowrap></td></tr>
<%If trim(Request.QueryString("end_import")) = "" Then%>
<tr><td align=right>&nbsp;<input type=button class="but_menu" value="תצוגה מקדימה של הקובץ" onclick="OpenPreview('<%="../../../download/import/" & "/" & newFileName%>')" style="width:170">&nbsp;</td></tr>
<tr><td height=10 nowrap></td></tr>
<tr><td>
<form action="Fileread.asp?import=1" method=post name=form1 id=form1>
<table border="0" width="500" bgcolor="#FFFFFF" cellspacing="1" cellpadding="1" align=center dir="<%=dir_obj_var%>">
<input type=hidden name=newFileName id=newFileName value="<%=vFix(newFileName)%>">
<%If IsArray(arrData) Then%>
<%
	If trim(lang_id) = "1" Then
		numOfTitles = Ubound(titlesArrayContact) : titlesArr = titlesArrayContact
	Else
		numOfTitles = Ubound(titlesArrayEContact) : titlesArr = titlesArrayEContact
	End If
%>
<%numOfFiedls = Ubound(arrData,1)%>
<%For i=0 To numOfTitles%>
<tr name="tr<%=i%>" id="tr<%=i%>" style="background:#e6e6e6"><td align="<%=align_var%>" width=100%>&nbsp;<%=titlesArr(i)%>&nbsp;</td>
<td align="center" width=250 nowrap>
<%is_sel = false
is_equal = -1%>
<select dir="<%=dir_obj_var%>" class="norm" style="width:220;font-size:10pt" id="<%=columnArrayContact(i)%>" name="<%=columnArrayContact(i)%>" onchange="setColor(this.value,<%=i%>,<%=is_equal%>)">
<option value="-1">ללא שיוך</option>
<%For j=0 To numOfFiedls%>
<option value="<%=j%>" 
<%
is_equal = -1
If is_sel = false Then 
    'arr_title = Split(arrXlsTitles(j),Chr(32))
    'For count = 0 To Ubound(arr_title)	
		If InStr(arrXlsTitles(j),titlesArrayContact(i)) > 0 Then
			is_equal = InStr(arrXlsTitles(j),titlesArrayContact(i))
			is_sel = true
			'Exit For
		ElseIf InStr(Lcase(arrXlsTitles(j)), Lcase(trim(titlesArrayEContact(i)))) > 0 Then
			is_equal = InStr(Lcase(arrXlsTitles(j)), Lcase(trim(titlesArrayEContact(i))))
			is_sel = true
			'Exit For
		End If				
	'Next	
End If
If 	is_equal > 0 Then%> selected<%End If%>>
<%=arrXlsTitles(j)%></option>
<%Next%>
</select></td>
</tr>
<script language=javascript>
<!--
	setColor(window.document.all("<%=columnArrayContact(i)%>"),"<%=i%>","<%=is_equal%>");	
//-->
</script>
<%Next%>
<%
	If trim(lang_id) = "1" Then
		numOfTitlesCompany = Ubound(titlesArrayCompany) : titlesArr = titlesArrayCompany
	Else
		numOfTitlesCompany = Ubound(titlesArrayECompany) : titlesArr = titlesArrayECompany
	End If
%>
<!---------------------------------------Company---------------------------------------------------->
<%For i=0 To numOfTitlesCompany%>
<%If IsArray(titlesArr) Then%>

<tr name="tr<%=numOfTitles+i+1%>" id="tr<%=numOfTitles+i+1%>" style="background:#e6e6e6"><td align="<%=align_var%>" width=100%>&nbsp;<%=titlesArr(i)%>&nbsp;</td>
<td align="center" width=250 nowrap>
<%is_sel = false
is_equal = -1%>
<select dir="<%=dir_obj_var%>" class="norm" style="width:220;font-size:10pt" id="<%=columnArrayCompany(i)%>" name="<%=columnArrayCompany(i)%>" onchange="setColor(this.value,<%=numOfTitles+i+1%>,<%=is_equal%>)">
<option value="-1">ללא שיוך</option>
<%For j=0 To numOfFiedls%>
<option value="<%=j%>" 
<%
is_equal = -1
If is_sel = false Then 
    'arr_title = Split(arrXlsTitles(j),Chr(32))
    'For count = 0 To Ubound(arr_title)	
		If InStr(arrXlsTitles(j),titlesArrayCompany(i)) > 0 Then
			is_equal = InStr(arrXlsTitles(j),titlesArrayCompany(i))
			is_sel = true
			'Exit For
		ElseIf InStr(Lcase(arrXlsTitles(j)), Lcase(trim(titlesArrayECompany(i)))) > 0 Then
			is_equal = InStr(Lcase(arrXlsTitles(j)), Lcase(trim(titlesArrayECompany(i))))
			is_sel = true
			'Exit For
		End If				
	'Next	
End If
If 	is_equal > 0 Then%> selected<%End If%>>
<%=arrXlsTitles(j)%></option>
<%Next%>
</select></td>
</tr>
<script language=javascript>
<!--
	setColor(window.document.all("<%=columnArrayCompany(i)%>"),"<%=numOfTitles+i+1%>","<%=is_equal%>");
//-->
</script>
<%End If%>
<%Next%>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
<tr>
<td colspan=2 style="background:#e6e6e6">
<table cellpadding=0 cellspacing=0 width=400 align=center dir="<%=dir_var%>">
<tr>
<td width=50% align="center"><input type=button class="but_menu" style="width:90px" onclick="window.close();" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50 nowrap></td>
<td width=50% align=center><input type=submit class="but_menu" style="width:90px" onclick="return CheckFields();" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
<%End If%>
</table>
</form>
</td></tr>
<%Else%>
<tr><td height=30 nowrap></td></tr>
<tr><td class="normalB" align=center><font color=red>ייבוא נתונים בוצע בהצלחה</font></td></tr>
<tr><td height=30 nowrap></td></tr>
<tr><td align=center><a href="../../../download/import/double_<%=trim(OrgID)%>.csv" class=link1 target=_blank>לחץ להורדת רשימת הכפולים</a></td></tr>
<tr><td height=30 nowrap colspan=2></td></tr>
<tr>
<td width=100% align="center" colspan=2>
<input type=button class="but_menu" style="width:90px" onclick="window.opener.document.location.reload(true);window.close();" value="<%=arrButtons(2)%>"></td>
</tr>
<%End If%>
</table>
</body>


