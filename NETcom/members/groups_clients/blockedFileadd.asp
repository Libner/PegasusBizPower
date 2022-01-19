<%SERVER.ScriptTimeout=3000000%>
<%Response.Buffer = False%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<script language="JavaScript">
<!--
    function remLoadMessage(){
        document.getElementById("loadMessage").style.display = "none";
    }
    function closeWin()
    {
		window.opener.location.reload(true);
		window.close();
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
.but_menu
{
    BORDER-RIGHT: #696969 1px solid;    
    BORDER-LEFT: #ffffff 1px solid;
    BORDER-TOP: #ffffff 1px solid;
    BORDER-BOTTOM: #696969 1px solid;
    PADDING-LEFT: 5px;
    PADDING-RIGHT: 5px;
    FONT-WEIGHT: bolder;
    FONT-SIZE: 12px;    
    WIDTH: 100%;
    CURSOR: hand;
    COLOR: #ffffff;    
    FONT-FAMILY: Arial;
    LETTER-SPACING: 1px;
    LINE-HEIGHT: 20px;
    BACKGROUND-COLOR: #736ba6;
    TEXT-DECORATION: none;
    text-align:center
}
</style>
</head>
<body onLoad="remLoadMessage" bgcolor="#E6E6E6">
<div style="background-color:#E6E6E6" id="loadMessage">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
    <tr><td height="30"></td></tr>
	<tr>
		<td align="center" nowrap class="normalB" dir=rtl>תהליך טעינת הנתונים לאתר מתבצעת כעת, התהליך יארך מספר דקות</td>
	</tr>		
	<tr><td height="10"></td></tr>	
	<tr>
		<td  align="center" nowrap class="normalB" dir=rtl><font color="red">אין לסגור חלון זה ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</font></td>
	</tr>	
	<tr><td height="20"></td></tr>
	<tr><td align=center><img src="../../images/butt_act.gif" border=0 vspace=0 hspace=0></td></tr>
	<tr><td height="10"></td></tr>
</table>
</div>
</body>
</html>
<%
    UserId = trim(Request.Cookies("bizpegasus")("UserID"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    
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
  
	set upl = Server.CreateObject("SoftArtisans.FileUp") 
	upl.path=server.mappath("../../../download/")		
	groupid = upl.Form("groupID")	
	If trim(groupid) <> "" Then
		sqlstr = "Select GROUPNAME From GROUPS Where ORGANIZATION_ID = "& OrgID & " and groupId =" & groupid 
		set rs_name = con.getRecordSet(sqlstr)
		If not rs_name.eof Then
			GroupName = rs_name(0)
		Else
			GroupName = ""
		End If
		set rs_name = Nothing		
	Else
		GroupName = ""
	End If
		   	
	newFileName = OrgID & "_" & second(now) & "_blocked_mails.xls"
	upl.Form("UploadFile").SaveAs newFileName 
			
	'reading records from excell file		
	Const adOpenStatic = 3
	Const adLockPessimistic = 2

	Dim cnnExcel
	Dim rstExcel
	Dim I
	Dim iCols

	Set cnnExcel = Server.CreateObject("ADODB.Connection")
	cnnExcel.CommandTimeout = 0
	cnnExcel.ConnectionTimeout = 0

	cnnExcel.Open "DBQ=" & Server.MapPath("../../../download/"& newFileName ) & ";" & _
	"DRIVER={Microsoft Excel Driver (*.xls)};"
	
' Same as any other data source.
' FYI: mytable is my named range in the Excel file
Set rstExcel = Server.CreateObject("ADODB.Recordset")
rstExcel.Open "SELECT Email FROM Peoples where rtrim(ltrim(Email)) <> '';", cnnExcel, _
	adOpenStatic, adLockPessimistic, adCmdText

' Get a count of the fields and subtract one since we start
' counting from 0.
iCols = rstExcel.RecordCount
'Response.Write "iCols=" & iCols & "<br>"
if not rstExcel.EOF then
	'rstExcel.MoveNext ' this is row of selolar number and name on hebrew, don't add.
	arrBlockedMails = rstExcel.GetRows()
end if
rstExcel.Close
Set rstExcel = Nothing

'//loop over the excell record
'//insert just the correct record which have valid Email and nort null name
totalCount  = 0
If IsArray(arrBlockedMails) Then
	For i=1 To Ubound(arrBlockedMails,2)
		excell_Email = LCase(trim(arrBlockedMails(0,i)))				
		If excell_Email <> "" And IsEmailValid(excell_Email) Then			
		    sqlstr = "Select Count(People_ID) FROM PEOPLES WHERE PEOPLE_EMAIL LIKE '" & sFix(excell_Email) &_
			"' AND ORGANIZATION_ID = "& OrgID			
			If trim(groupid) <> "" Then
			    sqlstr = sqlstr & " and groupId =" & groupid 
			End If 
			set rs_count = con.getRecordSet(sqlstr)
			If not rs_count.eof Then
				numOfPeoples = rs_count(0)
			Else
			    numOfPeoples = 0
			End If	
			set rs_count = Nothing
			totalCount =  totalCount + numOfPeoples
			
			sqlstr = "DELETE FROM PRODUCT_CLIENT WHERE PEOPLE_EMAIL LIKE '" & sFix(excell_Email) &_
			"' AND ORGANIZATION_ID = "& OrgID
			If trim(groupid) <> "" Then
			    sqlstr = sqlstr & " AND groupId =" & groupid 
			End If 					
			con.executeQuery(sqlstr)		
			
			sqlstr = "DELETE FROM PEOPLES WHERE PEOPLE_EMAIL LIKE '" & sFix(excell_Email) &_
			"' AND ORGANIZATION_ID = "& OrgID
			If trim(groupid) <> "" Then
			    sqlstr = sqlstr & " AND groupId =" & groupid 
			End If 						
			con.executeQuery(sqlstr)
						
		End If		
	Next
'end of reading records from excell file 
End If
 
cnnExcel.Close
Set cnnExcel = Nothing


set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
filetocopy=server.MapPath("../../../download/" & newFileName)

if fs.FileExists(filetocopy) then
	fs.DeleteFile filetocopy,true
end if

set validate=nothing
%>
<script language=javascript>
<!--
    remLoadMessage()
//-->
</script>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align=center>
    <tr><td height="30"></td></tr>
	<tr>
		<td align="center" nowrap class="normalB" dir=rtl>תהליך טעינת הנתונים הסתיים</td>
	</tr>		
	<tr><td height="10"></td></tr>	
	<tr>
		<td align="center" nowrap class="normalB" dir=rtl>הוסרו <font color=red><%=totalCount%></font>
		<%If Len(GroupName) > 0 Then%>
		נמענים מקבוצת <font color=red><%=GroupName%></font>
		<%Else%>
		נמענים מכל הקבוצות
		<%End If%>
		</td>
	</tr>
	<tr><td height="20"></td></tr>
	<tr><td align=center><input type=button value="סגור חלון" onclick="closeWin()" class="but_menu" style="width:100"></td></tr>
	<tr><td height="10"></td></tr>
</table>
