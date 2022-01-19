<%SERVER.ScriptTimeout=3000000%>
<%Response.Buffer = False%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<script language="JavaScript">
<!--
    function remLoadMessage(){
        document.getElementById("loadMessage").style.display = "none";
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
</body>
</html>
<%
    UserId = trim(Request.Cookies("bizpegasus")("UserID"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    double_phone = 0
    
    Function IsCellPhoneValid(strCellPhone)  

		Dim blnIsItValid 
	   
		' assume the phone number is correct   
		blnIsItValid = True 
		
		If IsNull(strCellPhone) Then
			blnIsItValid = False 
			IsCellPhoneValid = blnIsItValid 
			Exit Function
		End If
	
		If Len(strCellPhone) < 10 Then 
			blnIsItValid = False 
			IsCellPhoneValid = blnIsItValid 
			Exit Function 
		End If 
		
		For cc = 1 To Len(strCellPhone)
			c = LCase(Mid(strCellPhone, cc, 1)) 
			' if there is an illegal character in the part 
			If Not IsNumeric(c) Then 
				blnIsItValid = False 
				IsEmailValid = blnIsItValid 
				Exit Function 
			End If 
		Next 		
	   
		' finally it's OK   
		IsCellPhoneValid = blnIsItValid 
     
End Function 

groupid = Request.QueryString("groupid")
If trim(groupid) <> "" Then
	sqlstr = "SELECT group_name From sms_groups WHERE (ORGANIZATION_ID = "& OrgID & ") AND (group_id =" & groupid & ")"
	set rs_name = con.getRecordSet(sqlstr)
	If not rs_name.eof Then
		group_name = rs_name(0)
	Else
		group_name = ""
	End If
	set rs_name = Nothing		
Else
	group_name = ""
End If

	set upl = Server.CreateObject("SoftArtisans.FileUp") 
	upl.path = server.mappath("../../../download/smsgroups/")
	newFileName = Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
	newFileName = groupid & "_" & second(now) & "_data_temp.xls"
	
	''Response.Write newFileName
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

cnnExcel.Open "DBQ=" & Server.MapPath("../../../download/smsgroups/"& newFileName ) & ";" & _
"DRIVER={Microsoft Excel Driver (*.xls)};"
	
' Same as any other data source.
' FYI: mytable is my named range in the Excel file
Set rstExcel = Server.CreateObject("ADODB.Recordset")
rstExcel.Open "SELECT Cell, Name, Company, Position, Phone, Fax From Peoples WHERE Cell is Not NULL;", cnnExcel, _
	adOpenStatic, adLockPessimistic, adCmdText	

' Get a count of the fields and subtract one since we start
' counting from 0.
iCols = rstExcel.RecordCount
'Response.Write "iCols=" & iCols & "<br>"
if not rstExcel.EOF then
	'rstExcel.MoveNext ' this is row of selolar number and name on hebrew, don't add.
	arrPeoples = rstExcel.GetRows()
end if
rstExcel.Close
Set rstExcel = Nothing

'//loop over the excell record
'//insert just the correct record which have valid cellular and nort null name
If IsArray(arrPeoples) Then
numOfPeoples = 0
For i=1 To Ubound(arrPeoples,2)
'Do while not rstExcel.EOF 
	
	excell_Cell = LCase(trim(arrPeoples(0,i)))
	excell_Cell = Replace(excell_Cell, Chr(160), "")
	excell_Cell = Replace(excell_Cell, "-", "")	
	excell_Field1 = trim(arrPeoples(1,i))
	excell_Field2 = trim(arrPeoples(2,i))
	excell_Field3 = trim(arrPeoples(3,i))
	excell_Field4 = trim(arrPeoples(4,i))
	excell_Field5 = trim(arrPeoples(5,i))
			
	If trim(excell_Cell) = "" Then
	
		strxls = strxls & excell_Cell & "," & excell_Field1 & "," & excell_Field2 & "," & excell_Field3 &_
		"," & excell_Field4 & "," & excell_Field5 & ",טלפון נייד ריק" & vbCrLf
		double_phone = double_phone + 1
		
	ElseIf Not IsCellPhoneValid(excell_Cell) Then
	
		strxls = strxls & excell_Cell & "," & excell_Field1 & "," & excell_Field2 & "," & excell_Field3 &_
		"," & excell_Field4 & "," & excell_Field5 & ",טלפון נייד לא חוקי" & vbCrLf
		double_phone = double_phone + 1
		
	Else	
			
		sqlstr = "SELECT PEOPLE_ID FROM sms_peoples WHERE PEOPLE_CELL LIKE ('" & sFix(excell_Cell) &_
		"') AND (ORGANIZATION_ID = "& OrgID & ") AND (group_id =" & groupid & ")"
		set rs_check = con.getRecordSet(sqlstr)
		if rs_check.eof then
				
		excell_Field1 = sfix(left(excell_Field1,200))
		excell_Field2 = sfix(left(excell_Field2,200))
		excell_Field3 = sfix(left(excell_Field3,200))
		excell_Field4 = sfix(left(excell_Field4,200))		
		excell_Field5 = sfix(left(excell_Field5,200))		
		
		sqlStr = "INSERT INTO sms_peoples (PEOPLE_CELL, PEOPLE_NAME, PEOPLE_COMPANY, "&_
		" PEOPLE_OFFICE, PEOPLE_PHONE, PEOPLE_FAX, User_ID,ORGANIZATION_ID, group_id) VALUES ('" &_
		sFix(excell_Cell) & "','" & excell_Field1 & "','" & excell_Field2 & "','" & excell_Field3 & "','" & excell_Field4 & "','" & _
		excell_Field5 & "'," & UserID & "," & OrgID & "," & groupid & ")"
		con.ExecuteQuery(sqlStr)
		numOfPeoples = numOfPeoples + 1
		Else
			double_phone = double_phone + 1
			strxls = strxls & excell_Cell & "," & excell_Field1 & "," & excell_Field2 & "," & excell_Field3 &_
			"," & excell_Field4 & "," & excell_Field5 & ",טלפון נייד כפול" & vbCrLf
		end if
		set rs_check = nothing		
	End If
Next

'end of reading records from excell file			
 end if
 
cnnExcel.Close
Set cnnExcel = Nothing

set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
filetocopy=server.MapPath(str_path & newFileName)
if fs.FileExists(filetocopy) then
	fs.DeleteFile filetocopy,true
end if

strUrl = "group.asp?groupId=" & groupId %>
<script language=javascript>
<!--
      remLoadMessage();
      function closeWin()
      {
		window.opener.location.href = "<%=strUrl%>";
		window.close();
	  }	   
//-->
</script>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align=center>
    <tr><td height="20"></td></tr>
	<tr>
		<td align="center" nowrap class="normalB" dir=rtl>תהליך טעינת הנתונים הסתיים</td>
	</tr>		
	<tr><td height="10"></td></tr>	
	<tr>
		<td align="center" nowrap class="normalB" dir=rtl>נוספו <font color=red><%=numOfPeoples%></font> נמענים לקבוצת <font color=red><%=group_name%></font></td>
	</tr>
	<tr>
		<td align="center" nowrap class="normalB" dir=rtl><font color=red><%=double_phone%></font> טלפונים ניידים כפולים בקובץ Excel</td>
	</tr>
<%set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!  
	filestring="../../../download/import/wrong_sms_" & trim(groupid) & "_" & Second(Now()) & ".csv"
	'Response.Write server.mappath(filestring)
	'Response.End
	set XLSfile=fs.CreateTextFile(server.mappath(filestring)) ' Create text file as text file 
	strLine = Date()
	XLSfile.writeline strLine	   
	XLSfile.writeline ""
 
	strLine = "רשימת טלפונים ניידים שגוים / ריקים "
	XLSfile.writeline strLine
	XLSfile.writeline ""	    
	strLine = "טלפון נייד , שם, חברה, תפקיד, טלפון, פקס, שגיאה"
	XLSfile.writeline strLine
	XLSfile.writeline strxls	
	XLSfile.Close()
	Set XLSfile = Nothing %>
<tr><td align=center><a href="<%=filestring%>" class=normalB target=_blank>לחץ להורדת מספרי טלפון שגוים / ריקים</a></td></tr>
<tr><td height="20"></td></tr>
<tr><td align=center><input type=button value="סגור חלון" onclick="closeWin()" class="but_menu" style="width:100"></td></tr>
<tr><td height="10"></td></tr>
</table>
</body>
</html>
<%Set con = Nothing%>