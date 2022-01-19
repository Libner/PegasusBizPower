<%@ Language=VBScript%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
if Request.Cookies("bizpegasus")("AddInsurance")="1" then
			sql_InsuranceCount="Select InsuranceSend_Counter from Users where USER_ID="& UserId
			set rs_InsuranceCount=con.getRecordSet(sql_InsuranceCount)
				if not rs_InsuranceCount.eof then
					SendInsuranceCount=rs_InsuranceCount("InsuranceSend_Counter")
				end if
	    end if

	    %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="stylesheet" type="text/css">
		<script src="http://code.jquery.com/jquery-1.9.1.js"></script>
	
		<script language="javascript" type="text/javascript">
<!--
	var oPopup = window.createPopup();

function openDetails(IDate,IName,LeadID)
{


  var posx = 0;

 var posy = 0;

  if (!e) var e = window.event;

if (e.pageX || e.pageY)

{

   posx = e.pageX;

   posy = e.pageY+document.body.scrollTop;

   }

 else if (e.clientX || e.clientY)

  {

    posx = e.clientX;
    posy = e.clientY+document.body.scrollTop;

 }
var Layertext =""
Layertext="<table border=0 cellspacing=2 cellpadding=2 align=right width=200><tr><TD align=right dir=rtl>הביטוח נשלח ע''  "+"  " + IName +"</td></tr><TR><TD align=right> בתאריך " + IDate+"</td></tr><TR><TD align=right>"+LeadID +" :מספר הפניה </td></tr></table>"


	document.getElementById('pos').innerHTML = Layertext;

	document.getElementById('pos').style.left = posx;

document.getElementById('pos').style.top = posy;

document.getElementById('pos').style.display='block';

}
function CloseDiv()
{
document.getElementById('pos').style.display='none';

}
function SendInsuranceConfirmGroup()
{
//alert("sendInsuranceConfirmGroup")
var Layertext =""
Layertext="<table cellpadding=0 cellspacing=0 style=height:200px;width:320px;padding:10px><tr><td height=20></td></tr><tr><td valign=top><table cellpadding=0 cellspacing=0><tr><td align =right  dir=rtl colspan=2 style=font-size:12pt;color:#ffffff>שים לב, <br>"
Layertext=Layertext+ "זוהי תזכורת שבה הנך מחוייב/ת לשלוח פרטים רק במידה וקיבלת אישור מהלקוח שהוא מאשר שמוכן לקבל פרטים בעבור עשיית ביטוח נסיעות.</td></tr>"
Layertext=Layertext+ "<tr><td align=right colspan=2 style=font-size:12pt;color:#ffffff dir=rtl>אין בשום אופן אישור לשלוח פרטי הלקוח לחברת הביטוח כאשר הלקוח אינו מעוניין בכך ויש להתנהל לפי הנוהל הנ''ל ללא יוצא מהכלל</td><tr>"
Layertext=Layertext+ "<tr><td align=center height=40 colspan=2></td></tr>"
Layertext=Layertext+ "<tr><td align=center ><a href=javascript:DeletePermision('delete') class=button_edit_1 style=width:106px> ביטול</a></td><td align=center>"
Layertext=Layertext+ "<a href=# onclick='javascript:SendInsuranceGroup()' class=button_edit_1 style=width:106px>אני מאשר/ת</a></td></tr>"
Layertext=Layertext+ "</td></tr><tr><td align=center height=20></td></tr>"
Layertext=Layertext+ "</table></td></tr></table>"
//alert(Layertext)
	document.getElementById('confirmBlock').innerHTML = Layertext;
		document.getElementById("confirmBlock").style.left ="680px";
	document.getElementById("confirmBlock").style.top = "400px";
    document.getElementById("confirmBlock").style.display='block';
}
function sendInsuranceConfirm(ContactId,companyID,AppealId)
{
var Layertext =""
Layertext="<table cellpadding=0 cellspacing=0 style=height:200px;width:320px;padding:10px><tr><td height=20></td></tr><tr><td valign=top><table cellpadding=0 cellspacing=0><tr><td align =right  dir=rtl colspan=2 style=font-size:12pt;color:#ffffff>שים לב, <br>"
Layertext=Layertext+ "זוהי תזכורת שבה הנך מחוייב/ת לשלוח פרטים רק במידה וקיבלת אישור מהלקוח שהוא מאשר שמוכן לקבל פרטים בעבור עשיית ביטוח נסיעות.</td></tr>"
Layertext=Layertext+ "<tr><td align=right colspan=2 style=font-size:12pt;color:#ffffff dir=rtl>אין בשום אופן אישור לשלוח פרטי הלקוח לחברת הביטוח כאשר הלקוח אינו מעוניין בכך ויש להתנהל לפי הנוהל הנ''ל ללא יוצא מהכלל</td><tr>"
Layertext=Layertext+ "<tr><td align=center height=40 colspan=2></td></tr>"
Layertext=Layertext+ "<tr><td align=center ><a href=javascript:DeletePermision('delete') class=button_edit_1 style=width:106px> ביטול</a></td><td align=center>"
Layertext=Layertext+ "<a href=javascript:sendInsurance('"+ ContactId + "','" + companyID +"','" + AppealId + "') class=button_edit_1 style=width:106px>אני מאשר/ת</a></td></tr>"
Layertext=Layertext+ "</td></tr><tr><td align=center height=20></td></tr>"
Layertext=Layertext+ "</table></td></tr></table>"
//alert(Layertext)
	document.getElementById('confirmBlock').innerHTML = Layertext;
		document.getElementById("confirmBlock").style.left ="680px";
	document.getElementById("confirmBlock").style.top = "400px";
    document.getElementById("confirmBlock").style.display='block';


//h = parseInt(530);
//w = parseInt(470);
//window.open("sendInsuranceConfirm.asp?quest_id=16735&ContactId="+ ContactId+ "&CompanyId="+ companyID + "&appid="+AppealId, "T_Wind" ,"scrollbars=1,toolbar=0,top=100,left=20,width="+w+",height="+h+",align=center,resizable=0");
}

function sendInsurance(ContactId,companyID,AppealId)
{
	document.getElementById('confirmBlock').style.display='none';

	var r = confirm("האם הלקוח אישר שפרטיו ישחלו לחברת הביטוח\n    ?ליצירת קשר להצעת פוליסת ביטוח נסיעות");
	if (r == true)
	{
	

h = parseInt(530);
w = parseInt(470);
window.open("../companies/Insurance_Send.asp?quest_id=16735&ContactId="+ ContactId+ "&CompanyId="+ companyID + "&appid="+AppealId, "T_Wind" ,"scrollbars=1,toolbar=0,top=100,left=20,width="+w+",height="+h+",align=center,resizable=0");
}
}
	function DeletePermision(obj)
	{
	//alert ("שים לב,אחרי שינוי הרשאות יש להכנס מחדש למערכת");
	document.getElementById("confirmBlock").style.display='none';

	window.open("../companies/UpdatePermisionInsurence.asp", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");
//	 document.location.href='UpdatePermisionInsurence.asp'
	}
-->
		</script>
	</head>
	<%sqlPer="SELECT  bar_id  FROM bar_users WHERE (user_id =" & UserID &") and bar_id=47 and is_visible=1"
 set rs_Per = con.getRecordSet(sqlPer)
			if not rs_Per.eof then	
			SMS="1"
			else
			SMS="0"
		end if
		set rs_Per= nothing   %>
	<%app_status = trim(Request("app_status"))
app_TourId=trim(request("app_TourId"))
PinpsearchDep=request("inpsearchDep")
if app_TourId="" then app_TourId=0

app_GuideId=trim(request("app_GuideId"))
if app_GuideId="" then app_GuideId=0
app_DepId=trim(request("app_DepId"))
if app_DepId="" then app_DepId=0


app_Country = trim(Request("app_Country"))
if app_Country="" then app_Country=0
'response.Write app_Country
	arch = trim(Request("arch"))
	if arch = "" then
		arch = 0
	end if

	if arch = "0" then
		set_arch = 1
	else
		set_arch = 0
	end if

	where_arch = " and appeal_deleted=" & arch

if (Request("dt")<>nil) then
dt=month(Request("dt")) &"/"& day(Request("dt")) &"/"& year(Request("dt"))
end if 
'response.Write "dt="& dt


	if (Request("prodId") <> nil) then 'and Request.QueryString("delProd") = nil) then
		prodId = Request("prodId")
		where_product_id = " and questions_id=" & prodId
		set prod = con.GetRecordSet("Select product_name, ADD_CLIENT from products where product_id=" & prodId)
		if not prod.eof then
			title_prodName = trim(prod("product_name"))
			ADD_CLIENT = trim(prod("ADD_CLIENT"))
		end if
		set prod = nothing
	else
		sqlstr = "Select Top 1 product_id, product_name, ADD_CLIENT from Products Where "&_
		" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
		set rs_prod = con.getRecordSet(sqlstr)
		If not rs_prod.eof Then
			prodId = rs_prod(0)
			title_prodName = rs_prod(1)
			ADD_CLIENT = trim(prod("ADD_CLIENT"))
		End If
		set rs_prod = Nothing
	end if

	if Request.QueryString("clientId") <> nil then
		clientId = Request.QueryString("clientId")
		where_client_id = " and contact_id=" & clientId
	else
		where_client_id = ""	
	end if

	if trim(Request("orgname")) <> "" then
		orgname = trim(Request("orgname")) 
		where_orgname = " and LOWER(COMPANY_NAME) LIKE '%"& LCase(trim(Request.Form("orgname"))) &"%'"		
	else
		where_orgname = ""
	end if
	
	if trim(Request("clname")) <> "" then
		clname = trim(Request("clname"))
		where_clname = " and LOWER(CONTACT_NAME) LIKE '%"& LCase(trim(Request.Form("clname"))) &"%'"		
	else
		where_clname = ""
	end if
	
	if trim(Request("usname")) <> "" then
		usname = trim(Request("usname"))		
	else
		usname = ""
	end if
	
	if trim(Request("orderOwner")) <> "" then
		orderOwner = trim(Request("orderOwner"))		
	else
		orderOwner = ""
	end if	

	if trim(Request("search_id")) <> "" then
		search_id = trim(Request("search_id"))
		where_id = " and APPEAL_ID = "& trim(Request.Form("search_id"))
	else
		where_id = ""
	end if

If IsNumeric(prodId) Then
   set rs_fields=con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type FROM FORM_FIELD WHERE PRODUCT_ID=" & prodId &" And ORGANIZATION_ID = "& OrgID &" AND FORM_FIELD.Field_Key = 1 Order by Field_Order DESC")
   If Not rs_fields.EOF Then
	    arr_fields = rs_fields.getRows()
   End If
   Set rs_fields = Nothing
   if IsArray(arr_fields) Then
		count_fields = Ubound(arr_fields, 2)
   Else
		count_fields = 0
   End If
   
  urlSort = "appealsDetails16735.asp?dep=" & request("dep")& "&arch=" & arch & "&prodId=" & prodId &"&dt="& Request("dt") & "&app_Country=" & app_Country & "&app_TourId=" & app_TourId & "&app_GuideId=" & app_GuideId & "&app_DepId=" & app_DepId & "&app_status=" & app_status & "&search_id=" & search_id & _
  "&clname=" & Server.URLEncode(clname) & "&orgname=" & Server.URLEncode(orgname) & _
  "&usname=" & Server.URLEncode(usname) & "&orderOwner=" & Server.URLEncode(orderOwner) 
  
	'dynamic fields search 
	 if IsArray(arr_fields) Then
	If count_fields >= 0 Then
		str_where_values = ""				
		If Request.Form.Count > 0 Then
			For ss=0 To count_fields							
				srch_val = ""
				field_type = trim(arr_fields(2, ss))
				field_title = trim(arr_fields(1, ss))
				
				if Not isNULL( Request.Form("inpsearch" & trim(arr_fields(0, ss))) ) Then
					If Len(Request.Form("inpsearch" & trim(arr_fields(0, ss)))) > 0 Then
						srch_val = trim(Request.Form("inpsearch" & trim(arr_fields(0, ss))))
					End If	
				End If
				Execute("srch_val" & trim(arr_fields(0, ss)) & "= """ & vFix(srch_val) & """")
				urlSort = urlSort & "&srch_val" & trim(arr_fields(0, ss)) & "=" & Server.URLEncode(Eval("srch_val" & trim(arr_fields(0, ss))))
				If trim(srch_val) <> "" Then
					str_where_values = str_where_values & " AND Appeal_ID IN (SELECT Appeal_ID FROM  FORM_VALUE WHERE FIELD_ID = " & trim(arr_fields(0, ss)) 
					If trim(field_type) = "5" And trim(srch_val) = trim(field_title) Then
						str_where_values = str_where_values & " AND FIELD_VALUE = 'on' )"
					Else
						str_where_values = str_where_values & " AND FIELD_VALUE Like '%" &  sFix(srch_val) & "%')"						
					End If
				End If	
				'Response.Write(Eval("srch_val" & trim(arr_fields(0, ss))) & "<br>") 
			Next
			Else
				For ss=0 To count_fields	
					srch_val = ""
					field_type = trim(arr_fields(2, ss))
					field_title = trim(arr_fields(1, ss))

					if Not isNULL( Request.QueryString("srch_val" & trim(arr_fields(0, ss))) ) Then
						If Len(Request.QueryString("srch_val" & trim(arr_fields(0, ss)))) > 0 Then
							srch_val = trim(Request.QueryString("srch_val" & trim(arr_fields(0, ss))))
						End If	
					End If
					Execute("srch_val" & trim(arr_fields(0, ss)) & "= """ & vFix(srch_val) & """")
					urlSort = urlSort & "&srch_val" & trim(arr_fields(0, ss)) & "=" & Server.URLEncode(Eval("srch_val" & trim(arr_fields(0, ss))))
					If trim(srch_val) <> "" Then
						str_where_values = str_where_values & " AND Appeal_ID IN (SELECT Appeal_ID FROM  FORM_VALUE WHERE FIELD_ID = " & trim(arr_fields(0, ss)) 
						If trim(field_type) = "5" And trim(srch_val) = trim(field_title) Then
							str_where_values = str_where_values & " AND FIELD_VALUE = 'on' )"
						Else
							str_where_values = str_where_values & " AND FIELD_VALUE Like '%" &  sFix(srch_val) & "%')"						
						End If					
					End if	
					Next
				End if 
			End if 
		End if 
	End If
	
	dim sortby(16)	
	sortby(1) = " appeal_date, appeal_status, appeal_id DESC"
	sortby(2) = " appeal_date DESC, appeal_status, appeal_id DESC"
	sortby(3) = " appeal_id"
	sortby(4) = " appeal_id DESC"
	sortby(5) = " U.FIRSTNAME, U.LASTNAME, appeal_id DESC"
	sortby(6) = " U.FIRSTNAME DESC, U.LASTNAME DESC, appeal_id DESC"
	sortby(7) = " U1.FIRSTNAME, U1.LASTNAME, appeal_id DESC"
	sortby(8) = " U1.FIRSTNAME DESC, U1.LASTNAME DESC, appeal_id DESC"
	sortby(9) = " CONTACT_NAME , appeal_id DESC"
	sortby(10) = "CONTACT_NAME DESC, appeal_id DESC"
	sortby(11) = " COMPANY_NAME, appeal_id DESC"
	sortby(12) = " COMPANY_NAME DESC, appeal_id DESC"
    sortby(13) = " Insurance_Status,appeal_id DESC"
    sortby(14) = " Insurance_Status DESC,appeal_id DESC"
    
    sortby(15) = " DepartmentName"
    sortby(16) = " DepartmentName DESC"

	sort = Request("sort")	
	if sort = nil then
		sort = 2
	end if
	
	
	
'response.end
' ------ delete appeals --------


' ------ start deleting appeals ----
if (Request.QueryString("delProd") <> "") And chief then
	delappid = Request.QueryString("appId")			
	'------ start deleting the new message from XML file ------
	xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
	set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	objDom.async = false			
	if objDOM.load(server.MapPath(xmlFilePath)) then
		set objNodes = objDOM.getElementsByTagName("FORM")		
		for j=0 to objNodes.length-1
			set objTask = objNodes.item(j)
			node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
			if trim(Request.QueryString("appId")) = trim(node_app_id) Then							
				objDOM.documentElement.removeChild(objTask)
				exit for
			else
				set objTask = nothing
			end if
		next
		Set objNodes = nothing
		set objTask = nothing
		'objDom.save server.MapPath(xmlFilePath)
	end if
	set objDOM = nothing
	
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	'-----deleting files----
	sqlstr = "SELECT D.document_file FROM dbo.appeals_documents AD INNER JOIN  dbo.documents D " & _
	" ON AD.document_id = D.document_id WHERE appeal_id = " & delappid
	'Response.Write sqlstr
	'Response.End
	set files = con.getRecordSet(sqlstr)
	do while not files.eof
		file_path="../../../download/documents/" & trim(files("document_file"))
		if fs.FileExists(server.mappath(file_path)) then
			set a = fs.GetFile(server.mappath(file_path))
			a.delete '------------------------------		
		end if	
	files.movenext
	loop
	set files =nothing
	set fs = nothing	
	
	con.executeQuery("DELETE FROM DOCUMENTS WHERE document_id IN (Select document_id From appeals_documents Where Appeal_ID = " & delappid & ")")
	con.executeQuery("DELETE FROM appeals_documents WHERE Appeal_ID = " & delappid)	
	con.ExecuteQuery("UPDATE tasks SET appeal_id = NULL WHERE appeal_id =" & delappid )
	con.ExecuteQuery("DELETE FROM Form_Value WHERE appeal_id =" & delappid )
	con.ExecuteQuery("DELETE FROM contact_to_forms WHERE appeal_id =" & delappid )	
	
	'--insert into changes table
	sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
	"SELECT 'טופס', 'שם: ' + P.Product_Name + ', איש קשר:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, 'מחיקה', getDate(), " & UserId & _
	" FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
	" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = " & delappid & ")	"
	con.executeQuery(sqlstr)	
	
	con.ExecuteQuery("DELETE FROM appeals WHERE appeal_id =" & delappid)
	
	Response.Redirect "appealsDetails16735.asp?prodId=" & prodId &"&dt="& request("dt") & "&arch=" & arch &"&dep="& request("dep")
end if 

' ------ end deleting appeals --------

' ------ transfering to archive --------
'response.Write Request.Form("trapp")
'
'response.Write  ( Request.Form("delete_flag") &":"& Request.Form("change_status_flag") &":"& Request.Form("status_flag"))
'response.end
if trim(Request.Form("trapp")) <> "" And trim(Request.Form("delete_flag")) = "0" And trim(Request.Form("change_status_flag")) = "0"  And trim(Request.Form("site_flag")) = "0" then
'response.end

	sqlstr = "UPDATE appeals SET appeal_deleted=" & set_arch & " WHERE appeal_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	con.ExecuteQuery(sqlstr)
	Response.Redirect urlSort '"appealsDetails16735.asp?prodId=" & prodId & "&arch=" & arch
end if
' ------ transfering to archive --------
if trim(Request.Form("site_flag")) = "1" and  trim(Request.Form("trapp")) <> ""  then
'response.Write "site="& Request.Form("site_flag")
	sqlstr = "SELECT appeal_id,Tour_Id,Departure_Id,Guide_Id,File_ThankLetter FROM appeals WHERE appeal_id IN (" & Request.Form("trapp") & ")"
    set rs_site = con.getRecordSet(sqlstr)
    	 do while not rs_site.EOF 
    '	***************************
    appeal_id=rs_site("appeal_id")
   ' response.Write "appeal_id="&appeal_id
    if not IsNull(rs_site("Tour_Id")) then
    Tour_Id=rs_site("Tour_Id")
    else
    Tour_Id=0
    end if
    
      if not IsNull(rs_site("Departure_Id")) then
    Departure_Id=rs_site("Departure_Id")
    else
    Departure_Id=0
    end if
       if not IsNull(rs_site("Guide_Id")) then
    Guide_Id=rs_site("Guide_Id")
    else
    Guide_Id=0
    end if
    
   

     File_ThankLetter=rs_site("File_ThankLetter")
     File_Path="http://pegasus.bizpower.co.il/download/thanksLetter"
   sqlFiled="select  dbo.Get_FieldValue40767(appeal_id) as DescName, dbo.Get_FieldValue40769(appeal_id) as F1, dbo.Get_FieldValue40770(appeal_id) as F2,  dbo.Get_FieldValue40771(appeal_id) as F3  from FORM_VALUE where appeal_id="&appeal_id
  '   response.Write sqlFiled
  '   response.end
       set rs_Field = con.getRecordSet(sqlFiled)
       if not rs_Field.eof then
			Description=rs_Field("DescName")
			F1=rs_Field("F1")
			F2=rs_Field("F2")
			F3=rs_Field("F3")
       end if

   
   
       	cmdInsert = "Exec dbo.ThanksLetter_insert  " &  Trim(appeal_id) & "," &  Trim(Tour_Id) & "," & Trim(Departure_Id) & "," &  Trim(Guide_Id) & ",'" & Trim(File_Path) &"','"& Trim(File_ThankLetter) &"','"& sfix(Trim(Description)) &"','"&  Trim(F1) &"','"&  Trim(F2) &"','"&  Trim(F3) &"'"
	'response.Write cmdInsert
	'response.end
	con.ExecuteQuery(cmdInsert)		
 
    '   ******************************
    rs_site.MoveNext
    loop
	Set rs_site = Nothing
	sqlUpd=  "update appeals set  ThankLetter_SendDate =GetDate() WHERE appeal_id IN (" & Request.Form("trapp") & ")"
con.ExecuteQuery(sqlUpd)

'	Response.Redirect "appealsDetails16735.asp?prodId=" & prodId & "&arch=" & arch	
end if

'------ start changing appeals statuses --
if trim(Request.Form("trapp")) <> "" And trim(Request.Form("change_status_flag")) = "1" then
	If isNumeric(Request.Form("cmb_change_status")) And trim(Request.Form("cmb_change_status")) <> "" Then
		cmb_change_status = cInt(Request.Form("cmb_change_status"))
	Else
		cmb_change_status = 1
	End If	
	sqlstr = "UPDATE APPEALS SET APPEAL_STATUS=" & cmb_change_status & " WHERE appeal_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	con.ExecuteQuery(sqlstr)
	Response.Redirect urlSort '"appealsDetails16735.asp?prodId=" & prodId & "&arch=" & arch
end if
'------ end changing appeals statuses ---


if trim(Request.Form("delete_flag")) = "1" then
	sqlstr = "SELECT appeal_id FROM appeals WHERE appeal_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	set rs_delete = con.getRecordSet(sqlstr)
	While not rs_delete.eof
		delappid = rs_delete.Fields(0)		
		
		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
		'-----deleting files----
		sqlstr = "SELECT D.document_file FROM dbo.appeals_documents AD INNER JOIN  dbo.documents D " & _
		" ON AD.document_id = D.document_id WHERE appeal_id = " & delappid
		'Response.Write sqlstr
		'Response.End
		set files = con.getRecordSet(sqlstr)
		do while not files.eof
			file_path="../../../download/documents/" & trim(files("document_file"))
			if fs.FileExists(server.mappath(file_path)) then
				set a = fs.GetFile(server.mappath(file_path))
				a.delete '------------------------------		
			end if	
		files.movenext
		loop
		set files =nothing
		set fs = nothing
			
		'------ start deleting the new message from XML file ------
		xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false			
		if objDOM.load(server.MapPath(xmlFilePath)) then
			set objNodes = objDOM.getElementsByTagName("FORM")		
			for j=0 to objNodes.length-1
				set objTask = objNodes.item(j)
				node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
				if trim(delappid) = trim(node_app_id) Then							
					objDOM.documentElement.removeChild(objTask)
					exit for
				else
					set objTask = nothing
				end if
			next
			Set objNodes = nothing
			set objTask = nothing
			'objDom.save server.MapPath(xmlFilePath)
		end if
		set objDOM = nothing
		
		con.executeQuery("DELETE FROM DOCUMENTS WHERE document_id IN (Select document_id From appeals_documents Where Appeal_ID = " & delappid & ")")
		con.executeQuery("DELETE FROM appeals_documents WHERE Appeal_ID = " & delappid)			
		con.ExecuteQuery("UPDATE tasks SET appeal_id = NULL WHERE appeal_id =" & delappid )
		con.ExecuteQuery("DELETE FROM Form_Value WHERE appeal_id =" & delappid )
		con.ExecuteQuery("DELETE FROM contact_to_forms WHERE appeal_id =" & delappid )
		
		'--insert into changes table
		sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		"SELECT 'טופס', 'שם: ' + P.Product_Name + ', איש קשר:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, 'מחיקה', getDate(), " & UserId & _
		" FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
		" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = " & delappid & ")	"
		con.executeQuery(sqlstr)		
		
		con.ExecuteQuery("DELETE FROM appeals WHERE appeal_id =" & delappid)
	rs_delete.moveNext
	Wend
	Set rs_delete = Nothing
	Response.Redirect "appealsDetails16735.asp?prodId=" & prodId &"&dt="& request("dt") & "&arch=" & arch	&"&dep="& request("dep")
end if
' ------ delete appeals --------

	arr_Status = session("arr_Status")
	if lang_id = "1" then
	    arr_StatusT = Array("","חדש","בטיפול","סגור")	
    else
		arr_StatusT = Array("","new","active","close")	
    end if 	
  
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 18 Order By word_id"				
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
	set rstitle = Nothing	%>
	<script LANGUAGE="javascript" type="text/javascript">
<!--
var oPopup = window.createPopup();

function CheckDelProd() {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המשוב"
     Else
		str_confirm = "Are you sure want to delete the form?"
     End If   
  %>
  return (confirm("<%=str_confirm%>"))    
}

function cball_onclick() {
	var strid = new String(document.form1.ids.value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.form1.elements['cb'+ arrid[i]].checked = document.form1.cb_all.checked ;
	
}
function SendInsuranceGroup()
{
	var fl = 0;
	document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
		}
		}}
			
				
 if (fl!=1)
 {
 <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים על מנת לשלוח הודעה גורפת"
			Else
				str_confirm = "Please select forms to send Insurance details !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		}
		else
		
{
  <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך לשלוח פרטי ביטוח עבור הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to send Insurance details for the selected forms?"
			End If   
		    %>
if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				document.getElementById("Insurance_flag").value = "1";
				h = parseInt(520);
				w = parseInt(520);
				window.open("../companies/SendInsuranceGroup.asp?appealsId=" + document.form1.trapp.value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
			//alert("yyy")
				//return true;
			}
	
		}
}


function checkGroupMail(){
		var fl = 0;
		
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך לשלוח הודעה גורפת עבור הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to send Email for the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
			document.getElementById("sms_flag").value = "1";
				h = parseInt(560);
				w = parseInt(600);
				window.open("sendMailGroup.asp?quest_id=<%=prodId%>&appealsId=" + document.form1.trapp.value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים על מנת לשלוח הודעה גורפת"
			Else
				str_confirm = "Please select forms to send EMail !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}	
	}


	
	function checkSite()
	{
			var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
			    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך להעביר את הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to transfer the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
			//	window.document.all("delete_flag").value = "0";
				document.getElementById("site_flag").value = "1";
			//	alert("dddd")
		//	alert(document.getElementById("site_flag").value)	
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים להעברה"
			Else
				str_confirm = "Please select forms to transfer !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}	
	return false;	
	}

function checktransf(){
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
							{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך להעביר את הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to transfer the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
			//	window.document.all("delete_flag").value = "0";
				document.getElementById("site_flag").value = "0";
				document.getElementById("delete_flag").value = "0";
				document.getElementById("change_status_flag").value = "0";
					return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים להעברה"
			Else
				str_confirm = "Please select forms to transfer !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}	
	return false;	
}

function checkDelete(){
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך למחוק את הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to delete the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				document.getElementById("delete_flag").value = "1";
				
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים למחיקה"
			Else
				str_confirm = "Please select forms to delete !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}
	
	function checkAddTasks()
	{
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך ליצור משימה גורפת עבור הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to add tasks for the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
			var reccount=checkDataClient()
			//alert("reccount="+reccount)
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 2);
			//	alert(document.form1.trapp.value)
				document.getElementById("add_tasks_flag").value = "1";
				h = parseInt(520);
				w = parseInt(920);
				alert(reccount)
				if (reccount!=0)
				{window.open("../tasks/Checktasks.asp?appealsId=" + document.form1.trapp.value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
				}
				else
				{
				window.open("../tasks/addtasks.asp?appealsId=" + document.form1.trapp.value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
}
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים על מנת ליצור משימה גורפת"
			Else
				str_confirm = "Please select forms to add to tasks !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}	
	}
	
	function checkDataClient()
	{
	var rec=0;
	document.form1.trapp.value=document.form1.trapp.value +'0'
//alert ("data="+document.form1.trapp.value)
	    $.ajax
             ({
                type: 'GET',
                url: 'CheckDataClient.asp?appIds='+document.form1.trapp.value ,
                async:   false,
              	success: function (result) {
					//alert("result="+result);
							 rec=result;
							
					}
             });
   //  alert ("rec="+rec);
          return rec;
       }
            
	
	function checkChangeStatus()
	{
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.form1.elements['cb'+ arrid[i]].checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך להעביר את הטפסים המסומנים לסטאטוס הנבחר"
			Else
				str_confirm = "Are you sure want to move the selected forms to the selected status?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				document.form1.action = "<%=urlSort%>";
				//window.alert(document.form1.action);
				document.getElementById("change_status_flag").value = "1";
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים "
			Else
				str_confirm = "Please select forms !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}

var prod;
function openPreview(prodId)
{
	prod = window.open("../products/check_form.asp?prodId="+prodId,"Product","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=700, height=400, left=10, top=20");
	if((prod.document != null) && (prod.document.body != null))
	{
		prod.document.title = "Product Preview";
		prod.document.body.style.margintop  = 20;
	}	
	return false;
}

//-->
	</script>
	<body>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table1">
		
	
			<tr>
				<td><table cellpadding="0" cellspacing="0" width="100%" ID="Table2">
						<tr>
								<td class="page_title" style="border-left:none" width="100%">מסך רישום חתום יומי
	</td>
						</tr>
<%sqlstrDep = "SELECT  departmentName  FROM Departments   where departmentId="&request("dep")
set rs_Dep= con.getRecordSet(sqlstrDep)
    Dep_Name =rs_Dep("departmentName")
rs_Dep.close
		set rs_Dep=Nothing	
%>
	<tr><td class="page_title" dir=rtl>&nbsp;&nbsp;&nbsp;<%=Request("dt")%>&nbsp;&nbsp;&nbsp;<%=Dep_Name%></td></tr>

					</table>
				</td>
			</tr>
			<tr>
				<td valign="top">
					<FORM action="appealsDetails16735.asp?prodId=<%=prodId%>&dt=<%=request("dt")%>&dep=<%=request("dep")%>&sort=<%=sort%>&arch=<%=arch%>" method=POST id="form1" name="form1" target="_self">
						<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" ID="Table3">
							<tr>
								<td width="100%" valign="top" align="center">
									<table width="100%" cellspacing="0" cellpadding="0" align="center" border="0" bgcolor="#ffffff" ID="Table4">
										<tr>
											<td width="100%" align="center" valign="top">
												<!-- start code -->
												<%if Request("Page")<>"" then
		Page=request("Page")
	else
		Page=1
	end if
	If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
		 PageSize = RowsInList
	Else	
    	 PageSize = 10
	End If
'response.Write sortby(sort) 
IF trim(PinpsearchDep)<>"" then
'if str_where_values="" then
' str_where_values= " AND Departure_Code Like '%" &  sFix(PinpsearchDep) & "%'"	
'else
str_where_values = str_where_values & " AND Departure_Code Like '%" &  sFix(PinpsearchDep) & "%'"	

'end if
'response.Write (str_where_values)
'response.end
end if
if trim(dt)<>"" then
str_where_values = str_where_values & "  AND DATEDIFF(dd, '"& dt &"', A.appeal_date)= 0 "
'response.Write  str_where_values
end if
if trim(request("dep"))<>"" then
str_where_values = str_where_values & "  AND A.Department_Id="& request("dep")
end if


    sqlstr = "exec dbo.get_appeals_paging @Page=" & Page & ", @RecsPerPage=" & PageSize & ", @company_Name ='" & sFix(orgname) & _
    "', @contact_name='" & sFix(clname) & "', @project_name='" & project_name & "', @appeal_Country='" & app_Country & "', @appeal_TourId='" & app_TourId  & "', @appeal_GuideId='" & app_GuideId& "', @appeal_DepId='" & app_DepId  & "', @appeal_status='" & app_status & _
    "', @OrgID='" & OrgID & "', @sort='" & sortby(sort) & "', @start_date='" & start_date & "', @end_date='" & end_date  & _
    "', @company_id='" & CompanyId & "', @contact_id='" & ContactId & "', @project_id='" & ProjectId &_
    "', @appeal_id='" & search_id & "', @product_id='" & prodId & "', @UserId='" & UserID & "', @archive='" & arch & _
    "', @user_name='" & sFix(usname) & "', @orderOwner='" & sFix(orderOwner) & "', @IsGroups='" & is_groups & _
    "', @str_where_values='" & sFix(str_where_values) & "'"	
 'Response.Write sqlStr
'    Response.End
	set app=con.GetRecordSet(sqlStr)	%>
												<!-- start search row -->
												<input type="hidden" name="trapp" value="" ID="trapp"> <input type="hidden" name="delete_flag" value="0" ID="delete_flag">
												<input type="hidden" name="site_flag" value="0" ID="site_flag"> <input type="hidden" name="add_tasks_flag" value="0" ID="add_tasks_flag">
												<input type="hidden" name="sms_flag" value="0" ID="sms_flag"> <input type="hidden" name="Insurance_flag" value="0" ID="Insurance_flag">
												<input type="hidden" name="change_status_flag" value="0" ID="change_status_flag">
												<table border="0" cellspacing="1" cellpadding="0" width="100%" ID="Table5">
													<tr style="background-color: #8A8A8A">
														<td dir="<%=dir_obj_var%>" align="left" colspan="3" style="padding-left: 5px;"><input type="submit" value="חפש" class="but_menu" style="width: 50px" ID="btnSubmit" NAME="btnSubmit">&nbsp;&nbsp;<input type="button" 
		value="הצג הכל" class="but_menu" style="width: 60px" onclick="document.location.href='appealsDetails16735.asp?prodId=<%=prodId%>&dt=<%=Request("dt")%>&dep=<%=request("dep")%>&sort=<%=sort%>&arch=<%=arch%>'" ID="btnShowAll" NAME="btnShowAll"></td>
														<td width="85"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80px;" value="<%=vFix(usname)%>" name="usname" ID="usname" maxlength="50"></td>
														<%If ADD_CLIENT <> "" Then 'טופס מקושר לפרטים קשרי לקוחות%>
														<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(clname)%>" name="clname" ID="clname" maxlength="50"></td>
														<%End If%>
														<%If ADD_CLIENT = "2" Then 'טופס מקושר לפרטים מלאים קשרי לקוחות%>
														<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width: 99%;" value="<%=vFix(orgname)%>" name="orgname" ID="orgname" maxlength="50"></td>
														<%End If%>
														<%If trim(prodId) = "16735" Then%>
														<td width="100" align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width: 99%;" value="<%=vFix(orderOwner)%>" name="orderOwner" ID="orderOwner" maxlength="50"></td>
														<td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>">
														
																</td>
														
														<%if Request.Cookies("bizpegasus")("AddInsurance")="1" then%>
														<td><!--ביטוח--></td>
														<%end if%>
														<%End If%>
														<%If count_fields >= 0 And IsArray(arr_fields) Then%>
														<%For ff=0 To count_fields%>
														<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><input id="inpsearch<%=arr_fields(0, ff)%>" name="inpsearch<%=arr_fields(0, ff)%>" value="<%=vFix(Eval("srch_val" & trim(arr_fields(0, ff))))%>" class="search" dir="<%=dir_obj_var%>" style="width: 100%" maxlength="50"></td>
														<%Next %>
														<%End If%>
														<%if trim(prodId) ="16504" then%>
														<td  align="center" dir="<%=dir_obj_var%>">
															<select ID="app_Country" dir="<%=dir_obj_var%>" 
		 NAME="app_Country">
																<option dir="<%=dir_obj_var%>" value=""><!--הכל--><%=arrTitles(19)%></option>
																<%sqlCountry="select Country_Name,Country_Id From dbo.Countries order by Country_Name"
		set recCountry=con.GetRecordSet(sqlCountry)
		do while not recCountry.EOF %>
																<option value="<%=recCountry("Country_Id")%>" <%If recCountry("Country_Id")= cInt(app_Country) Then%> selected <%End If%>><%=recCountry("Country_Name")%></option>
																<%recCountry.moveNext()
			loop
		set recCountry=nothing%>
															</select>
														</td>
														<%end if%>
														<%if trim(prodid)="17857" then%>
														<td nowrap align="center" dir="<%=dir_obj_var%>">&nbsp;</td>
														<%end if%>
														<%if trim(prodid)="17857" or trim(prodid)="17001"  or   trim(prodid)= "17057" then%>
														<td nowrap>&nbsp;</td>
														<td  align="center" dir="<%=dir_obj_var%>">
															<select name="app_GuideId" id="app_GuideId" class="sel" dir="rtl" style="width:100px;font-size:10pt;display:block;">
																<OPTION value="">בחר מדריך</OPTION>
																<%
sqlstrGuide = "SELECT Guide_Id, (Guide_FName + '  ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY Guide_LName, Guide_FName"
    

	'Response.Write sqlstr
	'Response.End
	set rs_Guide= conPegasus.Execute(sqlstrGuide)
   do while not rs_Guide.EOF 
		Guide_Id =rs_Guide("Guide_Id")
		Guide_Name =rs_Guide("Guide_Name")%>
																<OPTION value="<%=Guide_Id%>" <%If trim(app_GuideId) = trim(Guide_Id) Then%> selected <%End If%>><%=Guide_Name%>
																</OPTION>
																<%rs_Guide.moveNext
		loop
		rs_Guide.close
		set rs_Guide=Nothing	
%>
															</select></td>
														<td nowrap align="center" dir="<%=dir_obj_var%>">
															<select ID="app_TourId" dir="<%=dir_obj_var%>" class="sel"  NAME="app_TourId" style="width:200px">
																<option dir="<%=dir_obj_var%>" value=""><!--הכל--> בחר טיול</option>
																<%
sqlstrTour = "SELECT Tour_Id, (Category_Name + ' - ' + Tour_Name) as Tour_Name  FROM dbo.Tours T " & _
        " INNER JOIN dbo.Tours_Categories TC ON T.Category_Id = TC.Category_Id  " & _
        " INNER JOIN dbo.Tours_SubCategories TSC ON T.SubCategory_Id = TSC.SubCategory_Id " & _
        " WHERE (Tour_Vis = 1) AND (SubCategory_Vis = 1) AND (Category_Vis = 1) " & _
        " ORDER BY Category_Order, Tour_Name"
    

	'Response.Write sqlstr
	'Response.End
	set rs_tours= conPegasus.Execute(sqlstrTour)
   do while not rs_tours.EOF 
		Tour_Id =rs_tours("Tour_Id")
		Tour_name =rs_tours("Tour_Name")%>
																<OPTION value="<%=Tour_Id%>" <%If trim(app_TourId) = trim(Tour_Id) Then%> selected <%End If%>><%=Tour_name%>
																</OPTION>
																<%rs_tours.moveNext
		loop
		rs_tours.close
		set rs_tours=Nothing	
%>
															</select>
														</td>
														<%end if%>
														<%if  trim(prodid)= "16735" then%>
														<td align="center"><input id="inpsearchDep" name="inpsearchDep" value="<%=PinpsearchDep%>" class="search" dir="<%=dir_obj_var%>" style="width: 80%" maxlength="50"></td>
														<td  align="center" dir="<%=dir_obj_var%>">
															<select name="app_GuideId" id="Select1" class="sel" dir="rtl" style="width:100px;font-size:10pt;display:block;">
																<OPTION value="">בחר מדריך</OPTION>
																<%
sqlstrGuide = "SELECT Guide_Id, (Guide_FName + '  ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY Guide_LName, Guide_FName"
    

	'Response.Write sqlstr
	'Response.End
	set rs_Guide= conPegasus.Execute(sqlstrGuide)
   do while not rs_Guide.EOF 
		Guide_Id =rs_Guide("Guide_Id")
		Guide_Name =rs_Guide("Guide_Name")%>
																<OPTION value="<%=Guide_Id%>" <%If trim(app_GuideId) = trim(Guide_Id) Then%> selected <%End If%>><%=Guide_Name%>
																</OPTION>
																<%rs_Guide.moveNext
		loop
		rs_Guide.close
		set rs_Guide=Nothing	
%>
															</select></td>
														<td nowrap align="center" dir="<%=dir_obj_var%>">
															<select ID="Select2" dir="<%=dir_obj_var%>" class="sel"  NAME="app_TourId" style="width:200px">
																<option dir="<%=dir_obj_var%>" value=""><!--הכל--> בחר טיול</option>
																<%
sqlstrTour = "SELECT Tour_Id, (Category_Name + ' - ' + Tour_Name) as Tour_Name  FROM dbo.Tours T " & _
        " INNER JOIN dbo.Tours_Categories TC ON T.Category_Id = TC.Category_Id  " & _
        " INNER JOIN dbo.Tours_SubCategories TSC ON T.SubCategory_Id = TSC.SubCategory_Id " & _
        " WHERE (Tour_Vis = 1) AND (SubCategory_Vis = 1) AND (Category_Vis = 1) " & _
        " ORDER BY Category_Order, Tour_Name"
    

	'Response.Write sqlstr
	'Response.End
	set rs_tours= conPegasus.Execute(sqlstrTour)
   do while not rs_tours.EOF 
		Tour_Id =rs_tours("Tour_Id")
		Tour_name =rs_tours("Tour_Name")%>
																<OPTION value="<%=Tour_Id%>" <%If trim(app_TourId) = trim(Tour_Id) Then%> selected <%End If%>><%=Tour_name%>
																</OPTION>
																<%rs_tours.moveNext
		loop
		rs_tours.close
		set rs_tours=Nothing	
%>
															</select>
														</td>
														<%end if%>
														<td width="40" nowrap align="center"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:40px;" value="<%=vFix(search_id)%>" name="search_id" ID="Text4" maxlength="50"></td>
														<td colspan="2" nowrap align="center" dir="<%=dir_obj_var%>">
															<select ID="app_status" dir="<%=dir_obj_var%>" 
		onchange="document.location.href='appealsDetails16735.asp?prodId=<%=prodId%>&dep=<%=request("dep")%>&sort=<%=sort%>&arch=<%=arch%>&app_status=' + this.value;&app_Country=' + <%=app_Country%>" NAME="app_status">
																<option dir="<%=dir_obj_var%>" value="" ><!--הכל--><%=arrTitles(19)%></option>
																<%For i=1 To Ubound(arr_Status)%>
																<option value="<%=arr_Status(i, 0)%>" <%If trim(arr_Status(i, 0)) = trim(app_status) Then%> selected <%End If%> ><%=arr_Status(i, 1)%></option>
																<%Next%>
															</select></td>
													</tr>
													<!-- end search row -->
													<%	if not app.eof then %>
													<tr>
														<td width="55" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><!--משימות--><%=arrTitles(4)%></td>
														<td width="65" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center">משימה 
															אחרונה</td>
														<td width="55" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">&nbsp;<!--תאריך--><%=arrTitles(6)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<td width="85"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" title="מיון" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">&nbsp;<!--עובד--><%=arrTitles(7)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<%If ADD_CLIENT <> "" Then 'טופס מקושר לפרטים קשרי לקוחות%>
														<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=9 or sort=10 then%>_act<%end if%>"><a class="title_sort" title="מיון" HREF="<%=urlSort%>&sort=<%if sort=10 then%>9<%elseif sort=9 then%>10<%else%>10<%end if%>" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="9" then%>down<%elseif trim(sort)="10" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<%End If%>
														<%If ADD_CLIENT = "2" Then 'טופס מקושר לפרטים מלאים קשרי לקוחות%>
														<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=11 or sort=12 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=12 then%>11<%elseif sort=11 then%>12<%else%>12<%end if%>" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>down<%elseif trim(sort)="12" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<%End If%>
														<%If trim(prodId) = "16735" Then%>
														<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=7 or sort=8 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=7 then%>8<%elseif sort=8 then%>7<%else%>7<%end if%>" target="_self">&nbsp;למי 
																שייך הרישום&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>down<%elseif trim(sort)="8" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=15 or sort=16 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=15 then%>16<%elseif sort=16 then%>15<%else%>15<%end if%>" target="_self">&nbsp;מחלקה
																&nbsp;<img src="../../images/arrow_<%if trim(sort)="15" then%>down<%elseif trim(sort)="16" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
													
														<%if Request.Cookies("bizpegasus")("AddInsurance")="1" then%>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="title_sort<%if sort=13 or sort=14 then%>_act<%end if%>"><a class="title_sort" title="ביטוח" HREF="<%=urlSort%>&sort=<%if sort=13 then%>14<%elseif sort=14 then%>13<%else%>14<%end if%>" target="_self">&nbsp;ביטוח&nbsp;<img src="../../images/arrow_<%if trim(sort)="13" then%>down<%elseif trim(sort)="14" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a>
														</td>
														<%end if%>
														<%End If%>
														<%If count_fields >= 0 And IsArray(arr_fields)  Then
			For i=0 To count_fields	
				Field_Id = arr_fields(0, i)
				Field_Title = trim(arr_fields(1, i))
				Field_Type = trim(arr_fields(2, i))
				If Len(Field_Title) > 22 Then
					Field_Title_S = Left(Field_Title,20) & ".."
				Else
					Field_Title_S = Field_Title
				End If	%>
														<%if Field_Type <> "10" then  'question%>
														<td class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=middle title="<%=vFix(Field_Title)%>" width="<%=Fix(100/5)-5%>%" nowrap><%=Field_Title_S%></td>
														<%end if%>
														<%Next%>
														<%end if%>
														<%if trim(prodId) ="16504" then%>
														<td class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" >יעד נסיעה</td>
														<%end if%>
														<%if trim(prodid)="17857"   then%>
														<td nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort">ת. העברה לאתר</td>
														<%end if%>
														<%if trim(prodid)="16735" then%>
														<td  nowrap align="center" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>" align="center" valign=middle>יציאה</td>
														<%end if%>
														<%if trim(prodid)="17857" or trim(prodid)="17001"  or   trim(prodid)= "17057"   then%>
														<td align="center" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>" align="center" valign=middle>יציאה</td>
														<td width="100"  class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>" align="center">מדריך</td>
														<td  nowrap align="center" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>" align="center">טיול</td>
														<%end if%>
														<%if trim(prodid)="16735" then%>
														<td width="100"  class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>" align="center">מדריך</td>
														<td  nowrap align="center" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>" align="center">טיול</td>
														<%end if%>
														<td width="40" nowrap align="center" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<td width="48" nowrap  class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" >&nbsp;<!--'סט--><%=arrTitles(9)%>&nbsp;</td>
														<td width="20" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"></td>
														</tr>
													<%recCount = app("CountRecords")
	 aa = 0
	 do while not app.eof
		If aa Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If	
	
		appid=app("appeal_id")		
		COMPANY_NAME = app("COMPANY_NAME")
		If trim(COMPANY_NAME) = "" Or IsNull(COMPANY_NAME) Then
			If trim(lang_id) = "1" Then
			COMPANY_NAME = "אין"
			Else
			COMPANY_NAME = "No"
			End If
		End If	
		CONTACT_NAME = app("CONTACT_NAME")
		If trim(CONTACT_NAME) = "" Or IsNull(CONTACT_NAME) Then
		    If trim(lang_id) = "1" Then
			CONTACT_NAME = "אין"
			Else
			CONTACT_NAME = "No"
			End If
		End If	
		User_Name = app("User_Name")
		If trim(User_Name) = "" Or IsNull(User_Name) Then
			If trim(lang_id) = "1" Then
			User_Name = "אינטרנט"
			Else
			User_Name = "No"
			End If
		End If	
		ids = ids & appid 		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		contactID = trim(app("contact_id"))	
		mes_new = trim(app("mes_new"))
		mes_work = trim(app("mes_work"))
		mes_close = trim(app("mes_close"))
		appeal_status = trim(app("appeal_status"))	
		appeal_status_name = trim(app("appeal_status_name"))	
		appeal_status_color = trim(app("appeal_status_color"))	
		Order_Owner = trim(app("Order_Owner"))	 
		TaskUserName = trim(app("Task_USER_NAME"))	
		task_date = trim(app("task_date"))	
		Country_CRM = trim(app("Country_CRM"))	
		Guide_Name= trim(app("Guide_Name"))	
		Tour_Name=trim(app("Tour_Name"))
		Departure_Name=trim(app("Departure_Name"))
		departmentName=trim(app("departmentName"))
		ThankLetter_SendDate=trim(app("ThankLetter_SendDate"))
		
		  %>
													<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
														<td align="center" valign="top">
															<a href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" style="width:10pt;" class="task_status_num3" title="<%=arr_StatusT(3)%>">
																<%=mes_close%>
															</a><a href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" style="width:10pt;" class="task_status_num2" title="<%=arr_StatusT(2)%>">
																<%=mes_work%>
															</a><a href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" style="width:10pt;" class="task_status_num1" title="<%=arr_StatusT(1)%>">
																<%=mes_new%>
															</a>
														</td>
														<td align="right" dir="rtl" valign="top" style="padding: 0px 3px;"><%=TaskUserName%><br>
															<%=task_date%>
														</td>
														<td align="center" valign="top"><a class="link_categ" href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self"><%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%></a></td>
														<td align="<%=align_var%>" valign="top"><a class="link_categ" HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>"><%=User_NAME%></a></td>
														<%If ADD_CLIENT <> "" Then 'טופס מקושר לפרטים קשרי לקוחות%>
														<td align="<%=align_var%>" valign="top" <%if app("FlagProblemContact")>=3 then%>style="background-color:#FFFF00"<%end if%> ><a class="link_categ" HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>"><%=CONTACT_NAME%></a></td>
														<%End If%>
														<%If ADD_CLIENT = "2" Then 'טופס מקושר לפרטים מלאים קשרי לקוחות%>
														<td align="<%=align_var%>" valign="top"><a class="link_categ" HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>"><%=COMPANY_NAME%></a></td>
														<%End If%>
														<%If trim(prodId) = "16735" Then%>
														<td align="<%=align_var%>" valign="top"><a class="link_categ" HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>"><%=Order_Owner%></a></td>
														<td align="<%=align_var%>" valign="top"  dir=rtl><%=departmentName%></td>
														<%if Request.Cookies("bizpegasus")("AddInsurance")="1" then%>
														<%if app("Insurance_Status")=1 then%>
														<td nowrap align="center">
															<a onclick="javascript:openDetails('<%=app("Insurance_Date")%>','<%=app("Worker_Name")%>','<%=app("LeadID")%>')"  onmouseout ="javascript:CloseDiv()" title="פרטים נשלחו"  class="status_num2" >
																פרטים נשלחו</a>
														</td>
														<%else%>
														<td align="center" dir="<%=dir_var%>" >
															<%if false then ' SendInsuranceCount mod 50 =0 then%>
															<a title="שלח" class="status_numIns" href="javascript:sendInsuranceConfirm('<%=ContactId%>','<%=CompanyId%>','<%=appid%>')">
																שלח</a>
															<%else%>
															<a href="javascript:sendInsurance('<%=ContactId%>','<%=CompanyId%>','<%=appid%>')"  title="שלח" class="status_numIns">
																שלח</a>
															<%end if%>
															<%if false then %>
															<a href="javascript:sendInsurance('<%=ContactId%>','<%=CompanyId%>','<%=appid%>')" onclick="return window.confirm('האם לקוח אישר שפרטיו ישחלו לחברת הביטוח ליצירת קשר להצעת פוליסת ביטוח נסיעות')" title="שלח" class="status_numIns">
																שלח</a><%end if%>
															<%End If%>
															<%End If%>
															<%End If%>
															<!--#include file="../appeals/key_fields.asp"-->
															<%if trim(prodId) ="16504" then%>
														<td align="right" valign="top" dir="rtl"><%=Country_CRM%></td>
														<%end if%>
														<%if trim(prodid)="17857"   then%>
														<td nowrap align="center" dir="<%=dir_obj_var%>"  valign="top" style="border-right:solid 1px #ffffff">
															<%if ThankLetter_SendDate<>"" then%>
															<%=day(ThankLetter_SendDate)%>/<%=month(ThankLetter_SendDate)%>/<%=mid(year(ThankLetter_SendDate),3,2)%><%end if%>
															<%'=ThankLetter_SendDate%>
														</td>
														<%end if%>
														<%if  trim(prodid)="16735" then%>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff"><%=Departure_Name%></td>
														<%end if%>
														<%if trim(prodid)="17857"  or trim(prodid)="17001"  or   trim(prodid)="17057" then%>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff" nowrap><%=Departure_Name%></td>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff"><%=Guide_Name%></td>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff"><%=Tour_Name%></td>
														<%end if%>
														<%if  trim(prodid)="16735" then%>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff"><%=Guide_Name%></td>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff"><%=Tour_Name%></td>
														<%end if%>
														<td align="center" valign="top" style="border-right:solid 1px #ffffff"><a class="link_categ" HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>"><%=appid%></a></td>
														<td align="center" valign="top"><a class="status_num" style="background-color:<%=trim(appeal_status_color)%>" href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0"><%=appeal_status_name%></a></td>
														<td align="center" valign="top"></td>
													
													</tr>
													<%		app.movenext
		j=j+1 : aa = aa + 1
		
		loop  %>
												</table>
												<input type="hidden" name="ids" value="<%=ids%>" ID="ids"></td>
										</tr>
										<%	NumberOfPages = Fix((recCount / PageSize)+0.9)	%>
										<%	End If 	%>
										<% if NumberOfPages > 1 then%>
										<tr class="card">
											<td width="100%" align="center" nowrap class="card" dir="ltr">
												<table border="0" cellspacing="0" cellpadding="2" ID="Table6">
													<% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If %>
													<tr>
														<%if numOfRow <> 1 then%>
														<td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" name=22 title="לדפים הקודמים"><b><<</b></a></td>
														<%end if%>
														<td><font size="2" color="#060165">[</font></td>
														<%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
														<td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
														<%else%>
														<td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
														<%end if
	                  end if
	               next%>
														<td><font size="2" color="#060165">]</font></td>
														<%if NumberOfPages > cint(num * numOfRow) then%>
														<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" name=23 title="לדפים הבאים"><b>>></b></a></td>
														<%end if%>
													</tr>
												</table>
											</td>
										</tr>
										<%End If%>
										
    <%set app = nothing%>
									</table>
								</td>
							</tr> </table></form></td></tr> </table> </center></div>
			 
	</body>
</html>
<%set con=nothing%>
<div id="pos" style="position:absolute; display:none;background-color:#ffffff;border:solid 1px #22B14C;width:200px;height:100px;">test</div>
<div id="confirmBlock" style="position:absolute; display:none;background-color:#8A8A8A;color:#ffffff;border:solid 1px #ff0000;width:320px;height:200px;">
	confirm
</div>
