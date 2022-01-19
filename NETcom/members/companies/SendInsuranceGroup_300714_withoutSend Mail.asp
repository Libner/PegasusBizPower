<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%Response.CharSet = "windows-1255"
	orgName="pegasus"
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	UserID=trim(trim(Request.Cookies("bizpegasus")("UserID")))
	CurrUserName=trim(trim(Request.Cookies("bizpegasus")("UserName")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))  
'send  sms
'	If Request.Form("add") <> nil Then	
	appealsId= trim(Request("appealsId"))  
'end if
'response.write appealsId
'response.end

sqlstr="select appeal_id,CONTACT_NAME,phone,cellular,email from APPEALS A left join dbo.CONTACTS ON A.CONTACT_ID = CONTACTS.CONTACT_ID  WHERE appeal_id IN (" & appealsId & ")"
set rs_Contacts = con.getRecordSet(sqlstr)	
do while not rs_Contacts.eof
      CONTACT_NAME = trim(rs_Contacts("CONTACT_NAME"))
      email = trim(rs_Contacts("email"))
      phone = trim(rs_Contacts("phone"))
      cellular = trim(rs_Contacts("cellular"))
      appid=rs_Contacts("appeal_id")
		   if cellular<>"" then
				phoneSend=left(cellular,3) &" " & Mid(cellular,4,len(cellular)) &";"
		   end if
	       if phoneSend="" and len(phone)>3 then
				phoneSend=left(phone,3) &" " & Mid(phone,4,len(phone)) &";"
		   end if
  '----------------------
    urlWS="http://services.passportcard.co.il/PC_Site_DMZ_WebService/PC_Mobile_WebService.asmx/Mobile_LeadoMat_App_Handler_Vendor"
  postWS = ""
	postWS = postWS & "EventType=4"
	postWS = postWS & "&LeadoMatNumber=46402"
	postWS = postWS & "&ParentPersonalID=0"
	postWS = postWS & "&DistributorID=0"
	postWS = postWS & "&PhoneNumber="& phoneSend '054 1234560;" '& trim(PhoneNumber)
	postWS = postWS & "&Email="& email 'TEST@TEST.COM" ' & trim(Email)
	postWS = postWS & "&FirstName="'" & trim(FirstName)
	postWS = postWS & "&LastName=" & trim(CONTACT_NAME)
	postWS = postWS & "&FlightDate="
	postWS = postWS & "&Comment="
	postWS = postWS & "&ObjTypeID=243"
	postWS = postWS & "&MobileDeviceType=4"
	postWS = postWS & "&otherVerification=" & user_name
	postWS = postWS & "&Objid=UT/VRCK5SuOb6bEb10F6WQ=="
	postWS = postWS & "&DateStart=" 'null"
	postWS = postWS & "&DateEnd=" 'null"
	'response.Write postWS
'	  postWS = ""
	'response.end
		'	on error resume next

		
	Set oxmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
	oxmlHTTP.open "POST", urlWS,  false
	oxmlHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
'	oxmlHTTP.setRequestHeader "Content-length", params.length;
 
oxmlHTTP.send postWS
'
strStatus = oxmlHTTP.Status
strRetval = oxmlHTTP.responseText
set oxmlHTTP = nothing
  Set xml = Server.CreateObject("MSXML2.DOMDocument.3.0") 
    xml.LoadXml(strRetval)

 
 
    For Each node in xml.selectSingleNode("/LeadomatMobile/LeadNumber").childNodes
        sResult = sResult & node.xml
    Next

LeadNumber= Server.HTMLEncode(sResult) 
if strStatus=200 then
	sqlstr = "UPDATE APPEALS SET Insurance_Status=1,  Insurance_Date = getDate(), " & _
		" worker_id = " & UserID & ",LeadID="& LeadNumber & " WHERE APPEAL_Id = " & appid
		'Response.Write(sqlstr)
		'Response.End
		con.ExecuteQuery(sqlstr)
	sqlstr = "UPDATE Users SET InsuranceSend_Counter=InsuranceSend_Counter+1 where User_Id="& UserID
	con.ExecuteQuery(sqlstr)
	
		
		end if
  '----------------------
  
  sResult=""
  LeadNumber=""
 
rs_Contacts.MoveNext
	loop
	
%>
<SCRIPT LANGUAGE=javascript>
	<!--	
	opener.focus();
		opener.window.location.reload(true);
		self.close();
	-->
	</SCRIPT>
