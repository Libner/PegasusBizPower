<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->


<%ContactId = trim(Request("ContactId"))
  companyID = trim(Request("companyID"))
  questid = trim(Request("quest_id"))
  appid= trim(Request("appid"))
    set listContact=con.GetRecordSet("EXEC dbo.contacts_contact_details @ContactId=" & ContactId & ", @OrgID=" & OrgID)
   if not listContact.EOF then 
      ContactId = cLng(listContact("contact_ID"))
       CONTACT_NAME = trim(listContact("CONTACT_NAME"))
  	'  first_name_E = trim(listContact("first_name_E"))
	'  last_name_E = trim(listContact("last_name_E"))
      email = trim(listContact("email"))
      phone = trim(listContact("phone"))
      cellular = trim(listContact("cellular"))
 end if
 if cellular<>"" then
 phoneSend=left(cellular,3) &" " & Mid(cellular,4,len(cellular)) &";"
 end if
 if phoneSend="" and len(phone)>3 then
 phoneSend=left(phone,3) &" " & Mid(phone,4,len(phone)) &";"
  end if
 
 
'  response.Write phoneSend &":"&  user_name &":"& CONTACT_NAME &":"& first_name_E&":"& cellular
'  response.end
 
 ' response.Write ContactId &":"&  companyID &":"& questid &":"& appid&":"& UserID
 ' response.end
  if 1=1 then
  	sqlstr = "UPDATE APPEALS SET Insurance_Status=1,  Insurance_Date = getDate(), " & _
		" worker_id = " & UserID & " WHERE APPEAL_Id = " & appid
		'Response.Write(sqlstr)
		'Response.End
		'con.ExecuteQuery(sqlstr)
	
  
  end if
  

 ' FirstName="cyber"
 ' LastName="serve"
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

response.Write "פרטים נשלחו בהצלחה"
	'Set oXMLHTTP = CreateObject("msxml2.ServerXMLHTTP")     
	'oXMLHTTP.setRequestHeader "Content-Type","text/xml"

	'oXMLHTTP.open "POST", urlWS, false      
'	oXMLHTTP.send postWS  
	'dim wsResponse
'	wsResponse = oXMLHTTP.responseText    
'response.Write "strRetval ="& strRetval
  %>
<SCRIPT LANGUAGE=javascript>
	<!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	//-->
	</SCRIPT>
