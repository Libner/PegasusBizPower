<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->


<%



ContactId = trim(Request("ContactId"))
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
  'before P2990: urlWS="http://services.passportcard.co.il/PC_Site_DMZ_WebService/PC_Mobile_WebService.asmx/Mobile_LeadoMat_App_Handler_Vendor"
  urlWS="http://services.passportcard.co.il/PC_Site_DMZ_WebService/PC_Mobile_WebService.asmx/Mobile_LeadoMat_App_Handler" 
  postWS = ""
	''before P2990: postWS = postWS & "EventType=4"
	postWS = postWS & "EventType=2"
	postWS = postWS & "&LeadoMatNumber=70120" 'before 02/02/2020 : 46402
	''before P2990:postWS = postWS & "&ParentPersonalID=0"
	postWS = postWS & "&ParentPersonalID=3000"
	''before P2990: postWS = postWS & "&DistributorID=0"
	'before 12102020:
	'postWS = postWS & "&DistributorID=46943"
	'after 12102020:
	postWS = postWS & "&DistributorID=55310"
	postWS = postWS & "&PhoneNumber="& phoneSend '054 1234560;" '& trim(PhoneNumber)
	postWS = postWS & "&Email="& email 'TEST@TEST.COM" ' & trim(Email)
	postWS = postWS & "&FirstName="'" & trim(FirstName)
	postWS = postWS & "&LastName=" & trim(CONTACT_NAME)
	postWS = postWS & "&FlightDate="
	postWS = postWS & "&Comment="
	''before P2990: postWS = postWS & "&ObjTypeID=243"
	postWS = postWS & "&ObjTypeID=5"
	postWS = postWS & "&MobileDeviceType=4"
	''before P2990: postWS = postWS & "&otherVerification=" & user_name
	''before P2990: postWS = postWS & "&Objid=UT/VRCK5SuOb6bEb10F6WQ=="
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

 xml.save(Server.MapPath( "/download/xml_forms/testPasportCard.xml"))
 
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
''send mail to with result30.07.14---
	BodyText=""

	BodyText = BodyText &  "<html><head><meta charset=windows-1255></head>"
	BodyText = BodyText &  "<style>"
	BodyText = BodyText &  "td.content_cl"
	BodyText = BodyText &  "{"
	BodyText = BodyText &  vbTab & "FONT-WEIGHT: bold;"
	BodyText = BodyText &  vbTab & "FONT-SIZE: 10pt;"
	BodyText = BodyText &  vbTab & "COLOR: #022c63;"
	BodyText = BodyText &  vbTab & "BACKGROUND-COLOR: #e0e0f4;"
	BodyText = BodyText &  vbTab & "FONT-FAMILY: Arial;"
	BodyText = BodyText &  "}"
	BodyText = BodyText &  "td.title_cl"
	BodyText = BodyText &  "{"
	BodyText = BodyText &  vbTab & "FONT-WEIGHT: 600;"
	BodyText = BodyText &  vbTab & "FONT-SIZE: 11pt;"
	BodyText = BodyText &  vbTab & "COLOR: #6f6da6;"
	BodyText = BodyText &  vbTab & "FONT-FAMILY: Arial;"
	BodyText = BodyText &  "}"
	BodyText = BodyText &  "td.content_b"
	BodyText = BodyText &  "{"
	BodyText = BodyText &  vbTab & "FONT-WEIGHT: 500;"
	BodyText = BodyText &  vbTab & "FONT-SIZE: 10pt;"
	BodyText = BodyText &  vbTab & "BACKGROUND-COLOR: #f7f7f7;"
	BodyText = BodyText &  vbTab & "COLOR: #022c63;"
	BodyText = BodyText &  vbTab & "FONT-FAMILY: Arial;"
	BodyText = BodyText &  "}"
	BodyText = BodyText &  "</style>"
	BodyText = BodyText &  "<body>"


	'*************************************************************************************

	BodyText = BodyText & "<table width=60% cellpadding=3 cellspacing=2 align='center' border='0'>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td colspan=2 class='title_cl' dir=rtl align='right' >" & "פנייה לחברת פספורט כארד BizPower" & "</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right' >" & trim(CONTACT_NAME) & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right' nowrap>:שם</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & phoneSend & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right'>:טלפון</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b'  align='right'><a href=mailto:" & email & " target=_blank>" & email & "</a></td>"
	BodyText = BodyText & "<td class='content_cl' align='right'>:אימייל</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & user_name & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align=right>:מוקדן</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & LeadNumber & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right'>:מספר הפניה</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "</table>"
	BodyText = BodyText & "</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "</table>"

	'*************************************************************************************

	BodyText = BodyText & "</body>"
	BodyText = BodyText & "</html>"


Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
			Msg.From = "support@pegasusisrael.co.il"
			'Msg.To = "faina@cyberserve.co.il"				
			Msg.To = "Pegasus342@pegasusisrael.co.il"				
			Msg.Bcc = "yoelziv5@gmail.com"
				Msg.MimeFormatted = true
				Msg.Subject = "פנייה לחברת פספורט כארד"
					Msg.HTMLBody = BodyText
			
				Msg.Send()						
			Set Msg = Nothing

'end send mail to with result30.07.14---
  
  

end if
  'test
  
Dim Msgtest
			Set Msgtest = Server.CreateObject("CDO.Message")
				Msgtest.BodyPart.Charset = "windows-1255"
			Msgtest.From = "mila@cyberserve.co.il"	
			Msgtest.To = "mila@cyberserve.co.il"
				Msgtest.MimeFormatted = true
				Msgtest.Subject = "פנייה לחברת פספורט כארד"
					Msgtest.HTMLBody =  "urlWS=" & urlWS & "<br>postWS=" &  postWS & "<br><br>" & strRetval
			
				Msgtest.Send()						
			Set Msgtest = Nothing
			
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
		//opener.location.href = opener.location.href
		opener.window.location.reload(true);
		self.close();
	//-->
	</SCRIPT>
