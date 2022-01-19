<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<script>
function Send(postWS)
{
var params = postWS;
if(window.XMLHttpRequest && !(window.ActiveXObject))
	 var xmlHTTP = new XMLHttpRequest();
    else
		if(window.ActiveXObject)
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	if (!xmlHTTP)
		var xmlHTTP = new ActiveXObject("Msxml2.XMLHTTP");
	if (xmlHTTP)
	{
		xmlHTTP.open("POST", 'http://services.passportcard.co.il/PC_Site_DMZ_WebService/PC_Mobile_WebService.asmx/Mobile_LeadoMat_App_Handler_Vendor', false);
		xmlHTTP.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
	xmlHTTP.setRequestHeader("Content-length", params.length);
 	xmlHTTP.send(postWS);	 
		result = new String(xmlHTTP.responseText);
		alert (result)
		}
}

</script>



<%ContactId = trim(Request("ContactId"))
  companyID = trim(Request("companyID"))
  questid = trim(Request("quest_id"))
  appid= trim(Request("appid"))
  
 ' response.Write ContactId &":"&  companyID &":"& questid &":"& appid&":"& UserID
 ' response.end
  if 1=1 then
  	sqlstr = "UPDATE APPEALS SET Insurance_Status=1,  Insurance_Date = getDate(), " & _
		" worker_id = " & UserID & " WHERE APPEAL_Id = " & appid
		'Response.Write(sqlstr)
		'Response.End
		'con.ExecuteQuery(sqlstr)
	
  
  end if
  

  FirstName="cyber"
  LastName="serve"
  urlWS="http://services.passportcard.co.il/PC_Site_DMZ_WebService/PC_Mobile_WebService.asmx/Mobile_LeadoMat_App_Handler_Vendor"
  postWS = ""
	postWS = postWS & "EventType=4"
	postWS = postWS & "&LeadoMatNumber=46402"
	postWS = postWS & "&ParentPersonalID=0"
	postWS = postWS & "&DistributorID=0"
	postWS = postWS & "&PhoneNumber=054 1234560;" '& trim(PhoneNumber)
	postWS = postWS & "&Email=TEST@TEST.COM" ' & trim(Email)
	postWS = postWS & "&FirstName=" & trim(FirstName)
	postWS = postWS & "&LastName=" & trim(LastName)
	postWS = postWS & "&FlightDate=null"
	postWS = postWS & "&Comment=nn"
	postWS = postWS & "&ObjTypeID=243"
	postWS = postWS & "&MobileDeviceType=4"
	postWS = postWS & "&otherVerification=נציג"
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

	'Set oXMLHTTP = CreateObject("msxml2.ServerXMLHTTP")     
	'oXMLHTTP.setRequestHeader "Content-Type","text/xml"

	'oXMLHTTP.open "POST", urlWS, false      
'	oXMLHTTP.send postWS  
	'dim wsResponse
'	wsResponse = oXMLHTTP.responseText    
response.Write "strRetval ="& strRetval
  %>
 <%if false then%> <a href="javascript:void(0)" onclick="javascript:Send('<%=postWS%>')">send</a><%end if%>

