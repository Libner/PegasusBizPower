<!--#include file="../../connect_withoutCodePage.asp"-->
<%Server.ScriptTimeout=20000
  Response.Buffer=True
  lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
		
  If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
  End If
   	
  If trim(lang_id) = "1" Then
     Response.CharSet = "windows-1255"      
  Else
     Response.CharSet = "Windows-1252"    
  End If

  Response.ContentType = "application/vnd.ms-excel"  
  'Response.ContentType = "application/save"
%>
<!--#include file="../../reverse.asp"-->
<%  
   
   OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))    
   filestring="export_" & trim(OrgID) & ".csv"   
   Response.AddHeader "Content-Disposition", "filename="&filestring&";"
   
   'set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   'filestring="../../../download/reports/"
   'Response.Write server.mappath(filestring)
   'Response.End
   'fs.DeleteFile server.mappath(filestring) & "/*.*" ,False   
  
   'filestring="../../../download/reports/export_" & trim(OrgID) & ".csv"
   'Response.Write server.mappath(filestring)
   'Response.End
   'set XLSfile=fs.CreateTextFile(server.mappath(filestring),True,False) ' Create text file as text file 
  
   strLine = "Name" & "," & "Job Title" & "," & "Phone" & "," & "Fax" & "," & "Mobile phone" & "," & "Email" & "," & "Company" & "," & "Business phone" & "," & "Business fax" & "," & "Business email" & "," & "Web Page" & "," & "Business address" & "," & "Business city"
   Response.Write strLine & vbCrlf	

    SQL = "SELECT CONTACTS.CONTACT_NAME, CONTACTS.phone, CONTACTS.fax, CONTACTS.cellular, "&_
    " CONTACTS.email, CONTACTS.Messanger_Name, COMPANIES.City_Name, COMPANIES.COMPANY_NAME, COMPANIES.address, "&_
    " COMPANIES.phone1, COMPANIES.fax1, COMPANIES.url, COMPANIES.Email FROM CONTACTS "&_
    " INNER JOIN COMPANIES ON CONTACTS.COMPANY_ID = COMPANIES.COMPANY_ID "&_   
    " WHERE companies.ORGANIZATION_ID = " & OrgID & " Order BY CONTACTS.CONTACT_NAME"
    'Response.Write sql
    'Response.End
    set listContact=con.GetRecordSet(SQL)
    If not listContact.EOF Then   
	  contArray = listContact.GetRows()	  	  
	  recCount =  listContact.RecordCount 		
	  listContact.close
      set listContact=Nothing
      count = 0
      do while count < recCount
		name = trim(contArray(0,count))
		If Len(trim(contArray(1,count))) > 3 Then
			phone = trim(contArray(1,count))
		Else
			phone = ""	
		End If
		If Len(trim(contArray(2,count))) > 3 Then
			fax = trim(contArray(2,count))
		Else
			fax = ""	
		End If
		If Len(trim(contArray(3,count))) > 3 Then
			cellurar = trim(contArray(3,count)) 
		Else
			cellurar = ""	
		End If     
		email = contArray(4,count)            
		duty = trim(contArray(5,count))
		City_Name = trim(contArray(6,count))
		COMPANY_NAME = trim(contArray(7,count))
		address = trim(contArray(8,count))
		If Len(trim(contArray(9,count))) > 3 Then
			company_phone = trim(contArray(9,count))
		Else
			company_phone = ""	
		End If
		If Len(trim(contArray(10,count))) > 3 Then
			company_fax = trim(contArray(10,count))
		Else
			company_fax = ""	
		End If	
		url = trim(contArray(11,count))
		company_email = trim(contArray(12,count))

		strLine = breaks(name) & "," & breaks(duty) & "," & breaks(phone) & "," & breaks(fax) & "," & breaks(cellurar) & "," & breaks(email) & ","  &_		
		breaks(COMPANY_NAME) & "," & breaks(company_phone) & "," & breaks(company_fax) & "," & breaks(company_email) & "," & breaks(url) & "," & breaks(address) & "," & breaks(City_Name)
	    
		Response.Write strLine & vbCrlf
	      
		count=count+1
      loop
    End If   	
	
	
	set con=Nothing		
	'Response.Redirect filestring
	Response.End

 %>
