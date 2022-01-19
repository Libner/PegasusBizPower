<%Server.ScriptTimeout=20000
  Response.Buffer=True
  Response.CharSet = "windows-1255"
%>
<!--#include file="../../connect.asp"-->
<%  
   OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))    
   set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   filestring="../../../download/reports/"
   'Response.Write server.mappath(filestring)
   'Response.End
   fs.DeleteFile server.mappath(filestring) & "/*.*" ,False   
  
   filestring="../../../download/reports/export_" & trim(OrgID) & ".txt"
   'Response.Write server.mappath(filestring)
   'Response.End
   set XLSfile=fs.CreateTextFile(server.mappath(filestring)) ' Create text file as text file 
  
   strLine = "שם" & vbTab & "תפקיד" & vbTab & "טלפון אישי" & vbTab & "פקס אישי" & vbTab & "נייד" & vbTab & "דואר אלקטרוני אישי" & vbTab & "חברה" & vbTab & "טלפון בעבודה" & vbTab & "פקס בעבודה" & vbTab & "דואר אלקטרוני" & vbTab & "אתר" & vbTab & "כתובת" & vbTab & "עיר"
   XLSfile.writeline strLine		

    SQL = "SELECT CONTACTS.CONTACT_NAME, CONTACTS.phone, CONTACTS.fax, CONTACTS.cellular, "&_
    " CONTACTS.email, messangers.item_name, COMPANIES.City_Name, COMPANIES.COMPANY_NAME, COMPANIES.address, "&_
    " COMPANIES.phone1, COMPANIES.fax1, COMPANIES.url, COMPANIES.Email FROM CONTACTS "&_
    " INNER JOIN COMPANIES ON CONTACTS.COMPANY_ID = COMPANIES.COMPANY_ID "&_
    " LEFT OUTER JOIN messangers ON CONTACTS.messanger_id = messangers.item_ID "&_  
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

		strLine = name & vbTab & duty & vbTab & phone & vbTab & fax & vbTab & cellurar & vbTab & email & vbTab  &_		
		COMPANY_NAME & vbTab & company_phone & vbTab & company_fax & vbTab & company_email & vbTab & url & vbTab & address & vbTab & City_Name
	    
		XLSfile.writeline strLine      
	      
		count=count+1
      loop
    End If 
    	
	XLSfile.close
	set con=Nothing
	 
	Response.Redirect(filestring) 'Open the text file in the browser
 %>
