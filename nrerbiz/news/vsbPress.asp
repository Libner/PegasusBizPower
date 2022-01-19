<!--#include file="../../include/connect.asp"-->

<%catID=Request.QueryString("catID")
  newID=Request.QueryString("idsite")

if Request.QueryString("idSite")<>nil then
		s="update News set New_Site_Visible=1-New_Site_Visible where New_ID=" & newID
		con.execute (s)
		
		sqlstr = "Select New_Title, New_Site_Visible, New_Date From news WHERE New_ID=" & newID
		set rs_new = con.execute(sqlstr)
		if not rs_new.eof then
			New_Title = trim(rs_new(0))
			New_Site_Visible = trim(rs_new(1))
			New_Date = trim(rs_new(2))
		end if
		set rs_new = nothing
		
		xmlFilePath = "../../download/xml_news/bizpower_news.xml"
		'------ start deleting the new message from XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false			
			if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.documentElement.childNodes 
				for j=0 to objNodes.length-1
					set objTask = objNodes.item(j)
					node_mess_id = objTask.attributes.getNamedItem("ID").text
					'if trim(newID) = trim(node_mess_id) And trim(New_Site_Visible) = "False" then
					If trim(New_Site_Visible) = "True" Then
						objDOM.documentElement.removeChild(objTask)
						'exit for
					'else
						'set objTask = nothing
					end if
				next
				Set objNodes = nothing
				set objTask = nothing
				objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		' ------ end  deleting the new message from XML file ------
		
		If trim(New_Site_Visible) = "True" Then
		'----start adding to XML file ------
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false
		
		if objDOM.load(server.MapPath(xmlFilePath)) then
		Set objRoot = objDOM.documentElement
		'create new node
		  Set objChield = objDOM.createElement("NEW")
		'append new node
 		  set objNewField=objRoot.appendChild(objChield) 
 		'create new attribute USERNAME
		  Set objNewAttribute = objDOM.createAttribute("ID")
		  objNewAttribute.text = newID
		'append attribute to new node
 		  objNewField.attributes.setNamedItem(objNewAttribute) 
		  Set objNewAttribute=nothing
		'create new attribute DATE
		  Set objNewAttribute = objDOM.createAttribute("DATE")
		  objNewAttribute.text = New_Date		  
		'append attribute to new node
 		  objNewField.attributes.setNamedItem(objNewAttribute) 
 		 'create new attribute title
		  Set objNewAttribute = objDOM.createAttribute("TITLE")
		  objNewAttribute.text = New_Title		  
		'append attribute to new node
 		  objNewField.attributes.setNamedItem(objNewAttribute)  
 		  Set objNewAttribute=nothing
		  objDom.save server.MapPath(xmlFilePath)
				    
		set objNewField=nothing
		Set objChield = nothing
	end if ' objDOM.load	
	set objDOM = nothing
'----end adding to XML file ------
	End if	
end if 'if request("addmess") <> nil


if Request.QueryString("idhome")<>nil then
		s="update News set New_Home_Visible=1-New_Home_Visible where New_ID=" & Request.QueryString("idhome")
		con.execute (s)
end if%>

<meta http-equiv="refresh" content="0;url=admin.asp?catID=<%=catID%>">  
<%set rec=Nothing
con.close
set con=nothing%>
