<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%  
	If Request.QueryString("Competitor_Id") <> nil Then
		if cint(trim(Request.QueryString("Competitor_Id")))>0 then
			Competitor_Id = trim(Request.QueryString("Competitor_Id"))
		end if
	end if
	
    If Request.QueryString("add") <> nil Then
    
		set upl = Server.CreateObject("SoftArtisans.FileUp") 
		upl.path=server.mappath("../../../download/competitors")
		Competitor_Name=upl.Form("Competitor_Name")
		Competitor_Phone=upl.Form("Competitor_Phone")
		if upl.Form("Competitor_Rank")<>"" then
			Competitor_Rank=upl.Form("Competitor_Rank")
		else
			Competitor_Rank="NULL"
		end if
		Competitor_SearchTerms=upl.Form("Competitor_SearchTerms")
		
		
		If trim(Competitor_Id) = "" Then ' add Competitor
			sqlstr = "Insert into Compare_Competitors (Competitor_Name, Competitor_Phone, Competitor_Rank, Competitor_SearchTerms) values ("
			sqlstr=sqlstr & "'"  & sFix(Competitor_Name) & "','"  & sFix(Competitor_Phone) & "',"  & Competitor_Rank & ",'"  & sFix(Competitor_SearchTerms) & "')"
		   con.executeQuery(sqlstr)
		   sql="SELECT top 1 Competitor_Id  from Compare_Competitors  order by Competitor_Id desc"
			set rs_tmp = con.getRecordSet(sql)
			if not rs_tmp.eof then
					Competitor_Id = rs_tmp("Competitor_Id")
			end if
			set rs_tmp = Nothing
			 %>
<% Else ' update Competitor
			sqlstr="Update Compare_Competitors set Competitor_Name = '" & sFix(Competitor_Name) & "'"
			sqlstr=sqlstr & ",Competitor_Phone = '" & sFix(Competitor_Phone) & "'"
			sqlstr=sqlstr & ",Competitor_Rank = " & Competitor_Rank 
			sqlstr=sqlstr & ",Competitor_SearchTerms = '" & sFix(Competitor_SearchTerms) & "'"
			sqlstr=sqlstr & " Where Competitor_Id = " & Competitor_Id
			con.executeQuery(sqlstr)
				 %>
<%End If%>
<%if upl.UserFileName<>"" then		
			upl.Form("UploadFile1").SaveAs upl.path & "/" & upl.UserFileName
			sqlstr="Update Compare_Competitors set Competitor_Logo = '" & sFix(upl.UserFileName) & "'"
			sqlstr=sqlstr & " Where Competitor_Id = " & Competitor_Id
			con.executeQuery(sqlstr)
end if%>
<%end if

%>
<html>
<head>

<SCRIPT LANGUAGE="javascript">
						<!--
							opener.focus();
							opener.window.location.reload();
							self.close();
						//-->
</SCRIPT>
</head>
</html>