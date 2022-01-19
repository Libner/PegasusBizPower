<!-- הוספת חשבון משתמש באתר תדמיתי פגסוס-->
<% Server.ScriptTimeout=10000 %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%contactID = trim(Request("contactID"))%>
<%If isNumeric(trim(UserID)) = true And IsNumeric(trim(OrgID)) Then%>
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY class="body_admin">
<div id="div_save" bgcolor="#e8e8e8" style="position:absolute; left:0px; top:0px; width:100%; height:100%; " >  												
  <table bgcolor="#e8e8e8" height="100%" width="100%" cellspacing="2" cellpadding="2">  
     <tr><td bgcolor="#e8e8e8" align="center" >
         <table bgcolor="#ebebeb" border="0" height="100" width="400" cellspacing="1" cellpadding="1">
            <tr>  
              <td align="center" bgcolor="#d0d0d0">
              <font style="font-size:14px;color:#FF0000;"><b>מתבצעת שמירת נתונים</b></font>
              <br>
              <font style="font-size:14px;color:#000000;">... אנא המתן</font>
              </td>
            </tr>
         </table>
         </td>
     </tr>
  </table>
</div>
<%Function RandomPassword(PasswordLength)

    Dim lNumberOfLowerCases
    Dim lNumberOfUpperCases
    Dim lNumberOfNumbers
    Dim l, j
    
    ReturnedPassword = ""
    If PasswordLength < 3 Then        
        Exit Function
    End If

    'Get the number of digits for each type of characters
    Randomize
    lNumberOfLowerCases = 0 ' CInt((PasswordLength - 3) * Rnd) + 1
    'lNumberOfUpperCases = CInt((PasswordLength - lNumberOfLowerCases - 2) * Rnd) + 1
    lNumberOfUpperCases = 0
    lNumberOfNumbers = PasswordLength - lNumberOfLowerCases - lNumberOfUpperCases

       
    ReturnedPassword = ""
    For l = 1 To PasswordLength
        Randomize
        j = CInt(2 * Rnd + 1)
        Select Case j
        Case 1 'Lower Case
            If lNumberOfLowerCases > 0 Then
                ReturnedPassword = ReturnedPassword & Chr(CInt(25 * Rnd) + 97)
                lNumberOfLowerCases = lNumberOfLowerCases - 1
            Else
                l = l - 1 'Re-do the loop
            End If
        Case 2 'Upper Case
            If lNumberOfUpperCases > 0 Then
                ReturnedPassword = ReturnedPassword & Chr(CInt(25 * Rnd) + 65)
                lNumberOfUpperCases = lNumberOfUpperCases - 1
            Else
                l = l - 1 'Re-do the loop
            End If
        Case 3 'Number
            If lNumberOfNumbers > 0 Then
                ReturnedPassword = ReturnedPassword & CInt(9 * Rnd)
                lNumberOfNumbers = lNumberOfNumbers - 1
            Else
                l = l - 1 'Re-do the loop
            End If
        End Select
        For j = 1 To 100
            'Give the seed some time
        Next
    Next
    
     '   response.Write("ReturnedPassword="& ReturnedPassword)
      '  response.end
    RandomPassword = ReturnedPassword
End Function 

	ItemId = 0
	
   sql = "SELECT email FROM dbo.CONTACTS WHERE (contact_ID=" & contactID & ") AND (organization_id = " & OrgID & ")"
   Set listContact=con.GetRecordSet(sql)
   If not listContact.EOF Then
      email = trim(listContact("email"))
   End if   
   Set  listContact = Nothing      
	
	Password = RandomPassword(6) 
	LoginName = email
	
	sqlIns="SET NOCOUNT ON; INSERT INTO [dbo].[Contacts_History]  ([Contact_Id],[Password],[LoginName],[Insert_Date],[User_Id])" & _
	" VALUES (" & contactID & ", '" & Password & "', '" & LoginName & "', getDate(), " & UserID & "); SELECT @@IDENTITY AS ItemId"
    'Response.Write(sqlIns)
    'Response.End     
	Set rs_tmp = con.getRecordSet(sqlIns)
		ItemId = rs_tmp.Fields("ItemId").value
	Set rs_tmp = Nothing	  
	
  Set con = Nothing%>
<%End If%>
<script language="javascript">
  <!--
   document.location.href="<%=Application("SiteUrl")%>/members/AMemberAdd.aspx?ItemId=<%=ItemId%>";
  //-->
</script>
</BODY>
</HTML>