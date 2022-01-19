<!-- יצירת טופס רישום חתום באתר עבור איש הקשר -->
<% Server.ScriptTimeout=10000 %>
<% Response.Buffer = False %>
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
<%ItemId = 0 
	
	sqlIns="SET NOCOUNT ON; INSERT INTO [dbo].[Contacts_Forms]  ([Contact_Id],[Insert_Date],[User_Id])" & _
	" VALUES (" & contactID & ", getDate(), " & UserID & "); SELECT @@IDENTITY AS ItemId"
    'Response.Write(sqlIns)
    'Response.End     
	Set rs_tmp = con.getRecordSet(sqlIns)
		ItemId = rs_tmp.Fields("ItemId").value
	Set rs_tmp = Nothing	  
	
  Set con = Nothing%>
<%End If%>
<script language="javascript">
  <!--
   document.location.href="<%=Application("SiteUrl")%>/tours/form.aspx?contactID=<%=contactID%>&UserID=<%=UserID%>";
  //-->
</script>
</BODY>
</HTML>