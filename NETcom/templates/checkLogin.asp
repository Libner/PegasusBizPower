<!--#INCLUDE FILE="members/companies/_ScriptLibrary/RS.ASP"--> 
<!--#include file="connect.asp"-->
<!--#include file="reverse.asp"-->
<% RSDispatch %> 

<script language="JavaScript" runat="server"> 

var public_description = new MainMethod(); 
function MainMethod(){ 
   this.ifFoundUserName = Function('strString','return foundUserName (strString)'); 
   this.ifFoundEmail = Function('strString','return foundEmail(strString)'); 
} 
</script> 


<script language="VBScript" runat="server"> 

function foundUserName(strString) 
		'//ckeck nickName that there is no member used this nickname 
				sqlStr = "SELECT USER_ID FROM USERS WHERE LOGINNAME ='"& sFix(strString) & "'"
				''Response.Write sqlStr
				set rs_is_found_UserName = con.GetRecordSet(sqlStr)
				if  not rs_is_found_UserName.eof then
						foundUserName = "True"
				else
						foundUserName = "False"							
				end if
				
				set rs_is_found_UserName = nothing
		'//<<<<<< end if delete the picture uploades						
end function 

function foundEmail(strString) 
		'//ckeck nickName that there is no member used this nickname 
				sqlStr = "SELECT USER_ID from USERS where EMAIL='" & sFix(strString) & "'"
				''Response.Write sqlStr
				set rs_is_found_Email = con.GetRecordSet(sqlStr)
				if  not rs_is_found_Email.eof then
						foundEmail = "True"
				else
						foundEmail = "False"													
				end if
				set rs_is_found_Email = nothing
		'//<<<<<< end if delete the picture uploades						
end function 
</script> 
<%set con = nothing%>    