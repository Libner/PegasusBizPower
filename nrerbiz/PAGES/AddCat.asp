<!--#include file="../../include/connect.asp"-->
<!--#include file="../../include/reverse.asp"-->

<script language="JavaScript">
<!--                                    
function CheckFields() {
	  if (document.f1.name.value==''){
		alert('יש למלא שדה טקסט');
		return false;
		}
	  else
		document.f1.submit();
		return false;
}
<!--End-->
//-->
</script>

<html>

<head>
<title>CyberServe ניהול תבניות תוכן</title>
<meta charset="windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>
<body  class="body_admin">
<div align="right">
<%
'on error resume next
    catId=Request("catId")
'isRightMenu=Request.querystring("isright")
  maincat=Request("maincat")
  siteid=session("siteid")

 innerparent=request("innerparent")
subcat=request("subcat") 
set m=con.execute("SELECT * FROM Main_Categories WHERE Main_Category_ID="&maincat&" " )

  if catId<>nil then
	sqlSt="SELECT * FROM Publication_Categories WHERE Publication_Category_ID="&catId&" "
	set pr=con.Execute(sqlSt)
		perName=pr("Publication_Category_Name")
		catURL=pr("Category_URL")
		mapcat=pr("Map_Category_Id")
  end if%>
<form name="f1" action="admin.asp?innerparent=<%=innerparent%>&subcat=<%=subcat%>&catId=<%=catId%>&maincat=<%=maincat%>" method="post">
<table class="table_admin_2" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
 <tr>
    <td class="a_title_big" width="100%" dir="rtl"><%if catId=nil then%>הוספת תת קטגוריה<%else%>עדכון תת קטגוריה "<%=m("Main_Category_Name")%>"<%end if%>
    </td>
  </tr>
<tr>
  <td class="td_admin_4" height="10">
  </td>
</tr>
<tr >
   <td class="td_admin_4" align="left"  valign="middle" nowrap>
		<table class="table_admin_2" width="100%" border="0"  cellpadding="0" cellspacing="0">
			<tr>
				<td align="left" >
					<a  class="button_admin_e" href="" onClick="return CheckFields()" >שלח</a>
				</td>
				<td align="left" valign="middle" nowrap>
					<a  class="button_admin_1" href="../choose.asp?siteid=<%=siteid%>">
						חזרה לניהול קטגוריות ראשיות
					</a>
				</td>
				<td align="left" valign="middle" nowrap>
					<a  class="button_admin_1" href="admin.asp?siteid=<%=siteid%>&maincat=<%=maincat%>">
					חזרה לניהול תת קטגוריות
					</a>
				</td>
				
				<td width="100%">&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
  <tr>
       <td class="td_admin_4" align="center" valign="middle" nowrap>&nbsp;
       </td>
  </tr>
</table>
<table class="table_admin_2" width="100%" border="0"  cellpadding="2" cellspacing="0">

      <tr>
         <td align="right" width="70%"  class =  "td_admin_4_right_align"><font face="Arial (Hebrew)"><textarea dir="rtl" name="name" rows="1" cols="45"><%=perName%></textarea></font></td>
         <td align="left" width="30%"  class =  "td_admin_4" nowrap>שם תת קטגוריה</td>
      </tr>
<tr>
<td align=right   class =  "td_admin_4_right_align"><input type=text  name="catURL" value="<%=catURL%>" size=70></td>
<td align="left"   class =  "td_admin_4">כתובת של אתר אחר URL</td>
</tr>
<%strsql="SELECT isSiteMap FROM Main_Categories WHERE Main_Category_Id=" & maincat
	set temp=con.execute(strsql)
	if temp("isSiteMap") then%>
<tr>
<td align=right   class =  "td_admin_4_right_align">
	<select name="mapid" size="1" dir="rtl">
	<%strsql="SELECT Map_Category_Id, Map_Category_Name FROM SiteMap_Categories WHERE Main_Category_Id=" & maincat
			set map=con.execute(strsql)
	%>
		<option value="" <%if mapcat=nil then%>selected<%end if%>>----בחר מפה התמצאות----</option>
		<%do while not map.eof%>
			<option value="<%=map("Map_Category_Id")%>" <%if map("Map_Category_Id")=mapcat then%>selected<%end if%>><%=map("Map_Category_Name")%></option>
			<%map.movenext
			loop
		%>
	</select>
</td>
<td align="left"   class =  "td_admin_4">מפות התמצאות</td>
</tr>
	
	<%end if%>
<tr>
         <td align="right" colspan=2 class =  "td_admin_4_right_align">&nbsp;</td>
</tr>
</table>
</form>

</div>
</body>

</html>
