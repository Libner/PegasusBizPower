<%facultyId=Request.QueryString("catID")%>
<!--#include file="../../include/connect.asp"-->
<!--#include file="../../include/reverse.asp"-->
<!--#INCLUDE FILE="../checkAuWorker.asp"-->
<html>
<head>
<title>Administration Site</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckDel(str) {
  return (confirm("? האם ברצונך למחוק את האירוע"))    
}
<!--End-->
</script>  
<%'
  catID=Request.QueryString("catID")
 if Request.QueryString("delId")<>nil then
	xmlFilePath = "../../download/xml_news/bizpower_news.xml"
	'------ start deleting the new message from XML file ------
	set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	objDom.async = false			
	if objDOM.load(server.MapPath(xmlFilePath)) then
		set objNodes = objDOM.documentElement.childNodes 
		for j=0 to objNodes.length-1
			set objTask = objNodes.item(j)
			node_new_id = objTask.attributes.getNamedItem("ID").text			
			if trim(Request.QueryString("delId")) = trim(node_new_id) Then					
				objDOM.documentElement.removeChild(objTask)
				exit for
			else
				set objTask = nothing
			end if
		next
		Set objNodes = nothing
		set objTask = nothing
		objDom.save server.MapPath(xmlFilePath)
	end if
	set objDOM = nothing
	' ------ end  deleting the new message from XML file ------
	
	set pr=con.execute("SELECT New_Picture FROM News WHERE New_ID=" & Request.QueryString("delId"))
	d=0
	con.execute "delete from news where New_Id=" & Request.QueryString("delId"),d
	if d<>0 then
	if not pr.eof then
		fileName1=pr("New_Picture")
		if fileName1<>"" then
			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
			fileString= Server.MapPath("../../download/news/"& fileName1 ) 'deleting of existing file
			if fs.FileExists(fileString) then
				set f=fs.GetFile(fileString) 
				f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
			end if
			
			fileString= Server.MapPath("../../download/news/small/"& fileName1 ) 'deleting of existing file
			if fs.FileExists(fileString) then
				set f=fs.GetFile(fileString) 
				f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
			end if			
	    end if
	    pr.movenext	
	end if
	pr.close
	end if
end if

newId=Request.QueryString("newId")
elmOrd=Request.QueryString("elmOrd")

if Request.QueryString("down")<>nil then
stam=CInt(elmOrd)+1
	con.Execute("update news set New_order=-10 WHERE New_order="&stam&" and category_ID="& catID )
	con.Execute("update news set New_order="&stam&" WHERE New_order="&elmOrd&" and category_ID="& catID )
	con.Execute("update news set New_order="&elmOrd&" WHERE New_order=-10 and category_ID="& catID )
end if

if Request.QueryString("up")<>nil then
stam=CInt(elmOrd)-1
	sqlStr = "update news set New_order=-10 WHERE New_order="&stam&" and category_ID="& catID
	''Response.Write sqlStr
	con.Execute(sqlStr)
	con.Execute("update news set New_order="&stam&" WHERE New_order="&elmOrd&" and category_ID="& catID )
	con.Execute("update news set New_order="&elmOrd&" WHERE New_order=-10 and category_ID="& catID )
end if

set mcnt=con.Execute("select count(*) as myCont from news WHERE category_ID="& catID)
	cnt=mcnt("myCont")
mcnt.close
set myNum=con.Execute("select New_order,New_id from news WHERE  category_ID="& catID & " ORDER BY New_order ")
if not myNum.eof then
 myNum.moveFirst
 p=1
	do while p<=cnt
	 pgId=myNum("New_id")
	 stst="update news set New_order="&p&" where New_id=" & pgId
	 con.Execute(stst)
 myNum.moveNext
 p=p+1
	loop
end if
myNum.close

if catid<>nil and catid<>"" then
	set rs_category = con.execute("SELECT Category_Name FROM News_Categories WHERE Category_ID=" &catid)
	category_Name = rs_Category("Category_Name")
	rs_category.close
	set rs_category=Nothing
end if
%>

<body class="body_admin">
<div align="right">

<table border="0" width="100%" align="center"  cellspacing="0" cellpadding="0" ID="Table1">

  <tr><td class="a_title_big" width="100%" dir=rtl><font size="4" color="#ffffff"><strong>"<%=category_Name%>" דף ניהול אירועים</strong></font></td></tr>
  <tr>
    <td>
       <table class="table_admin_1" border="0" width="100%" align="center"  cellspacing="0" cellpadding="0" ID="Table2">
         <tr>
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="Aparaadd.asp?catID=<%=catID%>">הוספת אירוע&nbsp;</a></td>             
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>             
			<td width="100%"></td>
         </tr>
       </table>
    </td>
  </tr>
  <tr><td height="5"></td></tr>
</table>
<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" ID="Table4">
<%numItem=0
 set counpubl=con.Execute("SELECT COUNT(new_ID) AS itemCnt FROM news WHERE category_ID="& catID &" " )
 if not counpubl.eof then
	numItem=counpubl("itemCnt")
 end if
 counpubl.close
 set counpubl=Nothing
%>
<tr>
   <td align="center" class="td_admin_5" colspan="2">&nbsp;</td>
   <td class="td_admin_5" align="center" valign="top" nowrap>&nbsp;<b>מופיע בראשי</b>&nbsp;</td>   
   <td class="td_admin_5" align="center" valign="top" nowrap>&nbsp;<b>מופיע באתר</b>&nbsp;</td>       
    <%if 0 then%>
   <td width="70" class="td_admin_5" align="center" nowrap><b>שעת הורדה</b></td>
   <td width="70" class="td_admin_5" align="center" nowrap><b>שעת הופעה</b></td>
   <%end if%>
   <td width="70" class="td_admin_5" align="center" nowrap><b>תאריך</b></td>   
   <td width="60%" class="td_admin_5" width="50%" align="center" nowrap><font size="3"><strong>שם אירוע</strong></font></td>
   <%if 0 then%>
   <td align="center" class="td_admin_5" colspan="2">&nbsp;</td>
   <%end if%>
</tr>
<%
sqlStr = "select New_ID,new_Order,New_Title,new_Date,new_Time_on,new_Time_off,New_Home_Visible, New_Site_Visible from News WHERE category_ID="& catID &" ORDER BY New_Order"
set rs_news=con.execute(SqlStr)

i=0
do while not rs_news.eof
    i=i+1
    id=rs_news("New_Id")
    elmOrd=rs_news("new_Order")
	New_Title=rs_news("New_Title")
	new_Time_on = rs_news("new_Time_on")
    new_Time_off = rs_news("new_Time_off")	
    perHome=rs_news("New_Home_Visible")
    perSite=rs_news("New_Site_Visible")    
	new_Date=rs_news("new_Date")    
%>
	<tr>
     	<td class="td_admin_4" align="left" valign="middle"   nowrap><a class="button_delete_1" href="admin.asp?catID=<%=catID%>&delId=<%=id%>" ONCLICK="return CheckDel()">מחיקה</a></td>
		<td class="td_admin_4" align="left" valign="middle"   nowrap><a class="button_edit_1" href="Aparaadd.asp?catID=<%=catID%>&Id=<%=id%>">עדכון</b></a></td>
	    <td class="td_admin_4" align="center" valign="middle" nowrap><a class="link_categ" href="vsbPress.asp?catID=<%=catID%>&idhome=<%=id%>"><%if perHome=false then%><img src="../images/lamp_off.gif" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../images/lamp_on.gif" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>
	    <td class="td_admin_4" align="center" valign="middle" nowrap><a class="link_categ" href="vsbPress.asp?catID=<%=catID%>&idSite=<%=id%>"><%if perSite=false then%><img src="../images/lamp_off.gif" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../images/lamp_on.gif" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></td>		
		<%if 0 then%>
		<td class="td_admin_4" align="center" valign="middle" ><font style="font-weight:400;font-size:11px;"><%=new_Time_off%></font></td>
		<td class="td_admin_4" align="center" valign="middle" ><font style="font-weight:400;font-size:11px;"><%=new_Time_on%></font></td>
		<%end if%>
		<td width="70" class="td_admin_4" align="center" valign="middle" ><font style="font-weight:400;font-size:11px;"><%=new_Date%></font></td>		
		<td width="60%" class="td_admin_4" align="right" valign="middle" dir=rtl><%=New_Title%>&nbsp;</td>
		<%if 0 then%>
		<td width="1%" align="center" valign="bottom" class="td_admin_2">
	       <%if i > 1 then%>
		      <a href="admin.asp?catID=<%=catID%>&newID=<%=id%>&elmOrd=<%=elmOrd%>&up=1"><img src="../images/up.gif" border="0" width="11"></a>
	       <%else%>
	          &nbsp;
	       <%end if%>
	    </td>
	    <td width="1%" align="center" valign="top" class="td_admin_2">
	       <%if i < numItem then%>
		      <a href="admin.asp?catID=<%=catID%>&newID=<%=id%>&elmOrd=<%=elmOrd%>&down=1"><img src="../images/down.gif" border="0" width="11"></a>
	       <%else%>
	          &nbsp;
	       <%end if%>
	    </td>
	    <%end if%>
	</tr>
<%	
rs_news.movenext
loop
rs_news.close
set rs_news = nothing
%>
</table>
</div>
</body>
</html>
<%
con.close
set con=Nothing
%>
