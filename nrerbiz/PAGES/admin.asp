<%SERVER.ScriptTimeout=3000%>
<!--#include file="../../include/connect.asp"-->
<!--#include file="../../include/reverse.asp"-->
<%
if session("admin.username") <> nil then	
	username=session("admin.username")
	password=session("admin.password")
else 
	Response.Redirect(strLocal & "/nrerbiz/default.asp")	
end if
%>
<html>

<head>
<title>CyberServe</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>

<script LANGUAGE="JavaScript">
<!--
function warningFunc()
{
  str='             WARNING !!! \r\n\n A process of FILE REBUILDING will be started. \r\n This process can take up to an HOUR or more. \r\n Stopping this process will lead to incorrect information \r\n and irreversible problems with your work. \r\n\n Are you sure you want to continue? \r\n'
	return confirm(str)
}
function CheckDel() {
  return (confirm("?האם ברצונך למחוק את הדף"))    
}

function CheckProcess() {
  return (confirm("?האם ברצונך להעביר לארכיון"))    
}

function mapwindow(catid)
{
	url="copypage.asp?catid=" + catid
	cpagewin=window.open(url,"cpagewin","left=100,top=50,height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes");
	return false
}
function mapwindowCopy(mainid,catid)
{
	url="copypageDubl.asp?mainid="+mainid+"&catid=" + catid
	cpagewin1=window.open(url,"cpagewin1","left=100,top=50,height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes");
	return false
}
function  updvisibility(catid)
{
	document.ff.catIdVis.value=catid;
	document.ff.submit();
}

<!--End-->
</script>  

<%
'on error resume next

sqlstring="SELECT * from workers where loginName='" & userName & "' and  password='"& password &"'"
set worker=con.Execute(sqlstring)
if worker.EOF then%>
<p><center><font color="red" size="4"><b>You are not authorized to enter the SITE staff zone !!!</b></font></center>
<%else
  session("admin.username")=username
  session("admin.password")=password

maincat=Request("maincat")
siteid=session("siteid")

place=Request.QueryString("place")
categId=Request.QueryString("catId")
catOrd=Request.QueryString("catOrd")
'isRightMenu=Request.querystring("isright")

if Request.QueryString("DEL")="1" then 
con.execute "delete from Publication_Categories WHERE Publication_Category_Id=" & categId 
con.execute "delete from pagesTav WHERE Category_Id=" & categId 
end if

if Request.QueryString("down")<>nil then
	set ord=con.execute("select TOP 1 Publication_Category_Id,Category_Order FROM Publication_Categories WHERE Category_Order>" & catOrd & " AND Main_Category_ID=" & maincat)
	if not ord.eof then
	nextord=ord("Category_Order")
	nextcatId=ord("Publication_Category_Id")
	con.Execute("update Publication_Categories set Category_Order="&nextord&" WHERE Publication_Category_Id="&categId)
	con.Execute("update Publication_Categories set Category_Order="&catOrd&" WHERE Publication_Category_Id="&nextcatId)
	end if
end if

if Request.QueryString("up")<>nil then
	set ord=con.execute("select TOP 1 Publication_Category_Id,Category_Order FROM Publication_Categories WHERE Category_Order<"&catOrd & " AND Main_Category_ID=" & maincat & " ORDER BY Category_Order DESC")
	if not ord.eof then
		prevord=ord("Category_Order")
		prevcatId=ord("Publication_Category_Id")
		con.Execute("update Publication_Categories set Category_Order="&prevord&" WHERE Publication_Category_Id="&categId)
		con.Execute("update Publication_Categories set Category_Order="&catOrd&" WHERE Publication_Category_Id="&prevcatId)
	end if
end if

if Request.form("catIdVis")<>"" AND  Request.form("catIdVis")<>nil then 
  con.execute ("update Publication_Categories set Category_Vis=1-Category_Vis WHERE Publication_Category_Id=" & Request.form("catIdVis"))
end if

if Request("name")<>nil then
	if categId<>nil and categId<>"" then
		if Request("mapid")<>nil then
			mapid=Request("mapid")
		else
			mapid="NULL"
		end if
		con.Execute("update Publication_Categories set Publication_Category_Name='" & sFix(request("name")) & "',Category_URL='"& sFix(request("catURL")) &"',Map_Category_id=" & mapid & " WHERE Publication_Category_Id="&categId)
	else
		set ord=con.execute("select Max(Category_Order) as maxord FROM Publication_Categories WHERE Main_Category_ID=" & maincat)
		if isNull(ord("maxord")) then
			maxord=0
		else
			maxord=ord("maxord")
		end if
		str="insert into Publication_Categories (Publication_Category_Name,Category_Order,Main_Category_ID,Category_URL,Category_Vis) VALUES('" & sFix(request("name")) & "'," & maxord+1 & "," & maincat & ",'" & sFix(request("catURL")) & "',1)"
		con.Execute(str)
	end if
end if
set main=con.execute("SELECT Main_Category_Name FROM Main_Categories WHERE Main_Category_ID=" & maincat)
 
%>
<body class="body_admin">
<div align="right">
<table class="table_admin_2" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <%set sites=con.execute("SELECT Site_Name FROM Sites WHERE Site_Id=" & siteid)
  if not sites.EOF then
       siteName=sites("Site_Name")  
    end if
    sites.close
    set sites=Nothing%>
 <tr><td class="a_title_big" width="100%" dir=rtl><center> מנגנון ניהול אתר "<%=siteName%>"</center></td></tr>
 <tr><td class="td_admin_4" height="10"></td></tr>
<tr >
   <td class="td_admin_4" align="left"  valign="middle" nowrap>
			<table class="table_admin_2" width="100%" border="0"  cellpadding="1" cellspacing="0">
					<tr>
						<td align="left" valign="middle" nowrap><a  class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>
						<td width="100%">&nbsp;</td>
					</tr>
			</table>
	</td>
</tr>
  <tr><td class="td_admin_4" align="center" valign="middle" nowrap>&nbsp;</td></tr>
<tr >
   <td class="title_admin_1" align="left"  valign="middle" nowrap>
			<table  width="100%" border="0"  cellpadding="1" cellspacing="0">
					<tr>
							<td align="left" valign="middle" nowrap>
								<a class="button_admin_1" href="AddCat.asp?maincat=<%=maincat%>">הוספת תת קטגוריה חדשה</a>
							</td>
							<td align="left" valign="middle" nowrap>
								<!--a class="button_admin_1" href="../PagesArch/admin.asp?maincat=<%=maincat%>">
									ןויכרא
								</a-->
							</td>
							<td class="title_admin_1" align="right" valign="middle" width="100%" dir="rtl" nowrap>&nbsp;<%=main("Main_Category_Name")%></td>
					</tr>
			</table>
	</td>
</tr>
</table>

<table class="table_admin_2" width="100%" cellpadding="0" cellspacing="0">
<tr>
     <td class="td_admin_4_left_align" width="100%" align="right" height="10">
         <table  border="0" cellpadding="1" cellspacing="0" width="100%">
            <tr>
                <td  class="td_line_between" align="right" colspan="10"></td>
            </tr>
            <form name="ff" action="admin.asp?maincat=<%=maincat%>" method="post">
            <input type="hidden" name="catIdVis" value="">
<%Set publ=con.Execute("SELECT Publication_Categories.* FROM Publication_Categories  WHERE Publication_Categories.Main_Category_ID=" & maincat & " order by Publication_Categories.Category_Order")
 Set counpubl=con.Execute("SELECT COUNT(*) AS catcount FROM Publication_Categories WHERE Main_Category_ID=" & maincat & "")
 if not counpubl.eof then
	numcat=counpubl("catcount")
 end if
 for i=1 to numcat
  catId=publ("Publication_Category_Id")
  catName=publ("Publication_Category_Name")
  catOrd=publ("Category_Order")
  catURL=publ("Category_URL")
  'catmap=publ("Map_Category_Name")
  catVisible=publ("Category_Vis")
%>
<tr>
	<td nowrap><a class="button_delete_1" href="admin.asp?catId=<%=catId%>&DEL=1&isright=<%=isRightMenu%>&maincat=<%=maincat%>" ONCLICK="return CheckDel()">מחיקה</a></td>
	<td nowrap><a class="button_edit_1" href = "AddCat.asp?catId=<%=catId%>&isright=<%=isRightMenu%>&maincat=<%=maincat%>">עדכון</a></td>
	<!--td  nowrap><a  class="button_copy" href="admin.asp?arch=1&arcatId=<%=catID%>&catId=<%=catID%>">לארכיון</a></td>
	<td align="left" valign="middle" nowrap><a class="button_copy"  href="" onclick="return mapwindow('<%=catId%>')">העברה</a></td>
	<td align="left" valign="middle"  nowrap><a class="button_copy"  href="" onclick="return mapwindowCopy('<%=maincat%>','<%=catId%>')">העתק</a></td-->
	<!--td class="td_admin_4_right_align" align="right" valign="top" nowrap>&nbsp;&nbsp;
	   <input type="checkbox" <%if catVisible then%> checked onclick="updvisibility(<%=catId%>)" <%else%> onclick="updvisibility(<%=catId%>)" <%end if%> id=checkbox1 name=checkbox1>&nbsp;הירוגטקכ גצה תרתוכ&nbsp;&nbsp;
    </td-->
	<td class="td_admin_4_right_align" align="right" valign="top" nowrap>
		<%if catmap<>nil then%>
		   <%=catmap%> :הפמ
		<%end if%>
	</td>
	<td class="td_admin_4_right_align" align="right" valign="top" width="100%" nowrap dir="rtl">
		<%if not isNull(catURL) and trim(catURL)<>"" then%>
			<a class="link_categ" href="<%=catURL%>" target="_blank"><%=catName%></a>
		<%else%>
		    <a class="link_categ" href="publications.asp?catId=<%=catId%>&isright=<%=isRightMenu%>&maincat=<%=maincat%>"><%=catName%>&nbsp;</a>
		<%end if%>
	</td>

	<%if i>1 then%>
		<td class="td_admin_2" width="1%" align="center" valign="bottom"><a href="admin.asp?catId=<%=catId%>&isright=<%=isRightMenu%>&maincat=<%=maincat%>&catOrd=<%=catOrd%>&up=1"><img src="images/up.gif" alt="להעביר את הקטגוריה למעלה" border="0"></a></td>
	<%else%>
	    <td class="td_admin_2" width="1%" align="center" valign="bottom">&nbsp;</td>
	<%end if%>
	<%if i<numcat then%>
		<td class="td_admin_2" width="1%" align="center" valign="top"><a href="admin.asp?catId=<%=catID%>&isright=<%=isRightMenu%>&maincat=<%=maincat%>&catOrd=<%=catOrd%>&down=1"><img src="images/down.gif" alt="להעביר את הקטגוריה למטה" border="0"></a></td>
	<%else%>
	    <td class="td_admin_2" width="1%" align="center" valign="top">&nbsp;</td>
	<%end if%>
</tr>
<tr>
    <td class="td_line_between" align="right" colspan="10">
    </td>
</tr>
<%publ.movenext
	next
	publ.close %>
<tr>
  <td class="td_admin_4" colspan="10" height="10">
  </td>
</tr>
</table>
</td>
</form>
</table>
</div>
</body>
<%end if%>
</html>
